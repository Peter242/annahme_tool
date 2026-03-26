param(
  [int]$ParentPid = 0
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = [System.Text.UTF8Encoding]::new()
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$script:ExcelOpenRequiredMessage = ('Fehler: Annahme muss ge' + [char]246 + 'ffnet sein. Bitte ' + [char]246 + 'ffnen und erneut versuchen')
$script:ForbiddenAttachOnlyMessage = 'FORBIDDEN: attempted to start Excel or open workbook'
$script:Excel = $null
$script:CachedWorkbookPath = ''
$script:CachedWorkbook = $null
$script:CachedSheetKey = ''
$script:CachedSheet = $null
$script:CurrentWhere = 'init'
$script:CurrentDetail = ''
$script:LastStep = ''
$script:LastWorkbookFullName = ''
[Console]::Out.WriteLine('{"type":"worker","msg":"started"}')

function Log-Step([string]$msg) {
  $ts = (Get-Date).ToString('HH:mm:ss.fff')
  $script:LastStep = [string]$msg
  $obj = @{ type = 'step'; ts = $ts; msg = $msg }
  $json = ($obj | ConvertTo-Json -Compress)
  [Console]::Out.WriteLine($json)
}

function Normalize-LogHead {
  param(
    [string]$text,
    [int]$maxLen = 200
  )
  $normalized = [string]$text
  $normalized = $normalized -replace "`r", ' '
  $normalized = $normalized -replace "`n", ' '
  $normalized = $normalized.Trim()
  if ($normalized.Length -le $maxLen) {
    return $normalized
  }
  return ($normalized.Substring(0, $maxLen) + '...')
}

function Write-Diag([string]$msg) {
  [Console]::Error.WriteLine($msg)
}

function Get-OpenWorkbookNames {
  if ($null -eq $script:Excel) {
    return @()
  }
  $names = New-Object System.Collections.Generic.List[string]
  try {
    foreach ($candidate in $script:Excel.Workbooks) {
      try {
        $name = [string]$candidate.FullName
        if (-not [string]::IsNullOrWhiteSpace($name)) {
          $names.Add($name)
        }
      } catch {
        # ignore workbook enumeration errors
      }
    }
  } catch {
    return @()
  }
  return $names.ToArray()
}

function Set-ObjectPropertyValue {
  param(
    [object]$target,
    [string]$name,
    [object]$value
  )
  if ($null -eq $target -or [string]::IsNullOrWhiteSpace($name)) {
    return
  }
  if ($target.PSObject.Properties.Name -contains $name) {
    $target.$name = $value
    return
  }
  $target | Add-Member -NotePropertyName $name -NotePropertyValue $value -Force
}

if ($ParentPid -gt 0) {
  try {
    $script:ParentWatch = New-Object System.Timers.Timer
    $script:ParentWatch.Interval = 1000
    $script:ParentWatch.AutoReset = $true
    Register-ObjectEvent -InputObject $script:ParentWatch -EventName Elapsed -Action {
      if (-not (Get-Process -Id $using:ParentPid -ErrorAction SilentlyContinue)) {
        [Environment]::Exit(0)
      }
    } | Out-Null
    $script:ParentWatch.Start()
  } catch {
    # ignore parent watch setup errors
  }
}

function Set-ErrorContext {
  param(
    [string]$where,
    [string]$detail = ''
  )
  $script:CurrentWhere = [string]$where
  $script:CurrentDetail = [string]$detail
}

function Mark-Time {
  param(
    [hashtable]$timings,
    [System.Diagnostics.Stopwatch]$sw,
    [string]$name
  )
  $timings[$name] = [int][Math]::Round($sw.Elapsed.TotalMilliseconds)
}

function Elapsed-Since {
  param(
    [hashtable]$timings,
    [string]$start,
    [string]$ending
  )
  $startMs = if ($timings.ContainsKey($start)) { [int]$timings[$start] } else { 0 }
  $endMs = if ($timings.ContainsKey($ending)) { [int]$timings[$ending] } else { $startMs }
  return [int][Math]::Max(0, ($endMs - $startMs))
}

function Convert-CellValueToText {
  param([object]$value)
  if ($null -eq $value) { return '' }
  return [string]$value
}

function Normalize-CounterCellValue {
  param([object]$value)

  if ($null -eq $value) {
    return ''
  }

  if ($value -is [string]) {
    return $value.Trim()
  }

  $culture = [System.Globalization.CultureInfo]::InvariantCulture
  if (
    $value -is [byte] -or
    $value -is [sbyte] -or
    $value -is [int16] -or
    $value -is [uint16] -or
    $value -is [int32] -or
    $value -is [uint32] -or
    $value -is [int64] -or
    $value -is [uint64]
  ) {
    return $value.ToString($culture)
  }

  if ($value -is [decimal] -or $value -is [double] -or $value -is [single]) {
    try {
      $decimalValue = [decimal]$value
      if ($decimalValue -eq [decimal]::Truncate($decimalValue)) {
        return ([decimal]::Truncate($decimalValue)).ToString($culture)
      }
      return $decimalValue.ToString('0.############################', $culture)
    } catch {
      return ([string]$value).Trim()
    }
  }

  return ([string]$value).Trim()
}

function Build-TodayPrefix {
  param([datetime]$now)
  return $now.ToString('ddMMyy') + '8'
}

function ParseDateOrNull {
  param([object]$s)
  if ($null -eq $s) { return $null }
  $raw = [string]$s
  if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
  $value = $raw.Trim()
  $culture = [System.Globalization.CultureInfo]::InvariantCulture
  $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces
  $formats = @(
    'yyyy-MM-dd',
    'dd.MM.yyyy',
    'yyyy-MM-ddTHH:mm:ss',
    'yyyy-MM-ddTHH:mm:ssK',
    'yyyy-MM-ddTHH:mm:ss.fff',
    'yyyy-MM-ddTHH:mm:ss.fffK',
    'o'
  )
  foreach ($format in $formats) {
    try {
      return [datetime]::ParseExact($value, $format, $culture, $styles)
    } catch {
      # continue
    }
  }
  try {
    return [datetime]::Parse($value, $culture)
  } catch {
    throw "Invalid date format: $value"
  }
}

function Format-GermanDateOnly {
  param([object]$value)
  if ([string]::IsNullOrWhiteSpace([string]$value)) { return '' }
  $parsed = ParseDateOrNull -s $value
  if ($null -eq $parsed) { return '' }
  return ('{0:dd.MM.yyyy}' -f $parsed)
}

function Format-GermanTermin {
  param([string]$ymd)
  if ([string]::IsNullOrWhiteSpace($ymd)) { return '' }
  $parsed = ParseDateOrNull -s $ymd
  if ($null -eq $parsed) { return '' }
  $weekdays = @('So','Mo','Di','Mi','Do','Fr','Sa')
  $wd = $weekdays[[int]$parsed.DayOfWeek]
  return ('{0} {1:dd.MM.yyyy}' -f $wd, $parsed)
}

function Build-HeaderI {
  param([object]$order)
  $lines = @()
  $kunde = if ([string]::IsNullOrWhiteSpace([string]$order.auftraggeberKurz)) { [string]$order.kunde } else { [string]$order.auftraggeberKurz }
  $ansprechpartner = [string]$order.ansprechpartner
  $projektnummer = [string]$order.projektnummer
  $projektname = if ([string]::IsNullOrWhiteSpace([string]$order.projektname)) { [string]$order.projekt } else { [string]$order.projektname }
  $rawProbenahmedatum = if ([string]::IsNullOrWhiteSpace([string]$order.probenahmedatum)) { [string]$order.probenEingangDatum } else { [string]$order.probenahmedatum }
  $formattedProbenahmedatum = Format-GermanDateOnly -value $rawProbenahmedatum
  $projektNrLine = if ([string]::IsNullOrWhiteSpace($projektnummer)) { '' } else { 'Projekt Nr: ' + [string]$projektnummer }
  $projektLine = if ([string]::IsNullOrWhiteSpace($projektname)) { '' } else { 'Projekt: ' + [string]$projektname }
  $probenahmeLine = if ([string]::IsNullOrWhiteSpace($formattedProbenahmedatum)) { '' } else { 'Probenahme: ' + [string]$formattedProbenahmedatum }
  foreach ($item in @($kunde, $ansprechpartner, $projektNrLine, $projektLine, $probenahmeLine)) {
    if (-not [string]::IsNullOrWhiteSpace($item)) {
      $lines += [string]$item
    }
  }
  return (($lines | ForEach-Object { [string]$_ }) -join "`n")
}

function Build-HeaderJ {
  param(
    [object]$order,
    [object]$termin,
    [bool]$writeAddressBlock = $true
  )
  $lines = @()
  $kuerzel = if (-not [string]::IsNullOrWhiteSpace([string]$order.kuerzel)) { [string]$order.kuerzel } else { [string]$order.erfasstKuerzel }
  $terminValue = if (-not [string]::IsNullOrWhiteSpace([string]$order.terminDatum)) { [string]$order.terminDatum } else { [string]$termin }
  $terminText = if ([string]::IsNullOrWhiteSpace($terminValue)) { '' } else { 'Termin: ' + (Format-GermanTermin -ymd $terminValue) }
  $firstLineParts = @()
  if (-not [string]::IsNullOrWhiteSpace($kuerzel)) { $firstLineParts += [string]$kuerzel }
  if ($order.eilig -eq $true) { $firstLineParts += 'EILIG' }
  if (-not [string]::IsNullOrWhiteSpace($terminText)) { $firstLineParts += [string]$terminText }
  if ($firstLineParts.Count -gt 0) {
    $lines += (($firstLineParts | ForEach-Object { [string]$_ }) -join ' ')
  }
  $kopfBemerkung = [string]$order.kopfBemerkung
  if (-not [string]::IsNullOrWhiteSpace($kopfBemerkung)) {
    $lines += $kopfBemerkung.Trim()
  }
  if ($writeAddressBlock -eq $true) {
    $adresseBlock = [string]$order.adresseBlock
    if (-not [string]::IsNullOrWhiteSpace($adresseBlock)) {
      $adresseLines = @()
      foreach ($line in ($adresseBlock -replace "`r`n?", "`n").Split("`n")) {
        $trimmedLine = [string]$line
        if (-not [string]::IsNullOrWhiteSpace($trimmedLine)) {
          $adresseLines += $trimmedLine.Trim()
        }
      }
      if ($adresseLines.Count -gt 0) {
        $lines += (($adresseLines | ForEach-Object { [string]$_ }) -join "`n")
      }
    }
  }
  if (-not [string]::IsNullOrWhiteSpace([string]$order.email)) {
    $lines += ('Mail: ' + [string]$order.email)
  }
  return (($lines | ForEach-Object { [string]$_ }) -join "`n")
}

function Build-ProbeJ {
  param([object]$probe)
  $provided = [string]$probe.probeJ
  if (-not [string]::IsNullOrWhiteSpace($provided)) {
    return $provided.Trim()
  }
  $parts = @()
  if ($null -ne $probe.gewicht -and -not [string]::IsNullOrWhiteSpace([string]$probe.gewicht)) {
    $parts += ('Gewicht: ' + [string]$probe.gewicht + ' kg')
  }
  $geruchRaw = if (-not [string]::IsNullOrWhiteSpace([string]$probe.geruch)) { [string]$probe.geruch } else { [string]$probe.geruchAuffaelligkeit }
  if (-not [string]::IsNullOrWhiteSpace($geruchRaw)) {
    $parts += ('Geruch: ' + [string]$geruchRaw)
  }
  $bemerkungValue = [string]$probe.bemerkung
  if (-not [string]::IsNullOrWhiteSpace($bemerkungValue)) {
    $parts += $bemerkungValue.Trim()
  }
  return (($parts | ForEach-Object { [string]$_ }) -join '; ')
}

function Normalize-ParameterTextForCompare {
  param([object]$value)
  if ($null -eq $value) { return '' }
  $text = [string]$value
  if ([string]::IsNullOrWhiteSpace($text)) { return '' }
  $normalized = $text -replace "`r`n", "`n"
  $normalized = $normalized -replace "`r", "`n"
  return $normalized.Trim()
}

function Get-ColAHash50 {
  param([object]$sheet)
  $lines = New-Object System.Collections.Generic.List[string]
  for ($row = 1; $row -le 50; $row++) {
    $value = (Convert-CellValueToText ($sheet.Cells.Item($row, 1).Value2)).Trim()
    $lines.Add([string]$value)
  }
  $joined = [string]::Join("`n", $lines)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($joined)
  $sha1 = [System.Security.Cryptography.SHA1]::Create()
  try {
    $hash = $sha1.ComputeHash($bytes)
  } finally {
    $sha1.Dispose()
  }
  return -join ($hash | ForEach-Object { $_.ToString('x2') })
}

function Get-SheetState {
  param(
    [object]$sheet,
    [datetime]$now
  )
  $usedRows = [Math]::Max([int]$sheet.UsedRange.Rows.Count, 1)
  $lastUsedRow = 0
  $maxLabNumber = 0
  $maxOrderSeqToday = 0
  $todayPrefix = Build-TodayPrefix -now $now
  for ($row = 1; $row -le $usedRows; $row++) {
    $hasContent = $false
    for ($col = 1; $col -le 10; $col++) {
      $v = Convert-CellValueToText ($sheet.Cells.Item($row, $col).Value2)
      if ($v.Trim() -ne '') {
        $hasContent = $true
        break
      }
    }
    if ($hasContent) {
      $lastUsedRow = $row
    }
    $a = (Convert-CellValueToText ($sheet.Cells.Item($row, 1).Value2)).Trim()
    if ($a -eq '') { continue }
    $orderCoreMatch = [regex]::Match($a, '^(\d{6}8\d{2})')
    if ($orderCoreMatch.Success) {
      $core = [string]$orderCoreMatch.Groups[1].Value
      if ($core.StartsWith($todayPrefix, [System.StringComparison]::Ordinal)) {
        $seq = [int]$core.Substring($core.Length - 2, 2)
        if ($seq -gt $maxOrderSeqToday) { $maxOrderSeqToday = $seq }
      }
      continue
    }
    $labMatch = [regex]::Match($a, '^(\d{5,6})([A-Za-z]|-\d+)?$')
    if ($labMatch.Success) {
      $lab = [int]$labMatch.Groups[1].Value
      if ($lab -gt $maxLabNumber) { $maxLabNumber = $lab }
    }
  }
  return @{
    lastUsedRow = $lastUsedRow
    maxLabNumber = $maxLabNumber
    maxOrderSeqToday = $maxOrderSeqToday
  }
}

function Find-FirstCompletelyEmptyRow {
  param(
    [object]$sheet,
    [int]$startRow
  )
  $row = [Math]::Max($startRow, 1)
  while ($true) {
    $hasContent = $false
    for ($col = 1; $col -le 10; $col++) {
      $v = Convert-CellValueToText ($sheet.Cells.Item($row, $col).Value2)
      if ($v.Trim() -ne '') {
        $hasContent = $true
        break
      }
    }
    if (-not $hasContent) { return $row }
    $row++
  }
}

function Resolve-FullPath {
  param([string]$pathValue)
  $targetPath = [System.IO.Path]::GetFullPath([string]$pathValue)
  try {
    $resolved = Resolve-Path -LiteralPath $targetPath -ErrorAction Stop
    if ($null -ne $resolved) {
      $targetPath = [string]$resolved.Path
    }
  } catch {
    # keep as-is
  }
  return $targetPath
}

function Normalize-PathFast {
  param([string]$pathValue)
  if ([string]::IsNullOrWhiteSpace([string]$pathValue)) {
    return ''
  }
  try {
    return [System.IO.Path]::GetFullPath([string]$pathValue).TrimEnd('\').ToLowerInvariant()
  } catch {
    return ([string]$pathValue).Trim().TrimEnd('\').ToLowerInvariant()
  }
}

function Reset-WorkbookCache {
  $script:CachedWorkbookPath = ''
  $script:CachedWorkbook = $null
  $script:CachedSheetKey = ''
  $script:CachedSheet = $null
}

function Get-ExcelApplication {
  param([bool]$allowAutoOpen)
  Set-ErrorContext -where 'excel.connect' -detail 'get/open excel application'
  if ($allowAutoOpen -eq $true) {
    throw $script:ForbiddenAttachOnlyMessage
  }
  if ($null -ne $script:Excel) {
    try {
      $null = $script:Excel.Version
    } catch {
      $script:Excel = $null
      Reset-WorkbookCache
    }
  }
  if ($null -eq $script:Excel) {
    try {
      $script:Excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
      $script:Excel = $null
    }
  }
  if ($null -eq $script:Excel) {
    [Console]::Error.WriteLine('[worker] attach failed: no running Excel')
    throw $script:ExcelOpenRequiredMessage
  }
  return $script:Excel
}

function New-ExcelNotReadyError {
  $message = 'Excel ist nicht bereit oder wartet auf ein Dialogfenster. Bitte Excel in den Vordergrund holen und Dialog schließen.'
  $exception = New-Object System.Exception($message)
  $exception.Data['errorCode'] = 'EXCEL_NOT_READY'
  throw $exception
}

function Invoke-WithExcelAccess {
  param(
    [object]$excel,
    [scriptblock]$action
  )

  $isReady = $false
  $isInteractive = $false
  try {
    $isReady = ($excel.Ready -eq $true)
  } catch {
    $isReady = $false
  }
  try {
    $isInteractive = ($excel.Interactive -eq $true)
  } catch {
    $isInteractive = $false
  }
  if (-not $isReady -or -not $isInteractive) {
    New-ExcelNotReadyError
  }

  $previousDisplayAlerts = $true
  try {
    $previousDisplayAlerts = ($excel.DisplayAlerts -eq $true)
  } catch {
    $previousDisplayAlerts = $true
  }

  try {
    try {
      $excel.DisplayAlerts = $false
    } catch {
      # ignore
    }
    return & $action
  } finally {
    try {
      $excel.DisplayAlerts = $previousDisplayAlerts
    } catch {
      # ignore
    }
  }
}

function Get-Workbook {
  param(
    [object]$excel,
    [string]$targetPath,
    [string]$targetPathNormalized,
    [string]$targetName,
    [bool]$allowAutoOpen
  )
  Set-ErrorContext -where 'workbook.find' -detail "targetPath=$targetPath"
  return Invoke-WithExcelAccess -excel $excel -action {
    if ($null -eq $excel.Workbooks) {
      throw 'Excel.Workbooks ist null'
    }

    if (-not [string]::IsNullOrWhiteSpace($script:CachedWorkbookPath) -and $null -ne $script:CachedWorkbook) {
      try {
        $cachedFullName = [string]$script:CachedWorkbook.FullName
        $null = $script:CachedWorkbook.Worksheets.Count
        $cachedPathNormalized = Normalize-PathFast -pathValue $cachedFullName
        if ([string]::Equals($cachedPathNormalized, $targetPathNormalized, [System.StringComparison]::OrdinalIgnoreCase)) {
          return $script:CachedWorkbook
        }
      } catch {
        Reset-WorkbookCache
      }
    }

    $wb = $null
    foreach ($candidate in $excel.Workbooks) {
      try {
        $candidateFullName = [string]$candidate.FullName
        if ([string]::IsNullOrWhiteSpace($candidateFullName)) { continue }
        $candidatePath = Normalize-PathFast -pathValue $candidateFullName
        if ([string]::Equals($candidatePath, $targetPathNormalized, [System.StringComparison]::OrdinalIgnoreCase)) {
          $wb = $candidate
          break
        }
      } catch {
        # continue
      }
    }

    if ($null -eq $wb) {
      foreach ($candidate in $excel.Workbooks) {
        try {
          if ([string]::Equals([string]$candidate.Name, $targetName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $wb = $candidate
            break
          }
        } catch {
          # continue
        }
      }
    }

    if ($null -eq $wb) {
      [Console]::Error.WriteLine(('[worker] attach failed: workbook not open: ' + [string]$targetPath))
      throw $script:ExcelOpenRequiredMessage
    }

    $script:CachedWorkbookPath = $targetPath
    $script:CachedWorkbook = $wb
    $script:CachedSheetKey = ''
    $script:CachedSheet = $null
    return $wb
  }
}

function Get-Sheet {
  param(
    [object]$wb,
    [string]$sheetName
  )
  Set-ErrorContext -where 'sheet.get' -detail "yearSheetName=$sheetName"
  return Invoke-WithExcelAccess -excel $script:Excel -action {
    $wbPath = Resolve-FullPath -pathValue ([string]$wb.FullName)
    $key = "$wbPath|$sheetName"

    if ([string]::Equals($script:CachedSheetKey, $key, [System.StringComparison]::OrdinalIgnoreCase) -and $null -ne $script:CachedSheet) {
      try {
        $null = $script:CachedSheet.Name
        return $script:CachedSheet
      } catch {
        $script:CachedSheetKey = ''
        $script:CachedSheet = $null
      }
    }

    $sheet = $wb.Worksheets.Item($sheetName)
    if ($null -eq $sheet) {
      throw "Jahresblatt $sheetName nicht gefunden"
    }
    $script:CachedSheetKey = $key
    $script:CachedSheet = $sheet
    return $sheet
  }
}

function Build-ReadOnlyErrorResult {
  param(
    [string]$targetPath,
    [object]$wb,
    [bool]$isWriteReserved
  )
  $message = 'Annahme.xlsx ist schreibgeschützt oder gesperrt. Bitte Datei schreibbar öffnen (nicht Schreibgeschützt) und erneut versuchen.'
  $workbookFullName = ''
  $readOnly = $false
  try { $workbookFullName = [string]$wb.FullName } catch { $workbookFullName = '' }
  try { $readOnly = ($wb.ReadOnly -eq $true) } catch { $readOnly = $false }
  return @{
    ok = $false
    saved = $false
    writer = 'com'
    errorCode = 'WORKBOOK_READONLY'
    error = $message
    message = $message
    targetPath = $targetPath
    workbookFullName = $workbookFullName
    workbookReadOnly = $readOnly
    workbookWriteReserved = $isWriteReserved
    openWorkbooks = @(Get-OpenWorkbookNames)
  }
}

function Build-SaveFailedErrorResult {
  param(
    [string]$targetPath,
    [object]$wb,
    [bool]$isWriteReserved,
    [string]$errorMessage
  )
  $workbookFullName = ''
  $readOnly = $false
  try { $workbookFullName = [string]$wb.FullName } catch { $workbookFullName = '' }
  try { $readOnly = ($wb.ReadOnly -eq $true) } catch { $readOnly = $false }
  $message = if ([string]::IsNullOrWhiteSpace([string]$errorMessage)) { 'Speichern fehlgeschlagen' } else { [string]$errorMessage }
  return @{
    ok = $false
    saved = $false
    writer = 'com'
    errorCode = 'SAVE_FAILED'
    error = $message
    message = $message
    targetPath = $targetPath
    workbookFullName = $workbookFullName
    workbookReadOnly = $readOnly
    workbookWriteReserved = $isWriteReserved
    openWorkbooks = @(Get-OpenWorkbookNames)
  }
}

function Build-ExcelNotReadyDialogErrorResult {
  $message = 'Excel wartet auf ein Dialogfenster (Datei gesperrt, Warnung, etc). Bitte schließe das Dialogfenster und versuche erneut.'
  return @{
    ok = $false
    saved = $false
    writer = 'com'
    errorCode = 'EXCEL_NOT_READY'
    error = $message
    message = $message
  }
}

function Wait-ExcelReady {
  param(
    [object]$excel,
    [int]$timeoutMs = 1500,
    [int]$pollMs = 100
  )
  try {
    return (($excel.Ready -eq $true) -and ($excel.Interactive -eq $true))
  } catch {
    return $false
  }
}

function Get-ExcelReadyDebug {
  param([object]$excel)
  $debug = @{
    ready = $null
    calculationState = $null
    interactive = $null
    displayAlerts = $null
  }
  try { $debug.ready = $excel.Ready } catch { $debug.ready = $null }
  try { $debug.calculationState = $excel.CalculationState } catch { $debug.calculationState = $null }
  try { $debug.interactive = $excel.Interactive } catch { $debug.interactive = $null }
  try { $debug.displayAlerts = $excel.DisplayAlerts } catch { $debug.displayAlerts = $null }
  return $debug
}

function Assert-ExcelReadyNow {
  param([object]$excel)
  $debug = Get-ExcelReadyDebug -excel $excel
  $isBusy = ($debug.ready -ne $true) -or ($debug.interactive -ne $true)
  if (-not $isBusy -and $null -ne $debug.calculationState) {
    try {
      $isBusy = ([int]$debug.calculationState -ne 0)
    } catch {
      $isBusy = $true
    }
  }
  if ($isBusy) {
    Write-Diag ("[excel-not-ready] ready={0} interactive={1} calculationState={2} displayAlerts={3}" -f $debug.ready, $debug.interactive, $debug.calculationState, $debug.displayAlerts)
    New-ExcelNotReadyError -excel $excel
  }
}

function New-ExcelNotReadyError {
  param([object]$excel = $null)
  $message = 'Excel ist nicht bereit oder wartet auf ein Dialogfenster. Bitte Excel in den Vordergrund holen, Dialog schließen und erneut versuchen.'
  $exception = New-Object System.Exception($message)
  $exception.Data['errorCode'] = 'EXCEL_NOT_READY'
  $exception.Data['debug'] = if ($null -ne $excel) { Get-ExcelReadyDebug -excel $excel } else { @{} }
  throw $exception
}

function Invoke-WithExcelAccess {
  param(
    [object]$excel,
    [scriptblock]$action
  )
  Assert-ExcelReadyNow -excel $excel

  $previousDisplayAlerts = $true
  try {
    $previousDisplayAlerts = ($excel.DisplayAlerts -eq $true)
  } catch {
    $previousDisplayAlerts = $true
  }

  try {
    try {
      $excel.DisplayAlerts = $false
    } catch {
      # ignore
    }
    return & $action
  } finally {
    try {
      $excel.DisplayAlerts = $previousDisplayAlerts
    } catch {
      # ignore
    }
  }
}

function Process-Commit {
  param([object]$payload)
  $APPEND_SCAN_MAX_ROWS = 1200
  $APPEND_SCAN_BUDGET_MS = 2000
  $NUMBER_SCAN_MAX_ROWS = 1500
  $timings = @{}
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  Set-ErrorContext -where 'init' -detail ''
  Log-Step "commit start requestId=$($script:CurrentRequestId)"
  Mark-Time -timings $timings -sw $sw -name 'start'

  Set-ErrorContext -where 'payload.parse' -detail 'reading payload json'
  Mark-Time -timings $timings -sw $sw -name 'payload.parsed'

  $allowAutoOpenExcel = $false
  if ($payload.PSObject.Properties.Name -contains 'allowAutoOpenExcel') {
    try {
      $allowAutoOpenExcel = [bool]$payload.allowAutoOpenExcel
    } catch {
      $allowAutoOpenExcel = $false
    }
  }
  $debugCom = $false
  if ($payload.PSObject.Properties.Name -contains 'debugCom') {
    try {
      $debugCom = [bool]$payload.debugCom
    } catch {
      $debugCom = $false
    }
  }

  Log-Step 'attach excel begin'
  $excel = Get-ExcelApplication -allowAutoOpen $allowAutoOpenExcel
  Log-Step 'attach excel ok'
  Mark-Time -timings $timings -sw $sw -name 'excel.connect'
  Log-Step 'excel ready check begin'
  Assert-ExcelReadyNow -excel $excel
  try { $excel.DisplayAlerts = $false } catch { }
  Log-Step 'excel ready check ok'

  Set-ErrorContext -where 'path.resolve' -detail 'targetPath from payload.excelPath'
  $targetPath = Resolve-FullPath -pathValue ([string]$payload.excelPath)
  $targetPathNormalized = Normalize-PathFast -pathValue $targetPath
  $targetName = [System.IO.Path]::GetFileName($targetPath)
  Log-Step "find workbook begin target=$targetPath"
  Assert-ExcelReadyNow -excel $excel
  $wb = Get-Workbook -excel $excel -targetPath $targetPath -targetPathNormalized $targetPathNormalized -targetName $targetName -allowAutoOpen $allowAutoOpenExcel
  try { $script:LastWorkbookFullName = [string]$wb.FullName } catch { $script:LastWorkbookFullName = '' }
  Log-Step "find workbook ok fullName=$($script:LastWorkbookFullName)"
  Mark-Time -timings $timings -sw $sw -name 'workbook.attach'
  $isReadOnly = $false
  $isWriteReserved = $false
  try {
    $isReadOnly = ($wb.ReadOnly -eq $true)
  } catch {
    $isReadOnly = $false
  }
  try {
    if ($wb.PSObject.Properties.Name -contains 'WriteReserved') {
      $isWriteReserved = -not [string]::IsNullOrWhiteSpace([string]$wb.WriteReserved)
    }
  } catch {
    $isWriteReserved = $false
  }
  Log-Step "workbook state readOnly=$isReadOnly writeReserved=$isWriteReserved"
  if ($isReadOnly) {
    return (Build-ReadOnlyErrorResult -targetPath $targetPath -wb $wb -isWriteReserved $isWriteReserved)
  }
  if ($isWriteReserved) {
    Log-Step 'workbook writeReserved observed, continue'
  }

  Log-Step "get sheet begin name=$([string]$payload.yearSheetName)"
  Assert-ExcelReadyNow -excel $excel
  $sheet = Get-Sheet -wb $wb -sheetName ([string]$payload.yearSheetName)
  Log-Step 'get sheet ok'
  Mark-Time -timings $timings -sw $sw -name 'sheet.get'

  Set-ErrorContext -where 'state.compute' -detail 'Get-SheetState + append row'
  Log-Step 'compute appendRow begin'
  $now = [datetime]::Parse([string]$payload.now)
  $todayPrefix = Build-TodayPrefix -now $now
  $appendRow = $null
  $startLabNo = $null
  $nextSeq = $null
  $maxOrderSeqToday = 0
  $usedCacheHint = $false
  $cacheHint = if ($payload.PSObject.Properties.Name -contains 'cacheHint') { $payload.cacheHint } else { $null }
  if ($null -ne $cacheHint) {
    $hintPrefix = [string]$cacheHint.todayPrefix
    $hintAppendRow = 0
    $hintStartLabNo = 0
    $hintMaxSeq = 0
    $hintNextSeq = 0
    try { $hintAppendRow = [int]$cacheHint.appendRow } catch { $hintAppendRow = 0 }
    try { $hintStartLabNo = [int]$cacheHint.startLabNo } catch { $hintStartLabNo = 0 }
    try { $hintMaxSeq = [int]$cacheHint.maxOrderSeqToday } catch { $hintMaxSeq = 0 }
    try { $hintNextSeq = [int]$cacheHint.nextSeq } catch { $hintNextSeq = 0 }
    if ($hintAppendRow -gt 0 -and $hintStartLabNo -gt 0 -and $hintNextSeq -gt 0 -and [string]::Equals($hintPrefix, $todayPrefix, [System.StringComparison]::Ordinal)) {
      $hintRowA = (Convert-CellValueToText ($sheet.Cells.Item($hintAppendRow, 1).Value2)).Trim()
      if ($hintRowA -eq '') {
        $appendRow = $hintAppendRow
        $usedCacheHint = $true
      }
    }
  }
  if ($usedCacheHint -ne $true) {
    Log-Step 'compute appendRow strategy=tail-window'
    Log-Step 'compute appendRow usedRange begin'
    $usedRange = $sheet.UsedRange
    $usedRangeRows = [Math]::Max([int]$usedRange.Rows.Count, 1)
    $usedRangeCols = [Math]::Max([int]$usedRange.Columns.Count, 1)
    Log-Step "compute appendRow usedRange ok rows=$usedRangeRows cols=$usedRangeCols"
    Log-Step 'compute appendRow lastDataRow begin'
    $lastDataRow = [int]$sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
    if ($lastDataRow -lt 1) {
      $lastDataRow = $usedRangeRows
    }
    Log-Step "compute appendRow lastDataRow ok value=$lastDataRow"
    $scanStart = [Math]::Max(1, $lastDataRow - $APPEND_SCAN_MAX_ROWS + 1)
    $scanEnd = [Math]::Max($scanStart, $lastDataRow)
    Log-Step "compute appendRow inspect window begin start=$scanStart end=$scanEnd"
    $appendScanStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lastNonEmptyRow = 0
    $windowMaxLabNumber = 0
    $windowMaxOrderSeqToday = 0
    $scanBudgetExceeded = $false
    $rangeAddress = "A$scanStart:J$scanEnd"
    $windowRange = $sheet.Range($rangeAddress)
    $windowValues = $windowRange.Value2
    $windowRowCount = [Math]::Max($scanEnd - $scanStart + 1, 1)
    $windowRowMeta = @{}
    for ($offset = 1; $offset -le $windowRowCount; $offset++) {
      $row = $scanStart + $offset - 1
      if (($offset % 200) -eq 0) {
        Log-Step "compute appendRow inspect row=$row"
      }
      if (($offset % 50) -eq 0 -and $appendScanStopwatch.ElapsedMilliseconds -gt $APPEND_SCAN_BUDGET_MS) {
        Log-Step 'compute appendRow budget exceeded'
        $scanBudgetExceeded = $true
        break
      }
      $rowHasContent = $false
      $rowIsOccupied = $false
      $colA = ''
      for ($col = 1; $col -le 10; $col++) {
        $cellValue = if ($windowRowCount -eq 1) {
          $windowRange.Cells.Item(1, $col).Value2
        } else {
          $windowValues[$offset, $col]
        }
        $text = Convert-CellValueToText $cellValue
        if ($col -eq 1) {
          $colA = $text.Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($text)) {
          $rowHasContent = $true
        }
        if (($col -eq 1) -or ($col -eq 4) -or ($col -eq 6) -or ($col -eq 8) -or ($col -eq 9) -or ($col -eq 10)) {
          if (-not [string]::IsNullOrWhiteSpace($text)) {
            $rowIsOccupied = $true
          }
        }
      }
      $windowRowMeta[$row] = @{
        hasContent = $rowHasContent
        isOccupied = $rowIsOccupied
      }
      if ($rowHasContent) {
        $lastNonEmptyRow = $row
      }
      if ($colA -eq '') {
        continue
      }
      $orderCoreMatch = [regex]::Match($colA, '^(\d{6}8\d{2})')
      if ($orderCoreMatch.Success) {
        $core = [string]$orderCoreMatch.Groups[1].Value
        if ($core.StartsWith($todayPrefix, [System.StringComparison]::Ordinal)) {
          $seq = [int]$core.Substring($core.Length - 2, 2)
          if ($seq -gt $windowMaxOrderSeqToday) { $windowMaxOrderSeqToday = $seq }
        }
        continue
      }
      $labMatch = [regex]::Match($colA, '^(\d{5,6})([A-Za-z]|-\d+)?$')
      if ($labMatch.Success) {
        $lab = [int]$labMatch.Groups[1].Value
        if ($lab -gt $windowMaxLabNumber) { $windowMaxLabNumber = $lab }
      }
    }
    Log-Step 'compute appendRow inspect window ok'
    $lastOccupiedRow = 0
    for ($row = $scanEnd; $row -ge $scanStart; $row--) {
      $meta = $windowRowMeta[$row]
      if ($meta -and $meta.isOccupied -eq $true) {
        $lastOccupiedRow = $row
        break
      }
    }
    if ($lastOccupiedRow -lt 1) {
      $lastOccupiedRow = [Math]::Max($lastNonEmptyRow, $lastDataRow)
    }
    $blankRow = $lastOccupiedRow + 1
    Log-Step "compute appendRow lastOccupiedRow=$lastOccupiedRow"
    Log-Step "compute appendRow blankRow=$blankRow"
    if ($scanBudgetExceeded -eq $true) {
      $fallbackLastOccupiedRow = [Math]::Max([Math]::Max($lastOccupiedRow, $lastDataRow), 1)
      $appendRow = $fallbackLastOccupiedRow + 2
    } else {
      $appendRow = $lastOccupiedRow + 2
    }
  }
  Log-Step "compute appendRow result row=$appendRow"
  Mark-Time -timings $timings -sw $sw -name 'state.computed'

  Set-ErrorContext -where 'numbering.compute' -detail 'orderNo + lab numbers'
  Log-Step 'compute counters begin'
  $counterLastDataRow = [int]$sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
  if ($counterLastDataRow -lt 1) {
    $counterLastDataRow = [Math]::Max(($appendRow - 2), 1)
  }
  $counterScanEnd = [Math]::Max([Math]::Min(($appendRow - 2), $counterLastDataRow), 1)
  $counterScanStart = [Math]::Max(1, $counterScanEnd - $NUMBER_SCAN_MAX_ROWS + 1)
  Log-Step "compute counters scan window start=$counterScanStart end=$counterScanEnd max=$NUMBER_SCAN_MAX_ROWS"
  $lastOrderNo = ''
  $lastOrderNoRow = 0
  $lastLabNo = 0
  $lastLabNoRow = 0
  $counterLastInspectedValues = New-Object System.Collections.Generic.List[object]
  $counterClassifiedLabNos = New-Object System.Collections.Generic.List[object]
  $counterInspectLogCount = 0
  $counterLabLogCount = 0
  if ($counterScanEnd -ge $counterScanStart) {
    Log-Step 'compute counters read mode=cell-by-cell'
    for ($row = $counterScanEnd; $row -ge $counterScanStart; $row--) {
      $rawA = $sheet.Cells.Item($row, 1).Value2
      $rawPreview = Normalize-LogHead -text (Convert-CellValueToText $rawA) -maxLen 80
      $colA = Normalize-CounterCellValue $rawA
      if (-not [string]::IsNullOrWhiteSpace($colA)) {
        $counterLastInspectedValues.Add([pscustomobject]@{
          row = [int]$row
          raw = if ($rawPreview -ne '') { $rawPreview } else { $null }
          normalized = $colA
        })
        while ($counterLastInspectedValues.Count -gt 20) {
          $counterLastInspectedValues.RemoveAt(0)
        }
        if ($counterInspectLogCount -lt 15) {
          Log-Step ("compute counters inspect row={0} raw={1} normalized={2}" -f $row, $rawPreview, (Normalize-LogHead -text $colA -maxLen 80))
          $counterInspectLogCount++
        }
      }
      if ($lastOrderNoRow -eq 0) {
        $orderCoreMatch = [regex]::Match($colA, '^(\d{6}8\d{2})$')
        if ($orderCoreMatch.Success) {
          $lastOrderNo = [string]$orderCoreMatch.Groups[1].Value
          $lastOrderNoRow = $row
        }
      }
      $labMatch = [regex]::Match($colA, '^(\d{5,6})([A-Za-z]|-\d+)?$')
      if ($labMatch.Success) {
        $baseLabNo = [int]$labMatch.Groups[1].Value
        $labSuffix = [string]$labMatch.Groups[2].Value
        $isVariant = -not [string]::IsNullOrWhiteSpace($labSuffix)
        $counterClassifiedLabNos.Add([pscustomobject]@{
          row = [int]$row
          raw = if ($rawPreview -ne '') { $rawPreview } else { $null }
          normalized = $colA
          baseLabNo = [int]$baseLabNo
          isVariant = [bool]$isVariant
        })
        while ($counterClassifiedLabNos.Count -gt 20) {
          $counterClassifiedLabNos.RemoveAt(0)
        }
        if ($counterLabLogCount -lt 15) {
          Log-Step ("compute counters labNo candidate row={0} raw={1} base={2} variant={3}" -f $row, $rawPreview, $baseLabNo, $isVariant)
          $counterLabLogCount++
        }
        if ($baseLabNo -gt $lastLabNo) {
          $lastLabNo = [int]$baseLabNo
          $lastLabNoRow = [int]$row
        }
      }
      if ($lastOrderNoRow -gt 0 -and $lastLabNo -gt 0) {
        break
      }
    }
  }
  if ($lastOrderNoRow -gt 0) {
    Log-Step "compute counters lastOrderNo=$lastOrderNo row=$lastOrderNoRow"
    $maxOrderSeqToday = [int]$lastOrderNo.Substring($lastOrderNo.Length - 2, 2)
    $todayPrefix = $lastOrderNo.Substring(0, $lastOrderNo.Length - 2)
  } else {
    Log-Step 'compute counters no classified orderNo found'
    Log-Step 'compute counters fallback orderNo'
    $maxOrderSeqToday = 0
  }
  if ($lastLabNoRow -gt 0) {
    Log-Step "compute counters lastLabNo maxBase=$lastLabNo row=$lastLabNoRow"
    $startLabNo = [int]$lastLabNo + 1
  } else {
    Log-Step 'compute counters no classified labNo found'
    Log-Step 'compute counters fallback labNo'
    $startLabNo = 10000
  }
  $nextSeq = [int]$maxOrderSeqToday + 1
  $nextOrderNo = $todayPrefix + $nextSeq.ToString('00')
  Log-Step "compute counters nextOrderNo=$nextOrderNo"
  Log-Step "compute counters nextLabNo=$startLabNo"
  if ($nextSeq -gt 99) {
    throw "Maximale Tagessequenz erreicht fuer Prefix $todayPrefix"
  }
  $order = $payload.order
  $requestOrderNo = ''
  foreach ($candidate in @(
    [string]$payload.orderNo,
    [string]$payload.orderNumber,
    [string]$payload.auftragsnummer,
    [string]$order.orderNo,
    [string]$order.orderNumber,
    [string]$order.auftragsnummer
  )) {
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
      $requestOrderNo = $candidate.Trim()
      break
    }
  }
  $requestLabNo = 0
  foreach ($candidate in @(
    [string]$payload.firstLabNo,
    [string]$payload.nextLabNo,
    [string]$payload.labNo,
    [string]$order.firstLabNo,
    [string]$order.nextLabNo,
    [string]$order.labNo
  )) {
    if ([string]::IsNullOrWhiteSpace($candidate)) {
      continue
    }
    try {
      $parsedCandidate = [int]$candidate
      if ($parsedCandidate -gt 0) {
        $requestLabNo = $parsedCandidate
        break
      }
    } catch {
      # ignore invalid candidate
    }
  }
  $orderNo = $nextOrderNo
  $firstLabNo = [int]$startLabNo
  if (-not [string]::IsNullOrWhiteSpace($requestOrderNo) -and -not [string]::Equals($requestOrderNo, $orderNo, [System.StringComparison]::Ordinal)) {
    Log-Step "commit counters override requestOrderNo=$requestOrderNo liveOrderNo=$orderNo"
  }
  if ($requestLabNo -gt 0 -and $requestLabNo -ne $firstLabNo) {
    Log-Step "commit counters override requestLabNo=$requestLabNo liveLabNo=$firstLabNo"
  }
  Set-ObjectPropertyValue -target $payload -name 'orderNo' -value $orderNo
  Set-ObjectPropertyValue -target $payload -name 'orderNumber' -value $orderNo
  Set-ObjectPropertyValue -target $payload -name 'auftragsnummer' -value $orderNo
  Set-ObjectPropertyValue -target $payload -name 'firstLabNo' -value $firstLabNo
  Set-ObjectPropertyValue -target $payload -name 'nextLabNo' -value $firstLabNo
  Set-ObjectPropertyValue -target $payload -name 'labNo' -value $firstLabNo
  Set-ObjectPropertyValue -target $order -name 'orderNo' -value $orderNo
  Set-ObjectPropertyValue -target $order -name 'orderNumber' -value $orderNo
  Set-ObjectPropertyValue -target $order -name 'auftragsnummer' -value $orderNo
  Set-ObjectPropertyValue -target $order -name 'firstLabNo' -value $firstLabNo
  Set-ObjectPropertyValue -target $order -name 'nextLabNo' -value $firstLabNo
  Set-ObjectPropertyValue -target $order -name 'labNo' -value $firstLabNo
  $lastLabNoForCommit = $firstLabNo + [Math]::Max(($order.proben.Count - 1), 0)
  Log-Step "commit counters applied orderNo=$orderNo"
  Log-Step "commit counters applied firstLabNo=$firstLabNo"
  Log-Step "commit counters applied sampleCount=$($order.proben.Count)"
  Log-Step "commit counters applied lastLabNo=$lastLabNoForCommit"
  $sampleNos = @()
  $seenParameterTexts = @{}

  Set-ErrorContext -where 'header.write' -detail "appendRow=$appendRow"
  Log-Step 'header write begin'
  $pbTypValue = if ([string]::IsNullOrWhiteSpace([string]$order.pbTyp)) { 'PB' } else { [string]$order.pbTyp }
  $sheet.Cells.Item($appendRow, 1).Value2 = $orderNo
  $sheet.Cells.Item($appendRow, 2).Value2 = [string]'y'
  $sheet.Cells.Item($appendRow, 3).Value2 = [string]'y'
  $sheet.Cells.Item($appendRow, 4).Value2 = [string]''
  $sheet.Cells.Item($appendRow, 5).Value2 = [string]$pbTypValue
  $sheet.Cells.Item($appendRow, 9).Value2 = [string](Build-HeaderI -order $order)
  $sheet.Cells.Item($appendRow, 9).WrapText = $true
  $writeAddressBlock = $true
  if ($payload.PSObject.Properties.Name -contains 'excelWriteAddressBlock') {
    try {
      $writeAddressBlock = [bool]$payload.excelWriteAddressBlock
    } catch {
      $writeAddressBlock = $true
    }
  }
  $sheet.Cells.Item($appendRow, 10).Value2 = [string](Build-HeaderJ -order $order -termin $payload.termin -writeAddressBlock $writeAddressBlock)
  $sheet.Cells.Item($appendRow, 10).WrapText = $true
  $sheet.Cells.Item($appendRow, 10).Font.Bold = ($order.eilig -eq $true)
  Log-Step 'header write ok'
  Mark-Time -timings $timings -sw $sw -name 'write.header'

  Set-ErrorContext -where 'samples.write.loop' -detail "sampleCount=$($order.proben.Count)"
  Log-Step "samples write begin count=$($order.proben.Count)"
  for ($i = 0; $i -lt $order.proben.Count; $i++) {
    $probe = $order.proben[$i]
    $row = $appendRow + 1 + $i
    $lab = $firstLabNo + $i
    Log-Step "sample write i=$i begin labNo=$lab"
    $rawParameterText = [string]$probe.parameterTextPreview
    $normalizedParameterText = Normalize-ParameterTextForCompare -value $rawParameterText
    $parameterText = $rawParameterText
    if (-not [string]::IsNullOrWhiteSpace($normalizedParameterText)) {
      if ($seenParameterTexts.ContainsKey($normalizedParameterText)) {
        $parameterText = 'dito'
      } else {
        $seenParameterTexts[$normalizedParameterText] = $true
      }
    }
    $probeJ = Build-ProbeJ -probe $probe
    $sampleNos += $lab
    $sheet.Cells.Item($row, 1).Value2 = [int]$lab
    $sheet.Cells.Item($row, 4).Value2 = [string]$parameterText
    $sheet.Cells.Item($row, 4).WrapText = $true
    $sheet.Cells.Item($row, 6).Value2 = [string]$probe.probenbezeichnung
    $gValue = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$probe.tiefeOderVolumen)) {
      $gValue = [string]$probe.tiefeOderVolumen
    } elseif (-not [string]::IsNullOrWhiteSpace([string]$probe.tiefeVolumen)) {
      $gValue = [string]$probe.tiefeVolumen
    } elseif ($null -ne $probe.volumen -and -not [string]::IsNullOrWhiteSpace([string]$probe.volumen)) {
      $gValue = [string]$probe.volumen
    }
    $sheet.Cells.Item($row, 7).Value2 = [string]$gValue
    $material = if ($probe.material) { [string]$probe.material } else { '' }
    $gebinde = if ($probe.gebinde) { [string]$probe.gebinde } else { '' }
    $materialGebinde = if ($probe.materialGebinde) { [string]$probe.materialGebinde } else { ($material + ' ' + $gebinde).Trim() }
    if ([string]::IsNullOrWhiteSpace($materialGebinde)) { $materialGebinde = [string]$probe.matrixTyp }
    $sheet.Cells.Item($row, 8).Value2 = [string]$materialGebinde
    $sheet.Cells.Item($row, 10).Value2 = [string]$probeJ
    $sheet.Cells.Item($row, 10).WrapText = $true
    Log-Step "sample write i=$i ok"
  }
  Mark-Time -timings $timings -sw $sw -name 'write.samples'

  Set-ErrorContext -where 'rows.autofit' -detail "appendRow=$appendRow"
  $lastWrittenRow = $appendRow + $order.proben.Count
  for ($row = $appendRow; $row -le $lastWrittenRow; $row++) {
    $sheet.Rows.Item($row).EntireRow.AutoFit() | Out-Null
  }
  Mark-Time -timings $timings -sw $sw -name 'rows.autofit'

  Set-ErrorContext -where 'workbook.save' -detail 'Save only'
  $saveWorkbook = $false
  if ($payload.PSObject.Properties.Name -contains 'saveWorkbook') {
    try {
      $saveWorkbook = [bool]$payload.saveWorkbook
    } catch {
      $saveWorkbook = $false
    }
  }
  $saved = $false
  if ($saveWorkbook -eq $true) {
    Log-Step 'save begin'
    Assert-ExcelReadyNow -excel $excel
    try {
      $wb.Save()
      $saved = $true
    } catch {
      return (Build-SaveFailedErrorResult -targetPath $targetPath -wb $wb -isWriteReserved $isWriteReserved -errorMessage $_.Exception.Message)
    }
    Log-Step 'save ok'
  } else {
    Log-Step 'save skipped'
  }
  Mark-Time -timings $timings -sw $sw -name 'save.done'
  if ($debugCom -eq $true) {
    $savedOrderNo = (Convert-CellValueToText ($sheet.Cells.Item($appendRow, 1).Value2)).Trim()
    if (-not [string]::Equals($savedOrderNo, $orderNo, [System.StringComparison]::Ordinal)) {
      throw "Save verification failed: expected orderNo=$orderNo actual=$savedOrderNo row=$appendRow"
    }
  }
  Mark-Time -timings $timings -sw $sw -name 'verify.readback'

  $firstSampleNo = $null
  $lastSampleNo = $null
  if ($sampleNos.Count -gt 0) {
    $firstSampleNo = [int]$sampleNos[0]
    $lastSampleNo = [int]$sampleNos[$sampleNos.Count - 1]
  }
  $headerRow = $appendRow
  $firstSampleRow = if ($order.proben.Count -gt 0) { $appendRow + 1 } else { $null }
  $writtenHeaderCellA = ''
  $writtenFirstSampleCellA = ''
  $writtenSampleLabNos = @()
  try {
    $writtenHeaderCellA = (Convert-CellValueToText ($sheet.Cells.Item($headerRow, 1).Value2)).Trim()
  } catch {
    $writtenHeaderCellA = ''
  }
  if ($order.proben.Count -gt 0) {
    try {
      $writtenFirstSampleCellA = (Convert-CellValueToText ($sheet.Cells.Item($firstSampleRow, 1).Value2)).Trim()
    } catch {
      $writtenFirstSampleCellA = ''
    }
    for ($i = 0; $i -lt $order.proben.Count; $i++) {
      $sampleRow = $appendRow + 1 + $i
      try {
        $writtenSampleLabNos += (Convert-CellValueToText ($sheet.Cells.Item($sampleRow, 1).Value2)).Trim()
      } catch {
        $writtenSampleLabNos += ''
      }
    }
  }
  $endRowRange = "A$appendRow:J$lastWrittenRow"
  Mark-Time -timings $timings -sw $sw -name 'done'
  Log-Step 'commit done'

  $timingMs = @{
    computeSheetStateMs = (Elapsed-Since -timings $timings -start 'sheet.get' -ending 'state.computed')
    comConnectMs = (Elapsed-Since -timings $timings -start 'payload.parsed' -ending 'excel.connect')
    comAttachWorkbookMs = (Elapsed-Since -timings $timings -start 'excel.connect' -ending 'workbook.attach')
    connectAttachMs = (Elapsed-Since -timings $timings -start 'payload.parsed' -ending 'workbook.attach')
    rangeWriteHeaderMs = (Elapsed-Since -timings $timings -start 'state.computed' -ending 'write.header')
    rangeWriteSamplesMs = (Elapsed-Since -timings $timings -start 'write.header' -ending 'write.samples')
    autoFitMs = (Elapsed-Since -timings $timings -start 'write.samples' -ending 'rows.autofit')
    saveMs = (Elapsed-Since -timings $timings -start 'rows.autofit' -ending 'save.done')
    verificationMs = (Elapsed-Since -timings $timings -start 'save.done' -ending 'verify.readback')
    totalMs = (Elapsed-Since -timings $timings -start 'start' -ending 'done')
  }

  return @{
    ok = $true
    written = $true
    saved = $saved
    writer = 'com'
    saveMethodUsed = if ($saved) { 'Save' } else { 'Skipped' }
    targetPath = $targetPath
    orderNo = $orderNo
    auftragsnummer = $orderNo
    appendRow = $appendRow
    startLabNo = $startLabNo
    sampleNos = $sampleNos
    ersteProbennr = $firstSampleNo
    letzteProbennr = $lastSampleNo
    endRowRange = $endRowRange
    timingMs = $timingMs
    debug = @{
      computedLastOrderNo = if ($lastOrderNoRow -gt 0) { $lastOrderNo } else { $null }
      computedNextOrderNo = $nextOrderNo
      computedLastLabNo = if ($lastLabNoRow -gt 0) { [int]$lastLabNo } else { $null }
      computedNextLabNo = [int]$startLabNo
      appliedOrderNo = $orderNo
      appliedFirstLabNo = [int]$firstLabNo
      sampleCount = [int]$order.proben.Count
      appliedLastLabNo = if ($lastLabNoForCommit -ge $firstLabNo) { [int]$lastLabNoForCommit } else { $null }
      appendRow = [int]$appendRow
      headerRow = [int]$headerRow
      firstSampleRow = if ($null -ne $firstSampleRow) { [int]$firstSampleRow } else { $null }
      writtenHeaderCellA = $writtenHeaderCellA
      writtenFirstSampleCellA = if ($writtenFirstSampleCellA -ne '') { $writtenFirstSampleCellA } else { $null }
      writtenSampleLabNos = $writtenSampleLabNos
      counterScanWindowStart = [int]$counterScanStart
      counterScanWindowEnd = [int]$counterScanEnd
      counterLastInspectedValues = $counterLastInspectedValues.ToArray()
      counterClassifiedLabNos = $counterClassifiedLabNos.ToArray()
      requestOrderNo = if (-not [string]::IsNullOrWhiteSpace($requestOrderNo)) { $requestOrderNo } else { $null }
      requestFirstLabNo = if ($requestLabNo -gt 0) { [int]$requestLabNo } else { $null }
    }
  }
}

function Process-Warmup {
  param([object]$payload)
  $timings = @{}
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  Set-ErrorContext -where 'init' -detail ''
  Mark-Time -timings $timings -sw $sw -name 'start'

  $allowAutoOpenExcel = $false
  if ($payload.PSObject.Properties.Name -contains 'allowAutoOpenExcel') {
    try {
      $allowAutoOpenExcel = [bool]$payload.allowAutoOpenExcel
    } catch {
      $allowAutoOpenExcel = $false
    }
  }

  $excel = Get-ExcelApplication -allowAutoOpen $allowAutoOpenExcel
  Mark-Time -timings $timings -sw $sw -name 'excel.connect'
  $targetPath = Resolve-FullPath -pathValue ([string]$payload.excelPath)
  $targetPathNormalized = Normalize-PathFast -pathValue $targetPath
  $targetName = [System.IO.Path]::GetFileName($targetPath)
  $wb = Get-Workbook -excel $excel -targetPath $targetPath -targetPathNormalized $targetPathNormalized -targetName $targetName -allowAutoOpen $allowAutoOpenExcel
  Mark-Time -timings $timings -sw $sw -name 'workbook.attach'
  $null = Get-Sheet -wb $wb -sheetName ([string]$payload.yearSheetName)
  Mark-Time -timings $timings -sw $sw -name 'sheet.get'
  Mark-Time -timings $timings -sw $sw -name 'done'

  return @{
    ok = $true
    warmed = $true
    writer = 'com'
    timingMs = @{
      comConnectMs = (Elapsed-Since -timings $timings -start 'start' -ending 'excel.connect')
      comAttachWorkbookMs = (Elapsed-Since -timings $timings -start 'excel.connect' -ending 'workbook.attach')
      connectAttachMs = (Elapsed-Since -timings $timings -start 'start' -ending 'workbook.attach')
      sheetGetMs = (Elapsed-Since -timings $timings -start 'workbook.attach' -ending 'sheet.get')
      totalMs = (Elapsed-Since -timings $timings -start 'start' -ending 'done')
    }
  }
}

function Process-ReadSheetState {
  param([object]$payload)
  $MAX_SCAN_ROWS = 800
  $SCAN_BUDGET_MS = 2000
  Set-ErrorContext -where 'init' -detail ''
  Log-Step 'readSheetState start'
  $allowAutoOpenExcel = $false
  if ($payload.PSObject.Properties.Name -contains 'allowAutoOpenExcel') {
    try {
      $allowAutoOpenExcel = [bool]$payload.allowAutoOpenExcel
    } catch {
      $allowAutoOpenExcel = $false
    }
  }

  Log-Step 'readSheetState get excel begin'
  $excel = Get-ExcelApplication -allowAutoOpen $allowAutoOpenExcel
  Log-Step 'readSheetState get excel ok'
  Assert-ExcelReadyNow -excel $excel

  Set-ErrorContext -where 'path.resolve' -detail 'targetPath from payload.excelPath'
  $targetPath = Resolve-FullPath -pathValue ([string]$payload.excelPath)
  $targetPathNormalized = Normalize-PathFast -pathValue $targetPath
  $targetName = [System.IO.Path]::GetFileName($targetPath)
  Log-Step 'readSheetState workbook begin'
  Assert-ExcelReadyNow -excel $excel
  $wb = Get-Workbook -excel $excel -targetPath $targetPath -targetPathNormalized $targetPathNormalized -targetName $targetName -allowAutoOpen $allowAutoOpenExcel
  try { $script:LastWorkbookFullName = [string]$wb.FullName } catch { $script:LastWorkbookFullName = '' }
  Log-Step 'readSheetState workbook ok'
  Log-Step 'readSheetState sheet begin'
  Assert-ExcelReadyNow -excel $excel
  $sheet = Get-Sheet -wb $wb -sheetName ([string]$payload.yearSheetName)
  Log-Step 'readSheetState sheet ok'

  Set-ErrorContext -where 'state.compute' -detail 'Get-SheetState read-only'
  $now = if ($payload.PSObject.Properties.Name -contains 'now') { [datetime]::Parse([string]$payload.now) } else { Get-Date }
  Log-Step 'readSheetState usedRange begin'
  $usedRows = [Math]::Max([int]$sheet.UsedRange.Rows.Count, 1)
  Log-Step 'readSheetState usedRange ok'
  Log-Step 'readSheetState lastRow begin'
  $lastRow = $usedRows
  Log-Step "readSheetState lastRow ok value=$lastRow"
  if ($lastRow -gt 10000) {
    Log-Step "readSheetState largeSheet rows=$lastRow"
  }
  Log-Step 'readSheetState scan begin'
  $scanStartRow = [Math]::Max(1, $lastRow - $MAX_SCAN_ROWS + 1)
  $scanEndRow = $lastRow
  Log-Step "readSheetState scan window start=$scanStartRow end=$scanEndRow max=$MAX_SCAN_ROWS"
  $lastUsedRow = 0
  $maxLabNumber = 0
  $maxOrderSeqToday = 0
  $todayPrefix = Build-TodayPrefix -now $now
  $scanStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  for ($row = $scanStartRow; $row -le $scanEndRow; $row++) {
    if ((($row - $scanStartRow + 1) % 500) -eq 0) {
      Log-Step "readSheetState scan row=$row"
    }
    if (($row -gt $scanStartRow) -and (($row - $scanStartRow) % 50) -eq 0 -and $scanStopwatch.ElapsedMilliseconds -gt $SCAN_BUDGET_MS) {
      Log-Step 'readSheetState scan budget exceeded'
      break
    }
    $hasContent = $false
    for ($col = 1; $col -le 10; $col++) {
      $v = Convert-CellValueToText ($sheet.Cells.Item($row, $col).Value2)
      if ($v.Trim() -ne '') {
        $hasContent = $true
        break
      }
    }
    if ($hasContent) {
      $lastUsedRow = $row
    }
    $a = (Convert-CellValueToText ($sheet.Cells.Item($row, 1).Value2)).Trim()
    if ($a -eq '') { continue }
    $orderCoreMatch = [regex]::Match($a, '^(\d{6}8\d{2})')
    if ($orderCoreMatch.Success) {
      $core = [string]$orderCoreMatch.Groups[1].Value
      if ($core.StartsWith($todayPrefix, [System.StringComparison]::Ordinal)) {
        $seq = [int]$core.Substring($core.Length - 2, 2)
        if ($seq -gt $maxOrderSeqToday) { $maxOrderSeqToday = $seq }
      }
      continue
    }
    $labMatch = [regex]::Match($a, '^(\d{5,6})([A-Za-z]|-\d+)?$')
    if ($labMatch.Success) {
      $lab = [int]$labMatch.Groups[1].Value
      if ($lab -gt $maxLabNumber) { $maxLabNumber = $lab }
    }
  }
  $sheetState = @{
    lastUsedRow = $lastUsedRow
    maxLabNumber = $maxLabNumber
    maxOrderSeqToday = $maxOrderSeqToday
  }
  Log-Step 'readSheetState scan ok'
  $colAHash50 = Get-ColAHash50 -sheet $sheet

  Log-Step 'readSheetState response begin'
  Log-Step 'readSheetState response ok'
  return @{
    ok = $true
    writer = 'com'
    workbookFullName = [string]$wb.FullName
    workbookName = [string]$wb.Name
    yearSheetName = [string]$sheet.Name
    colAHash50 = $colAHash50
    sheetState = @{
      lastUsedRow = [int]$sheetState.lastUsedRow
      maxLabNumber = [int]$sheetState.maxLabNumber
      maxOrderSeqToday = [int]$sheetState.maxOrderSeqToday
    }
  }
}

while ($true) {
  Write-Diag '[worker-rx-wait]'
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) {
    Write-Diag '[worker-rx-eof]'
    break
  }
  $trimmed = [string]$line
  Write-Diag ("[worker-rx] len={0} head={1}" -f $trimmed.Length, (Normalize-LogHead -text $trimmed -maxLen 200))
  if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }

  $requestId = $null
  try {
    Write-Diag '[worker-parse-begin]'
    try {
      $request = $trimmed | ConvertFrom-Json
    } catch {
      Write-Diag ("[worker-parse-error] msg={0}" -f [string]$_.Exception.Message)
      throw
    }
    $requestId = $request.id
    Write-Diag ("[worker-parse-ok] id={0}" -f [string]$requestId)
    $script:CurrentRequestId = $requestId
    $script:LastStep = ''
    $script:LastWorkbookFullName = ''
    $payload = $request.payload
    $isPing = $false
    if ($payload.PSObject.Properties.Name -contains '__ping') {
      try {
        $isPing = [bool]$payload.__ping
      } catch {
        $isPing = $false
      }
    }
    $isReadSheetState = $false
    if ($payload.PSObject.Properties.Name -contains '__readSheetState') {
      try {
        $isReadSheetState = [bool]$payload.__readSheetState
      } catch {
        $isReadSheetState = $false
      }
    }
    $result = if ($isPing) {
      Write-Diag ("[worker-dispatch] id={0} kind=ping" -f [string]$requestId)
      @{
        ok = $true
        status = 'ready'
        writer = 'com'
      }
    } elseif ($isReadSheetState) {
      Write-Diag ("[worker-dispatch] id={0} kind=readSheetState" -f [string]$requestId)
      Process-ReadSheetState -payload $payload
    } else {
      Write-Diag ("[worker-dispatch] id={0} kind=commit" -f [string]$requestId)
      Process-Commit -payload $payload
    }
    $result.id = $requestId
    [Console]::WriteLine(($result | ConvertTo-Json -Compress -Depth 12))
  } catch {
    $lineInfo = if ($null -ne $_.InvocationInfo) { "line=$($_.InvocationInfo.ScriptLineNumber)" } else { '' }
    $detailParts = @()
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentDetail)) { $detailParts += [string]$script:CurrentDetail }
    if (-not [string]::IsNullOrWhiteSpace($lineInfo)) { $detailParts += $lineInfo }
    if (-not [string]::IsNullOrWhiteSpace($script:LastStep)) { $detailParts += ('lastStep=' + [string]$script:LastStep) }
    if (-not [string]::IsNullOrWhiteSpace($script:LastWorkbookFullName)) { $detailParts += ('workbookFullName=' + [string]$script:LastWorkbookFullName) }
    $openWorkbooks = @(Get-OpenWorkbookNames)
    if ($openWorkbooks.Count -gt 0) { $detailParts += ('openWorkbooks=' + ($openWorkbooks -join ', ')) }
    $detailText = ($detailParts | ForEach-Object { [string]$_ }) -join '; '
    $errorCode = [string]$_.Exception.Data['errorCode']
    $errorPayload = @{
      id = $requestId
      ok = $false
      saved = $false
      writer = 'com'
      errorCode = $errorCode
      code = $errorCode
      where = [string]$script:CurrentWhere
      detail = [string]$detailText
      error = $_.Exception.Message
      lastStep = [string]$script:LastStep
      workbookFullName = [string]$script:LastWorkbookFullName
      openWorkbooks = $openWorkbooks
      debug = $_.Exception.Data['debug']
    }
    [Console]::WriteLine(($errorPayload | ConvertTo-Json -Compress))
  }
}
