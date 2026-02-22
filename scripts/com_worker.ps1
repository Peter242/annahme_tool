param(
  [int]$ParentPid = 0
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
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
      $null = $script:Excel.Workbooks.Count
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

function Get-Workbook {
  param(
    [object]$excel,
    [string]$targetPath,
    [string]$targetPathNormalized,
    [string]$targetName,
    [bool]$allowAutoOpen
  )
  Set-ErrorContext -where 'workbook.find' -detail "targetPath=$targetPath"
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

function Get-Sheet {
  param(
    [object]$wb,
    [string]$sheetName
  )
  Set-ErrorContext -where 'sheet.get' -detail "yearSheetName=$sheetName"
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
    readOnly = $readOnly
    writeReserved = $isWriteReserved
  }
}

function Build-ExcelNotReadyDialogErrorResult {
  $message = 'Excel wartet auf ein Dialogfenster (Datei gesperrt, Warnung, etc). Bitte schließe das Dialogfenster und versuche erneut.'
  return @{
    ok = $false
    saved = $false
    writer = 'com'
    errorCode = 'EXCEL_NOT_READY_DIALOG'
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
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt $timeoutMs) {
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
    if ($isReady -and $isInteractive) {
      return $true
    }
    Start-Sleep -Milliseconds $pollMs
  }
  return $false
}

function Process-Commit {
  param([object]$payload)
  $timings = @{}
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  Set-ErrorContext -where 'init' -detail ''
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

  $excel = Get-ExcelApplication -allowAutoOpen $allowAutoOpenExcel
  Mark-Time -timings $timings -sw $sw -name 'excel.connect'
  if (-not (Wait-ExcelReady -excel $excel -timeoutMs 1500 -pollMs 100)) {
    return (Build-ExcelNotReadyDialogErrorResult)
  }

  Set-ErrorContext -where 'path.resolve' -detail 'targetPath from payload.excelPath'
  $targetPath = Resolve-FullPath -pathValue ([string]$payload.excelPath)
  $targetPathNormalized = Normalize-PathFast -pathValue $targetPath
  $targetName = [System.IO.Path]::GetFileName($targetPath)
  $wb = Get-Workbook -excel $excel -targetPath $targetPath -targetPathNormalized $targetPathNormalized -targetName $targetName -allowAutoOpen $allowAutoOpenExcel
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
  if ($isReadOnly -or $isWriteReserved) {
    return (Build-ReadOnlyErrorResult -targetPath $targetPath -wb $wb -isWriteReserved $isWriteReserved)
  }

  $sheet = Get-Sheet -wb $wb -sheetName ([string]$payload.yearSheetName)
  Mark-Time -timings $timings -sw $sw -name 'sheet.get'

  Set-ErrorContext -where 'state.compute' -detail 'Get-SheetState + append row'
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
        $startLabNo = $hintStartLabNo
        $maxOrderSeqToday = $hintMaxSeq
        $nextSeq = $hintNextSeq
        $usedCacheHint = $true
      }
    }
  }
  if ($usedCacheHint -ne $true) {
    $state = Get-SheetState -sheet $sheet -now $now
    $firstEmptyAtEnd = Find-FirstCompletelyEmptyRow -sheet $sheet -startRow ([int]$state.lastUsedRow + 1)
    $appendRow = $firstEmptyAtEnd
    $maxOrderSeqToday = [int]$state.maxOrderSeqToday
    $nextSeq = $maxOrderSeqToday + 1
    $startLabNo = [Math]::Max([int]$state.maxLabNumber, 9999) + 1
  }
  Mark-Time -timings $timings -sw $sw -name 'state.computed'

  Set-ErrorContext -where 'numbering.compute' -detail 'orderNo + lab numbers'
  if ($nextSeq -gt 99) {
    throw "Maximale Tagessequenz erreicht fuer Prefix $todayPrefix"
  }
  $orderNo = $todayPrefix + $nextSeq.ToString('00')
  $order = $payload.order
  $sampleNos = @()
  $seenParameterTexts = @{}

  Set-ErrorContext -where 'header.write' -detail "appendRow=$appendRow"
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
  Mark-Time -timings $timings -sw $sw -name 'write.header'

  Set-ErrorContext -where 'samples.write.loop' -detail "sampleCount=$($order.proben.Count)"
  for ($i = 0; $i -lt $order.proben.Count; $i++) {
    $probe = $order.proben[$i]
    $row = $appendRow + 1 + $i
    $lab = $startLabNo + $i
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
  }
  Mark-Time -timings $timings -sw $sw -name 'write.samples'

  Set-ErrorContext -where 'rows.autofit' -detail "appendRow=$appendRow"
  $lastWrittenRow = $appendRow + $order.proben.Count
  for ($row = $appendRow; $row -le $lastWrittenRow; $row++) {
    $sheet.Rows.Item($row).EntireRow.AutoFit() | Out-Null
  }
  Mark-Time -timings $timings -sw $sw -name 'rows.autofit'

  Set-ErrorContext -where 'workbook.save' -detail 'Save only'

  $wb.Save()
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
  $endRowRange = "A$appendRow:J$lastWrittenRow"
  Mark-Time -timings $timings -sw $sw -name 'done'

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
    saved = $true
    writer = 'com'
    saveMethodUsed = 'Save'
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

while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { break }
  $trimmed = [string]$line
  if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }

  $requestId = $null
  try {
    $request = $trimmed | ConvertFrom-Json
    $requestId = $request.id
    $payload = $request.payload
    $isWarmup = $false
    if ($payload.PSObject.Properties.Name -contains '__warmup') {
      try {
        $isWarmup = [bool]$payload.__warmup
      } catch {
        $isWarmup = $false
      }
    }
    $result = if ($isWarmup) {
      Process-Warmup -payload $payload
    } else {
      Process-Commit -payload $payload
    }
    $result.id = $requestId
    [Console]::WriteLine(($result | ConvertTo-Json -Compress -Depth 12))
  } catch {
    $lineInfo = if ($null -ne $_.InvocationInfo) { "line=$($_.InvocationInfo.ScriptLineNumber)" } else { '' }
    $detailParts = @()
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentDetail)) { $detailParts += [string]$script:CurrentDetail }
    if (-not [string]::IsNullOrWhiteSpace($lineInfo)) { $detailParts += $lineInfo }
    $detailText = ($detailParts | ForEach-Object { [string]$_ }) -join '; '
    $errorPayload = @{
      id = $requestId
      ok = $false
      saved = $false
      writer = 'com'
      where = [string]$script:CurrentWhere
      detail = [string]$detailText
      error = $_.Exception.Message
    }
    [Console]::WriteLine(($errorPayload | ConvertTo-Json -Compress))
  }
}
