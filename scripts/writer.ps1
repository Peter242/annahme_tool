param(
  [Parameter(Mandatory = $true)]
  [string]$PayloadPath
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$targetPath = $null
$script:CurrentWhere = 'init'
$script:CurrentDetail = ''

function Get-TypeName {
  param([object]$value)
  if ($null -eq $value) { return '<null>' }
  return $value.GetType().FullName
}

function Set-ErrorContext {
  param(
    [string]$where,
    [object]$detail
  )

  $script:CurrentWhere = [string]$where
  $script:CurrentDetail = if ($null -eq $detail) { '' } else { [string]$detail }
}

function Write-DebugType {
  param(
    [string]$name,
    [object]$value
  )

  $typeName = Get-TypeName -value $value
  $valueText = if ($null -eq $value) { '<null>' } else { [string]$value }
  Write-Host "[writer-debug] $name type=$typeName value=$valueText"
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

    $a = Convert-CellValueToText ($sheet.Cells.Item($row, 1).Value2)
    $a = $a.Trim()
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

    $labMatch = [regex]::Match($a, '^(\d+)')
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

    if (-not $hasContent) {
      return $row
    }

    $row++
  }
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

function Format-GermanDateOnly {
  param([object]$value)

  if ([string]::IsNullOrWhiteSpace([string]$value)) {
    return ''
  }

  $parsed = ParseDateOrNull -s $value
  if ($null -eq $parsed) { return '' }
  return ('{0:dd.MM.yyyy}' -f $parsed)
}

function Format-GermanTermin {
  param([string]$ymd)

  if ([string]::IsNullOrWhiteSpace($ymd)) {
    return ''
  }

  $parsed = ParseDateOrNull -s $ymd
  if ($null -eq $parsed) { return '' }

  $weekdays = @('So','Mo','Di','Mi','Do','Fr','Sa')
  $wd = $weekdays[[int]$parsed.DayOfWeek]
  return ('{0} {1:dd.MM.yyyy}' -f $wd, $parsed)
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
      # try next format
    }
  }

  try {
    return [datetime]::Parse($value, $culture)
  } catch {
    throw "Invalid date format: $value"
  }
}

function Build-HeaderJ {
  param(
    [object]$order,
    [object]$termin
  )

  $lines = @()
  $kuerzel = if (-not [string]::IsNullOrWhiteSpace([string]$order.kuerzel)) { [string]$order.kuerzel } else { [string]$order.erfasstKuerzel }

  $terminValue = if (-not [string]::IsNullOrWhiteSpace([string]$order.terminDatum)) { [string]$order.terminDatum } else { [string]$termin }
  Write-DebugType -name 'terminValue' -value $terminValue
  $terminText = if ([string]::IsNullOrWhiteSpace($terminValue)) { '' } else { 'Termin: ' + (Format-GermanTermin -ymd $terminValue) }
  $firstLineParts = @()
  if (-not [string]::IsNullOrWhiteSpace($kuerzel)) { $firstLineParts += [string]$kuerzel }
  if ($order.eilig -eq $true) { $firstLineParts += 'EILIG' }
  if (-not [string]::IsNullOrWhiteSpace($terminText)) { $firstLineParts += [string]$terminText }
  if ($firstLineParts.Count -gt 0) {
    $lines += (($firstLineParts | ForEach-Object { [string]$_ }) -join ' ')
  }
  if (-not [string]::IsNullOrWhiteSpace([string]$order.email)) {
    $lines += ('Mail: ' + [string]$order.email)
  }

  return (($lines | ForEach-Object { [string]$_ }) -join "`n")
}

function Build-ProbeJ {
  param([object]$probe)

  $gewichtValue = '-'
  if ($null -ne $probe.gewicht -and -not [string]::IsNullOrWhiteSpace([string]$probe.gewicht)) {
    $gewichtValue = ([string]$probe.gewicht + ' kg')
  }
  $geruchRaw = if (-not [string]::IsNullOrWhiteSpace([string]$probe.geruch)) { [string]$probe.geruch } else { [string]$probe.geruchAuffaelligkeit }
  $geruchValue = if ([string]::IsNullOrWhiteSpace($geruchRaw)) { '-' } else { [string]$geruchRaw }
  $bemerkungValue = if ([string]::IsNullOrWhiteSpace([string]$probe.bemerkung)) { '-' } else { [string]$probe.bemerkung }

  return "Gewicht: $gewichtValue`nGeruch: $geruchValue`nBemerkung: $bemerkungValue"
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

try {
  Set-ErrorContext -where 'payload.parse' -detail 'reading payload json'
  $payload = Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json
  Write-DebugType -name 'payload.excelPath' -value $payload.excelPath
  Write-DebugType -name 'payload.yearSheetName' -value $payload.yearSheetName
  Write-DebugType -name 'payload.now' -value $payload.now
  if ($null -ne $payload.order) {
    Write-DebugType -name 'order.probenEingangDatum' -value $payload.order.probenEingangDatum
    Write-DebugType -name 'order.terminDatum' -value $payload.order.terminDatum
  }

  Set-ErrorContext -where 'excel.connect' -detail 'get/open excel application'
  $excel = $null
  $createdExcelByScript = $false
  $openedByScript = $false

  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    $excel = $null
  }

  if ($null -eq $excel) {
    $excel = New-Object -ComObject Excel.Application
    $createdExcelByScript = $true
  }
  if ($null -eq $excel) {
    throw 'Excel.Application konnte nicht gestartet werden'
  }

  $excel.Visible = $true
  $excel.DisplayAlerts = $false

  Set-ErrorContext -where 'path.resolve' -detail 'targetPath from payload.excelPath'
  $targetPath = [System.IO.Path]::GetFullPath([string]$payload.excelPath)
  try {
    $resolved = Resolve-Path -LiteralPath $targetPath -ErrorAction Stop
    if ($null -ne $resolved) {
      $targetPath = [string]$resolved.Path
    }
  } catch {
    # keep full path as-is if resolve is not possible yet
  }
  $wb = $null

  Set-ErrorContext -where 'workbook.find' -detail "targetPath=$targetPath"
  if ($null -eq $excel.Workbooks) {
    throw 'Excel.Workbooks ist null'
  }

  foreach ($candidate in $excel.Workbooks) {
    $candidateFullName = [string]$candidate.FullName
    if ([string]::IsNullOrWhiteSpace($candidateFullName)) { continue }
    $candidatePath = [System.IO.Path]::GetFullPath($candidateFullName)
    if ([string]::Equals($candidatePath, $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
      $wb = $candidate
      break
    }
  }

  if ($null -eq $wb) {
    $targetName = [System.IO.Path]::GetFileName($targetPath)
    foreach ($candidate in $excel.Workbooks) {
      if ([string]::Equals([string]$candidate.Name, $targetName, [System.StringComparison]::OrdinalIgnoreCase)) {
        $wb = $candidate
        break
      }
    }
  }

  if ($null -eq $wb) {
    Set-ErrorContext -where 'workbook.open' -detail "targetPath=$targetPath"
    try {
      $wb = $excel.Workbooks.Open($targetPath, $null, $false)
      $openedByScript = $true
    } catch {
      $openMessage = $_.Exception.Message
      throw "Workbook open failed: $openMessage"
    }
  }

  if ($null -eq $wb) {
    Set-ErrorContext -where 'workbook.open' -detail "targetPath=$targetPath"
    throw 'Workbook not found/opened'
  }

  Set-ErrorContext -where 'sheet.get' -detail "yearSheetName=$([string]$payload.yearSheetName)"
  $sheet = $wb.Worksheets.Item([string]$payload.yearSheetName)
  if ($null -eq $sheet) {
    throw "Jahresblatt $($payload.yearSheetName) nicht gefunden"
  }

  Set-ErrorContext -where 'state.compute' -detail 'Get-SheetState + append row'
  $now = [datetime]::Parse([string]$payload.now)
  $state = Get-SheetState -sheet $sheet -now $now
  $firstEmptyAtEnd = Find-FirstCompletelyEmptyRow -sheet $sheet -startRow ([int]$state.lastUsedRow + 1)
  $appendRow = $firstEmptyAtEnd + 1
  Write-DebugType -name 'state.lastUsedRow' -value $state.lastUsedRow
  Write-DebugType -name 'appendRow' -value $appendRow

  Set-ErrorContext -where 'numbering.compute' -detail 'orderNo + lab numbers'
  $todayPrefix = Build-TodayPrefix -now $now
  $nextSeq = [int]$state.maxOrderSeqToday + 1
  if ($nextSeq -gt 99) {
    throw "Maximale Tagessequenz erreicht fuer Prefix $todayPrefix"
  }

  $orderNo = $todayPrefix + $nextSeq.ToString('00')
  $startLabNo = [Math]::Max([int]$state.maxLabNumber, 9999) + 1
  $order = $payload.order
  $sampleNos = @()
  $seenParameterTexts = @{}
  Write-DebugType -name 'todayPrefix' -value $todayPrefix
  Write-DebugType -name 'nextSeq' -value $nextSeq
  Write-DebugType -name 'orderNo' -value $orderNo
  Write-DebugType -name 'startLabNo' -value $startLabNo

  Set-ErrorContext -where 'header.write' -detail "appendRow=$appendRow"
  $pbTypValue = if ([string]::IsNullOrWhiteSpace([string]$order.pbTyp)) { 'PB' } else { [string]$order.pbTyp }
  $sheet.Cells.Item($appendRow, 1).Value2 = $orderNo
  $sheet.Cells.Item($appendRow, 2).Value2 = [string]'y'
  $sheet.Cells.Item($appendRow, 3).Value2 = [string]'y'
  $sheet.Cells.Item($appendRow, 4).Value2 = [string]$order.auftragsnotiz
  $sheet.Cells.Item($appendRow, 5).Value2 = [string]$pbTypValue
  $headerI = [string](Build-HeaderI -order $order)
  $headerJ = [string](Build-HeaderJ -order $order -termin $payload.termin)
  Write-DebugType -name 'headerI' -value $headerI
  Write-DebugType -name 'headerJ' -value $headerJ
  $sheet.Cells.Item($appendRow, 9).Value2 = $headerI
  $sheet.Cells.Item($appendRow, 9).WrapText = $true
  $sheet.Cells.Item($appendRow, 10).Value2 = $headerJ
  $sheet.Cells.Item($appendRow, 10).WrapText = $true
  $sheet.Cells.Item($appendRow, 10).Font.Bold = ($order.eilig -eq $true)

  Set-ErrorContext -where 'samples.write.loop' -detail "sampleCount=$($order.proben.Count)"
  for ($i = 0; $i -lt $order.proben.Count; $i++) {
    Set-ErrorContext -where 'sample.write' -detail "sampleIndex=$i"
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
    Write-DebugType -name "sample[$i].row" -value $row
    Write-DebugType -name "sample[$i].lab" -value $lab
    Write-DebugType -name "sample[$i].parameterText" -value $parameterText
    Write-DebugType -name "sample[$i].probeJ" -value $probeJ

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

  Set-ErrorContext -where 'rows.autofit' -detail "appendRow=$appendRow"
  $lastWrittenRow = $appendRow + $order.proben.Count
  for ($row = $appendRow; $row -le $lastWrittenRow; $row++) {
    $sheet.Rows.Item($row).EntireRow.AutoFit() | Out-Null
  }

  Set-ErrorContext -where 'workbook.save' -detail 'Save only'
  if ($wb.ReadOnly -eq $true) {
    @{ ok = $false; saved = $false; readOnly = $true; reason = 'read-only'; writer = 'com'; targetPath = $targetPath } | ConvertTo-Json -Compress
    exit 1
  }

  $wb.Save()
  if ($openedByScript -eq $true) {
    $wb.Close($true)
  }
  if ($createdExcelByScript -eq $true) {
    $excel.Quit()
  }

  $firstSampleNo = $null
  $lastSampleNo = $null
  if ($sampleNos.Count -gt 0) {
    $firstSampleNo = [int]$sampleNos[0]
    $lastSampleNo = [int]$sampleNos[$sampleNos.Count - 1]
  }
  $endRowRange = "A$appendRow:J$lastWrittenRow"

  @{
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
  } | ConvertTo-Json -Compress
  exit 0
} catch {
  $lineInfo = if ($null -ne $_.InvocationInfo) { "line=$($_.InvocationInfo.ScriptLineNumber)" } else { '' }
  $detailParts = @()
  if (-not [string]::IsNullOrWhiteSpace($script:CurrentDetail)) { $detailParts += [string]$script:CurrentDetail }
  if (-not [string]::IsNullOrWhiteSpace($lineInfo)) { $detailParts += $lineInfo }
  $detailText = ($detailParts | ForEach-Object { [string]$_ }) -join '; '
  @{
    ok = $false
    saved = $false
    writer = 'com'
    saveMethodUsed = 'Save'
    targetPath = $targetPath
    where = [string]$script:CurrentWhere
    detail = [string]$detailText
    error = $_.Exception.Message
  } | ConvertTo-Json -Compress
  exit 1
}
