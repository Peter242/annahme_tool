param(
  [Parameter(Mandatory = $true)]
  [string]$PayloadPath
)

$ErrorActionPreference = 'Stop'

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

  $usedRows = $sheet.UsedRange.Rows.Count
  $lastUsedRow = 0
  $maxLabNumber = 0
  $maxOrderSeqToday = 0
  $todayPrefix = Build-TodayPrefix -now $now
  $orderPattern = '^' + [regex]::Escape($todayPrefix) + '(\d{2})(?!\d)'

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

    $orderMatch = [regex]::Match($a, $orderPattern)
    if ($orderMatch.Success) {
      $seq = [int]$orderMatch.Groups[1].Value
      if ($seq -gt $maxOrderSeqToday) { $maxOrderSeqToday = $seq }
    }

    $isHeader = [regex]::IsMatch($a, '^\d{9}(?!\d)')
    if ($isHeader) { continue }

    $labMatch = [regex]::Match($a, '^(\d{5,6})(?!\d)')
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

function Build-HeaderI {
  param([object]$order)

  $lines = @()
  $values = @(
    $order.auftraggeberKurz,
    $order.kunde,
    $order.ansprechpartner,
    $order.projektnummer,
    $order.projektname,
    $order.projekt,
    $order.probenahmedatum,
    $order.probenEingangDatum
  )

  $auftraggeber = if ([string]::IsNullOrWhiteSpace([string]$order.auftraggeberKurz)) { [string]$order.kunde } else { [string]$order.auftraggeberKurz }
  $projektname = if ([string]::IsNullOrWhiteSpace([string]$order.projektname)) { [string]$order.projekt } else { [string]$order.projektname }
  $probenahmedatum = if ([string]::IsNullOrWhiteSpace([string]$order.probenahmedatum)) { [string]$order.probenEingangDatum } else { [string]$order.probenahmedatum }

  foreach ($item in @($auftraggeber, [string]$order.ansprechpartner, [string]$order.projektnummer, $projektname, $probenahmedatum)) {
    if (-not [string]::IsNullOrWhiteSpace($item)) {
      $lines += $item
    }
  }

  return [string]::Join("`n", $lines)
}

function Format-GermanTermin {
  param([string]$ymd)

  if ([string]::IsNullOrWhiteSpace($ymd)) { return '' }
  $parsed = $null
  if (-not [datetime]::TryParseExact($ymd.Trim(), 'yyyy-MM-dd', $null, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
    return $ymd
  }

  $weekdays = @('So','Mo','Di','Mi','Do','Fr','Sa')
  $wd = $weekdays[[int]$parsed.DayOfWeek]
  return ('{0} {1:dd.MM.yyyy}' -f $wd, $parsed)
}

function Build-HeaderJ {
  param(
    [object]$order,
    [string]$termin
  )

  $lines = @()
  if (-not [string]::IsNullOrWhiteSpace([string]$order.erfasstKuerzel)) {
    $lines += ('Erfasst: ' + [string]$order.erfasstKuerzel)
  }

  $terminValue = if (-not [string]::IsNullOrWhiteSpace([string]$order.terminDatum)) { [string]$order.terminDatum } else { [string]$termin }
  if (-not [string]::IsNullOrWhiteSpace($terminValue)) {
    $lines += ('Termin: ' + (Format-GermanTermin -ymd $terminValue))
  }

  if ($order.eilig -eq $true) { $lines += 'Eilauftrag' }
  if (-not [string]::IsNullOrWhiteSpace([string]$order.email)) { $lines += ('Email: ' + [string]$order.email) }
  if ($order.probeNochNichtDa -eq $true -or $order.sampleNotArrived -eq $true) { $lines += 'Probe noch nicht da' }

  return [string]::Join("`n", $lines)
}

function Build-ProbeJ {
  param([object]$probe)

  $lines = @()
  if ($null -ne $probe.gewicht -and -not [string]::IsNullOrWhiteSpace([string]$probe.gewicht)) {
    $unit = if (-not [string]::IsNullOrWhiteSpace([string]$probe.gewichtEinheit)) { ' ' + [string]$probe.gewichtEinheit } else { '' }
    $lines += ('Gewicht: ' + [string]$probe.gewicht + $unit)
  }

  if (-not [string]::IsNullOrWhiteSpace([string]$probe.geruchAuffaelligkeit)) {
    $lines += ('Geruch: ' + [string]$probe.geruchAuffaelligkeit)
  }

  if (-not [string]::IsNullOrWhiteSpace([string]$probe.bemerkung)) {
    $lines += ('Bemerkung: ' + [string]$probe.bemerkung)
  }

  return [string]::Join("`n", $lines)
}

function Get-LineCount {
  param([string]$text)
  if ([string]::IsNullOrEmpty($text)) { return 1 }
  return 1 + ([regex]::Matches($text, "`n")).Count
}

function Clamp {
  param([double]$value, [double]$min, [double]$max)
  if ($value -lt $min) { return $min }
  if ($value -gt $max) { return $max }
  return $value
}

try {
  $payload = Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json
  $excel = $null

  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    $excel = New-Object -ComObject Excel.Application
  }

  $targetPath = [System.IO.Path]::GetFullPath([string]$payload.excelPath)
  $wb = $null

  foreach ($candidate in $excel.Workbooks) {
    if ([string]::Equals([System.IO.Path]::GetFullPath([string]$candidate.FullName), $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
      $wb = $candidate
      break
    }
  }

  if ($null -eq $wb) {
    $wb = $excel.Workbooks.Open($targetPath)
  }

  $sheet = $wb.Worksheets.Item([string]$payload.yearSheetName)
  if ($null -eq $sheet) {
    throw "Jahresblatt $($payload.yearSheetName) nicht gefunden"
  }

  $now = [datetime]::Parse([string]$payload.now)
  $state = Get-SheetState -sheet $sheet -now $now
  $appendRow = [int]$state.lastUsedRow + 2

  $todayPrefix = Build-TodayPrefix -now $now
  $nextSeq = [int]$state.maxOrderSeqToday + 1
  if ($nextSeq -gt 99) {
    throw "Maximale Tagessequenz erreicht fuer Prefix $todayPrefix"
  }

  $orderNo = $todayPrefix + $nextSeq.ToString('00')
  $startLabNo = [int]$state.maxLabNumber + 1
  $order = $payload.order

  $sheet.Cells.Item($appendRow, 1).Value2 = $orderNo
  $sheet.Cells.Item($appendRow, 2).Value2 = 'y'
  $sheet.Cells.Item($appendRow, 3).Value2 = 'y'
  $sheet.Cells.Item($appendRow, 4).Value2 = [string]$order.auftragsnotiz
  $sheet.Cells.Item($appendRow, 5).Value2 = if ([string]::IsNullOrWhiteSpace([string]$order.pbTyp)) { 'PB' } else { [string]$order.pbTyp }
  $sheet.Cells.Item($appendRow, 9).Value2 = Build-HeaderI -order $order
  $sheet.Cells.Item($appendRow, 10).Value2 = Build-HeaderJ -order $order -termin ([string]$payload.termin)

  for ($i = 0; $i -lt $order.proben.Count; $i++) {
    $probe = $order.proben[$i]
    $row = $appendRow + 1 + $i
    $lab = $startLabNo + $i

    $parameterText = [string]$probe.parameterTextPreview
    $probeJ = Build-ProbeJ -probe $probe

    $sheet.Cells.Item($row, 1).Value2 = $lab
    $sheet.Cells.Item($row, 4).Value2 = $parameterText
    $sheet.Cells.Item($row, 4).WrapText = $true
    $sheet.Cells.Item($row, 6).Value2 = [string]$probe.probenbezeichnung

    $gValue = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$probe.tiefeVolumen)) {
      $gValue = [string]$probe.tiefeVolumen
    } elseif ($null -ne $probe.volumen -and -not [string]::IsNullOrWhiteSpace([string]$probe.volumen)) {
      $gValue = [string]$probe.volumen
    }
    $sheet.Cells.Item($row, 7).Value2 = $gValue

    $material = if ($probe.material) { [string]$probe.material } else { '' }
    $gebinde = if ($probe.gebinde) { [string]$probe.gebinde } else { '' }
    $materialGebinde = if ($probe.materialGebinde) { [string]$probe.materialGebinde } else { ($material + ' ' + $gebinde).Trim() }
    if ([string]::IsNullOrWhiteSpace($materialGebinde)) { $materialGebinde = [string]$probe.matrixTyp }
    $sheet.Cells.Item($row, 8).Value2 = $materialGebinde

    $sheet.Cells.Item($row, 10).Value2 = $probeJ
    $sheet.Cells.Item($row, 10).WrapText = $true

    $lineCount = [Math]::Max((Get-LineCount -text $parameterText), (Get-LineCount -text $probeJ))
    $rowHeight = Clamp -value (($lineCount * 15) + 5) -min 15 -max 300
    $sheet.Rows.Item($row).RowHeight = $rowHeight
  }

  $wb.Save()

  @{ ok = $true; orderNo = $orderNo; appendRow = $appendRow; startLabNo = $startLabNo } | ConvertTo-Json -Compress
  exit 0
} catch {
  @{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress
  exit 1
}
