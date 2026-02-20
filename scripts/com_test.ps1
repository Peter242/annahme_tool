param(
  [Parameter(Mandatory = $false)]
  [string]$PayloadPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$targetPath = $null
$mode = ''
$excelOpenRequiredMessage = ('Fehler: Annahme muss ge' + [char]246 + 'ffnet sein. Bitte ' + [char]246 + 'ffnen und erneut versuchen')
$forbiddenAttachOnlyMessage = 'FORBIDDEN: attempted to start Excel or open workbook'

function ToStr {
  param([object]$v)
  if ($null -eq $v) { return '' }
  if ($v -is [System.Array] -and -not ($v -is [string])) {
    return (($v | ForEach-Object { if ($null -eq $_) { '' } else { [string]$_ } }) -join '')
  }
  return [string]$v
}

try {
  $payloadJson = ''
  if (-not [string]::IsNullOrWhiteSpace($PayloadPath)) {
    $payloadJson = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8
  } else {
    $payloadJson = [Console]::In.ReadToEnd()
  }
  if ([string]::IsNullOrWhiteSpace($payloadJson)) {
    throw 'Payload fehlt (stdin/PayloadPath leer)'
  }
  $payload = $payloadJson | ConvertFrom-Json

  if ($payload.PSObject.Properties.Name -contains 'workbookFullName' -and -not [string]::IsNullOrWhiteSpace((ToStr $payload.workbookFullName))) {
    $targetPath = [System.IO.Path]::GetFullPath((ToStr $payload.workbookFullName))
  } elseif ($payload.PSObject.Properties.Name -contains 'excelPath' -and -not [string]::IsNullOrWhiteSpace((ToStr $payload.excelPath))) {
    $targetPath = [System.IO.Path]::GetFullPath((ToStr $payload.excelPath))
  } else {
    throw 'Kein gueltiger Zielpfad uebergeben (excelPath/workbookFullName)'
  }

  try {
    $resolved = Resolve-Path -LiteralPath $targetPath -ErrorAction Stop
    if ($null -ne $resolved) {
      $targetPath = [string]$resolved.Path
    }
  } catch {
    # keep normalized full path
  }

  $cellPath = ToStr $payload.cellPath
  if ([string]::IsNullOrWhiteSpace($cellPath) -or -not $cellPath.Contains('!')) {
    throw 'cellPath muss im Format Blatt!Zelle angegeben werden, z.B. 2026!Z1'
  }

  $parts = $cellPath.Split('!', 2)
  $sheetName = $parts[0].Trim()
  $cellAddress = $parts[1].Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($sheetName) -or [string]::IsNullOrWhiteSpace($cellAddress)) {
    throw 'Ungueltiger cellPath'
  }

  $excel = $null
  $wb = $null
  $allowAutoOpenExcel = $false
  if ($payload.PSObject.Properties.Name -contains 'allowAutoOpenExcel') {
    try {
      $allowAutoOpenExcel = [bool]$payload.allowAutoOpenExcel
    } catch {
      $allowAutoOpenExcel = $false
    }
  }
  if ($allowAutoOpenExcel -eq $true) {
    throw $forbiddenAttachOnlyMessage
  }

  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    $excel = $null
  }

  if ($null -eq $excel) {
    [Console]::Error.WriteLine('[worker] attach failed: no running Excel')
    throw $excelOpenRequiredMessage
  }

  foreach ($candidate in $excel.Workbooks) {
    $candidateFullName = ToStr $candidate.FullName
    if ([string]::IsNullOrWhiteSpace($candidateFullName)) { continue }
    $candidatePath = [System.IO.Path]::GetFullPath($candidateFullName)
    if ([string]::Equals($candidatePath, $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
      $wb = $candidate
      $mode = 'attached'
      break
    }
  }

  if ($null -eq $wb) {
    $targetName = [System.IO.Path]::GetFileName($targetPath)
    foreach ($candidate in $excel.Workbooks) {
      if ([string]::Equals((ToStr $candidate.Name), $targetName, [System.StringComparison]::OrdinalIgnoreCase)) {
        $wb = $candidate
        $mode = 'attached'
        break
      }
    }
  }

  if ($null -eq $wb) {
    [Console]::Error.WriteLine(('[worker] attach failed: workbook not open: ' + [string]$targetPath))
    throw $excelOpenRequiredMessage
  }

  $sheet = $wb.Worksheets.Item($sheetName)
  if ($null -eq $sheet) {
    throw "Blatt nicht gefunden: $sheetName"
  }

  $valueRaw = ToStr $payload.value
  $value = if ([string]::IsNullOrWhiteSpace($valueRaw)) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    "COM_OK_$timestamp"
  } else {
    $valueRaw
  }
  $sheet.Range($cellAddress).Value2 = [string]$value

  if ($wb.ReadOnly -eq $true) {
    @{
      ok = $false
      saved = $false
      readOnly = $true
      reason = 'read-only'
      targetPath = $targetPath
      mode = $mode
    } | ConvertTo-Json -Compress
    exit 1
  }

  $wb.Save()
  $readbackValue = ToStr $sheet.Range($cellAddress).Value2
  $workbookFullName = ToStr $wb.FullName
  $workbookName = ToStr $wb.Name
  $excelVersion = ToStr $excel.Version
  $excelHwnd = $null
  try {
    $excelHwnd = [int]$excel.Hwnd
  } catch {
    $excelHwnd = $null
  }

  @{
    ok = $true
    writtenValue = $value
    readbackValue = $readbackValue
    cellPath = $cellPath
    saved = $true
    saveMethodUsed = 'Save'
    targetPath = $targetPath
    workbookFullName = $workbookFullName
    workbookName = $workbookName
    excelVersion = $excelVersion
    excelHwnd = $excelHwnd
    mode = $mode
  } | ConvertTo-Json -Compress
  exit 0
} catch {
  @{
    ok = $false
    saved = $false
    error = $_.Exception.Message
    saveMethodUsed = 'Save'
    targetPath = $targetPath
    mode = $mode
  } | ConvertTo-Json -Compress
  exit 1
}

