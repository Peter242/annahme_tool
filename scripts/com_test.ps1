param(
  [Parameter(Mandatory = $true)]
  [string]$PayloadPath
)

$ErrorActionPreference = 'Stop'

try {
  $payload = Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json
  $excelPath = [System.IO.Path]::GetFullPath([string]$payload.excelPath)
  $cellPath = [string]$payload.cellPath

  if ([string]::IsNullOrWhiteSpace($cellPath) -or -not $cellPath.Contains('!')) {
    throw 'cellPath muss im Format Blatt!Zelle angegeben werden, z.B. 2026!Z1'
  }

  $parts = $cellPath.Split('!', 2)
  $sheetName = $parts[0].Trim()
  $cellAddress = $parts[1].Trim().ToUpperInvariant()

  if ([string]::IsNullOrWhiteSpace($sheetName) -or [string]::IsNullOrWhiteSpace($cellAddress)) {
    throw 'Ungueltiger cellPath'
  }

  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    throw 'Keine laufende Excel Instanz gefunden (GetActiveObject fehlgeschlagen)'
  }

  $wb = $null
  foreach ($candidate in $excel.Workbooks) {
    if ([string]::Equals([System.IO.Path]::GetFullPath([string]$candidate.FullName), $excelPath, [System.StringComparison]::OrdinalIgnoreCase)) {
      $wb = $candidate
      break
    }
  }

  if ($null -eq $wb) {
    throw "Zielarbeitsmappe ist nicht in der offenen Excel Instanz: $excelPath"
  }

  $sheet = $wb.Worksheets.Item($sheetName)
  if ($null -eq $sheet) {
    throw "Blatt nicht gefunden: $sheetName"
  }

  $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $value = "COM_OK_$timestamp"

  $sheet.Range($cellAddress).Value2 = $value
  $wb.Save()

  @{ ok = $true; writtenValue = $value } | ConvertTo-Json -Compress
  exit 0
} catch {
  @{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress
  exit 1
}
