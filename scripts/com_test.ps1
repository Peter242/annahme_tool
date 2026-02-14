param(
  [Parameter(Mandatory = $true)]
  [string]$PayloadPath
)

$ErrorActionPreference = 'Stop'

$targetPath = $null

try {
  $payload = Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json

  if ($payload.PSObject.Properties.Name -contains 'workbookFullName' -and -not [string]::IsNullOrWhiteSpace([string]$payload.workbookFullName)) {
    $targetPath = [System.IO.Path]::GetFullPath([string]$payload.workbookFullName)
  } elseif ($payload.PSObject.Properties.Name -contains 'excelPath' -and -not [string]::IsNullOrWhiteSpace([string]$payload.excelPath)) {
    $targetPath = [System.IO.Path]::GetFullPath([string]$payload.excelPath)
  } else {
    throw 'Kein gueltiger Zielpfad uebergeben (excelPath/workbookFullName)'
  }

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
    $excel = $null
  }

  $openedByScript = $false
  $createdExcelByScript = $false
  if ($null -eq $excel) {
    $excel = New-Object -ComObject Excel.Application
    $createdExcelByScript = $true
  }

  $wb = $null
  foreach ($candidate in $excel.Workbooks) {
    if ([string]::Equals([System.IO.Path]::GetFullPath([string]$candidate.FullName), $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
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
    if (-not $createdExcelByScript) {
      $excel = New-Object -ComObject Excel.Application
      $createdExcelByScript = $true
    }
    $wb = $excel.Workbooks.Open($targetPath, $null, $false)
    $openedByScript = $true
  }

  $sheet = $wb.Worksheets.Item($sheetName)
  if ($null -eq $sheet) {
    throw "Blatt nicht gefunden: $sheetName"
  }

  $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $value = "COM_OK_$timestamp"

  $sheet.Range($cellAddress).Value2 = $value
  if ($wb.ReadOnly -eq $true) {
    @{ ok = $false; saved = $false; readOnly = $true; reason = 'read-only'; targetPath = $targetPath } | ConvertTo-Json -Compress
    exit 1
  }

  if ($wb.Saved -eq $false) {
    $wb.Save()
  } else {
    $wb.Save()
  }

  if ($openedByScript -eq $true) {
    $wb.Close($true)
  }
  if ($createdExcelByScript -eq $true) {
    $excel.Quit()
  }

  @{ ok = $true; writtenValue = $value; saved = $true; saveMethodUsed = 'Save'; targetPath = $targetPath } | ConvertTo-Json -Compress
  exit 0
} catch {
  @{ ok = $false; saved = $false; error = $_.Exception.Message; saveMethodUsed = 'Save'; targetPath = $targetPath } | ConvertTo-Json -Compress
  exit 1
}
