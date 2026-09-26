$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$templateFiles = @(
    "E:\Ace\ExcelAddInCTtools\Resources\TenderReport_Regular.xlsx",
    "E:\Ace\ExcelAddInCTtools\Resources\CabinetTemplate.xlsx",
    "E:\Ace\draw-code\publish\wwwroot\downloads\报表模板.xlsx",
    "E:\Ace\draw-code\publish\wwwroot\downloads\苏北电影院配电箱采购清单240909.xlsx"
)

foreach ($f in $templateFiles) {
    if (Test-Path $f) {
        try {
            $wb = $excel.Workbooks.Open($f, [Type]::Missing, $true)
            Write-Host ("=== File: " + $f + " ===")
            foreach ($sh in $wb.Worksheets) {
                Write-Host ("  Sheet: " + $sh.Name)
            }
            $wb.Close($false)
        } catch {
            Write-Host ("Failed to open: " + $f)
        }
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

foreach ($d in $userDirs) {
    if (Test-Path $d) {
        Get-ChildItem -Path $d -Filter "*.xlsx" -ErrorAction SilentlyContinue | ForEach-Object {
            $candidateFiles += $_.FullName
        }
    }
}

foreach ($f in $candidateFiles) {
    if (Test-Path $f) {
        try {
            $wb = $excel.Workbooks.Open($f, [Type]::Missing, $true)
            Write-Host ("=== File: " + $f + " ===")
            foreach ($sh in $wb.Worksheets) {
                Write-Host ("  Sheet: " + $sh.Name)
            }
            $wb.Close($false)
        } catch {
            Write-Host ("Failed to open: " + $f)
        }
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null


