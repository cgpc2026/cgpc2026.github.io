$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open("d:\org\cgpc2026.github.io\CCF启明星计划（2026）-山东大学.xlsx")
$sheet = $workbook.Sheets.Item("5.08")
$usedRange = $sheet.UsedRange
$data = @()
for ($row = 1; $row -le $usedRange.Rows.Count; $row++) {
    $rowData = @()
    for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
        $rowData += $sheet.Cells.Item($row, $col).Text
    }
    $data += ($rowData -join "`t")
}
$workbook.Close($false)
$excel.Quit()
$data
