$processes=Get-Process
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.add()
$sheet = $workbook.worksheets.Item(1)
$workbook.Worksheets.item(3).delete()
$workbook.Worksheets.item(2).delete()
$workbook.Worksheets.item(1).name="Processes"
$sheet = $workbook.WorkSheets.Item("Processes")
$x = 2
$lineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
$colorIndex = "microsoft.office.interop.excel.xlColorIndex" -as [type]
$borderWeight = "microsoft.office.interop.excel.xlBorderWeight" -as [type]
$chartType = "microsoft.office.interop.excel.xlChartType" -as [type]
for($b = 1 ; $b -le 2 ; $b++)
{
 $sheet.cells.item(1,$b).font.bold = $true
 $sheet.cells.item(1,$b).borders.LineStyle = $lineStyle::xlDashDot
 $sheet.cells.item(1,$b).borders.ColorIndex = $colorIndex::xlColorIndexAutomatic
 $sheet.cells.item(1,$b).borders.weight = $borderWeight::xlMedium
}
$sheet.cells.item(1,1) = "Name of Process"
$sheet.cells.item(1,2) = "Working Set Size"
foreach($process in $processes)
{
 $sheet.cells.item($x, 1) = $process.name
 $sheet.cells.item($x,2) = $process.workingSet
 $x++
} #end foreach
$range = $sheet.usedRange
$range.EntireColumn.AutoFit() | out-null