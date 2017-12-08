#Export-Xls function was found on http://www.lucd.info/2010/05/29/beyond-export-csv-export-xls/
#requires -version 2
function Export-Xls{
<#
.SYNOPSIS
Saves Microsoft .NET Framework objects to a worksheet in an XLS file
.DESCRIPTION
The Export-Xls function allows you to save Microsoft .NET Framework objects
to a named worksheet in an Excel file (type XLS). The position of the
worksheet can be specified.
.NOTES
Author:  Luc Dekens
.PARAMETER InputObject
Specifies the objects to be written to the worksheet. The parameter accepts
objects through the pipeline.
.PARAMETER Path
Specifies the path to the XLS file.
.PARAMETER WorksheetName
The name for the new worksheet. If not specified the name will
be "Sheet" followed by the "Ticks" value
.PARAMETER SheetPosition
Specifies where the new worksheet will be inserted in the series of
existing worksheets. You can specify "begin" or "end". The default
is "begin".
.PARAMETER ChartType
Specifies the type of chart you want add to the worksheet.
All types in the [microsoft.Office.Interop.Excel.XlChartType]
enumeration are accepted.
.PARAMETER NoTypeInformation
Omits the type information from the worksheet. The default is to
include the "#TYPE" line.
.PARAMETER AppendWorksheet
Specifies if the worksheet should keep or remove the existing
worksheet in the spreadsheet. The default is to append.
.EXAMPLE
PS> $data = Get-Process | Select-Object Name, Id, WS
PS> Export-Xls $data C:\Reports\MyWkb.xls -WorksheetName "WS" -AppendWorksheet:$false
.EXAMPLE
PS> $data = Get-Process | Select-Object Name, Id, WS
PS> Export-Xls $data C:\Reports\MyWkb.xls -SheetPosition "end"
.EXAMPLE
PS> $data = Get-Process | Select-Object Name, Id, WS
PS> Export-Xls $data C:\Reports\MyWkb.xls -WorksheetName "WS" -ChartType "xlColumnClustered"
#>
param(
[parameter(ValueFromPipeline = $true,Position=1)]
[ValidateNotNullOrEmpty()]
$InputObject,
[parameter(Position=2)]
[ValidateNotNullOrEmpty()]
[string]$Path,
[string]$WorksheetName = ("Sheet " + (Get-Date).Ticks),
[string]$SheetPosition = "begin",
[PSObject]$ChartType,
[switch]$NoTypeInformation = $true,
[switch]$AppendWorksheet = $true
)
begin{
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")
if($ChartType){
[microsoft.Office.Interop.Excel.XlChartType]$ChartType = $ChartType
}
function Set-ClipBoard{
param(
[string]$text
)
process{
Add-Type -AssemblyName System.Windows.Forms
$tb = New-Object System.Windows.Forms.TextBox
$tb.Multiline = $true
$tb.Text = $text
$tb.SelectAll()
$tb.Copy()
}
}
function Add-Array2Clipboard {
param (
[PSObject[]]$ConvertObject,
[switch]$Header
)
process{
$array = @()
if ($Header) {
$line =""
$ConvertObject | Get-Member -MemberType Property,NoteProperty,CodeProperty | Select -Property Name | %{
$line += ($_.Name.tostring() + "`t")
}
$array += ($line.TrimEnd("`t") + "`r")
}
else {
foreach($row in $ConvertObject){
$line =""
$row | Get-Member -MemberType Property,NoteProperty | %{
$Name = $_.Name
if(!$Row.$Name){$Row.$Name = ""}
$line += ([string]$Row.$Name + "`t")
}
$array += ($line.TrimEnd("`t") + "`r")
}
}
Set-ClipBoard $array
}
}
$excelApp = New-Object -ComObject "Excel.Application"
$originalAlerts = $excelApp.DisplayAlerts
$excelApp.DisplayAlerts = $false
if(Test-Path -Path $Path -PathType "Leaf"){
$workBook = $excelApp.Workbooks.Open($Path)
}
else{
$workBook = $excelApp.Workbooks.Add()
}
$sheet = $excelApp.Worksheets.Add($workBook.Worksheets.Item(1))
if(!$AppendWorksheet){
$workBook.Sheets | where {$_ -ne $sheet} | %{$_.Delete()}
}
$sheet.Name = $WorksheetName
if($SheetPosition -eq "end"){
$nrSheets = $workBook.Sheets.Count
2..($nrSheets) |%{
$workbook.Sheets.Item($_).Move($workbook.Sheets.Item($_ - 1))
}
}
$sheet.Activate()
$array = @()
}
process{
$array += $InputObject
}
end{
Add-Array2Clipboard $array -Header:$True
$selection = $sheet.Range("A1")
$selection.Select() | Out-Null
$sheet.Paste()
$Sheet.UsedRange.HorizontalAlignment = [microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignCenter
Add-Array2Clipboard $array
$selection = $sheet.Range("A2")
$selection.Select() | Out-Null
$sheet.Paste() | Out-Null
$selection = $sheet.Range("A1")
$selection.Select() | Out-Null
$sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
$workbook.Sheets.Item(1).Select()
if($ChartType){
$sheet.Shapes.AddChart($ChartType) | Out-Null
}
$workbook.SaveAs($Path)
$excelApp.DisplayAlerts = $originalAlerts
$excelApp.Quit()
Stop-Process -Name "Excel"
}
}




function GenerateVMList([parameter(Mandatory=$true)][string]$VMHostList ,[parameter(Mandatory=$true)][string]$ExportPath){

$VMHostCredentials = (get-credential)

    if(Test-Path $VMHostList){
    
        Get-Content $VMHostList |  Foreach-Object {
        Connect-VIServer -Server $_ -Credential $VMHostCredentials  
        }       
         
    $ExportVmList =  Get-VM | Select -Property VMHost, name ,@{N='IP';E={[string]::Join(',',$_.Guest.IPAddress)}},@{N="DNS";E={[string]::Join(',',($_.ExtensionData.Guest.IpStack.DnsConfig.IpAddress))}} 
    Export-Xls $ExportVmList $ExportPath
    }
}
