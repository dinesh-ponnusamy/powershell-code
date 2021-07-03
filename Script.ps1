#Current Folder path 
$current_folder_path = "C:\Temp" 
cd $current_folder_path 

#Variable var
$var = (1,2),(2,2),(3,2),(4,2),(5,2),(6,2) 

#Variable var1
[array]$var1 = @(1..6) | ForEach-Object { ,$(@($_,2)) }

write-host -f green $var.GetType()
write-host -f green $var
write-host -f green $var1.GetType()
write-host -f green $var1
Compare-Object $var $var1 -IncludeEqual

#Opening Excel
$xl = New-Object -ComObject Excel.Application -Property @{visible = $true}
$xl.DisplayAlerts = $false
$wbt = $xl.Workbooks.Open("$current_folder_path\INPUT.TXT")
$wst = $wbt.Worksheets.Item(1)
$colA=$wst.range("A1").EntireColumn
$colrange=$wst.range("A1")

#Direct value 
#[void]$colA.texttocolumns($colrange,1,-4142,$false,$false,$false,$false,$false,$true,"|",@((1,2),(2,2),(3,2),(4,2),(5,2),(6,2)))

#Variable Var 
#[void]$colA.texttocolumns($colrange,1,-4142,$false,$false,$false,$false,$false,$true,"|",$var)

#variable Var1
[void]$colA.texttocolumns($colrange,1,-4142,$false,$false,$false,$false,$false,$true,"|",$var1)

[void]$wbt.SaveAs("$current_folder_path\Output.xlsx",51)

#Closing Excel Properly
$xl.Quit()
Get-Process Excel
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)