# Install the PowerShell module from the gallery

`Install-Module -Name ImportExcel`

Now you’re ready to run the script and get the results.
The output of `Get-Process`, `Get-Service`, and `Get-ChildItem` are piped to Excel, to the same workbook `$xlfile.`
By not specifying a `-WorkSheetName`, Export-Excel will insert the data to the same sheet, in this case the default `Sheet1`.
