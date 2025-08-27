# set folders
$downloadfolder = "C:\Users\"         # folder where the .xls files are
$uploadfolder   = "C:\Users\Upload"  # folder that uploads the .xlsx files
$backupfolder   = "C:\Users\Backup"  # folder that has .xls files as backup

# open and convert xls to xlsx
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # it is much faster if Excel is not visible

# loop through the .xls files and process them
$list = @()
$list = Get-ChildItem -Path $downloadfolder -Include ('*.xls') -Recurse -File
ForEach($file in $list) {
    try {
        $xlsfilename = $file.FullName
        #copy file to backup folder
        Copy-Item -Path $xlsfilename -Destination $backupfolder -Force
        # open the xls
        Write-Host "Converting $xlsfilename"
        $workbook = $excel.Workbooks.Open($xlsfilename)
        # save converted file (as xlsx) directly to the upload folder
        $newfilename = Join-Path -Path $uploadfolder -ChildPath ('{0}.xlsx' -f $file.BaseName)
        $workbook.SaveAs($newfilename, $xlFixedFormat)
        $workbook.Close()
        #remove old file
        Write-Host "Delete old file '$xlsfilename'"
        Remove-Item -Path $xlsfilename -Force
    }
    catch {
        # write out a warning as to why something went wrong
        Write-Warning "Could not convert '$xlsfilename':`r`n$($_.Exception.Message)"
    }
}
# close excel
$excel.Quit()
# garbage collection
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
