#Identify what directory to hash
Function Get-Folder($hashdirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select the directory to hash"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
$hashdirectory = Get-Folder

#Identify where to save the output files
Function Get-Folder($outputdirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select where to output results"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
$path = Get-Folder


$hashdirectory = $hashdirectory + '\'
$csvdirectory = $path + '\hashes.csv'
$xlsfile = $path + '\hashes.xlsx'



$job = Start-Job -ScriptBlock {Get-ChildItem $args[0] -Recurse |get-filehash -algorithm MD5| export-csv $args[1]} -ArgumentList $hashdirectory, $csvdirectory

Wait-Job $job


$XL = New-Object -ComObject Excel.Application
$XL.visible = $False

#Define variable: $Workbook
$Workbook = $XL.workbooks.open($csvdirectory)

#Define variable: $Worksheets
$Worksheets = $Workbook.worksheets

#Define variable "$Worksheet" as the first worksheet in "$Workbook"
$Worksheet = $Workbook.Worksheets.Item(1)

#Define variable "$usedarea"
$usedarea = $worksheet.UsedRange

#Save the file
#$Workbook.SaveAs($xlsfile,51)

#Delete first row and first column that are not needed for sorting
$worksheet.Cells.Item(1,1).EntireRow.Delete()
$Worksheet.Cells.Item(1,1).EntireColumn.Delete()

#sort data by hash value
$sortcolumn = $worksheet.Range("A1:A10000")
$usedarea.Sort($sortcolumn)

#auto fit the column size
$Worksheet.cells.entirecolumn.autofit()

#Sets loop variables
$counter = 1
$counterplus = $counter + 1
$counterless = $counter - 1

#Selects cell (A,1) in the current worksheet (row,column)
$Worksheet.Cells($counter,1).select()

Do 
{
    While ($Worksheet.Cells.Item($counter,1).value() -ne $Null)
   {
	    If ($Worksheet.Cells.Item($counter,1).Value() -ne $Worksheet.Cells.Item($counterplus,1).Value()) 
        {
            $Worksheet.Cells.Item($counter,1).EntireRow.Delete()
         }
         else
         { 
            While ($Worksheet.Cells.Item($counter,1).value() -ne $Null)
            {

	            If ($Worksheet.Cells.Item($counter,1).Value() -eq $Worksheet.Cells.Item($counterplus,1).Value() -or $Worksheet.Cells.Item($counter,1).Value() -eq $Worksheet.Cells.Item($counterless,1).Value()) 
                {
                    $counter=$counter + 1
                    $counterplus=$counterplus + 1
                    $counterless=$counterless + 1
                }
	            Else {$Worksheet.Cells.Item($counter,1).EntireRow.Delete()}
            } 
	    }
    }
        $counter=$counter + 1
        $counterplus=$counterplus + 1
        $counterless=$counterless + 1    
} until ($counter -ne 1)

#Insert row at top and label columns
$worksheet.cells.item(1,1).entireRow.activate()
$worksheet.cells.item(1,1).entireRow.insert()
$worksheet.cells.item(1,1) = 'File Hash'
$worksheet.cells.item(1,2) = 'File Path'

$Workbook.SaveAs($xlsfile,51)
$XL.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL)
Remove-Variable XL

