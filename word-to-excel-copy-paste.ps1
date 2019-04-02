# copies all text from a set of Word documents into Excel files of the same name
# original concept: VBA code at https://social.technet.microsoft.com/Forums/office/en-US/0a73df0b-096b-40eb-a65a-9b16bc6a57ed/need-macro-to-copy-text-from-word-to-excel-help?forum=word

# &('SCRIPTDIR\word-to-excel-copy-paste.ps1') -path $path -ext $ext -delrows $delrows

# require a path to be passed, will recurse over this directory
# other params optional
param (
  [string]$path = $(throw "-path is required"),
  [string]$ext = "doc",
  [switch]$autofit = $false,
  [switch]$delrows = $false
)

# creates an instance of Word as a COM object, so this is a memory-intensive way (basically like using a macro) 
$objWord = New-Object -ComObject word.application
# run in background
$objWord.visible = $false

# creates an instance of Excel as a COM object, so this is a memory-intensive way (basically like using a macro) 
$objExcel = New-Object -ComObject excel.application 
# run in background
$objExcel.visible = $false 

$filter = "*." + $ext

# include all the files of the given extension in this directory and its children
$wordFiles = Get-ChildItem -Path $path -recurse -include $filter 

# for each of the odt files create an excel doc with the same name and copied contents
foreach($wordFile in $wordFiles) 
{ 
 # may not be needed (for saving it later)
 $filepath = Join-Path -Path $path -ChildPath ($wordFile.BaseName + “.xlsx”) 
 Write-Host $filepath

 $wordDoc = $objWord.documents.open($wordFile.fullname) 

 # unchanged since open - https://docs.microsoft.com/en-us/office/vba/api/word.document.saved
 $wordDoc.Saved = $true 
 
 # copy the text from Word
 # see https://learn-powershell.net/2014/12/31/beginning-with-powershell-and-word/
 $selection = $objWord.Selection
 $selection.WholeStory()
 $selection.Copy()

 # create new spreadsheet - https://community.spiceworks.com/how_to/137203-create-an-excel-file-from-within-powershell
 $workbook = $objExcel.Workbooks.Add()

 # name the new worksheet where the data will be pasted
 $wksht = $workbook.Worksheets.Item(1)
 $wksht.Name = $wordFile.BaseName

 # paste the content
 $wksht.Paste()

 # auto-resize columns and rows so all data shows properly
 if($autofit) {
   $usedRange = $wksht.UsedRange
   $usedRange.EntireColumn.AutoFit() | Out-Null
   $usedRange.EntireRow.AutoFit() | Out-Null
 }

 # if set to delete empty rows at start then delete them
 # https://social.technet.microsoft.com/Forums/en-US/b0bd07a9-f437-456f-a656-ee896dd5fe80/deleting-empty-lines-in-an-excel-file?forum=winserverpowershell
 # https://devblogs.microsoft.com/scripting/powershell-looping-understanding-and-using-dountil/
 # TODO: get this working
 <# if($delrows) {
   $i = 1
   do {
     if($wksht.Cells.Item($i, 1).Formula -eq "") {
       $rowRange = $wksht.Cells.Item($i, 1).EntireRow
       $rowRange.Delete() > $null
     }
     $i++
    } until ($wksht.Cells.Item($i, 1).Formula -ne "")
    
  }
 } #>

 # identify the save
 Write-Host “saving $filepath” 

 $workbook.SaveAs($filepath) 
 
 # close all workbooks and documents before beginning the loop again
 $objExcel.Workbooks.close() 
}

# clean up after loop by closing Excel 
$objExcel.Quit()    

# close word too
$objWord.Quit()

# do more extensive cleanup
# https://mcpmag.com/articles/2018/05/24/getting-started-word-using-powershell.aspx
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objWord)
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable objWord
Remove-Variable objExcel
