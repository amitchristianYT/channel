#################################################################
#
# 2023-02-26
#
# https://www.youtube.com/amitchristian
# Please subscribe to my Youtube Channel for more tips and tricks
#
# Enable Powershell script execution using this command
#   set-executionpolicy remotesigned
#
# Make Sure that you set your folder path in first line
#
#################################################################
 $Folders = Get-ChildItem C:\Users\<Your-UserName>\desktop\test_doc -Directory

ForEach ($Folder in $Folders)
{
    $wdFormatPDF = 17
    $word = New-Object -ComObject word.application
    $word.visible = $false
    $folderpath = "$($Folder.FullName)\*"
    $fileTypes = "*.docx","*doc"
    Get-ChildItem -path $folderpath -include $fileTypes |
    foreach-object `
    {
     $path =  ($_.fullname).substring(0,($_.FullName).lastindexOf("."))
     "Converting $path to pdf ..."
     $doc = $word.documents.open($_.fullname)
     $doc.saveas([ref] $path, [ref]$wdFormatPDF)
     $doc.close()
    }
    $word.Quit()
}
