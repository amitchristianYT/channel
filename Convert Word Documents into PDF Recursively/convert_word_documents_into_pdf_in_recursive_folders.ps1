#################################################################
#
# 2023-02-26
#
# Watch the full tutorial demo of how to run this script here.
# https://www.youtube.com/watch?v=C6hzcyCZx4E
# Please subscribe to my Youtube Channel for more tips and tricks
#
# Enable Powershell script execution using this command
#   set-executionpolicy remotesigned
#
# STEPS
# 1. Download this script by copying and pasting or by clicking 'Raw' button on
#    top right and then right click and save as <Anyname>.ps1.
#    (IMP: Make sure that the extension of the saved file is in fact, ps1 and 
#    not .ps1.txt, which essentially is a text file)
# 2. Copy the ps1 file, to the root folder location where all your pptx files
#    are located.
# 3. Enable powersheel script execution using following command.
#      set-executionpolicy remotesigned
# 4. Make Sure that you set your folder path in first line of the script below.
# 5. Call the ps1 function from powershell, as shown in the video.
# 
# TROUBLESHOOTING
# - Please add a comment if you are getting errors running this script here.
#   https://www.youtube.com/watch?v=C6hzcyCZx4E
#
# Thank you.
# Cheers.
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
