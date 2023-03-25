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
# Steps
# 1. Download this script by copying and pasting or by clicking 'Raw' button on
#    top right and then right click and save as <Anyname>.ps1.
#    (IMP: Make sure that the extension of the saved file is in fact, ps1 and 
#    not .ps1.txt, which essentially is a text file)
# 2. Copy the ps1 file, to the root folder location where all your pptx files
#    are located.
# 3. Enable powersheel script execution using following command.
#      set-executionpolicy remotesigned
# 4. Call the ps1 function as below. 
#      call: .\Convert-PPTtoPDF_recursively.ps1 -PathToPPTFiles <path_to_ppt_files>
# 
# TROUBLESHOOTING
# - Please add a comment if you are getting errors running this script here.
#   https://www.youtube.com/watch?v=tAz2CDkTY2U
#
# Thank you.
# Cheers.
#
#################################################################


param (
    [parameter(Mandatory=$true)]
    [string]$PathToPPTFiles
)

$pptFiles = Get-ChildItem -Path $PathToPPTFiles -Filter *.ppt* -Recurse

foreach ($pptFile in $pptFiles) {
    $pdfFile = $pptFile.FullName -replace ".ppt", ".pdf"
    $ppt = New-Object -ComObject PowerPoint.Application
    # $ppt.Visible = $false
    $presentation = $ppt.Presentations.Open($pptFile.FullName)
    $presentation.SaveAs($pdfFile, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
    $presentation.Close()
    $ppt.Quit()
}
