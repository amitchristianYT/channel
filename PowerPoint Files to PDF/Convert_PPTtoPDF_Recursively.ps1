#################################################################
#
# 2023-02-26
#
# https://www.youtube.com/amitchristian
# Please subscribe to my Youtube Channel for more tips and tricks
#
# Enable Powershell script execution using this command
#   set-executionpolicy remotesigned

# call: .\Convert-PPTtoPDF_recursively.ps1 -PathToPPTFiles <path_to_ppt_files>
# 
#
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
