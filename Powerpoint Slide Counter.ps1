# PowerPoint File and Slide Counter Script
# Author: Peter Yates (pyates@gmail.com)
# Date: 2018-05-18

# Load modules
Import-Module Microsoft.PowerShell.Management
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Office
Add-Type -AssemblyName Microsoft.Office.Interop.Powerpoint

# Recursively identify all PPT and PPTX files
$strDir = "YOUR_DIRECTORY_HERE!!!"
$colFiles = Get-ChildItem $strDir -Recurse -File -Include *.ppt*| Select-Object FullName

# Create a powerpoint instance
$Application = New-Object -ComObject PowerPoint.Application

# Cycle through file list and count number of files and total number of slides
$intPages = 0
$arraySlideCount = @()
foreach($objFile in $colFiles)
{
    #$objDoc = $Application.presentations.Open($colPath + $objFile.name, 2, $false, $false)
    $objDoc = $Application.presentations.Open($objFile.FullName, 2, $false, $false)
    $arraySlideCount += $objDoc.slides.count
    $intPages = $intPages + $objDoc.slides.count
    $application.ActivePresentation.Close
}
$Application.Quit()

# Sort array so largest is in 0 location
$arraySlideCount = $arraySlideCount | Sort-Object -Descending

# Print results
"Results for Directory " + $strDir
"Total Presentations: " + $colFiles.Length
"Total Slides: " + $intPages
"Avg. Slides per Presentation: " + $intPages/$colFiles.Length
"Longest Deck: " + $arraySlideCount[0] + " slides"