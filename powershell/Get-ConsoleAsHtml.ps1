<#
	.NOTES
	===========================================================================
	 Created on:   	05_08_2021_22:03
	 Created by:   	CailleauThierry
	 Organization: 	Private
	 Filename: 		Get-ColorsFromTerminal.ps1 Posted by :https://devblogs.microsoft.com/powershell/colorized-capture-of-console-screen-in-html-and-rtf/
	===========================================================================
   .DESCRIPTION
   In the previous post we demonstrated how we can use console host API to capture screen buffer contents as text. But what if we want some colors. Would not it be nice to publish colorized copy of console in HTML or insert it as part of Microsoft Word document. For this to work we need to add some modifications to original script. Colors of each character are available as properties of System.Management.Automation.Host.BufferCell object:

PS E:\MyScripts> $bufferWidth = $host.ui.rawui.BufferSize.Width                                                         
PS E:\MyScripts> $bufferHeight = $host.ui.rawui.CursorPosition.Y                                                        
PS E:\MyScripts> $rec = new-object System.Management.Automation.Host.Rectangle 0,0,($bufferWidth - 1),$bufferHeight     
PS E:\MyScripts> $buffer = $host.ui.rawui.GetBufferContents($rec)                                                       
PS E:\MyScripts> $buffer[1,1]                                                                                           
                                                                                                                        
                    Character               ForegroundColor               BackgroundColor                BufferCellType 
                    ---------               ---------------               ---------------                -------------- 
                            o                    DarkYellow                   DarkMagenta                      Complete 
                                                                                                                        
                                                                                                                        
PS E:\MyScripts>                                                                                                        
All we need to do is to iterate through the screen buffer array, keeping track of the cell colors and generate HTML spans or RTF blocks with varying color attributes as soon as they change.

Why implementing both formats? While HTML is sufficient for Web applications, we will get bad results if we will try to use it in word processing publications. MS Word is much better at pasting and rendering RTF, rather than HTML. The e-mail editor in Microsoft Outlook also produce better results with RTF. By implementing console capture in both formats we will cover much bigger range of applications.

Usage of the scripts is fairly easy. The following example demonstrates how to put both scripts to a quick test:

Windows PowerShell V2 (Community Technology Preview - Features Subject to Change)                                       
Copyright (C) 2008 Microsoft Corporation. All rights reserved.                                                          
                                                                                                                        
PS C:\Users\Vladimir> cd E:\MyScripts                                                                                   
PS E:\MyScripts> $htmlFileName = "$env:temp\ConsoleBuffer.html"                                                         
PS E:\MyScripts> .\Get-ConsoleAsHtml | out-file $htmlFileName -encoding UTF8                                            
PS E:\MyScripts> $notnull = [System.Diagnostics.Process]::Start("$htmlFileName")                                           
PS E:\MyScripts>                                                                                                        
PS E:\MyScripts>                                                                                                        
PS E:\MyScripts> $rtfFileName = "$env:temp\test.rtf"                                                                    
PS E:\MyScripts> .\Get-ConsoleAsRTF | out-file $rtfFileName -encoding ascii                                             
PS E:\MyScripts> $notnull = [System.Diagnostics.Process]::Start("$rtfFileName")                                            
PS E:\MyScripts>                                                                                                        
Needless to say, the scripts can be further modified to include configurable parameters such as font name and size. The purpose here is to demonstrate the basic techniques for automated generation of simple HTML and RTF documents.

Hope you will find it useful,
Vladimir Averkin
Windows PowerShell Team 
   .EXAMPLE
   To conver all the .xls file present in the current folder and all its sub-folders (Note: this does not create a list of files that were converted)

   Get-ChildItem -Path ./  -Recurse -Include *.xls | Convert-Xls2Xlsx

   Which you could follow by removing the old .xls files

   Get-ChildItem -Path ./  -Recurse -Include *.xls | Remove-Item    

#>

<# 
############################################################################################################
# Get-ConsoleAsHtml.ps1
#
# The script captures console screen buffer up to the current cursor position and returns it in HTML format.
#
# Returns: UTF8-encoded string.
#
# Example:
#
# $htmlFileName = "$env:temp\ConsoleBuffer.html"
# .\Get-ConsoleAsHtml | out-file $htmlFileName -encoding UTF8
# $notnull = [System.Diagnostics.Process]::Start("$htmlFileName")
#
#>

# Check the host name and exit if the host is not the Windows PowerShell console host.
if ($host.Name -ne 'ConsoleHost')
{
  write-host -ForegroundColor Red "This script runs only in the console host. You cannot run this script in $($host.Name)."
  exit -1
}

# The Windows PowerShell console host redefines DarkYellow and DarkMagenta colors and uses them as defaults.
# The redefined colors do not correspond to the color names used in HTML, so they need to be mapped to digital color codes.
#
function Set-HtmlColor ($color)
{
  if ($color -eq "DarkYellow") { $color = "#eeedf0" }
  if ($color -eq "DarkMagenta") { $color = "#012456" }
  return $color
}

# Create an HTML span from text using the named console colors.
#
function Set-HtmlSpan ($text, $forecolor = "DarkYellow", $backcolor = "DarkMagenta")
{
  $forecolor = Set-HtmlColor $forecolor
  $backcolor = Set-HtmlColor $backcolor

  # You can also add font-weight:bold tag here if you want a bold font in output.
  return "<span style='font-family:Courier New;color:$forecolor;background:$backcolor'>$text</span>"
}

# Generate an HTML span and Add it to HTML string builder
#
function Add-HtmlSpan
{
  $spanText = $spanBuilder.ToString()
  $spanHtml = Set-HtmlSpan $spanText $currentForegroundColor $currentBackgroundColor
  $notnull = $htmlBuilder.Add($spanHtml)
}

# Add line break to HTML builder
#
function Add-HtmlBreak
{
  $notnull = $htmlBuilder.Add("<br>")
}

# Initialize the HTML string builder.
$htmlBuilder = new-object system.text.stringbuilder
$notnull = $htmlBuilder.Add("<pre style='MARGIN: 0in 10pt 0in;line-height:normal';font-size:10pt>")

# Grab the console screen buffer contents using the Host console API.
$bufferWidth = $host.ui.rawui.BufferSize.Width
$bufferHeight = $host.ui.rawui.CursorPosition.Y
$rec = new-object System.Management.Automation.Host.Rectangle 0,0,($bufferWidth - 1),$bufferHeight
$buffer = $host.ui.rawui.GetBufferContents($rec)

# Iterate through the lines in the console buffer.
for($i = 0; $i -lt $bufferHeight; $i++)
{
  $spanBuilder = new-object system.text.stringbuilder

  # Track the colors to identify spans of text with the same formatting.
  $currentForegroundColor = $buffer[$i, 0].Foregroundcolor
  $currentBackgroundColor = $buffer[$i, 0].Backgroundcolor

  for($j = 0; $j -lt $bufferWidth; $j++)
  {
    $cell = $buffer[$i,$j]

    # If the colors change, generate an HTML span and Add it to the HTML string builder.
    if (($cell.ForegroundColor -ne $currentForegroundColor) -or ($cell.BackgroundColor -ne $currentBackgroundColor))
    {
      Add-HtmlSpan

      # Reset the span builder and colors.
      $spanBuilder = new-object system.text.stringbuilder
      $currentForegroundColor = $cell.Foregroundcolor
      $currentBackgroundColor = $cell.Backgroundcolor
    }

    # Substitute characters which have special meaning in HTML.
    switch ($cell.Character)
    {
      '>' { $htmlChar = '&gt;' }
      '<' { $htmlChar = '&lt;' }
      '&' { $htmlChar = '&amp;' }
      default
      {
        $htmlChar = $cell.Character
      }
    }

    $notnull = $spanBuilder.Add($htmlChar)
  }

  Add-HtmlSpan
  Add-HtmlBreak
}

# Add HTML ending tag.
$notnull = $htmlBuilder.Add("</pre>")

return $htmlBuilder.ToString()