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
PS E:\MyScripts> $null = [System.Diagnostics.Process]::Start("$htmlFileName")                                           
PS E:\MyScripts>                                                                                                        
PS E:\MyScripts>                                                                                                        
PS E:\MyScripts> $rtfFileName = "$env:temp\test.rtf"                                                                    
PS E:\MyScripts> .\Get-ConsoleAsRTF | out-file $rtfFileName -encoding ascii                                             
PS E:\MyScripts> $null = [System.Diagnostics.Process]::Start("$rtfFileName")                                            
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
###########################################################################################################
# Get-ConsoleAsRtf.ps1
#
# The script captures console screen buffer up to the current cursor position and returns it in RTF format.
#
# Returns: ASCII-encoded string.
#
# Example:
#
# $rtfFileName = "$env:temp\ConsoleBuffer.rtf"
# .\Get-ConsoleAsRtf | out-file $rtfFileName -encoding ascii
# $null = [System.Diagnostics.Process]::Start("$rtfFileName")
#

# Check the host name and exit if the host is not the Windows PowerShell console host.
c:\Users\tcailleau\Documents\WindowsPowerShell\Scripts\FromPowerShellTeam\powershell\Get-ConsoleAsRtf.ps1
This script runs only in the console host. You cannot run this script in Visual Studio Code Host. 
#>

if ($host.Name -ne 'ConsoleHost')
{
  write-host -ForegroundColor Red "This script runs only in the console host. You cannot run this script in $($host.Name)."
  exit -1
}

# Maps console color name to RTF color index.
# The index of \cf is referencing the color definition in RTF color table.
#
function Get-RtfColorIndex ([string]$color)
{
  switch ($color)
  {
    'Black' { $index = 17 }
    'DarkBlue' { $index = 2 }
    'DarkGreen' { $index = 3 }
    'DarkCyan' { $index = 4 }
    'DarkRed' { $index = 5 }
    'DarkMagenta' { $index = 6 }
    'DarkYellow' { $index = 7 }
    'Gray' { $index = 8 }
    'DarkGray' { $index = 9 }
    'Blue' { $index = 10 }
    'Green' { $index = 11 }
    'Cyan' { $index = 12 }
    'Red' { $index = 13 }
    'Magenta' { $index = 14 }
    'Yellow' { $index = 15 }
    'White' { $index = 16 }
    default
    {
      $index = 0
    }
  }
  return $index
}

# Create RTF block from text using named console colors.
#
function Add-RtfBlock ($text)
{
  $foreColorIndex = Get-RtfColorIndex $currentForegroundColor
  $null = $rtfBuilder.Add("{\cf$foreColorIndex")

  # You can also add \ab* tag here if you want a bold font in the output.

  $backColorIndex = Get-RtfColorIndex $currentBackgroundColor
  $null = $rtfBuilder.Add("\chshdng0\chcbpat$backColorIndex")

  $text = $blockBuilder.ToString()
  $null = $rtfBuilder.Add(" $text}")
}

# Add line break to RTF builder
#
function Add-Break
{
  $backColorIndex = Get-RtfColorIndex $currentBackgroundColor
  $null = $rtfBuilder.Add("\shading0\cbpat$backColorIndex\par`r`n")
}

# Initialize the RTF string builder.
$rtfBuilder = new-object system.text.stringbuilder

# Set the desired font
$fontName = 'Lucida Console'
# Add RTF header
$null = $rtfBuilder.Add("{\rtf1\fbidis\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 $fontName;}}")
$null = $rtfBuilder.Add("`r`n")
# Add RTF color table which will contain all Powershell console colors.
$null = $rtfBuilder.Add('{\colortbl;red0\green0\blue128;\red0\green128\blue0;\red0\green128\blue128;\red128\green0\blue0;\red1\green36\blue86;\red238\green237\blue240;\red192\green192\blue192;\red128\green128\blue128;\red0\green0\blue255;\red0\green255\blue0;\red0\green255\blue255;\red255\green0\blue0;\red255\green0\blue255;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue0;}')
$null = $rtfBuilder.Add("`r`n")
# Add RTF document settings.
$null = $rtfBuilder.Add('\viewkind4\uc1\pard\ltrpar\f0\fs23 ')
 
# Grab the console screen buffer contents using the Host console API.
$bufferWidth = $host.ui.rawui.BufferSize.Width
$bufferHeight = $host.ui.rawui.CursorPosition.Y
$rec = new-object System.Management.Automation.Host.Rectangle 0,0,($bufferWidth - 1),$bufferHeight
$buffer = $host.ui.rawui.GetBufferContents($rec)

# Iterate through the lines in the console buffer.
for($i = 0; $i -lt $bufferHeight; $i++)
{
  $blockBuilder = new-object system.text.stringbuilder

  # Track the colors to identify spans of text with the same formatting.
  $currentForegroundColor = $buffer[$i, 0].Foregroundcolor
  $currentBackgroundColor = $buffer[$i, 0].Backgroundcolor

  for($j = 0; $j -lt $bufferWidth; $j++)
  {
    $cell = $buffer[$i,$j]

    # If the colors change, generate an RTF span and Add it to the RTF string builder.
    if (($cell.ForegroundColor -ne $currentForegroundColor) -or ($cell.BackgroundColor -ne $currentBackgroundColor))
    {
      Add-RtfBlock

      # Reset the block builder and colors.
      $blockBuilder = new-object system.text.stringbuilder
      $currentForegroundColor = $cell.Foregroundcolor
      $currentBackgroundColor = $cell.Backgroundcolor
    }

    # Substitute characters which have special meaning in RTF.
    switch ($cell.Character)
    {
      "`t" { $rtfChar = '\tab' }
      '\' { $rtfChar = '\\' }
      '{' { $rtfChar = '\{' }
      '}' { $rtfChar = '\}' }
      default
      {
        $rtfChar = $cell.Character
      }
    }

    $null = $blockBuilder.Add($rtfChar)
  }

  Add-RtfBlock
  Add-Break
}

# Add RTF ending brace.
$null = $rtfBuilder.Add('}')

return $rtfBuilder.ToString()