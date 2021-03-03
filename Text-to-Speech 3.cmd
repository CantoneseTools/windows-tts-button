:: -------------------------------------------------------------
:: Batch Script - for performing Text-to-Speech (TTS) on whatever text is currently selected by a user.

:: - This is intended to be a Button, pinned to the Task Manager for easy access. 

:: - To change the Voice used: 
:: -- 0. (Note) Windows has 2 TTS systems: 'Control Panel - Speech Recognition' (older), and 'Settings App - Speech' (newer). This script is restricted to the older 'Control Panel - Speech Recognition' system.
:: -- 1. Go to 'Control Panel > Ease of Access > Speech Recognition > Text to Speech > Voice Selection.' 
:: -- 2. (Optional) More languages can be added to this dropdown through 'Settings App > Time & Language > Language > Preffered Languages.' (Ignore the 'Speech' Section in the 'Settings App', this script does not use that.)

:: - NOTE: This methods is limited to speaking out the first ~ 900 characters (~160 English words.)

:: -------------------------------------------------------------

@echo off
:: Change the Code Page to UTF-8 (To process CJK chars)
chcp 65001 
:: echo -----------------------------------------

:: Alt-Tab back to existing app, and Copy (the user's current Selected Text) to Clipboard
PowerShell -Command^
 "$wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('%%{TAB}'); Start-Sleep -m 1000;  $wshell.SendKeys('^c');"

:: echo -----------------------------------------

:: Process Clipboard
PowerShell -Command "Get-Clipboard" > ~speaktemp.txt
:: Convert to a single line, which can be read by Command Prompt
PowerShell -Command "(Get-Content '~speaktemp.txt') -join ' . ' | Set-Content '~speaktemp.txt'"

:: set /p myclipboard=<~speaktemp.txt 
set /p myclipboard=<~speaktemp.txt

:: v FAILS if there's a '|', e.g.  --- Translate | 翻譯
:: echo CLIPBOARD: %myclipboard%
:: v Works if there's a '|'
echo CLIPBOARD: "%myclipboard%"

del ~speaktemp.txt

:: echo -----------------------------------------

:: Escape certain characters (Brute-Force) :: Can try PowerShell replace commands as well, in the future - https://www.educba.com/powershell-string-replace/, etc.

:: Warning: This is a Very Unstable way of cleaning the user-selection-text. In the future, we should find something that's more stable.
:: (e.g. Right now, we're manually being mindful of special characters in the Command Prompt: '|', '&', ''', '"') 

@echo off
set myclipboard="%myclipboard%"
set myclipboard=%myclipboard:|=-%

set myclipboard=%myclipboard:'=%
set myclipboard=%myclipboard:‘=%
set myclipboard=%myclipboard:’=%
set myclipboard=%myclipboard:“=%
set myclipboard=%myclipboard:”=%

set myclipboard=%myclipboard:‛=%
set myclipboard=%myclipboard:‟=%
set myclipboard=%myclipboard:′=%
set myclipboard=%myclipboard:″=%

set myclipboard=%myclipboard:´=%
set myclipboard=%myclipboard:˝=%
set myclipboard=%myclipboard:˛=%
set myclipboard=%myclipboard:`=%
set myclipboard=%myclipboard:❛=%
set myclipboard=%myclipboard:❜=%
set myclipboard=%myclipboard:❝=%
set myclipboard=%myclipboard:❞=%

set myclipboard=%myclipboard:&=and%
set myclipboard="%myclipboard:"=%"

:: These lines have been commented, as they cause glitches.
rem set myclipboard=%myclipboard¨=%
:: .

echo PC Saying: "%myclipboard%"


:: echo -----------------------------------------

:: Speak out the clipboard text
PowerShell -Command "Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('%myclipboard%'); "

:: If there's an issue, uncomment the final 'pause' command to debug.
:: pause