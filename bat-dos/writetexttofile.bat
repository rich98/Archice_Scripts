echo off

REM NUM lock state 

echo set WshShell = CreateObject("WScript.Shell")>"C:\Documents and Settings\%username%\Start Menu\Programs\Startup\numlock.vbs"

echo WshShell.SendKeys "{NUMLOCK}">>"C:\Documents and Settings\%username%\Start Menu\Programs\Startup\numlock.vbs"

