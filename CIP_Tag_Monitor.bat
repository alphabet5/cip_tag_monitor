@ECHO OFF

SET EXEName=CIP_Tag_Monitor.exe
SET EXEFullPath=C:\CIP_Tag_Monitor\CIP_Tag_Monitor.exe

TASKLIST | FINDSTR /I "%EXEName%"
IF ERRORLEVEL 1 GOTO :StartTagMonitor
GOTO EOF

:StartTagMonitor
START "" "%EXEFullPath%"
GOTO EOF