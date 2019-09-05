@echo off
:Start
set time=9

:loop
ping localhost -n 2 >nul
set /a time=%time%-1
if %time% EQU 0 goto Exit
goto loop

:Exit
@echo off
cd /
Taskkill /f /t /im excel.exe