
rem @echo off

rem if exist C:\Users\%username%\Desktop\dirs.txt erase C:\Users\%username%\Desktop\dirs.txt
rem for /d %%i in ("C:\Users\%username%\AppData\Local\Temp\VertiPa*.*" ) do echo %%i >> C:\Users\%username%\Desktop\dirs.txt

for /d %%i in ("C:\Users\%username%\AppData\Local\Temp\VertiPa*.*") do rd /s /q "%%i"

del /q "C:\Users\%username%\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.mso\*.*"