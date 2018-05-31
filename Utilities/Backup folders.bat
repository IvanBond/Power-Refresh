for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
set today=%MyDate:~0,4%-%MyDate:~4,2%-%MyDate:~6,2%

"<C:\Program Files\7-Zip\7z.exe>" a "<C:\Temp\%today% Power Refresh.zip>" "<C:\Power Refresh\*>"

"<C:\Program Files\7-Zip\7z.exe>" a "C:\Temp\%today% BYD Tasks.zip" "C:\Windows\System32\Tasks\*"