@ECHO OFF
cd /d "%~dp0"
for /f "delims=" %%# in ('dir /b Ez*.cmd') do set "script=%%#"
call "%script%" /k38 /silent
attrib -R -A -S -H *.*
RMDIR /S /Q "%WINDIR%\Setup\Scripts"
exit
