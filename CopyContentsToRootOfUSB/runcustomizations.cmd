rem runcustomizations.cmd

set "LOG_DIR=C:\Recovery\OEM\logs"
for /f "tokens=2 delims==" %%d in ('wmic os get localdatetime /value') do set "LDT=%%d"
set "LOG_FILE=%LOG_DIR%\runcustomizations.cmd_%LDT:~0,4%-%LDT:~4,2%-%LDT:~6,2%_%LDT:~8,2%-%LDT:~10,2%-%LDT:~12,2%.log"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

:: Run and log everything
copy /y "C:\Recovery\OEM\unattend.xml" "C:\Windows\Panther\unattend.xml" >> "%LOG_FILE%" 2>&1
echo Copy completed at %DATE% %TIME% >> "%LOG_FILE%"

EXIT 0