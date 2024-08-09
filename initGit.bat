REM isntall git hook pour vba code

@echo off
echo Installation git Hooks en cours...
xcopy  /s /i /Y ".hooks\*" ".git\hooks\" >nul
echo Copie terminée. Installation git Hooks DONE.
pause