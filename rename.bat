@echo OFF

(set \n=^^^
%=empty, do not delete this line =%
^
%=empty, do not delete this line =%
)
setlocal enabledelayedexpansion


echo Début du processus de renommage des fichiers bruts et nets...

rem Tableaux simulés de fichiers sources et noms de fichiers de destination
set "fileMappings[1]=DB.xls export_DAME_Brut_12.xlsx"
set "fileMappings[2]=DN.xls export_DAME_Net_12.xlsx"
set "fileMappings[3]=MB.xls export_HOMME_Brut_1.xlsx"
set "fileMappings[4]=MN.xls export_HOMME_Net_1.xlsx"

rem Variables pour le rapport
setlocal enabledelayedexpansion
set "successCount=0"

for %%i in (1 2 3 4) do (
    for /f "tokens=1,2" %%a in ("!fileMappings[%%i]!") do (
        rename "%%a" "%%b" > nul
        if !errorlevel! equ 0 (
            set /a "successCount+=1"
            set report=!report!%%a en %%b %\n%
        )
    )
)

if %successCount% gtr 0 (
    echo Le processus de renommage est terminé. Les fichiers ont été renommés avec succès :
    echo !report!
) else (
    echo Aucun fichier n'a été renommé. Vérifiez les noms de fichiers sources.
)
endlocal