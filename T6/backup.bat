@echo off
setlocal enableDelayedExpansion

REM Obtient la date au format AAAA-MM-JJ
for /f "tokens=1-3 delims=/-" %%a in ('date /t') do (
    set "year=%%c"
    set "month=%%a"
    set "day=%%b"
)

REM Crée le nom du répertoire
set "backup_folder=backup_%year%-%month%-%day%"

REM Vérifie si le répertoire existe, sinon le crée
if not exist "%backup_folder%" (
    REM Crée le répertoire s'il n'existe pas
    mkdir "%backup_folder%"
    echo Répertoire de sauvegarde créé : %backup_folder%
) else (
    echo Le répertoire de sauvegarde existe déjà : %backup_folder%
)

REM Liste des noms de fichiers à déplacer
set "filesToMove=export_DAME_Brut_12.xlsx export_DAME_Net_12.xlsx export_HOMME_Brut_1.xlsx export_HOMME_Net_1.xlsx"

REM Variable pour suivre si au moins un fichier a été déplacé
set "filesMoved=false"

REM Variable pour stocker la liste des fichiers déplacés
set "movedFilesList="

REM Boucle pour déplacer les fichiers existants vers le répertoire de sauvegarde
for %%f in (%filesToMove%) do (
    if exist "%%f" (
        move "%%f" "%backup_folder%\%%f" > nul
        set "filesMoved=true"
        set "movedFilesList=!movedFilesList! %%f"
    )
)

REM Affiche un message si aucun fichier n'a été déplacé
if "%filesMoved%"=="false" (
    echo Aucun fichier existant à déplacer.
) else (
    echo Fichiers déplacés dans le répertoire de sauvegarde : %backup_folder%
    echo Liste des fichiers déplacés :
    for %%a in (!movedFilesList!) do (
        echo %%~a
    )
)

endlocal
