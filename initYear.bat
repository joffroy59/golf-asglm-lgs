@echo off
setlocal enableDelayedExpansion
set SCRIPT_INIT=init.bat
set SCRIPT_BACKUP=backup.bat

set /p "YEAR=Année: "

set "PATH_ASGLM_YEAR=..\ASGLM ^!YEAR^!"
for /f "delims=" %%A in ("!PATH_ASGLM_YEAR!") do set "PATH_ASGLM_YEAR=%%A"

echo V�rification de l'existence du r�pertoire %PATH_ASGLM_YEAR%


IF EXIST "%PATH_ASGLM_YEAR%" (
	echo %PATH_ASGLM_YEAR% exists.
	set "PATH_LGS=%PATH_ASGLM_YEAR%\LGS"
	echo V�rification de l'existence du r�pertoire !PATH_LGS!
	IF EXIST "!PATH_LGS!" (
		echo !PATH_LGS! exists.
		echo STOP
	) ELSE (
		echo Cr�ation du r�pertoire !PATH_LGS!
		mkdir "!PATH_LGS!" 2> nul

		echo Copy du script d'initialisation de l'application %SCRIPT_INIT% dans !PATH_LGS! et script de renomage des exports\
		copy /-Y "%SCRIPT_INIT%" "!PATH_LGS!\"
		copy /-Y "%SCRIPT_BACKUP%" "!PATH_LGS!\"
		dir
		dir T1
		echo Ex�cution du script d'initialisation de l'application %SCRIPT_INIT% dans !PATH_LGS!\
		cd "!PATH_LGS!\"
		call "%SCRIPT_INIT%"
		del "%SCRIPT_INIT%"
		del "%SCRIPT_BACKUP%"
	)
) ELSE (
	echo %PATH_ASGLM_YEAR% n'existe pas.
	echo STOP
)


