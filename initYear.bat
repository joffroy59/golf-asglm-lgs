@echo off
setlocal enableDelayedExpansion
set INIT_SCRIPT=init.bat

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

		echo Copy du script d'initialisation de l'application %INIT_SCRIPT% dans !PATH_LGS!\
		copy /-Y "%INIT_SCRIPT%" "!PATH_LGS!\"
		echo Ex�cution du script d'initialisation de l'application %INIT_SCRIPT% dans !PATH_LGS!\
		cd "!PATH_LGS!\"
		call "%INIT_SCRIPT%"
		del "%INIT_SCRIPT%"
	)
) ELSE (
	echo %PATH_ASGLM_YEAR% n'existe pas.
	echo STOP
)


