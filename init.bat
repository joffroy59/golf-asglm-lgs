@echo OFF
setlocal enableDelayedExpansion

set VERSION=2.11
set APPLICATION=Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v%VERSION%.xlsm

set APPLICATION_PATH=..\..\LGS_Application
set APPLICATION_FULL=%APPLICATION_PATH%\%APPLICATION%

set APPLICATION_README=readme.txt
set APPLICATION_README_FULL=%APPLICATION_PATH%\%APPLICATION_README%

set HELPER_PATH=fichier exemple nom export FFG
set HELPER_FULL=%APPLICATION_PATH%\%HELPER_PATH%

echo Cr�ation des r�pertoires
set list=T1 T2 T3 T4 T5 T6 T7 Finale Backup Poub 
(for %%x in (%list%) do ( 
  mkdir %%x 2> nul
))

rem get current year
for /f %%i in ('dir /B ..\..') do set CURRENT_YEAR=%%i
echo C�ation de l'application pour l'ann�e: %CURRENT_YEAR%

echo Installation de l'APPLICATION: %APPLICATION_FULL% --^>
set list=T1 T2 T3 T4 T5 T6 Finale 
(for %%a in (%list%) do ( 
  copy /-Y "%APPLICATION_FULL%" "Calcul La Grande Semaine - STROKEPLAY - %%a - HOMME_OU_DAME_v%VERSION%.xlsm"
))

copy /-Y "%APPLICATION_README_FULL%" .

xcopy /S /F "%HELPER_FULL%" "%HELPER_PATH%\"
