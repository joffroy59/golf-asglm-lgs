@echo OFF
setlocal enableDelayedExpansion

set APPLICATION=Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_1=Calcul La Grande Semaine - STROKEPLAY - T1 - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_2=Calcul La Grande Semaine - STROKEPLAY - T2 - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_3=Calcul La Grande Semaine - STROKEPLAY - T3 - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_4=Calcul La Grande Semaine - STROKEPLAY - T4 - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_5=Calcul La Grande Semaine - STROKEPLAY - T5 - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_6=Calcul La Grande Semaine - STROKEPLAY - T6 - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_FINAL=Calcul La Grande Semaine - STROKEPLAY - Finale - HOMME_OU_DAME_v2.9.xlsm
set APPLICATION_README=readme.txt
set APPLICATION_PATH=..\..\LGS_Application
set APPLICATION_FULL=%APPLICATION_PATH%\%APPLICATION%
set APPLICATION_README_FULL=%APPLICATION_PATH%\%APPLICATION_README%
set HELPER_PATH=fichier exemple nom export FFG
set HELPER_FULL=%APPLICATION_PATH%\%HELPER_PATH%

echo Cr‚ation des r‚pertoires
for /l %%x in (1, 1, 7) do mkdir T%%x 2> nul
mkdir Backup 2> nul
mkdir Poub 2> nul

rem get current year
for /f %%i in ('dir /B ..\..') do set CURRENT_YEAR=%%i
echo C‚ation de l'application pour l'ann‚e: %CURRENT_YEAR%

echo Installation de l'APPLICATION: %APPLICATION_FULL% --^>
copy /-Y "%APPLICATION_FULL%" "%APPLICATION_1%"
copy /-Y "%APPLICATION_FULL%" "%APPLICATION_2%"
copy /-Y "%APPLICATION_FULL%" "%APPLICATION_3%"
copy /-Y "%APPLICATION_FULL%" "%APPLICATION_4%"
copy /-Y "%APPLICATION_FULL%" "%APPLICATION_5%"
copy /-Y "%APPLICATION_FULL%" "%APPLICATION_6%"
copy /-Y "%APPLICATION_FULL%" "%APPLICATION_FINAL%"

copy /-Y "%APPLICATION_README_FULL%" .

xcopy /S /F "%HELPER_FULL%" "%HELPER_PATH%\"
