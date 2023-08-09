@echo OFF
setlocal enableDelayedExpansion

set VERSION=2.11
set BASE=
set END= 
set APPLICATION=Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v%VERSION%.xlsm

set APPLICATION_PATH=..\..\LGS_Application
set APPLICATION_FULL=%APPLICATION_PATH%\%APPLICATION%

set APPLICATION_README=readme.md
set APPLICATION_README_FULL=%APPLICATION_PATH%\%APPLICATION_README%

set HELPER_PATH=fichier exemple nom export FFG
set HELPER_FULL=%APPLICATION_PATH%\%HELPER_PATH%

set tours=T1 T2 T3 T4 T5 T6 Finale
set tours_plus=%tours% T7
set tech=Backup Poub

echo Creation des repertoires

set list=%tours_plus% %tech%
(for %%x in (%list%) do ( 
  mkdir %%x 2> nul
))

rem get current year
for /f %%i in ('dir /B ..\..') do set CURRENT_YEAR=%%i
echo Creation de l'application pour l'annee: %CURRENT_YEAR%

echo Installation de l'APPLICATION: %APPLICATION_FULL% --^>
(for %%a in (%tours%) do ( 
  copy /-Y "%APPLICATION_FULL%" "Calcul La Grande Semaine - STROKEPLAY - %%a - HOMME_OU_DAME_v%VERSION%.xlsm"
))

copy /-Y "%APPLICATION_README_FULL%" .

xcopy /S /F "%HELPER_FULL%" "%HELPER_PATH%\"
