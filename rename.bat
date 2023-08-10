@echo OFF
setlocal enableDelayedExpansion

set inputs=DB.xls DN.xls MB.xls MN.xls

echo Renomage Fichier Brut et Net dans le bon format
rename DB.xls export_DAME_Brut_12.xlsx
rename DN.xls export_DAME_Net_12.xlsx
rename MB.xls export_HOMME_Brut_1.xlsx
rename MN.xls export_HOMME_Net_1.xlsx