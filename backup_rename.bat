@echo OFF
echo Exécution du script de sauvegarde...
call backup.bat
echo Script de sauvegarde terminé.

echo.

echo Exécution du script de renommage...
call rename.bat
echo Script de renommage terminé.

echo.

echo Les deux scripts ont été exécutés avec succès.
pause
