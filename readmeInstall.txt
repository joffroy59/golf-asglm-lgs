################################
# automatique
################################
1. Aller dans le répertoire LGS_Application
2. Lancer initYear.bat


Note
	initYear.bat :
		1. création du répertoire ASGLM <YEAR>/LGS
		2. copie du script init.bat dans le répertoire créé
		3. lancement du fichier init.bat
	init.bat     : fichier d'initialisation du répertoire courant


################################
# a la main TOFINISH
################################
Initialisation de LGS
	<TODO>
	avoir une arborescence: 
		<ASGLM annee>/LGS/T1
		<ASGLM annee>/LGS/T2
		<ASGLM annee>/LGS/T3
		<ASGLM annee>/LGS/T4
		<ASGLM annee>/LGS/T5
		<ASGLM annee>/LGS/T6
		<ASGLM annee>/LGS/T7

	avoir le fichier de calcul vierge dans : 
		<ASGLM annee>/LGS/Calcul La Grande Semaine - STROKEPLAY - Final - HOMME_OU_DAME_v2.10 - ALL.xlsm

Procedure d'intégration des score de la journee

	produire les exports suivants avec RMS dans le répertoire <ASGLM annee>/LGS/:
		export_DAME_Brut_12.xlsx
		export_DAME_Brut_34.xlsx
		export_DAME_Net_12.xlsx
		export_DAME_Net_34.xlsx

		export_HOMME_Brut_1.xlsx
		export_HOMME_Brut_234.xlsx
		export_HOMME_Net_1.xlsx
		export_HOMME_Net_234.xlsx

	les fichiers doivent se trouver dans le repertoire Tx correspondant au Tour de LGS à enregistrer 