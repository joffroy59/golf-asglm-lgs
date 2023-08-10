# Application de gestion de "La Grande Semaine"

une application Excel pour gérer le calcul des score de LGS

## Installation pour une nouvelle Année

Lancer

    initYear.bat

## 🔥 Procedure d'intégration des score de la journée 🔥  TOFINISH

1. Produire les exports suivants avec RMS dans le répertoire {ASGLM annee}/LGS/T<numéro du tour>:

        export_DAME_Brut_12.xlsx
        export_DAME_Brut_34.xlsx
        export_DAME_Net_12.xlsx
        export_DAME_Net_34.xlsx

        export_HOMME_Brut_1.xlsx
        export_HOMME_Brut_234.xlsx
        export_HOMME_Net_1.xlsx
        export_HOMME_Net_234.xlsx

2. les fichiers doivent se trouver dans le repertoire Tx (correspondant au Tour de LGS à enregistrer)
3. Intégrer les résultat dans la feuille du jour

    par exemple

   - pour le 1er Tour:

         Calcul La Grande Semaine - STROKEPLAY - T1 - HOMME_OU_DAME_v2.11.xlsm
   - pour la Finale

         Calcul La Grande Semaine - STROKEPLAY - Finale - HOMME_OU_DAME_v2.11.xlsm

4. lancer l'application du jour
5. INTEGRER TOUT

    TODO verification, sinon faire a la main en partant d'une copie du jour précédent
6. Effacer All
7. Dame
   - choisir "Dame" dans la boite de selection
   - Cliquer "Importer tous les tours"
8. Nettoyage des imports précédents
   - Cliquer "Effacer l'import (forced)"
9. Homme
   - choisir "Homme" dans la boite de selection
   - Cliquer "Importer tous les tours"

----------

### Note

**init.bat**:  fichier d'initialisation du répertoire courant
**initYear.bat**: creation pour une nouvelle année

1. création du répertoire ASGLM {YEAR}/LGS
2. copie du script init.bat dans le répertoire créé
3. lancement du fichier init.bat

## Installation Manuel

### Initialisation de LGS

1. avoir une arborescence:

        {ASGLM annee}/LGS/T1
        {ASGLM annee}/LGS/T2
        {ASGLM annee}/LGS/T3
        {ASGLM annee}/LGS/T4
        {ASGLM annee}/LGS/T5
        {ASGLM annee}/LGS/T6
        {ASGLM annee}/LGS/T7

2. Feuille de calcul vierge :

        {ASGLM annee}/LGS/Calcul La Grande Semaine - STROKEPLAY - Final - HOMME_OU_DAME_v2.11 - ALL.xlsm
