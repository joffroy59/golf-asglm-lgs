# Application de gestion de "La Grande Semaine"

une application Excel pour gérer le calcul des score de LGS

## Installation pour une nouvelle Année

Lancer

    initYear.bat

## 🔥 Procedure d'intégration des score de la journée 🔥

ℹ️ le type d'export `2d. Extraction XLS globale.xlsx` doit contenir toutes les Séries pour Homme et Dame ensemble

1. **Produire le fichier d'export** (typez d'export `2d. Extraction XLS globale.xlsx`) suivants avec RMS dans le répertoire
   *[ASGLM annee]/LGS/T<numéro du tour>*

         2d. Extraction XLS globale.xls

2. le fichier doit se trouver dans le repertoire Tx (correspondant au Tour de LGS à enregistrer)
   dans le répertoire LGS/
3. Intégrer les résultat dans la feuille du jour

    par exemple

   - pour le 1er Tour:

         Calcul La Grande Semaine - STROKEPLAY - T1 - HOMME_OU_DAME_v2.xx.xlsm
   - pour la Finale

         Calcul La Grande Semaine - STROKEPLAY - Finale - HOMME_OU_DAME_v2.xx.xlsm

4. lancer l'application du jour
5. **Integrer 1 Tour**

   1. *Effacer All* (**Intégreation Résultats du premier jour pour partir d'une feuille vierge**)
   2. Tour à importer
      - ⚠️ choisir le tour à importer dans la boite de selection
   3. Nettoyage des imports précédents
      - Cliquer "Effacer l'import (forced)" ou "Effacer l'import"
   4. Importation
      - Cliquer "Importer Tour"
      - Choisir le fichier dans le Repertoire `Tn/Extraction XLS globale.xls`
      - Attendre l'importation se finisse (le curseur doit revenir en haut de la feuille en cours)
   5. Vérification en regardant dans les 2 onglets de résultats hommes et dames

## 🔥🔥 Mode tous les tours présent dans les repertoires 🔥🔥

...
5. ⚠️ **Integrer Tous Les Tours**

   1. *Effacer All* (**Intégreation Résultats du premier jour pour partir d'une feuille vierge**)
   2. *Nettoyage des imports précédents*
      - Cliquer "Effacer l'import (forced)" ou "Effacer l'import"
   3. *Importation*
      - Cliquer "Importer tous les tours"
      - Attendre l'importation se finisse (le curseur doit revenir en haut de la feuille en cours)
   4. *Vérification* en regardant dans les 2 onglets de résultats hommes et dames

----------

### Note

**init.bat**:  fichier d'initialisation du répertoire courant

**initYear.bat**: **creation pour une nouvelle année**

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

        {ASGLM annee}/LGS/Calcul La Grande Semaine - STROKEPLAY - Final - HOMME_OU_DAME_v2.xx - ALL.xlsm
