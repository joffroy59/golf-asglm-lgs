Sub EffacementImport()

Dim ColDebutTableau As Integer
Dim ColFinTableau As Integer
Dim LigDebutTableau As Integer
Dim NbLignes As Integer

    LigDebutTableau = Range("DebutTableauGeneralNet").Row + 1
    NbLignes = Range("NbLignesNet")
    If NbLignes = 0 Then
        NbLignes = Range("NbLignesBrut")
    End If
    ColDebutTableau = Range("DebutTableauGeneralNet").Column
    ColFinTableau = Range("ScratchBrut").Column
    Range(Cells(LigDebutTableau, ColDebutTableau), Cells(LigDebutTableau + NbLignes, ColFinTableau)).Select
    Selection.Clear
        
    Call recordToHistory("EffacementImport")
    
    RowDebutTableau = Range("DebutTableauGeneralNet").Row
    Cells(RowDebutTableau + 1, ColDebutTableau).Select

End Sub

Sub EffacementImportForced()

Dim ColDebutTableau As Integer
Dim ColFinTableau As Integer
Dim LigDebutTableau As Integer
Dim NbLignes As Integer

    LigDebutTableau = Range("DebutTableauGeneralNet").Row + 1
    NbLignes = 1000
    ColDebutTableau = Range("DebutTableauGeneralNet").Column
    ColFinTableau = Range("ScratchBrut").Column
    Range(Cells(LigDebutTableau, ColDebutTableau), Cells(LigDebutTableau + NbLignes, ColFinTableau)).Select
    Selection.Clear
    
    Call recordToHistory("EffacementImportForced")
    
    RowDebutTableau = Range("DebutTableauGeneralNet").Row
    Cells(RowDebutTableau + 1, ColDebutTableau).Select

End Sub
Sub EffacementResultatAll()
'TODO use select tableau and loop
EffacementResultat ("Resultat LGS (HOMME)")
EffacementResultat ("Resultat LGS (DAME)")
End Sub

Sub EffacementResultat(Optional playerTypeSheetName As String)

Dim ColDebutTableau As Integer
Dim ColFinTableau As Integer
Dim ColFinTableauFormula As Integer
Dim LigDebutTableau As Integer
Dim NbLignes As Integer
    'TODO reafctoring
    colEndClear = "AC1"
    colEndClearFormula = "AI1"
    
    Dim NomFeuilleCumulJoueur As String
    
    If (playerTypeSheetName = "") Then
        NomFeuilleCumulJoueur = Range("NomFeuilleCumuljoueur")
    Else
        NomFeuilleCumulJoueur = playerTypeSheetName
    End If

    Worksheets(NomFeuilleCumulJoueur).AutoFilter.ShowAllData
    LigDebutInsertion = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultat").Row + 1
    ''NbLignes = Range("NbLignesNet")
    NbLignes = 1000
    ColDebutTableau = Range("DebutTableauGeneralNet").Column
    'ColFinTableau = Range("ScratchBrut").Column
    'Worksheets("Notice d'utilisation").Activate
    OriginSheet = ActiveSheet.Name
    Worksheets(NomFeuilleCumulJoueur).Activate
    ColFinTableau = Range(colEndClear).Column
    ColFinTableauFormula = Range(colEndClearFormula).Column
        
    Range(Cells(LigDebutInsertion, ColDebutTableau), Cells(LigDebutInsertion + NbLignes, ColFinTableau)).Select
    Selection.ClearContents
    
    'Range(Cells(LigDebutInsertion, ColFinTableau + 1), Cells(LigDebutInsertion + NbLignes, ColFinTableauFormula)).Select
    'Selection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone

    Range(Cells(LigDebutInsertion, ColFinTableau + 1 + 2), Cells(LigDebutInsertion + NbLignes, ColFinTableauFormula)).Select
    Selection.ClearContents

    'Range(Cells(LigDebutInsertion, ColFinTableauFormula + 1), Cells(LigDebutInsertion + NbLignes, ColFinTableauFormula + 1 + 1)).Select
    'Selection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    
    Call recordToHistory("EffacementResultat", , NomFeuilleCumulJoueur)
    
    Cells(4, 2).Select
    
    Worksheets(OriginSheet).Activate
    'Cells(1, 1).Select
End Sub

Sub EffacementAll()
    Call recordToHistory("EffacementAll")
    
    EffacementImport
    EffacementResultatAll
    ' TODO refactor
    Worksheets("Import Resultats Tour").Activate
End Sub

Public Sub ImporterCleanBrutNet(Optional TaskType As String)
    Call AskRoundSelected
    
    Call ImporterBrutNet(TaskType, "", , True)
End Sub

Sub AskRoundSelected()
    tour = Range("TourSelected")
    playerType = Range("playerType")
    If (MsgBox("Avez-vous choisi le Tour et le Type de joueur ? " & vbCrLf & "Tour :" & tour & vbCrLf & "Type de joueur : " & playerType, vbYesNo) = vbNo) Then
        End
    End If
End Sub

Sub RetraitementFeuilleMatchFFGolf()
    Call AskRoundSelected
    
    Call RetraitementFeuilleMatchFFGolfFichier("", "Importation d'un fichier simple FFGolf pour 1 Tour (sans génération du tour)")
End Sub

Function RetraitementFeuilleMatchFFGolfFichier(NomFichierTour As String, TaskType As String, Optional TourImporte As Integer, Optional ByVal ScoreType As String = "Auto")
    Dim ColRang As Integer
    Dim ColNom As Integer
    Dim ColClub As Integer
    Dim ColIndex As Integer
    Dim ColScore As Integer
    Dim LigneRang As Integer
    Dim I As Integer
    Dim DerniereLigne As Integer
    Dim J As Integer
    Dim TableauJoueurs(100, 6) As Variant
    Dim tour As Integer
    Dim PremiereLigne As Integer
    Dim PremiereColonne As Integer
    Dim LigneTableau As Integer
    Dim Temp As String
    Dim DateRencontre As String
    Dim Net As Boolean
    Dim NomJoueur As String
    Dim Erreur As Boolean
    Dim MessageErreur As String
    Dim NomFeuilleCumulJoueur  As String

    '---------------------
    'Lecture des variables
    '---------------------
    If (IsMissing(TourImporte) Or TourImporte = 0) Then
        TourImporte = Right(Application.Range("Tour"), 1) + 1
    End If
    
    Net = False
    Erreur = False
    
    NomFeuilleCumulJoueur = Range("NomFeuilleCumuljoueur")
    
    If InStr(NomFichierTour, ":") = 0 Then
        '------------------------------------------
        'Ouverture du fichier de resultat d'un Tour
        '------------------------------------------
        NomFichierTour = Application.GetOpenFilename(Title:="Import du resultat d'un tour" & " " & ScoreType)
    End If
    Debug.Print NomFichierTour
    
    If InStr(NomFichierTour, ":") = 0 Then
        Erreur = True
        MessageErreur = "Vous n'avez pas sélectionné de feuille de résultat à importer. Fin de la procédure"
    End If
    If Erreur Then
        I = MsgBox(MessageErreur, vbOKOnly, "Import des résultats d'un tour")
        End
    Else
        Call recordToHistory(TaskType & " - Tour " & TourImporte, NomFichierTour, NomFeuilleCumulJoueur)
        
        Workbooks.Open (NomFichierTour)
        '-----------------------
        'Retour en haut ö gauche
        '-----------------------
        Range("A1").Select
        
        '------------------------
        'Recherche du type de feuille
        '------------------------
        NouvelleFeuilleScore = False
        Range("A1").Select
        On Error GoTo TypeFeuille
        Cells.Find(What:="Rang", After:=ActiveCell, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Activate
        If NouvelleFeuilleScore Then
            '---------------------------------
            'Recherche du type de score brut ou net
            '---------------------------------
            Net = True
            Range("A1").Select
            On Error GoTo TypeScore
            Cells.Find(What:=" - Net", After:=ActiveCell, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Activate
            '-------------------------
            'Recherche de la colonne Rang
            '-------------------------
            Range("A1").Select
            Cells.Find(What:="Pos.", After:=ActiveCell, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Activate
            ColRang = ActiveCell.Column
            '------------------------------------
            'Recherche de la premire ligne des r_sultats
            '------------------------------------
            LigneRang = ActiveCell.Row
            '-------------------------
            'Recherche de la colonne Nom
            '-------------------------
            ActiveCell.Offset(0, 2).Select
            ColNom = ActiveCell.Column
            '----------------------------
            'Recherche de la colonne Club
            '----------------------------
            ActiveCell.Offset(0, 1).Select
            ColClub = ActiveCell.Column
            '-------------------------------------------------------------
            'Recherche de la colonne Index du joueur pour classement s_rie
            '-------------------------------------------------------------
            ActiveCell.Offset(0, 1).Select
            ColIndex = ActiveCell.Column
            '-----------------------------
            'Recherche de la colonne Score Net
            '-----------------------------
            ActiveCell.Offset(0, 2).Select
            ColScore = ActiveCell.Column
        Else
             '----------------------------
            'Recherche de la colonne Rang
            '----------------------------
            Range("A1").Select
            Cells.Find(What:="Rang", After:=ActiveCell, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Activate
            ColRang = ActiveCell.Column
            '--------------------------------------------
            'Recherche de la premire ligne des r_sultatss
            '--------------------------------------------
            LigneRang = ActiveCell.Row
            '---------------------------
            'Recherche de la colonne Nom
            '---------------------------
            ActiveCell.Offset(0, 1).Select
            ColNom = ActiveCell.Column
            '----------------------------
            'Recherche de la colonne Club
            '----------------------------
            ActiveCell.Offset(0, 1).Select
            ColClub = ActiveCell.Column
            '-------------------------------------------------------------
            'Recherche de la colonne Index du joueur pour classement s_rie
            '-------------------------------------------------------------
            ActiveCell.Offset(0, 2).Select
            ColIndex = ActiveCell.Column
            '-----------------------------
            'Recherche de la colonne Score Net
            '-----------------------------
            ActiveCell.Offset(0, 1).Select
            ColScore = ActiveCell.Column
            ActiveCell.Offset(0, 1).Select
            ColScore = ActiveCell.Column
            If ActiveCell.Value = "Net" Then
                Net = True
            Else
                Net = False
            End If
        End If
        '--------------------------------
        'Recherche de la derni�re colonne
        '--------------------------------
        ActiveCell.SpecialCells(xlLastCell).Select
        DerniereLigne = ActiveCell.Row
        '---------------------------
        'Lecture du tableau de score
        '---------------------------
        J = 1
        For I = LigneRang + 1 To DerniereLigne
            If IsNumeric(Replace(Range(Cells(I, ColRang), Cells(I, ColRang)).Value, "T", "")) Then
                If Range(Cells(I, ColRang), Cells(I, ColRang)) > 0 Then
                     If IsNumeric(Range(Cells(I, ColScore), Cells(I, ColScore))) Then '_limination des joueurs absents ou forfait
                        NomJoueur = Range(Cells(I, ColNom), Cells(I, ColNom))
                        If InStr(NomJoueur, ",") > 0 Then
                            NomJoueur = Left(NomJoueur, InStr(NomJoueur, ",") - 1) & Right(NomJoueur, Len(NomJoueur) - InStr(NomJoueur, ","))
                        End If
                        TableauJoueurs(J, 1) = TourImporte
                        TableauJoueurs(J, 2) = Range(Cells(I, ColRang), Cells(I, ColRang))
                        TableauJoueurs(J, 3) = NomJoueur
                        TableauJoueurs(J, 4) = Range(Cells(I, ColClub), Cells(I, ColClub))
                        TableauJoueurs(J, 5) = Range(Cells(I, ColIndex), Cells(I, ColIndex))
                        TableauJoueurs(J, 6) = Range(Cells(I, ColScore), Cells(I, ColScore))
                        J = J + 1
                    End If
                End If
            End If
        Next I
        
        '-------------------------------
        'Fermeture du fichier des scores
        '-------------------------------
        ActiveWorkbook.Close
        '---------------------------------------------------------------------------
        'Mise à jour de la date de la rencontre dans le classeur de feuille de match
        '---------------------------------------------------------------------------
        Application.Range("DateRencontre") = DateRencontre
        Application.Range("DateImport") = Date
        '-------------------------------------------------------------------------
        ' Recherche premier ligne et premiere colonne du tableau gneral score net
        '-------------------------------------------------------------------------
        If Net Then
            ' NET
            PremiereLigne = Range("DebutTableauGeneralNet").Row + Range("NbLignesNet").Value
            PremiereColonne = Range("DebutTableauGeneralNet").Column
            ColIndex = Range("ColIndexNet").Column
        Else
            ' BRUT
            PremiereLigne = Range("DebutTableauGeneralBrut").Row + Range("NbLignesBrut").Value
            PremiereColonne = Range("DebutTableauGeneralBrut").Column
            ColIndex = Range("ColIndexBrut").Column
        End If
        Range("A1:A1").Select
        '---------------------------------
        'Alimentation du classement global
        '---------------------------------
        Dim index As String
        
        For I = 1 To J - 1
            LigneTableau = PremiereLigne + I
            Range(Cells(LigneTableau, PremiereColonne), Cells(LigneTableau, PremiereColonne)) = TableauJoueurs(I, 1)
            Range(Cells(LigneTableau, PremiereColonne + 1), Cells(LigneTableau, PremiereColonne + 1)) = TableauJoueurs(I, 2)
            Range(Cells(LigneTableau, PremiereColonne + 2), Cells(LigneTableau, PremiereColonne + 2)) = TableauJoueurs(I, 3)
            Range(Cells(LigneTableau, PremiereColonne + 3), Cells(LigneTableau, PremiereColonne + 3)) = TableauJoueurs(I, 4)
            Range(Cells(LigneTableau, PremiereColonne + 4), Cells(LigneTableau, PremiereColonne + 4)) = TableauJoueurs(I, 5)
            index = TableauJoueurs(I, 5)
            
            'Range(Cells(LigneTableau, PremiereColonne + 5), Cells(LigneTableau, PremiereColonne + 5)) = getSerie(1.5)
            'Range(Cells(LigneTableau, PremiereColonne + 5), Cells(LigneTableau, PremiereColonne + 5)) = getSerie(Range(Cells(LigneTableau, PremiereColonne + 4), Cells(LigneTableau, PremiereColonne + 4)))
            Range(Cells(LigneTableau, PremiereColonne + 5), Cells(LigneTableau, PremiereColonne + 5)) = getSerie(index)
            'Range(Cells(LigneTableau, PremiereColonne + 5), Cells(LigneTableau, PremiereColonne + 5)).Activate
            'ActiveCell.FormulaR1C1 = "=IF(RC[-1]<=IndexMaxBB,""BB"",""JR"")"
            Range(Cells(LigneTableau, PremiereColonne + 6), Cells(LigneTableau, PremiereColonne + 6)) = TableauJoueurs(I, 6)
            Range(Cells(LigneTableau, PremiereColonne + 7), Cells(LigneTableau, PremiereColonne + 7)).Select
            'Ajout d'une liste deroulante avec les valeur Oui / Non pour les joueurs scratches
            If Net Then
                'With Selection.Validation
                '    .Delete
                '    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=LstOuiNon"
                '    .IgnoreBlank = True
                '    .InCellDropdown = True
                '    .ShowInput = True
                '    .ShowError = True
                'End With
            Else
                'Ajout d'une formule de report du scratch des joueurs du net vers le brut
                ActiveCell.FormulaR1C1 = "=IF(VLOOKUP(RC[-5],C[-14]:C[-9],6,FALSE)=0,"""",VLOOKUP(RC[-5],C[-14]:C[-9],6,FALSE))"
            End If
        Next I
        '--------------------------------------------------------------------------
        'Transformation des index en nombres (en format texte la feuille de score )
        '--------------------------------------------------------------------------
        Range(Cells(PremiereLigne + 1, ColIndex), Cells(PremiereLigne + 5000, ColIndex)).Select
        Selection.TextToColumns Destination:=Range(Cells(PremiereLigne + 1, ColIndex), Cells(PremiereLigne + 1, ColIndex)), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
    End If
    RetraitementFeuilleMatchFFGolfFichier = NomFichierTour
TypeFeuille:
    NouvelleFeuilleScore = True
Resume Next
TypeScore:
    Net = False
Resume Next
    
End Function


Function getSerie(indexString As String) As String
    Dim index As Double
    If StrComp(indexString, "PRO") = 0 Then
        getSerie = Application.Range("serie_1")
    Else
        index = CDbl(indexString)
    End If
    If index >= Application.Range("serie1IndexMin") And index <= Application.Range("serie1IndexMax") Then
        getSerie = Application.Range("serie_1")
    End If
    If index >= Application.Range("serie2IndexMin") And index <= Application.Range("serie2IndexMax") Then
        getSerie = Application.Range("serie_2")
    End If
    If index >= Application.Range("serie3IndexMin") And index <= Application.Range("serie3IndexMax") Then
        getSerie = Application.Range("serie_3")
    End If
    If index >= Application.Range("serie4IndexMin") And index <= Application.Range("serie4IndexMax") Then
        getSerie = Application.Range("serie_4")
    End If
    If index >= Application.Range("serie5IndexMin") And index <= Application.Range("serie5IndexMax") Then
        getSerie = Application.Range("serie_5")
    End If
End Function

Sub CalculTour(Optional tour As Integer, Optional CleanResult As Boolean = False)
    Dim ColTour As Integer
    Dim ColNom As Integer
    Dim ColSerie As Integer
    Dim ColScore As Integer
    Dim PremiereLigne As Integer
    Dim NbLignes As Integer
    Dim I As Integer
    'Dim tour As Integer
    Dim Nom As String
    Dim Serie As String
    Dim Score As Integer
    Dim TableauComplet(9, 35000) As Variant
    'tableau utilis_ pour construire le cumul des joueurs sur la saison
    'Dimension 1 :
    '1 : nom - pr_nom
    '2 : série
    '3 : tour
    '4 : rang brut
    '5 : score brut
    '6 : rang net
    '7 : score net
    '8 : Club
    '9 : Index
    Dim J As Integer
    Dim K As Integer
    Dim Z As Integer
    Dim NomFeuilleResultatTour As String
    Dim ColDebutTableauCumul As Integer
    Dim ColFinTableauCumul As Integer
    Dim LigneTour As Integer
    Dim LigneScore As Integer
    Dim NbTour As Integer
    Dim NbJoueurs As Integer
    Dim ListeJoueur As String
    Dim L As Integer
    Dim NomEnregFichier As String
    Dim NomFeuilleCumulJoueur As String

    ' TODO use global
    NomFeuilleResultatTour = Range("NomFeuilleResultatTour")
    NomFeuilleCumulJoueur = Range("NomFeuilleCumuljoueur")
    Worksheets(NomFeuilleResultatTour).Activate
    PremiereLigne = Range("C6").Row
    PremiereLigne = Range("DebutTableauGeneralNet").Row

    If CleanResult Then
        Call EffacementResultat
    End If
    
    '------------------
    'Lecture resultats nets
    '---------------------
    'Lecture des variables
    '---------------------
    If IsMissing(tour) Or tour = 0 Then
        tour = Range("TourSelected")
    End If
    Call recordToHistory("Calcul du Tour " & tour, , NomFeuilleCumulJoueur)
    '---------------------
    'Lecture des variables
    '---------------------
    ColClub = Range("ClubNet").Column
    ColTour = Range("DebutTableauGeneralNet").Column
    'ColIndex = Range("IndexNet").Column
    ColIndex = Range("SerieNet").Column - 1
    ColNom = Range("NomNet").Column
    ColSerie = Range("SerieNet").Column
    ColScore = Range("ScoreNet").Column
    ColRang = Range("RangNet").Column
    ColScratch = Range("ScratchNet").Column
    PremiereLigne = Range("C6").Row
    PremiereLigne = Range("DebutTableauGeneralNet").Row
    'todo get from sheet cells
    NbTour = 7
    NbLignes = Range("NbLignesNet").Value
    J = 1
    L = 1
    Z = 1
    For I = 1 To NbLignes
        '------------------------------------
        'Lecture des informations de la ligne
        '------------------------------------
        Club = Range(Cells(I + PremiereLigne, ColClub), Cells(I + PremiereLigne, ColClub))
        '''Tour = Range(Cells(I + PremiereLigne, ColTour), Cells(I + PremiereLigne, ColTour))
        'tour = Range("TourSelected")
        Nom = Range(Cells(I + PremiereLigne, ColNom), Cells(I + PremiereLigne, ColNom))
        Rang = Range(Cells(I + PremiereLigne, ColRang), Cells(I + PremiereLigne, ColRang))
        Serie = Range(Cells(I + PremiereLigne, ColSerie), Cells(I + PremiereLigne, ColSerie))
        Score = Range(Cells(I + PremiereLigne, ColScore), Cells(I + PremiereLigne, ColScore))
        Scratch = Range(Cells(I + PremiereLigne, ColScratch), Cells(I + PremiereLigne, ColScratch))
        index = Range(Cells(I + PremiereLigne, ColIndex), Cells(I + PremiereLigne, ColIndex))
        TableauComplet(1, Z) = Nom
        TableauComplet(2, Z) = Serie
        TableauComplet(3, Z) = tour
        TableauComplet(6, Z) = Rang
        TableauComplet(7, Z) = Score
        TableauComplet(8, Z) = Club
        TableauComplet(9, Z) = index
        Z = Z + 1
    Next I

    '------------------
    'Lecture resultats bruts
    '---------------------
    'Lecture des variables
    '---------------------
    ColClub = Range("ClubBrut").Column
    ColTour = Range("DebutTableauGeneralBrut").Column
    ColNom = Range("NomBrut").Column
    ColSerie = Range("SerieBrut").Column
    ColIndex = Range("SerieBrut").Column - 1
    ColScore = Range("ScoreBrut").Column
    ColRang = Range("RangBrut").Column
    ColScratch = Range("ScratchBrut").Column
    PremiereLigne = Range("DebutTableauGeneralBrut").Row
    NbLignes = Range("NbLignesBrut").Value
    J = 1
    ListeEquipe = ""
    For I = 1 To NbLignes
        '------------------------------------
        'Lecture des informations de la ligne
        '------------------------------------
        Club = Range(Cells(I + PremiereLigne, ColClub), Cells(I + PremiereLigne, ColClub))
        '''Tour = Range(Cells(I + PremiereLigne, ColTour), Cells(I + PremiereLigne, ColTour))
        'tour = Range("TourSelected")
        Nom = Range(Cells(I + PremiereLigne, ColNom), Cells(I + PremiereLigne, ColNom))
        Rang = Range(Cells(I + PremiereLigne, ColRang), Cells(I + PremiereLigne, ColRang))
        Serie = Range(Cells(I + PremiereLigne, ColSerie), Cells(I + PremiereLigne, ColSerie))
        Score = Range(Cells(I + PremiereLigne, ColScore), Cells(I + PremiereLigne, ColScore))
        Scratch = Range(Cells(I + PremiereLigne, ColScratch), Cells(I + PremiereLigne, ColScratch))
        index = Range(Cells(I + PremiereLigne, ColIndex), Cells(I + PremiereLigne, ColIndex))
        TableauComplet(1, Z) = Nom
        TableauComplet(2, Z) = Serie
        TableauComplet(3, Z) = tour
        TableauComplet(4, Z) = Rang
        TableauComplet(5, Z) = Score
        TableauComplet(8, Z) = Club
        TableauComplet(9, Z) = index
        Z = Z + 1
    Next I
    
    '-------------------------------
    'Mise a jour du cumul individuel
    '-------------------------------
    Worksheets(NomFeuilleCumulJoueur).Activate
    '-------------------------------------------
    'Constitution de la liste des joueurs deja existants
    '-------------------------------------------
    Dim LigNotEmpty As Long
    debutTableau = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultat").Row
    LignStartPlayer = debutTableau + 1
    LigNotEmpty = LignStartPlayer
    Do While Not IsEmpty(Range("B" & LigNotEmpty))
        LigNotEmpty = LigNotEmpty + 1
    Loop
    lastr_row = LigNotEmpty
    
    NbJoueurs = LigNotEmpty - (LignStartPlayer)
    
    ColNom = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultat").Column
    For I = LignStartPlayer To LignStartPlayer + NbJoueurs
        Nom = Range(Cells(I, ColNom), Cells(I, ColNom))
        If InStr(ListeJoueur, Nom) = 0 Then
            ListeJoueur = ListeJoueur + Format(I, "0000") + " " + Nom + "/ "
        End If
    Next I
    
    '------------------------------
    'Init Formulas
    '------------------------------
    'Call initFormula
    playerType = Range("PlayerType")
    Select Case playerType
        Case "HOMME"
            FormulaBestNett = "=IF(MIN(RC[-24],RC[-20],RC[-16],RC[-12],RC[-8],RC[-4])<>0,MIN(RC[-24],RC[-20],RC[-16],RC[-12],RC[-8],RC[-4]),"""")"
            FormulaBestBrut = "=IF(MIN(RC[-23],RC[-19],RC[-15],RC[-11],RC[-7],RC[-3])<>0,MIN(RC[-23],RC[-19],RC[-15],RC[-11],RC[-7],RC[-3]),"""")"
        Case "DAME"
            FormulaBestNett = "=IF(MIN(RC[-24],RC[-20],RC[-16],RC[-12],RC[-8],RC[-4])<>0,MIN(RC[-24],RC[-20],RC[-16],RC[-12],RC[-8],RC[-4]),"""")"
            FormulaBestBrut = "=IF(MIN(RC[-23],RC[-19],RC[-15],RC[-11],RC[-7],RC[-3])<>0,MIN(RC[-23],RC[-19],RC[-15],RC[-11],RC[-7],RC[-3]),"""")"
        Case Else
            FormulaBestNett = "=IF(MIN(RC[-24],RC[-20],RC[-16],RC[-12],RC[-8],RC[-4])<>0,MIN(RC[-24],RC[-20],RC[-16],RC[-12],RC[-8],RC[-4]),"""")"
            FormulaBestBrut = "=IF(MIN(RC[-23],RC[-19],RC[-15],RC[-11],RC[-7],RC[-3])<>0,MIN(RC[-23],RC[-19],RC[-15],RC[-11],RC[-7],RC[-3]),"""")"
    End Select
    FormulaTotalNett = "=IF(RC[-6]="""","""",IF(ISBLANK(RC[-4]),""En cours"",RC[-6]+RC[-4]))"
    FormulaTotalBrut = "=IF(RC[-6]="""","""",IF(ISBLANK(RC[-3]),""En cours"",RC[-6]+RC[-3]))"
    
    '------------------------------
    'Insertion des joueurs et des scores
    '------------------------------
    offsetTour = (tour - 1) * 4
    startResultatCol = offsetTour + ColNom + 4
    
    'Finale
    If (tour = NbTour) Then
        offsetTour = offsetTour + 2
        startResultatCol = offsetTour + ColNom + 4
    End If
    For I = 1 To Z - 1
        '-----------------------------------------------------------------------
        'Lecture des informations compil_es lors de la lecture des feuilles de score
        '-----------------------------------------------------------------------
        Nom = TableauComplet(1, I)
        Serie = TableauComplet(2, I)
        tour = TableauComplet(3, I)
        RangBrut = TableauComplet(4, I)
        ScoreBrut = TableauComplet(5, I)
        RangNet = TableauComplet(6, I)
        ScoreNet = TableauComplet(7, I)
        Club = TableauComplet(8, I)
        index = TableauComplet(9, I)
        '-----------------------------------------------------------------------
        'Recherche si le joueur est deja dans le tableau de cumul des joueurs, insertion sinon
        '-----------------------------------------------------------------------
        endCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatEnd").Column
        bestTourNettCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatMaxNet").Column
        bestTourBrutCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatMaxBrut").Column
        totalTourNettCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatTotalNet").Column
        totalTourBrutCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatTotalBrut").Column
        
        If tour > 0 Then
            If InStr(ListeJoueur, Nom) > 0 Then
                LigneScore = Mid(ListeJoueur, InStr(ListeJoueur, Nom) - 5, 4)
            Else
                LigneScore = LignStartPlayer + NbJoueurs
                'Ajout du joueur a la liste des joueurs
                NbJoueurs = NbJoueurs + 1
                ListeJoueur = ListeJoueur + Format(LigneScore, "0000") + " " + Nom + "/ "
                'Insertion du nom
                Range(Cells(LigneScore, ColNom), Cells(LigneScore, ColNom)) = Nom
                
                ' Max
                'Range(Cells(LigneScore, ColNom + bestTourNettCol), Cells(LigneScore, ColNom + bestTourNettCol)).Select
                'Selection.value = "=MAX(F4;J4;N4;R4;V4;Z4)"
                'Range(Cells(LigneScore, ColNom + bestTourNettCol), Cells(LigneScore, ColNom + bestTourNettCol)) = "MAX(F4;J4;N4;R4;V4;Z4)"
                
                'ajout du quadrillage
                Range(Cells(LigneScore, ColNom), Cells(LigneScore, endCol)).Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            End If
            '---------------------------------------
            'Insertion des informations dans le tableau
            '---------------------------------------
            Range(Cells(LigneScore, ColNom + 1), Cells(LigneScore, ColNom + 1)) = Club
            Range(Cells(LigneScore, ColNom + 2), Cells(LigneScore, ColNom + 2)) = index
            Range(Cells(LigneScore, ColNom + 3), Cells(LigneScore, ColNom + 3)) = Serie
            ColScore = offsetTour
            ColBestNet = NbTour * 4 + 1
            ColBestBrut = ColBestNet + 1

            If RangBrut > 0 Then
                Range(Cells(LigneScore, startResultatCol + 2), Cells(LigneScore, startResultatCol + 2)).Select
                Range(Cells(LigneScore, startResultatCol + 2), Cells(LigneScore, startResultatCol + 2)) = ScoreBrut
                Range(Cells(LigneScore, startResultatCol + 3), Cells(LigneScore, startResultatCol + 3)) = RangBrut
            End If
            If RangNet > 0 Then
                Range(Cells(LigneScore, startResultatCol), Cells(LigneScore, startResultatCol)) = ScoreNet
                Range(Cells(LigneScore, startResultatCol + 1), Cells(LigneScore, startResultatCol + 1)) = RangNet
            End If
            '---------------------------------------
            'Insertion des formules de calcul du meilleur score et Total de la semaine
            '---------------------------------------
            
            Range(Cells(LigneScore, bestTourNettCol), Cells(LigneScore, bestTourNettCol)).Select
            Call setFormula(FormulaBestNett)
            Range(Cells(LigneScore, bestTourBrutCol), Cells(LigneScore, bestTourBrutCol)).Select
            Call setFormula(FormulaBestBrut)
            Range(Cells(LigneScore, totalTourNettCol), Cells(LigneScore, totalTourNettCol)).Select
            Call setColorConditional(True)
            Call setFormula(FormulaTotalNett)
            Range(Cells(LigneScore, totalTourBrutCol), Cells(LigneScore, totalTourBrutCol)).Select
            Call setColorConditional(True)
            Call setFormula(FormulaTotalBrut)
        End If
    Next I
    Range(Cells(LignStartPlayer, startResultatCol), Cells(LignStartPlayer, startResultatCol)).Select
End Sub
Sub setFormula(formula)
    'Debug.Print "formula 1: " & Selection.Formula
    'Debug.Print "formula 1: " & Selection.FormulaR1C1
    
    Selection.FormulaR1C1 = formula
    
    'Debug.Print "formula 3: " & Selection.Formula
    'Debug.Print "formula 3: " & Selection.FormulaR1C1
End Sub


Sub setColor(colorType, isBold)
    ColorInProgress = RGB(255, 0, 0)
    ColorDone = RGB(72, 148, 31)
    
    Select Case colorType
        Case "inProgress"
            ColorToSet = ColorInProgress
        Case "done"
            ColorToSet = ColorDone
        Case Else
            ColorToSet = ColorInProgress
    End Select
    
    'Debug.Print "style 3: Bold :" & Selection.Font.bold
    'Debug.Print "style 3: Color :" & Selection.Font.Color
    'Debug.Print "style 3: " & Selection.FormatConditions.Count
    Selection.Font.Color = ColorToSet
    Selection.Font.bold = isBold
End Sub
Sub setColorConditional(isBold)
    ColorInProgress = RGB(255, 0, 0)
    ColorDone = RGB(72, 148, 31)
    
    With Selection.FormatConditions.Delete
    End With
    With Selection.FormatConditions.Delete
    End With
    With Selection.FormatConditions _
        .Add(xlCellValue, xlEqual, "En cours")
        With .Font
         .bold = isBold
         .Color = ColorInProgress
        End With
    End With
    With Selection.FormatConditions _
        .Add(xlCellValue, xlNotEqual, "En cours")
        With .Font
         .bold = isBold
         .Color = ColorDone
        End With
    End With
    
End Sub

Public Sub setGender(gender As Integer)
    Worksheets("Import Resultats Tour").Activate
    
    Range("F13").Value = gender

End Sub

Public Sub GetScoresFromFFGolfHommeDame(Optional scoreFolder As String, Optional ByVal Clean As Boolean = True, Optional TaskType As String = "Importation et generation de tous les tours depuis un repertoire HOMME/DAME")
    Set rg = ActiveSheet.ListObjects("TableauType").DataBodyRange
    For r = 2 To rg.Rows.Count
        genre = rg(r - 1, 1).Value
        genreIdx = r - 1
        setGender (genreIdx)
        Call GetScoresFromFFGolf(scoreFolder, Clean, TaskType)
        Call EffacementImportForced
    'Next Counter
    Next
End Sub


Public Sub GetScoresFromFFGolf(Optional scoreFolder As String, Optional ByVal Clean As Boolean = True, Optional TaskType As String = "Importation et generation de tous les tours depuis un repertoire")
    'TODO use global
    NbTour = 7
    iTourFinal = 7
    tourFolderPatternPrefix = "T"
    tourFolderFinale = "Finale"
    ShowMissingFolder = True
    
    Dim scoreFolderRoot As String
    scoreFolderRoot = scoreFolder
    
    If scoreFolderRoot = "" Then
        scoreFolderRoot = GetScoreFolder("")
    End If
    scoreFolder = scoreFolderRoot
    
    Call recordToHistory(TaskType, scoreFolderRoot)
    
    'Dim NomFichierBrut As String
    Dim tour As Integer
    Dim scoreFolderTour As String
    For iTour = 1 To NbTour
        If iTour = iTourFinal Then
            scoreFolderTour = scoreFolderRoot & "\" & tourFolderFinale
        Else
            scoreFolderTour = scoreFolderRoot & "\" & tourFolderPatternPrefix & iTour
        End If
        
        If Path_Exists(scoreFolderTour) Then
            Dim fileList(1) As String
            'fileList(0) = scoreFolderTour & "\" & "export_DAME_Brut_12.xlsx"
            fileList(0) = scoreFolderTour & "\" & Range("export_Brut_Strokeplay_filename")
            fileList(1) = scoreFolderTour & "\" & Range("export_Brut_Stabelford_filename")
            
            'NomFichierBrut = fileList(1)
            'export_Brut_Stabelford_filename
            'Call ImporterBrutNet(fileList(1), False)
            tour = iTour
            Dim CleanResult As Boolean
            CleanResult = Range("cleanResult").Value And (iTour = 1)
            
            Call ImporterBrutNet(TaskType, fileList(0), tour, True, CleanResult)
            fileList(1) = FixMissingBrut(fileList(1))
            Call ImporterBrutNet(TaskType, fileList(1), tour, False, False)
            
        Else
            If ShowMissingFolder Then
                MsgBox scoreFolderTour & " n'existe pas"
            End If
        End If
    Next iTour
End Sub

Public Function GetScoreFolder(Optional scoreFolder As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    If scoreFolder <> "" Then GoTo NextCode2
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetScoreFolder = sItem
    Set fldr = Nothing
    Exit Function
NextCode2:
    GetScoreFolder = scoreFolder
    Set fldr = Nothing
    Exit Function
End Function


Function Path_Exists(Path As String)

'Dim Path As String
Dim Folder As String
Dim Answer As VbMsgBoxResult
    'Path = "C:\Users\LG\Desktop\VBA\S12"
    Folder = Dir(Path, vbDirectory)
    If Folder = vbNullString Then
        Path_Exists = False
    Else
        Path_Exists = True
    End If
End Function

Function File_Is_Brut(FileName As String)
If InStr(LCase(FileName), "brut") <> 0 Then
    File_Is_Brut = True
Else
    File_Is_Brut = False
End If
End Function

Public Function FixMissingBrut(inputFile As String)
    If Not Path_Exists(inputFile) Then
        FixMissingBrut = GetFileNameNetFromBrut(inputFile)
    Else
        FixMissingBrut = inputFile
    End If
End Function

Public Sub ImporterBrutNet(Optional TaskType As String = "", Optional inputFile As String = "", Optional tour As Integer, Optional Clean As Boolean = False, Optional CleanResult As Boolean = False)
    'TODO refactoring
    Worksheets("Import Resultats Tour").Activate
    If Path_Exists(inputFile) Then
        If Clean Then
            EffacementImport
        End If
        If TaskType = "" Then
            TaskType = "Importation des fichiers Brut et Net (couple) FFGolf pour 1 Tour"
        End If
        Call ImporterBrutNetFromFiles(inputFile, TaskType & " (Clean import =" & Clean & ")", tour, CleanResult)
    End If
End Sub

Sub ImporterBrutNetFromFiles(NomFichierBrutBase As String, TaskType As String, Optional tour As Integer, Optional CleanResult As Boolean = False)
    Dim NomFichierBrut As String
    Dim NomFichierNet As String
    NomFichierBrut = "" & NomFichierBrutBase
    NomFichierBrut = RetraitementFeuilleMatchFFGolfFichier(NomFichierBrutBase, TaskType, tour, "Brut")
    If File_Is_Brut(NomFichierBrut) Then
        NomFichierNet = GetFileNameNetFromBrut(NomFichierBrut)
        NomFichierNet = RetraitementFeuilleMatchFFGolfFichier(NomFichierNet, TaskType, tour, "Net")
    End If
    
    Call CalculTour(tour, CleanResult)
    
    Worksheets("Import Resultats Tour").Activate
End Sub

Function GetFileNameNetFromBrut(NomFichierBrut)
    Dim NomFichierNet As String
    NomFichierNet = Replace(NomFichierBrut, "Brut", "Net")
    'NomFichierNet = Replace(NomFichierBrut, "BRUT", "NET")
    'NomFichierNet = Replace(NomFichierBrut, "brut", "net")

    GetFileNameNetFromBrut = NomFichierNet
End Function

Sub ClearHistory()
    Dim historySheetName As String
    Dim historyTableName As String
    historySheetName = "Historique Import"
    historyTableName = "Tableau1"
    
    Set historySheet = Worksheets(historySheetName)
    Set tbl = historySheet.ListObjects(historyTableName)
    Call ClearTable(tbl)
End Sub

Function getHistorySheet() As Worksheet
    ' DOES NOT WORK
    Dim historySheetName As String
    Dim historyTableName As String
    historyTableName = "Tableau1"
    historySheetName = "Historique Import"
    
    Set historySheet = Worksheets(historySheetName)

    getHistorySheet = historySheet.ListObjects(historyTableName)
End Function

Sub recordToHistory(TaskType As String, Optional reference As String = "Nan", Optional sheet As String = "Nan")
    Dim historySheetName As String
    Dim historyTableName As String
    historySheetName = "Historique Import"
    historyTableName = "Tableau1"
    playerType = Range("playerType")
    
    If (sheet = "") Then
        sheet = Range("NomFeuilleCumuljoueur")
    End If
    
    Set historySheet = Worksheets(historySheetName)
    Set tbl = historySheet.ListObjects(historyTableName)
    
    Debug.Print "write to history: " & TaskType & ", " & reference
    Set newrow = tbl.ListRows.Add
    With newrow
        .Range(1) = TaskType
        .Range(2) = playerType
        .Range(3) = sheet
        .Range(4) = reference
        .Range(5) = Now
    End With
    'End
End Sub

Sub ClearTable(ByVal tbl As ListObject)
    With tbl
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
End Sub
