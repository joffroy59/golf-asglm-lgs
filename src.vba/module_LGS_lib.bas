Sub integrate_01(genre As String, CleanResult As Boolean, Tour, ws)
    Dim TableauComplet(10, 35000) As Variant
    Dim NomFeuilleCumulJoueur As String
    
    NomFeuilleCumulJoueur = Mode_NomFeuilleCumulJoueur(genre)
    
    Call recordToHistory("Calcul du Tour " & Tour, , NomFeuilleCumulJoueur, "ALL - " & genre)
        
    If CleanResult Then
        Call EffacementResultat(GetSheetExportByGenre(genre))
    End If
    
    ws.Activate
    ActiveSheet.AutoFilterMode = False
    
    Z = 1
    
    '------------------------
    'Lecture resultats Nets
    '------------------------
    'Lecture des variables
    '------------------------
    ' Test git
    score_type = "Net"
    ColClub = Range("ClubNet").Column
    ColTour = Range("DebutTableauGeneralNet").Column
    ColNom = Range("NomNet").Column
    ColSerie = Range("SerieNet").Column
    colIndex = Range("IndexNet").Column
    ColScore = Range("ScoreNet").Column
    ColRang = Range("RangNet").Column
    colGenreNet = Range("GenreNet").Column
    colGenre = Range("GenreNet").Column
    PremiereLigne = Range("DebutTableauGeneralNet").Row
    NbLignes = Range("NbLignesNet").Value
    '------------------------------------
    'Lecture des informations
    '------------------------------------
    readAll TableauComplet, score_type, Z, I, PremiereLigne, Tour, genre, NbLignes, ColNom, ColSerie, ColRang, ColScore, ColClub, colIndex, colGenre

    '------------------------
    'Lecture resultats Bruts
    '------------------------
    'Lecture des variables
    '------------------------
    score_type = "Brut"
    ColClub = Range("ClubBrut").Column
    ColTour = Range("DebutTableauGeneralBrut").Column
    ColNom = Range("NomBrut").Column
    ColSerie = Range("SerieBrut").Column
    colIndex = Range("IndexBrut").Column
    ColScore = Range("ScoreBrut").Column
    ColRang = Range("RangBrut").Column
    colGenre = Range("GenreBrut").Column
    PremiereLigne = Range("DebutTableauGeneralBrut").Row
    NbLignes = Range("NbLignesBrut").Value
    '------------------------------------
    'Lecture des informations
    '------------------------------------
    readAll TableauComplet, score_type, Z, I, PremiereLigne, Tour, genre, NbLignes, ColNom, ColSerie, ColRang, ColScore, ColClub, colIndex, colGenre

    '-------------------------------
    'Mise a jour du cumul individuel
    '-------------------------------
    Worksheets(NomFeuilleCumulJoueur).Activate
    '-------------------------------------------
    'Constitution de la liste des joueurs deja existants
    '-------------------------------------------
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
    ' TODO FIX remove because each case identic
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
    offsetTour = (Tour - 1) * 4
    startResultatCol = offsetTour + ColNom + 4
    
    'Finale
    If (Tour = NbTour) Then
        offsetTour = offsetTour + 2
        startResultatCol = offsetTour + ColNom + 4
    End If
    For I = 1 To Z - 1
        '-----------------------------------------------------------------------
        'Lecture des informations compilees lors de la lecture des feuilles de score
        '-----------------------------------------------------------------------
        Nom = TableauComplet(0, I)
        Serie = TableauComplet(1, I)
        Tour = TableauComplet(2, I)
        RangBrut = TableauComplet(3, I)
        ScoreBrut = TableauComplet(4, I)
        RangNet = TableauComplet(5, I)
        ScoreNet = TableauComplet(6, I)
        Club = TableauComplet(7, I)
        index = TableauComplet(8, I)
        genre = TableauComplet(9, I)
        
        '-----------------------------------------------------------------------
        'Recherche si le joueur est deja dans le tableau de cumul des joueurs, insertion sinon
        '-----------------------------------------------------------------------
        endCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatEnd").Column
        bestTourNettCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatMaxNet").Column
        bestTourBrutCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatMaxBrut").Column
        totalTourNettCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatTotalNet").Column
        totalTourBrutCol = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultatTotalBrut").Column
        
        If Tour > 0 Then
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
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            End If
            
            '---------------------------------------
            'Insertion des informations dans le tableau
            '---------------------------------------
            Range(Cells(LigneScore, ColNom + 1), Cells(LigneScore, ColNom + 1)) = Club
            Range(Cells(LigneScore, ColNom + 2), Cells(LigneScore, ColNom + 2)) = index
            Range(Cells(LigneScore, ColNom + 3), Cells(LigneScore, ColNom + 3)) = GetPrefixSerie(Serie)
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