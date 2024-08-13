Sub InsertNameForPlayer(wsCumulJoueur, LigneScore, ColNom, nom)
    wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColNom), wsCumulJoueur.Cells(LigneScore, ColNom)) = nom
End Sub

Sub InsertScoreForPlayer(wsCumulJoueur, LigneScore, startResultatCol, ColNom, nom, Club, index, Serie, ScoreBrut, RangBrut, ScoreNet, RangNet)
    ColClub = ColNom + 1
    ColIndex = ColNom + 2
    ColSerie = ColNom + 3
    ColScore = offsetTour
    ColBestNet = NbTour * 4 + 1
    ColBestBrut = ColBestNet + 1
    
    ColScoreNet = startResultatCol
    ColRangNet = startResultatCol + 1
    ColScoreBrut = startResultatCol + 2
    ColRangBrut = startResultatCol + 3

    wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColClub), wsCumulJoueur.Cells(LigneScore, ColClub)) = Club
    wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColIndex), wsCumulJoueur.Cells(LigneScore, ColIndex)) = index
    wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColSerie), wsCumulJoueur.Cells(LigneScore, ColSerie)) = GetPrefixSerie(Serie)
    If RangBrut > 0 Then
        wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColScoreBrut), wsCumulJoueur.Cells(LigneScore, ColScoreBrut)) = ScoreBrut
        wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColRangBrut), wsCumulJoueur.Cells(LigneScore, ColRangBrut)) = RangBrut
    End If
    If RangNet > 0 Then
        wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColScoreNet), wsCumulJoueur.Cells(LigneScore, ColScoreNet)) = ScoreNet
        wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColRangNet), wsCumulJoueur.Cells(LigneScore, ColRangNet)) = RangNet
    End If

End Sub

Sub InsertFormaulasBestScore(wsCumulJoueur, playerType, LigneScore, bestTourNettCol, bestTourBrutCol, totalTourNettCol, totalTourBrutCol)
    Dim rng As Range
    
    InitFormulas (playerType)

    wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, bestTourNettCol), wsCumulJoueur.Cells(LigneScore, bestTourNettCol)).FormulaR1C1 = FormulaBestNett
    wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, bestTourBrutCol), wsCumulJoueur.Cells(LigneScore, bestTourBrutCol)).FormulaR1C1 = FormulaBestBrut
    
    Set rng = wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, totalTourNettCol), wsCumulJoueur.Cells(LigneScore, totalTourNettCol))
    Call setColorConditional(rng, True)
    wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, totalTourNettCol), wsCumulJoueur.Cells(LigneScore, totalTourNettCol)).FormulaR1C1 = FormulaTotalNett
    
    Set rng = wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, totalTourBrutCol), wsCumulJoueur.Cells(LigneScore, totalTourBrutCol))
    Call setColorConditional(rng, True)
    rng.FormulaR1C1 = FormulaTotalBrut
End Sub

Sub InitFormulas(playerType)
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
End Sub

Function GetLastRowFrom(wsCumulJoueur, LignStartPlayer)
    LigNotEmpty = LignStartPlayer
    Do While Not IsEmpty(wsCumulJoueur.Range("B" & LigNotEmpty))
        LigNotEmpty = LigNotEmpty + 1
    Loop
    GetLastRowFrom = LigNotEmpty
End Function

Public Function GetScoreFolder(Optional ScoreFolder As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    If ScoreFolder <> "" Then GoTo NextCode2
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
    GetScoreFolder = ScoreFolder
    Set fldr = Nothing
    Exit Function
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


Function GetPlayerScoreLine(wsCumulJoueur, ListeJoueur, nom, ByRef LignStartPlayer, ByRef NbJoueurs, ColNom, EndCol)
    If InStr(ListeJoueur, nom) > 0 Then
        LigneScore = Mid(ListeJoueur, InStr(ListeJoueur, nom) - 5, 4)
    Else
        LigneScore = LignStartPlayer + NbJoueurs
        
        'Ajout du joueur a la liste des joueurs
        NbJoueurs = NbJoueurs + 1
        ListeJoueur = ListeJoueur + Format(LigneScore, "0000") + " " + nom + "/ "
        
        'Insertion du nom
        'Range(Cells(LigneScore, ColNom), Cells(LigneScore, ColNom)) = Nom
        Call InsertNameForPlayer(wsCumulJoueur, LigneScore, ColNom, nom)
        
        'ajout du quadrillage
        Call AppliquerBordures(wsCumulJoueur, LigneScore, ColNom, EndCol)
        
    End If
    GetPlayerScoreLine = LigneScore
End Function



Function Mode_NomFeuilleCumulJoueur(genre As String)
    Dim ws As Worksheet
    Set ws = Worksheets("Import Resultats Tour")
    
    If genre = ws.Range("X19") Then
        Mode_NomFeuilleCumulJoueur = ws.Range("Z19")
    ElseIf genre = ws.Range("X20") Then
        Mode_NomFeuilleCumulJoueur = ws.Range("Z20")
    End If

End Function




