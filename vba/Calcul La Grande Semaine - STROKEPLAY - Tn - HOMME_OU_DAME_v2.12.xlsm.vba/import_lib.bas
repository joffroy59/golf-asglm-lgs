Public Sub ImporterBrutNetForTour(Optional TaskType As String = "", Optional inputFile As String = "", Optional Tour As Integer, Optional CleanImport As Boolean = False)
    'TODO refactoring
    Worksheets("Import Resultats Tour").Activate
    If Path_Exists(inputFile) Then
        If CleanImport Then
            EffacementImport
        End If
        If TaskType = "" Then
            TaskType = "Importation d ''un fichier Brut et Net (complet Homme Dame) FFGolf pour 1 Tour (2024) [auto]" & " (Clean import =" & CleanImport & ")"
        End If
        
        Call ImporterBrutNetFromFiles(inputFile, TaskType, Tour)
    End If
End Sub

Sub ImporterBrutNetFromFiles(NomFichierExportOrNull As String, TaskType As String, Optional Tour As Integer)
        Dim ExportFilename As String
        Dim NomFichierNet As String
    ExportFilename = "" & NomFichierExportOrNull
    ExportFilename = processGolfMatchSheetFromFile(NomFichierExportOrNull, TaskType, Tour)
    
    Call InsertTourFromImport(Tour)
    
    
End Sub

Sub InsertDataImported(TableauJoueurs As Variant, TableauJoueursIdx, scoreCount)
    Dim WsImportResultTour As Worksheet
    
    Set WsImportResultTour = Worksheets("Import Resultats Tour")
    
    PremiereLigneNet = WsImportResultTour.Range("DebutTableauGeneralNet").Row + WsImportResultTour.Range("NbLignesNet").Value
    PremiereColonneNet = WsImportResultTour.Range("DebutTableauGeneralNet").Column
    ColIndexNet = WsImportResultTour.Range("ColIndexNet").Column
    PremiereLigneNetCurrent = PremiereLigneNet
    
    PremiereLigneBrut = WsImportResultTour.Range("DebutTableauGeneralBrut").Row + WsImportResultTour.Range("NbLignesBrut").Value
    PremiereColonneBrut = WsImportResultTour.Range("DebutTableauGeneralBrut").Column
    ColIndexBrut = WsImportResultTour.Range("ColIndexBrut").Column
    PremiereLigneBrutCurrent = PremiereLigneBrut
    
    'ResetCellActive
    
    For I = 0 To scoreCount - 1
        ScoreType = TableauJoueurs(I, TableauJoueursIdx("score_type"))

        If IsNet(ScoreType) Then
            Increment PremiereLigneNetCurrent
            PremiereLigne = PremiereLigneNetCurrent
            ResultCol = PremiereColonneNet
        ElseIf IsBrut(ScoreType) Then
            Increment PremiereLigneBrutCurrent
            PremiereLigne = PremiereLigneBrutCurrent
            ResultCol = PremiereColonneBrut
        End If
        
        resultLine = PremiereLigne
        
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("tour"))
        Increment ResultCol
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("rang"))
        Increment ResultCol
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("name"))
        Increment ResultCol
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("club"))
        Increment ResultCol
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("index"))
        colIndexCurrent = ResultCol
        Increment ResultCol
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("serie"))
        Increment ResultCol
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("score"))
        Increment ResultCol
        WsImportResultTour.Range(WsImportResultTour.Cells(resultLine, ResultCol), WsImportResultTour.Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("genre"))
         
        '--------------------------------------------------------------------------
        'Transformation des index en nombres (en format texte la feuille de score )
        '--------------------------------------------------------------------------
        Call FixIndexToNumber(WsImportResultTour, resultLine, colIndexCurrent)
    Next I
    
    Call ResetCellActive(WsImportResultTour)

End Sub

Sub ResetCellActive(ws As Worksheet)
    ws.Range("A1:A1").Select
End Sub

Sub FixIndexToNumber(ws, resultLine, ColIndex)
    ws.Range(Cells(resultLine, ColIndex), ws.Cells(resultLine, ColIndex)).TextToColumns Destination:=ws.Range(ws.Cells(resultLine, ColIndex), Cells(resultLine, ColIndex)), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
End Sub

Function IsNet(ScoreType)
    IsNet = (ScoreType = "Net")
End Function

Function IsBrut(ScoreType)
    IsBrut = (ScoreType = "Brut")
End Function

Sub InsertTourFromImport(Optional Tour As Integer)
    Dim CleanResult As Boolean
    CleanResult = Range("cleanResult").Value
    
    Call InsertTourFromImportClean(Tour, CleanResult)
    
    
End Sub


Sub InsertTourFromImportClean(Optional Tour As Integer, Optional CleanResult As Boolean = False)
    Dim ColTour As Integer
    Dim ColNom As Integer
    Dim ColSerie As Integer
    Dim ColScore As Integer
    Dim ColDebutTableauCumul As Integer
    Dim ColFinTableauCumul As Integer
    
    Dim PremiereLigne As Integer
    Dim NbLignes As Integer
    
    Dim I As Integer
    
    Dim nom As String
    Dim Serie As String
    Dim Score As Integer

    Dim J As Integer
    Dim K As Integer
    Dim Z As Integer
    Dim L As Integer
    
    Dim NomFeuilleResultatTour As String
    
    Dim LigneTour As Integer
    Dim LigneScore As Integer
    
    Dim NbTour As Integer
    Dim NbJoueurs As Integer
    
    Dim ListeJoueur As String
    
    Dim NomEnregFichier As String
    
    Dim NomFeuilleCumulJoueur As String
    
    NomFeuilleResultatTour = Range("NomFeuilleResultatTour")
    Set FeuilleResultatTour = Worksheets(NomFeuilleResultatTour)
    
    CALCUL_MAM = True
    CALCUL_WOMAM = True
    
    NbTour = 7
    If IsMissing(Tour) Or Tour = 0 Then
        Tour = FeuilleResultatTour.Range("TourSelected")
    End If
    
    PremiereLigne = FeuilleResultatTour.Range("DebutTableauGeneralNet").Row
    
    Dim genre As String
    
    Call recordToHistory("Calcul du Tour " & Tour, , , "ALL")
    
    If CALCUL_WOMAM Then
        Call IntegrateScoreByGenre("Dames", CleanResult, Tour, FeuilleResultatTour)
    End If
    
    If CALCUL_MAM Then
        Call IntegrateScoreByGenre("Messieurs", CleanResult, Tour, FeuilleResultatTour)
    End If
    
End Sub
