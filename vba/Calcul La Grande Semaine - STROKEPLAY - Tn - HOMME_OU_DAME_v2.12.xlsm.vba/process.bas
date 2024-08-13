Function processGolfMatchSheetFromFile(NomFichierTour As String, TaskType As String, Optional TourImporte As Integer)
    Dim ColRang As Integer
    Dim ColNom As Integer
    Dim ColClub As Integer
    Dim ColIndex As Integer
    
    Dim LigneRang As Integer
    Dim I As Integer
    Dim DerniereLigne As Long
    Dim J As Integer
    Dim TableauJoueurs(400, 11) As Variant
    Dim Tour As Integer
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

    Dim wb As Workbook
    Dim WsImportResultTour As Worksheet
    
    '---------------------
    'Check 'Tour' is set
    '---------------------
    If (IsMissing(TourImporte) Or TourImporte = 0) Then
        TourImporte = Right(Application.Range("Tour"), 1) + 1
    End If
    
    '---------------------
    'Check export filename is provided
    '---------------------
    If InStr(NomFichierTour, ":") = 0 Then
        '------------------------------------------
        'Open window to choose Export file
        '------------------------------------------
        NomFichierTour = Application.GetOpenFilename(Title:="Import du resultat d'un tour" & " " & ScoreType)
    End If
    
    Debug.Print "Traitement du fichier: " & NomFichierTour;
    If InStr(NomFichierTour, ":") = 0 Then
        Erreur = True
        MessageErreur = "Vous n'avez pas sélectionné de feuille de résultat à importer. Fin de la procédure"
        I = MsgBox(MessageErreur, vbOKOnly, "Import des résultats d'un tour")
        End
    End If
    
    Call recordToHistory(TaskType & " - Tour " & TourImporte, NomFichierTour, "ALL", "ALL")
    
    Set TableauJoueursIdx = CreateObject("Scripting.Dictionary")
    TableauJoueursIdx.Add "date", 0
    TableauJoueursIdx.Add "competition", 1
    TableauJoueursIdx.Add "tour", 2
    TableauJoueursIdx.Add "rang", 3
    TableauJoueursIdx.Add "name", 4
    TableauJoueursIdx.Add "club", 5
    TableauJoueursIdx.Add "index", 6
    TableauJoueursIdx.Add "score", 7
    TableauJoueursIdx.Add "serie", 8
    TableauJoueursIdx.Add "serie_calc", 9
    TableauJoueursIdx.Add "score_type", 10
    TableauJoueursIdx.Add "genre", 11
    
    Set wb = Application.Workbooks.Open(NomFichierTour)
    Set WsImportResultTour = wb.Worksheets("Report")
    
    DerniereLigne = WsImportResultTour.Cells(WsImportResultTour.Rows.Count, "A").End(xlUp).Row
    
    '---------------------------
    'Lecture du tableau de score
    '---------------------------
    currentScoreIdx = 0
    For I = 1 To DerniereLigne
        nomComp = GetValueByColumnName(WsImportResultTour, "Nom competition", I)
        If nomComp <> "Nom competition" Then
            If IsNumeric(GetValueByColumnName(WsImportResultTour, "Score Tour 1", I)) Then '_limination des joueurs absents ou forfait
            
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("date")) = GetValueByColumnName(WsImportResultTour, "Date", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("competition")) = GetValueByColumnName(WsImportResultTour, "Nom competition", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("tour")) = TourImporte
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("rang")) = GetValueByColumnName(WsImportResultTour, "Rang", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("name")) = GetValueByColumnName(WsImportResultTour, "Nom / prenom", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("club")) = GetValueByColumnName(WsImportResultTour, "Club", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("index")) = GetValueByColumnName(WsImportResultTour, "Index Cpt", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("score")) = GetValueByColumnName(WsImportResultTour, "Score Tour 1", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("serie")) = GetValueByColumnName(WsImportResultTour, "Série d'index", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("serie_calc")) = getSerieMock(GetValueByColumnName(WsImportResultTour, "Index Cpt", I))
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("score_type")) = GetValueByColumnName(WsImportResultTour, "Brut / net", I)
                TableauJoueurs(currentScoreIdx, TableauJoueursIdx("genre")) = GetValueByColumnName(WsImportResultTour, "Sexe", I)
                Increment currentScoreIdx
            End If
        End If
    Next I
    
    scoreCount = currentScoreIdx
    
    wb.Close SaveChanges:=False
    
    Call InsertDataImported(TableauJoueurs, TableauJoueursIdx, scoreCount)

End Function
