Sub FillAllResult(ByRef TableauComplet, wsResultatTour, score_type, Z, Tour, genre)
    
    NbLignes = wsResultatTour.Range("NbLignes" & score_type).Value
    
    For L = 1 To NbLignes
        '------------------------------------
        'Lecture des informations de la ligne
        '------------------------------------
        Call FillResult(TableauComplet, wsResultatTour, score_type, Z, L, Tour, genre)
        Z = Z + 1
    Next L
End Sub

Sub FillResult(ByRef TableauComplet, wsResultatTour, score_type, Z, CurrentLineIdx, Tour, genre)
    
    ColClub = wsResultatTour.Range("Club" & score_type).Column
    ColTour = wsResultatTour.Range("DebutTableauGeneral" & score_type).Column
    ColNom = wsResultatTour.Range("Nom" & score_type).Column
    ColSerie = wsResultatTour.Range("Serie" & score_type).Column
    ColIndex = wsResultatTour.Range("Index" & score_type).Column
    ColScore = wsResultatTour.Range("Score" & score_type).Column
    ColRang = wsResultatTour.Range("Rang" & score_type).Column
    colGenre = wsResultatTour.Range("Genre" & score_type).Column
    PremiereLigne = wsResultatTour.Range("DebutTableauGeneral" & score_type).Row
    
    ResultLineIdx = CurrentLineIdx + PremiereLigne
    'Call InitialiserTableaux
    'MsgBox "La valeur pour 'nom' est " & TableauCompletIdx("nom")

    resuilt_gender = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, colGenre), wsResultatTour.Cells(ResultLineIdx, colGenre))
    
    If (resuilt_gender = genre) Then
        If TableauCompletIdx Is Nothing Then
            InitialiserTableaux
        End If
        TableauComplet(Z, TableauCompletIdx("tour")) = Tour
        
        TableauComplet(Z, TableauCompletIdx("nom")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColNom), wsResultatTour.Cells(ResultLineIdx, ColNom))
        TableauComplet(Z, TableauCompletIdx("genre")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, colGenre), wsResultatTour.Cells(ResultLineIdx, colGenre))
        TableauComplet(Z, TableauCompletIdx("club")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColClub), wsResultatTour.Cells(ResultLineIdx, ColClub))
        
        TableauComplet(Z, TableauCompletIdx("index")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColIndex), wsResultatTour.Cells(ResultLineIdx, ColIndex))
        TableauComplet(Z, TableauCompletIdx("serie")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColSerie), wsResultatTour.Cells(ResultLineIdx, ColSerie))
        
        If score_type = "Net" Then
            TableauComplet(Z, TableauCompletIdx("rangNet")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColRang), wsResultatTour.Cells(ResultLineIdx, ColRang))
            TableauComplet(Z, TableauCompletIdx("scoreNet")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColScore), wsResultatTour.Cells(ResultLineIdx, ColScore))
        Else
            TableauComplet(Z, TableauCompletIdx("rangBrut")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColRang), wsResultatTour.Cells(ResultLineIdx, ColRang))
            TableauComplet(Z, TableauCompletIdx("scoreBrut")) = wsResultatTour.Range(wsResultatTour.Cells(ResultLineIdx, ColScore), wsResultatTour.Cells(ResultLineIdx, ColScore))
        End If
    End If
    
End Sub

Sub IntegrateTableauInSheet(TableauComplet, I, wsCumulJoueur, ListeJoueur, LignStartPlayer, NbJoueurs)
    '-----------------------------------------------------------------------
    'Lecture des informations compilees lors de la lecture des feuilles de score
    '-----------------------------------------------------------------------
    nom = TableauComplet(I, TableauCompletIdx("nom"))
    Serie = TableauComplet(I, TableauCompletIdx("serie"))
    Tour = TableauComplet(I, TableauCompletIdx("tour"))
    RangBrut = TableauComplet(I, TableauCompletIdx("rangBrut"))
    ScoreBrut = TableauComplet(I, TableauCompletIdx("scoreBrut"))
    RangNet = TableauComplet(I, TableauCompletIdx("rangNet"))
    ScoreNet = TableauComplet(I, TableauCompletIdx("scoreNet"))
    Club = TableauComplet(I, TableauCompletIdx("club"))
    index = TableauComplet(I, TableauCompletIdx("index"))
    genre = TableauComplet(I, TableauCompletIdx("genre"))

End Sub
