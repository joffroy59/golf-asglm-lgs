
Sub readAll(ByRef TableauComplet, score_type, Z, I, PremiereLigne, Tour, genre, NbLignes, ColNom, ColSerie, ColRang, ColScore, ColClub, colIndex, colGenre)
    For I = 1 To NbLignes
        '------------------------------------
        'Lecture des informations de la ligne
        '------------------------------------
        FillAllResult TableauComplet, score_type, Z, I, PremiereLigne, Tour, genre, ColNom, ColSerie, ColRang, ColScore, ColClub, colIndex, colGenre
        Z = Z + 1
    Next I
End Sub

Sub FillAllResult(ByRef TableauComplet, score_type, Z, I, PremiereLigne, Tour, genre, ColNom, ColSerie, ColRang, ColScore, ColClub, colIndex, colGenre)
    
    resuilt_gender = Range(Cells(I + PremiereLigne, colGenre), Cells(I + PremiereLigne, colGenre))
    
    If (resuilt_gender = genre) Then
        TableauComplet(2, Z) = Tour
        
        TableauComplet(0, Z) = Range(Cells(I + PremiereLigne, ColNom), Cells(I + PremiereLigne, ColNom))
        TableauComplet(7, Z) = Range(Cells(I + PremiereLigne, ColClub), Cells(I + PremiereLigne, ColClub))
        TableauComplet(8, Z) = Range(Cells(I + PremiereLigne, colIndex), Cells(I + PremiereLigne, colIndex))
        TableauComplet(1, Z) = Range(Cells(I + PremiereLigne, ColSerie), Cells(I + PremiereLigne, ColSerie))
        TableauComplet(9, Z) = Range(Cells(I + PremiereLigne, colGenre), Cells(I + PremiereLigne, colGenre))
        
        If score_type = "Net" Then
            TableauComplet(5, Z) = Range(Cells(I + PremiereLigne, ColRang), Cells(I + PremiereLigne, ColRang))
            TableauComplet(6, Z) = Range(Cells(I + PremiereLigne, ColScore), Cells(I + PremiereLigne, ColScore))
            TableauComplet(9, Z) = Range(Cells(I + PremiereLigne, colGenre), Cells(I + PremiereLigne, colGenre))
        Else
            TableauComplet(3, Z) = Range(Cells(I + PremiereLigne, ColRang), Cells(I + PremiereLigne, ColRang))
            TableauComplet(4, Z) = Range(Cells(I + PremiereLigne, ColScore), Cells(I + PremiereLigne, ColScore))
        End If
    End If
    
End Sub

