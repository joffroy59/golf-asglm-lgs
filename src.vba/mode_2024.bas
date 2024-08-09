Function Mode_NomFeuilleCumulJoueur(genre As String)
    If (ModeExport = Range("Z23")) Then
        If genre = Range("X19") Then
            Mode_NomFeuilleCumulJoueur = Range("Z19")
        ElseIf genre = Range("X20") Then
            Mode_NomFeuilleCumulJoueur = Range("Z20")
        End If
    Else
        Mode_NomFeuilleCumulJoueur = Range("NomFeuilleCumuljoueur")
    End If
End Function