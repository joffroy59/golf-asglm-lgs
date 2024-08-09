Sub ResetCellActive()
    Range("A1:A1").Select
End Sub

Function IsNet(ScoreType)
    IsNet = (ScoreType = "Net")
End Function

Function IsBrut(ScoreType)
    IsBrut = (ScoreType = "Brut")
End Function

Sub fixIndexToNumber(resultLine, colIndex)
    Range(Cells(resultLine, colIndex), Cells(resultLine, colIndex)).Select
    Selection.TextToColumns Destination:=Range(Cells(resultLine, colIndex), Cells(resultLine, colIndex)), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
End Sub

Sub InsertDataImported(TableauJoueurs As Variant, TableauJoueursIdx, scoreCount)
    ' MsgBox "Insert Data in First table"
    
    PremiereLigneNet = Range("DebutTableauGeneralNet").Row + Range("NbLignesNet").Value
    PremiereColonneNet = Range("DebutTableauGeneralNet").Column
    ColIndexNet = Range("ColIndexNet").Column
    
    PremiereLigneBrut = Range("DebutTableauGeneralBrut").Row + Range("NbLignesBrut").Value
    PremiereColonneBrut = Range("DebutTableauGeneralBrut").Column
    ColIndexBrut = Range("ColIndexBrut").Column
    
    PremiereLigneNetCurrent = PremiereLigneNet
    PremiereLigneBrutCurrent = PremiereLigneBrut
    
    ResetCellActive
    
    
    ' TO TEST --> OK mais ca parcours meme les element "vide"
    'For I = LBound(TableauJoueurs, 1) To UBound(TableauJoueurs, 1)
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
        
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("tour"))
        Increment ResultCol
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("rang"))
        Increment ResultCol
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("name"))
        Increment ResultCol
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("club"))
        Increment ResultCol
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("index"))
        colIndexCurrent = ResultCol
        Increment ResultCol
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("serie"))
        Increment ResultCol
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("score"))
        Increment ResultCol
        Range(Cells(resultLine, ResultCol), Cells(resultLine, ResultCol)) = TableauJoueurs(I, TableauJoueursIdx("genre"))
         
        '--------------------------------------------------------------------------
        'Transformation des index en nombres (en format texte la feuille de score )
        '--------------------------------------------------------------------------
        fixIndexToNumber resultLine, colIndexCurrent
    Next I
    
    ResetCellActive

End Sub







