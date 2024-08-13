Function GetValueByColumnName(sheet As Worksheet, columnName As String, rowNumber As Integer) As Variant
    Dim colNumber As Integer
    Dim colRange As Range
    
    ' Recherche le numéro de colonne par son nom dans la première ligne
    Set colRange = sheet.Rows(1).Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not colRange Is Nothing Then
        colNumber = colRange.Column
        ' Récupérer la valeur dans la ligne spécifiée et la colonne trouvée
        GetValueByColumnName = sheet.Cells(rowNumber, colNumber).Value
    Else
        GetValueByColumnName = "Colonne non trouvée"
    End If
End Function

Function GetSheetExportByGenre(genre) As String
    If genre = Range("X19") Then
        GetSheetExportByGenre = Range("Z19")
    ElseIf genre = Range("X20") Then
        GetSheetExportByGenre = Range("Z20")
    End If
    
    
End Function

Sub AppliquerBordures(wsCumulJoueur, LigneScore, ColNom, EndCol)
    Dim rng As Range
    
    Set rng = wsCumulJoueur.Range(wsCumulJoueur.Cells(LigneScore, ColNom), wsCumulJoueur.Cells(LigneScore, EndCol))
    
    ' Supprimer les bordures diagonales
    With rng.Borders(xlDiagonalDown)
        .LineStyle = xlNone
    End With
    With rng.Borders(xlDiagonalUp)
        .LineStyle = xlNone
    End With
    
    ' Appliquer les bordures aux bords extérieurs et aux lignes internes
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Sub setColorConditional(rng As Range, isBold As Boolean)
    Dim ColorInProgress As Long
    Dim ColorDone As Long
    
    ' Définir les couleurs
    ColorInProgress = RGB(255, 0, 0) ' Rouge
    ColorDone = RGB(72, 148, 31) ' Vert

    ' Supprimer toutes les mises en forme conditionnelles existantes
    rng.FormatConditions.Delete
    
    ' Ajouter une nouvelle condition pour les cellules contenant "En cours"
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""En cours""")
        .Font.Bold = isBold
        .Font.Color = ColorInProgress
    End With
    
    ' Ajouter une nouvelle condition pour les cellules ne contenant pas "En cours"
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""En cours""")
        .Font.Bold = isBold
        .Font.Color = ColorDone
    End With
End Sub

Sub setFormula(formula)
    'Debug.Print "formula 1: " & Selection.Formula
    'Debug.Print "formula 1: " & Selection.FormulaR1C1
    
    Selection.FormulaR1C1 = formula
    
    'Debug.Print "formula 3: " & Selection.Formula
    'Debug.Print "formula 3: " & Selection.FormulaR1C1
End Sub


