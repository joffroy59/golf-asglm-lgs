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