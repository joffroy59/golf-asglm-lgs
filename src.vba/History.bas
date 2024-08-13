Sub recordToHistory(TaskType As String, Optional reference As String = "Nan", Optional sheet As String = "Nan", Optional playerType As String = "Nan")
    Dim historySheetName As String
    Dim historyTableName As String
    historySheetName = "Historique Import"
    historyTableName = "Tableau1"
    If (playerType = "") Then
        playerType = Range("playerType")
    End If
    
    If (sheet = "") Then
        sheet = Range("NomFeuilleCumuljoueur")
    End If
    
    Set historySheet = Worksheets(historySheetName)
    Set tbl = historySheet.ListObjects(historyTableName)
    
    Debug.Print "write to history: " & TaskType & ", " & reference
    Set newrow = tbl.ListRows.Add
    With newrow
        .Range(1) = TaskType
        .Range(2) = playerType
        .Range(3) = sheet
        .Range(4) = reference
        .Range(5) = Now
    End With
    'End
End Sub

Function getHistorySheet() As Worksheet
    ' DOES NOT WORK
    Dim historySheetName As String
    Dim historyTableName As String
    historyTableName = "Tableau1"
    historySheetName = "Historique Import"
    
    Set historySheet = Worksheets(historySheetName)

    getHistorySheet = historySheet.ListObjects(historyTableName)
End Function

Sub ClearHistory()
    Dim historySheetName As String
    Dim historyTableName As String
    historySheetName = "Historique Import"
    historyTableName = "Tableau1"
    
    Set historySheet = Worksheets(historySheetName)
    Set tbl = historySheet.ListObjects(historyTableName)
    Call ClearTable(tbl)
End Sub

Sub ClearTable(ByVal tbl As ListObject)
    With tbl
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
End Sub