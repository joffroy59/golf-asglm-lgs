Function Path_Exists(Path As String)

'Dim Path As String
Dim Folder As String
Dim Answer As VbMsgBoxResult
    'Path = "C:\Users\LG\Desktop\VBA\S12"
    Folder = Dir(Path, vbDirectory)
    If Folder = vbNullString Then
        Path_Exists = False
    Else
        Path_Exists = True
    End If
End Function

Sub Increment(ByRef var, Optional amount = 1)
    var = var + amount
End Sub

Function GetPrefixSerie(Serie)
    GetPrefixSerie = "Serie " & Serie
End Function