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

Public Function GetFolder(Optional Folder As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    If Folder <> "" Then GoTo NextCode2
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
    Exit Function
NextCode2:
    GetFolder = Folder
    Set fldr = Nothing
    Exit Function
End Function

