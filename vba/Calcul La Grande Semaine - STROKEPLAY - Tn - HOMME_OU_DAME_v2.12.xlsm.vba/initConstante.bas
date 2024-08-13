Public TableauCompletIdx As Object

Public FormulaBestNett As String
Public FormulaBestBrut As String
Public FormulaTotalNett As String
Public FormulaTotalBrut As String

Public modeExport As String
Public NbTour As Integer
Public TourFinalndex As Integer

 Public TourFolderPatternPrefix As String
 Public TourFolderFinale As String
 Public ShowMissingFolder As Boolean
 
Sub Auto_Open()
    InitAll
    modeExport = "XLS_2024"
    NbTour = 7
    TourFinalndex = NbTour
    
    TourFolderPatternPrefix = "T"
    TourFolderFinale = "Finale"
    
    ShowMissingFolder = True
End Sub

Sub InitAll()
    InitialiserTableaux
End Sub

'------------------------------------------------------------------------
Sub InitialiserTableaux()
    Call InitialiserTableauJoueursIdx
    Call InitialiserTableauCompletIdx
End Sub

Sub InitialiserTableauJoueursIdx()
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
End Sub

Sub InitialiserTableauCompletIdx()
    Set TableauCompletIdx = CreateObject("Scripting.Dictionary")
    TableauCompletIdx.Add "nom", 0
    TableauCompletIdx.Add "serie", 1
    TableauCompletIdx.Add "tour", 2
    TableauCompletIdx.Add "rangNet", 3
    TableauCompletIdx.Add "scoreNet", 4
    TableauCompletIdx.Add "rangBrut", 5
    TableauCompletIdx.Add "scoreBrut", 6
    TableauCompletIdx.Add "club", 7
    TableauCompletIdx.Add "index", 8
    TableauCompletIdx.Add "genre", 9
End Sub