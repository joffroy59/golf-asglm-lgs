Public ModeExport As String


Public Sub ImporterBrutNet_One_File(inputFile As String, Tour As Integer, Optional Clean As Boolean = False, Optional CleanResult As Boolean = False)
    'Worksheets("Import Resultats Tour").Activate
    If Path_Exists(inputFile) Then
        Call ImporterBrutNetFromFiles(inputFile, "test", Tour, False)
    End If
End Sub


Sub ImporterBrutNetFromFiles(NomFichier As String, TaskType As String, Tour As Integer, Optional CleanResult As Boolean = False)
    Dim task As String
    task = "Importation d ''un fichier Brut et Net (complet Homme Dame) FFGolf pour 1 Tour (2024) [manual]" & " (Clean import =" & CleanResult & ")"
    NomFichierBrut = processGolfMatchSheetFromFile(NomFichier, task, Tour)
    
    ' Call CalculTour(tour, CleanResult)
    
    ' Worksheets("Import Resultats Tour").Activate

End Sub

Sub extract()

    ModeExport = "XLS_2024"
    
    Dim inputFilename As String
    
    ' Lecture du nom de la feuille Excel à partir de la cellule T3
    inputFilename = Range("T3").Value

    Call ImporterBrutNet_One_File(inputFilename, 1)
End Sub


Sub extract1tour()

    Call ImporterBrutNet_2024

End Sub

Public Sub ImporterBrutNet_2024(Optional TaskType As String = "", Optional inputFile As String = "", Optional Tour As Integer, Optional CleanImport As Boolean = False)
    'TODO refactoring
    Worksheets("Import Resultats Tour").Activate
    If Path_Exists(inputFile) Then
        If CleanImport Then
            EffacementImport
        End If
        If TaskType = "" Then
            TaskType = "Importation d ''un fichier Brut et Net (complet Homme Dame) FFGolf pour 1 Tour (2024) [auto]" & " (Clean import =" & CleanImport & ")"
        End If
        
        
        Call ImporterBrutNetFromFiles_2024(inputFile, TaskType, Tour)
    End If
End Sub


Sub CalculTour_2024(Optional Tour As Integer)
    Dim CleanResult As Boolean
    CleanResult = Range("cleanResult").Value
    
    Call CalculTour(Tour, CleanResult)
    
    Worksheets("Import Resultats Tour").Activate
End Sub

Sub ImporterBrutNetFromFiles_2024(NomFichierExportOrNull As String, TaskType As String, Optional Tour As Integer)
        Dim ExportFilename As String
        Dim NomFichierNet As String
    ExportFilename = "" & NomFichierExportOrNull
    ExportFilename = processGolfMatchSheetFromFile(NomFichierExportOrNull, TaskType, Tour)
    
    Call CalculTour_2024(Tour)
    
    Worksheets("Import Resultats Tour").Activate
End Sub
' TODO remove doublon
Public Sub setGender(gender As Integer)
    Worksheets("Import Resultats Tour").Activate
    
    Range("F13").Value = gender

End Sub

Public Sub GetScoresFromFFGolfHommeDame_2024(Optional scoreFolder As String, Optional ByVal Clean As Boolean = True, Optional TaskType As String = "Importation et generation de tous les tours depuis un repertoire HOMME/DAME")
    Index_All = 3
    setGender (Index_All)
    Call GetScoresFromFFGolf_2024(scoreFolder, Clean, TaskType)
    Call EffacementImportForced
End Sub


Public Sub GetScoresFromFFGolf_2024(Optional scoreFolder As String, Optional ByVal Clean As Boolean = True, Optional TaskType As String = "Importation et generation de tous les tours depuis un repertoire")
    'TODO use global
    NbTour = 7
    iTourFinal = 7
    tourFolderPatternPrefix = "T"
    tourFolderFinale = "Finale"
    ShowMissingFolder = True
    
    Dim scoreFolderRoot As String
    scoreFolderRoot = scoreFolder
    
    If scoreFolderRoot = "" Then
        scoreFolderRoot = GetScoreFolder("")
    End If
    scoreFolder = scoreFolderRoot
    
    Call recordToHistory(TaskType, scoreFolderRoot)
    
    'Dim NomFichierBrut As String
    Dim Tour As Integer
    Dim scoreFolderTour As String
    For itour = 1 To NbTour
        If itour = iTourFinal Then
            scoreFolderTour = scoreFolderRoot & "\" & tourFolderFinale
        Else
            scoreFolderTour = scoreFolderRoot & "\" & tourFolderPatternPrefix & itour
        End If
        
        If Path_Exists(scoreFolderTour) Then
            Dim fileList(1) As String
            ' TODO
            fileList(0) = scoreFolderTour & "\" & "2d. Extraction XLS globale.xls"
            
            Tour = itour
            Range("cleanResult").Value = False
            Call EffacementImportForced
            Call ImporterBrutNet_2024(TaskType, fileList(0), Tour, Clean)
        Else
            If Not (scoreFolderTour Like "") Or Not (scoreFolderTour Like "*\T" & itour) Then
                MessageErreur = "Vous n'avez pas sélectionné de repertoire contenant les repertoire T1, .. T6, Finale. Fin de la procédure"
                I = MsgBox(MessageErreur, vbOKOnly, "Import des résultats de tous les tours")
                End
            End If
        
            If ShowMissingFolder Then
                MsgBox scoreFolderTour & " n'existe pas"
            End If
        End If
    Next itour
End Sub


