Sub ImportFileStaticOrAsk()
    Dim inputFilename As String
    Dim task As String
    
    ' Lecture du nom de la feuille Excel à partir de la cellule T3
    inputFilename = Range("T3").Value

    If Path_Exists(inputFilename) Then
        task = "Importation d ''un fichier Brut et Net (complet Homme Dame) FFGolf pour 1 Tour (2024) [manual]" & " (Clean import =" & CleanResult & ")"
        NomFichierBrut = processGolfMatchSheetFromFile(inputFilename, task, 1)
    End If
End Sub

Public Sub ImportExportFileForAllTour(Optional ScoreFolder As String, Optional ByVal clean As Boolean = True, Optional TaskType As String = "Importation et generation de tous les tours depuis un repertoire HOMME/DAME")
    Index_All = 3
    setGender (Index_All)
    Call ImporterBrutNetForAllTourFromFolder(ScoreFolder, clean, TaskType)
    Call EffacementImportForced
End Sub

Sub ImportExportFileForTour()
    Call ImporterBrutNetForTour
End Sub

Public Sub setGender(gender As Integer)
    Worksheets("Import Resultats Tour").Range("F13").Value = gender
End Sub

Public Sub ImporterBrutNetForAllTourFromFolder(Optional ScoreFolder As String, Optional ByVal clean As Boolean = True, Optional TaskType As String = "Importation et generation de tous les tours depuis un repertoire")
    Dim Tour As Integer
    
    Dim ScoreFolderTour As String
    Dim ScoreFolderRoot As String
    
    Dim fileList(1) As String
    
    ScoreFolderRoot = GetScoreRootFolder(ScoreFolder)
    
    Call recordToHistory(TaskType, ScoreFolderRoot)
    
    For tour_index = 1 To NbTour
        Tour = tour_index

        ScoreFolderTour = GetScoreTourFolder(ScoreFolderRoot, tour_index)
        
        If Path_Exists(ScoreFolderTour) Then
            fileList(0) = ScoreFolderTour & "\" & "2d. Extraction XLS globale.xls"
            
            ForceCleanResult
            
            Call EffacementImportForced
            Call ImporterBrutNetForTour(TaskType, fileList(0), Tour, clean)
        Else
            Call ErrorFolderTourNotExist(ScoreFolderTour, tour_index)
        End If
    Next tour_index
End Sub

Function GetScoreRootFolder(ScoreFolder)
    ScoreFolderRoot = ScoreFolder
    
    If ScoreFolderRoot = "" Then
        ScoreFolderRoot = GetFolder("")
    End If
    
    GetScoreRootFolder = ScoreFolderRoot
End Function

Function GetScoreTourFolder(ScoreFolderRoot, tour_index)
    If tour_index = TourFinalndex Then
        ScoreFolderTour = ScoreFolderRoot & "\" & TourFolderFinale
    Else
        ScoreFolderTour = ScoreFolderRoot & "\" & TourFolderPatternPrefix & tour_index
    End If
    
    GetScoreTourFolder = ScoreFolderTour
End Function

Function ErrorFolderTourNotExist(ScoreFolderTour, tour_index)
    If Not (ScoreFolderTour Like "") Or Not (ScoreFolderTour Like "*\" & TourFolderPatternPrefix & tour_index) Then
        MessageErreur = "Vous n'avez pas sélectionné de repertoire contenant les repertoire T1, .. T6, Finale. Fin de la procédure:"
        I = MsgBox(MessageErreur, vbOKOnly, "Import des résultats de tous les tours")
        End
    End If

    If ShowMissingFolder Then
        MsgBox ScoreFolderTour & " n'existe pas"
    End If
End Function


Function ForceCleanResult()
    Range("cleanResult").Value = False
End Function

Sub IntegrateScoreByGenre(genre As String, CleanResult As Boolean, Tour, wsResultatTour)
    Dim TableauComplet(1000, 10) As Variant
    
    Dim NomFeuilleCumulJoueur As String
    Dim wsCumulJoueur As Worksheet
    
    Dim LigneScore As Long
    Dim ColNom As Long
    Dim EndCol As Long
    
    Dim rng As Range
    
    NomFeuilleCumulJoueur = Mode_NomFeuilleCumulJoueur(genre)
    
    Call recordToHistory("Calcul du Tour " & Tour, , NomFeuilleCumulJoueur, "ALL - " & genre)
        
    If CleanResult Then
        Call EffacementResultat(GetSheetExportByGenre(genre))
    End If
    
    wsResultatTour.AutoFilterMode = False
    
    Z = 1
    
    scoreTypes = Array("Net", "Brut")
    For I = LBound(scoreTypes) To UBound(scoreTypes)
        '------------------------------------------------------------------------
        'Lecture des informations de la feuille Import Resultats Tour
        '------------------------------------------------------------------------
        Call FillAllResult(TableauComplet, wsResultatTour, scoreTypes(I), Z, Tour, genre)
    Next I
    
    '-------------------------------
    'Mise a jour du cumul individuel
    '-------------------------------
    Set wsCumulJoueur = Worksheets(NomFeuilleCumulJoueur)
    
    '-------------------------------------------
    'Constitution de la liste des joueurs deja existants
    '-------------------------------------------
    debutTableau = wsCumulJoueur.Range("TableauResultat").Row
    LignStartPlayer = debutTableau + 1
    lastr_row = GetLastRowFrom(wsCumulJoueur, LignStartPlayer)
    NbJoueurs = lastr_row - (LignStartPlayer)
    
    ColNom = wsCumulJoueur.Range("TableauResultat").Column
    For I = LignStartPlayer To LignStartPlayer + NbJoueurs
        nom = wsCumulJoueur.Range(wsCumulJoueur.Cells(I, ColNom), wsCumulJoueur.Cells(I, ColNom))
        If InStr(ListeJoueur, nom) = 0 Then
            ListeJoueur = ListeJoueur + Format(I, "0000") + " " + nom + "/ "
        End If
    Next I
    
    '------------------------------
    'Insertion des joueurs et des scores
    '------------------------------
    offsetTour = (Tour - 1) * 4
    startResultatCol = offsetTour + ColNom + 4
    
    NbTour = 7
    'Finale
    If (Tour = NbTour) Then
        offsetTour = offsetTour + 2
        startResultatCol = offsetTour + ColNom + 4
    End If
    
    EndCol = wsCumulJoueur.Range("TableauResultatEnd").Column
    bestTourNettCol = wsCumulJoueur.Range("TableauResultatMaxNet").Column
    bestTourBrutCol = wsCumulJoueur.Range("TableauResultatMaxBrut").Column
    totalTourNettCol = wsCumulJoueur.Range("TableauResultatTotalNet").Column
    totalTourBrutCol = wsCumulJoueur.Range("TableauResultatTotalBrut").Column
    
    For I = 1 To Z - 1
  
        '-----------------------------------------------------------------------
        'Lecture des informations compilees lors de la lecture des feuilles de score
        '-----------------------------------------------------------------------
        nom = TableauComplet(I, TableauCompletIdx("nom"))
        Serie = TableauComplet(I, TableauCompletIdx("serie"))
        Tour = TableauComplet(I, TableauCompletIdx("tour"))
        RangBrut = TableauComplet(I, TableauCompletIdx("rangBrut"))
        ScoreBrut = TableauComplet(I, TableauCompletIdx("scoreBrut"))
        RangNet = TableauComplet(I, TableauCompletIdx("rangNet"))
        ScoreNet = TableauComplet(I, TableauCompletIdx("scoreNet"))
        Club = TableauComplet(I, TableauCompletIdx("club"))
        index = TableauComplet(I, TableauCompletIdx("index"))
        genre = TableauComplet(I, TableauCompletIdx("genre"))
        
        '-----------------------------------------------------------------------
        'set colum for results
        '-----------------------------------------------------------------------
        
        If Tour > 0 Then
            LigneScore = GetPlayerScoreLine(wsCumulJoueur, ListeJoueur, nom, LignStartPlayer, NbJoueurs, ColNom, EndCol)
            
            '---------------------------------------
            'Insertion des informations dans le tableau
            '---------------------------------------
            Call InsertScoreForPlayer(wsCumulJoueur, LigneScore, startResultatCol, ColNom, nom, Club, index, Serie, ScoreBrut, RangBrut, ScoreNet, RangNet)
            
            '---------------------------------------
            'Insertion des formules de calcul du meilleur score et Total de la semaine
            '---------------------------------------
            playerType = Range("PlayerType")
            Call InsertFormaulasBestScore(wsCumulJoueur, playerType, LigneScore, bestTourNettCol, bestTourBrutCol, totalTourNettCol, totalTourBrutCol)

        End If
    Next I
End Sub
