Sub EffacementImport()

Dim ColDebutTableau As Integer
Dim ColFinTableau As Integer
Dim LigDebutTableau As Integer
Dim NbLignes As Integer

Dim wkImport

    Set wkImport = Worksheets("Import Resultats Tour")

    LigDebutTableau = wkImport.Range("DebutTableauGeneralNet").Row + 1
    NbLignes = wkImport.Range("NbLignesNet")
    If NbLignes = 0 Then
        NbLignes = wkImport.Range("NbLignesBrut")
    End If
    ColDebutTableau = wkImport.Range("DebutTableauGeneralNet").Column
    ColFinTableau = wkImport.Range("GenreBrut").Column
    wkImport.Range(wkImport.Cells(LigDebutTableau, ColDebutTableau), wkImport.Cells(LigDebutTableau + NbLignes, ColFinTableau)).Clear
        
    Call recordToHistory("EffacementImport")
    
    RowDebutTableau = wkImport.Range("DebutTableauGeneralNet").Row
    wkImport.Cells(RowDebutTableau + 1, ColDebutTableau).Select

End Sub

Sub EffacementImportForced()

Dim ColDebutTableau As Integer
Dim ColFinTableau As Integer
Dim LigDebutTableau As Integer
Dim NbLignes As Integer

    LigDebutTableau = Range("DebutTableauGeneralNet").Row + 1
    NbLignes = 1000
    ColDebutTableau = Range("DebutTableauGeneralNet").Column
    ColFinTableau = Range("GenreBrut").Column
    Range(Cells(LigDebutTableau, ColDebutTableau), Cells(LigDebutTableau + NbLignes, ColFinTableau)).Select
    Selection.Clear
    
    Call recordToHistory("EffacementImportForced")
    
    RowDebutTableau = Range("DebutTableauGeneralNet").Row
    Cells(RowDebutTableau + 1, ColDebutTableau).Select

End Sub
Sub EffacementResultatAll()
'TODO use select tableau and loop
EffacementResultat ("Resultat LGS (HOMME)")
EffacementResultat ("Resultat LGS (DAME)")
End Sub

Sub EffacementResultat(Optional playerTypeSheetName As String)

Dim ColDebutTableau As Integer
Dim ColFinTableau As Integer
Dim ColFinTableauFormula As Integer
Dim LigDebutTableau As Integer
Dim NbLignes As Integer

Dim wkCumulJoueur

    'TODO reafctoring
    colEndClear = "AC1"
    colEndClearFormula = "AI1"
    
    Dim NomFeuilleCumulJoueur As String

    If (playerTypeSheetName = "") Then
        NomFeuilleCumulJoueur = Range("NomFeuilleCumuljoueur")
    Else
        NomFeuilleCumulJoueur = playerTypeSheetName
    End If
    Set wkCumulJoueur = Worksheets(NomFeuilleCumulJoueur)
    
    If (wkCumulJoueur.AutoFilterMode) Then
        wkCumulJoueur.AutoFilter.ShowAllData
        wkCumulJoueur.AutoFilteMode = False
    End If
    LigDebutInsertion = Worksheets(NomFeuilleCumulJoueur).Range("TableauResultat").Row + 1
    ''NbLignes = Range("NbLignesNet")
    NbLignes = 1000
    ColDebutTableau = Range("DebutTableauGeneralNet").Column
    'ColFinTableau = Range("GenreBrut").Column
    
    ColFinTableau = wkCumulJoueur.Range(colEndClear).Column
    ColFinTableauFormula = wkCumulJoueur.Range(colEndClearFormula).Column
        
    wkCumulJoueur.Range(wkCumulJoueur.Cells(LigDebutInsertion, ColDebutTableau), wkCumulJoueur.Cells(LigDebutInsertion + NbLignes, ColFinTableau)).ClearContents
    
    'Range(Cells(LigDebutInsertion, ColFinTableau + 1), Cells(LigDebutInsertion + NbLignes, ColFinTableauFormula)).Select
    'Selection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone

    wkCumulJoueur.Range(wkCumulJoueur.Cells(LigDebutInsertion, ColFinTableau + 1 + 2), wkCumulJoueur.Cells(LigDebutInsertion + NbLignes, ColFinTableauFormula)).ClearContents

    'Range(Cells(LigDebutInsertion, ColFinTableauFormula + 1), Cells(LigDebutInsertion + NbLignes, ColFinTableauFormula + 1 + 1)).Select
    'Selection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    
    Call recordToHistory("EffacementResultat", , NomFeuilleCumulJoueur)

End Sub

Sub EffacementAll()
    Call recordToHistory("EffacementAll")
    
    EffacementImport
    EffacementResultatAll
End Sub
