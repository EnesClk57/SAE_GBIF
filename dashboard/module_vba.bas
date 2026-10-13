' =============================================================================
' MODULE VBA - Dashboard Biodiversité Grand Est
' SAE 501 - Eser Enes & Marfouk Omar
' =============================================================================
' 
' Ce module VBA permet de :
' - Importer les données CSV actualisées
' - Rafraîchir les graphiques
' - Créer des filtres interactifs
'
' INSTRUCTIONS :
' 1. Ouvrir le fichier Excel généré (Dashboard_Biodiversite.xlsx)
' 2. Enregistrer sous : Dashboard_Biodiversite.xlsm (avec macros)
' 3. Appuyer sur Alt+F11 pour ouvrir l'éditeur VBA
' 4. Insérer > Module
' 5. Copier-coller ce code
' 6. Fermer l'éditeur VBA
' 7. Utiliser les macros via Développeur > Macros
'
' =============================================================================

Option Explicit

' Chemin vers les fichiers CSV exportés
Const DATA_PATH As String = "C:\SAE_GBIF\data\export\"

' =============================================================================
' MACRO PRINCIPALE : Rafraîchir toutes les données
' =============================================================================
Sub RafraichirDonnees()
    Application.ScreenUpdating = False
    
    MsgBox "Début de l'import des données...", vbInformation, "Dashboard Biodiversité"
    
    ' Importer les fichiers CSV
    Call ImporterIndicateurs
    Call ImporterEspeces
    Call ImporterCroisementGES
    
    Application.ScreenUpdating = True
    
    MsgBox "Données importées avec succès !" & vbCrLf & vbCrLf & _
           "Les graphiques ont été mis à jour.", vbInformation, "Dashboard Biodiversité"
End Sub

' =============================================================================
' Import du fichier indicateurs_epci_annee.csv
' =============================================================================
Sub ImporterIndicateurs()
    Dim ws As Worksheet
    Dim fichier As String
    
    fichier = DATA_PATH & "indicateurs_epci_annee.csv"
    
    ' Vérifier si le fichier existe
    If Dir(fichier) = "" Then
        MsgBox "Fichier non trouvé : " & fichier, vbExclamation
        Exit Sub
    End If
    
    ' Créer ou récupérer la feuille
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Indicateurs EPCI")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Indicateurs EPCI"
    End If
    On Error GoTo 0
    
    ' Effacer le contenu existant
    ws.Cells.Clear
    
    ' Importer le CSV
    With ws.QueryTables.Add(Connection:="TEXT;" & fichier, Destination:=ws.Range("A3"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    ' Titre
    ws.Range("A1").Value = "Indicateurs de biodiversité par EPCI et par année"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    ' Mettre en forme les en-têtes
    Call FormaterEnTetes(ws, 3)
End Sub

' =============================================================================
' Import du fichier liste_especes.csv
' =============================================================================
Sub ImporterEspeces()
    Dim ws As Worksheet
    Dim fichier As String
    
    fichier = DATA_PATH & "liste_especes.csv"
    
    If Dir(fichier) = "" Then
        MsgBox "Fichier non trouvé : " & fichier, vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Top Espèces")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Top Espèces"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & fichier, Destination:=ws.Range("A3"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    ws.Range("A1").Value = "Top des espèces observées (2018-2022)"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    Call FormaterEnTetes(ws, 3)
End Sub

' =============================================================================
' Import du fichier croisement_biodiv_ges.csv
' =============================================================================
Sub ImporterCroisementGES()
    Dim ws As Worksheet
    Dim fichier As String
    
    fichier = DATA_PATH & "croisement_biodiv_ges.csv"
    
    If Dir(fichier) = "" Then
        MsgBox "Fichier non trouvé : " & fichier, vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Croisement GES")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Croisement GES"
    End If
    On Error GoTo 0
    
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & fichier, Destination:=ws.Range("A3"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    ws.Range("A1").Value = "Croisement Biodiversité × Émissions GES"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    Call FormaterEnTetes(ws, 3)
End Sub

' =============================================================================
' Formater les en-têtes d'un tableau
' =============================================================================
Sub FormaterEnTetes(ws As Worksheet, ligneEntete As Integer)
    Dim derniereColonne As Integer
    Dim i As Integer
    
    derniereColonne = ws.Cells(ligneEntete, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To derniereColonne
        With ws.Cells(ligneEntete, i)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(27, 77, 62)  ' Vert foncé
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' Ajuster largeur des colonnes
    ws.Columns("A:" & Chr(64 + derniereColonne)).AutoFit
End Sub

' =============================================================================
' Créer un tableau croisé dynamique pour analyse
' =============================================================================
Sub CreerTableauCroiseDynamique()
    Dim wsSource As Worksheet
    Dim wsTCD As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim plage As Range
    Dim derniereLigne As Long
    
    Set wsSource = ThisWorkbook.Sheets("Indicateurs EPCI")
    derniereLigne = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    Set plage = wsSource.Range("A3:G" & derniereLigne)
    
    ' Créer une nouvelle feuille pour le TCD
    On Error Resume Next
    Set wsTCD = ThisWorkbook.Sheets("Analyse TCD")
    If wsTCD Is Nothing Then
        Set wsTCD = ThisWorkbook.Sheets.Add
        wsTCD.Name = "Analyse TCD"
    Else
        wsTCD.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Créer le cache
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=plage)
    
    ' Créer le tableau croisé dynamique
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=wsTCD.Range("A3"), _
        TableName:="TCD_Biodiversite")
    
    ' Configurer le TCD
    With pt
        .PivotFields("year").Orientation = xlColumnField
        .PivotFields("lib_zone").Orientation = xlRowField
        .AddDataField .PivotFields("richesse_specifique"), "Moyenne Richesse", xlAverage
    End With
    
    wsTCD.Range("A1").Value = "Tableau Croisé Dynamique - Richesse par EPCI et Année"
    wsTCD.Range("A1").Font.Size = 14
    wsTCD.Range("A1").Font.Bold = True
    
    MsgBox "Tableau croisé dynamique créé dans la feuille 'Analyse TCD'", vbInformation
End Sub

' =============================================================================
' Filtrer les données par année
' =============================================================================
Sub FiltrerParAnnee()
    Dim annee As String
    Dim ws As Worksheet
    
    annee = InputBox("Entrez l'année à filtrer (2018-2022) :", "Filtre par année", "2022")
    
    If annee = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Indicateurs EPCI")
    
    ' Supprimer les filtres existants
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' Appliquer le filtre
    ws.Range("A3").AutoFilter Field:=2, Criteria1:=annee
    
    MsgBox "Filtre appliqué : année " & annee, vbInformation
End Sub

' =============================================================================
' Supprimer tous les filtres
' =============================================================================
Sub SupprimerFiltres()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Next ws
    
    MsgBox "Tous les filtres ont été supprimés.", vbInformation
End Sub
