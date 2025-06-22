Public Sub CheckDirectories(basePath As String, sheetName As String)
    Dim provider As String
    Dim ws As Worksheet
    Dim row As Integer
    Dim folderPath As String
    Dim s1Path As String
    Dim s2Path As String
    Dim pcaPath As String
    Dim fso As Object
    Dim folder As Object
    Dim found As Boolean
    
    ' Créer un objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Vérifier si la feuille existe déjà
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' Si la feuille n'existe pas, la créer
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
        
        ' En-têtes
        ws.Cells(1, 1).Value = "Prestataire"
        ws.Cells(1, 2).Value = "PCA"
        ws.Cells(1, 3).Value = "S1"
        ws.Cells(1, 4).Value = "Date_Envoi"
        ws.Cells(1, 5).Value = "Date_Relance1"
        ws.Cells(1, 6).Value = "Date_Relance2"
        ws.Cells(1, 7).Value = "Date_Relance3"
        ws.Cells(1, 8).Value = "S2"
        ws.Cells(1, 9).Value = "Date_Envoi(S2)"
        ws.Cells(1, 10).Value = "Date_Relance1(S2)"
        ws.Cells(1, 11).Value = "Date_Relance2(S2)"
        ws.Cells(1, 12).Value = "Date_Relance3(S2)"
    End If
    
    ' Effacer les anciennes données (sauf les en-têtes)
    ws.Rows("2:" & ws.Rows.Count).ClearContents
    
    ' Initialiser la ligne
    row = 2
    
    ' Parcourir chaque prestataire
    For Each folder In fso.GetFolder(basePath).SubFolders
        provider = folder.Name
        
        ' Chemins des sous-répertoires
        pcaPath = folder.Path & "\PCA"
        s1Path = folder.Path & "\Qualité\S1"
        s2Path = folder.Path & "\Qualité\S2"
        
        ' Vérifier la présence de fichiers
        found = False
        If Not IsEmpty(provider) Then
            ' Vérifier si le prestataire existe déjà dans la feuille
            For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
                If ws.Cells(i, 1).Value = provider Then
                    found = True
                    Exit For
                End If
            Next i
            
            ' Si le prestataire n'existe pas, ajouter une nouvelle ligne
            If Not found Then
                ' Écrire le nom du prestataire et ajouter un hyperlien
                ws.Cells(row, 1).Value = provider
                ws.Hyperlinks.Add Anchor:=ws.Cells(row, 1), Address:=folder.Path, TextToDisplay:=provider
                
                ws.Cells(row, 2).Value = IIf(fso.FolderExists(pcaPath) And fso.GetFolder(pcaPath).Files.Count > 0, 1, 0)
                ws.Cells(row, 3).Value = IIf(fso.FolderExists(s1Path) And fso.GetFolder(s1Path).Files.Count > 0, 1, 0)
                ' Ajoutez ici les valeurs pour les dates si nécessaire
                ws.Cells(row, 8).Value = IIf(fso.FolderExists(s2Path) And fso.GetFolder(s2Path).Files.Count > 0, 1, 0)
                ' Ajoutez ici les valeurs pour les dates S2 si nécessaire
                row = row + 1
            Else
                ' Si le prestataire existe, mettre à jour les valeurs
                ws.Cells(i, 2).Value = IIf(fso.FolderExists(pcaPath) And fso.GetFolder(pcaPath).Files.Count > 0, 1, 0)
                ws.Cells(i, 3).Value = IIf(fso.FolderExists(s1Path) And fso.GetFolder(s1Path).Files.Count > 0, 1, 0)
                ws.Cells(i, 8).Value = IIf(fso.FolderExists(s2Path) And fso.GetFolder(s2Path).Files.Count > 0, 1, 0)
                ' Mettez à jour les dates S2 si nécessaire
            End If
        End If
    Next folder
    
    ' Libérer les objets
    Set fso = Nothing
    MsgBox "Mise à jour terminée pour " & sheetName & " !"
End Sub

Public Sub CheckDirectories2023()
    CheckDirectories "P:\RE\COMMUN\Suivi_qualite_PSEE\2023", "2023"
End Sub

Public Sub CheckDirectories2024()
    CheckDirectories "P:\RE\COMMUN\Suivi_qualite_PSEE\2024", "2024"
End Sub

Public Sub CheckDirectories2025()
    CheckDirectories "P:\RE\COMMUN\Suivi_qualite_PSEE\2025", "2025"
End Sub
