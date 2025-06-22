' =============================================
' MODULE 1: DashboardModule
' =============================================

Option Explicit

' Variables globales
Public Const DASHBOARD_SHEET As String = "Dashboard"
Public Const DATA_START_ROW As Integer = 7
Public Const DATA_START_COL As Integer = 2

' Structure pour les métriques
Type MetricInfo
    Name As String
    dataColumn As String
    etpColumn As String
    Unit As String
End Type

' =============================================
' FONCTION PRINCIPALE - À APPELER POUR CRÉER LE DASHBOARD
' =============================================
Sub CreateDashboard()
    Dim ws As Worksheet
    
    ' Créer ou récupérer la feuille Dashboard
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = DASHBOARD_SHEET
    Else
        ws.Cells.Clear
    End If
    
    ' Créer l'interface du dashboard
    Call SetupDashboardLayout(ws)
    
    ' Initialiser avec le mois courant
    Call UpdateDashboard
    
    MsgBox "Dashboard créé avec succès!", vbInformation
End Sub

' =============================================
' MISE EN PAGE DU DASHBOARD
' =============================================
Sub SetupDashboardLayout(ws As Worksheet)
    With ws
        ' En-tête principal
        .Range("B2:J2").Merge
        .Range("B2").Value = "DASHBOARD DE SUIVI DE PRODUCTION"
        .Range("B2").Font.Size = 18
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlCenter
        .Range("B2").Interior.Color = RGB(102, 126, 234)
        .Range("B2").Font.Color = RGB(255, 255, 255)
        
        ' Contrôles
        .Range("B4").Value = "Mois :"
        .Range("C4").Validation.Delete
        .Range("C4").Validation.Add Type:=xlValidateList, _
            Formula1:="Janvier,Février,Mars,Avril,Mai,Juin,Juillet,Août,Septembre,Octobre,Novembre,Décembre"
        .Range("C4").Value = Format(Date, "mmmm")
        
        .Range("E4").Value = "Date spécifique :"
        .Range("F4").NumberFormat = "dd/mm/yyyy"
        
        ' Bouton de mise à jour
        Dim btn As Button
        Set btn = .Buttons.Add(.Range("H4").Left, .Range("H4").Top, 100, 25)
        btn.Text = "Mettre à jour"
        btn.OnAction = "UpdateDashboard"
        
        ' Section statistiques globales
        .Range("B6").Value = "STATISTIQUES GLOBALES"
        .Range("B6").Font.Bold = True
        .Range("B6").Font.Size = 14
        
        Call CreateStatsHeaders(ws, 7)
        
        ' Section détails des performances
        .Range("B12").Value = "DÉTAILS DES PERFORMANCES"
        .Range("B12").Font.Bold = True
        .Range("B12").Font.Size = 14
        
        Call CreatePerformanceHeaders(ws, 13)
        
        ' Section ETP par activité
        .Range("B25").Value = "RÉPARTITION ETP PAR ACTIVITÉ"
        .Range("B25").Font.Bold = True
        .Range("B25").Font.Size = 14
        
        Call CreateETPHeaders(ws, 26)
        
        ' Mise en forme générale
        .Columns("A:K").AutoFit
        .Range("B:K").HorizontalAlignment = xlCenter
    End With
End Sub

' =============================================
' CRÉATION DES EN-TÊTES - STATISTIQUES
' =============================================
Sub CreateStatsHeaders(ws As Worksheet, startRow As Integer)
    With ws
        .Cells(startRow, 2).Value = "ETP Jour"
        .Cells(startRow, 3).Value = "ETP Semaine (Moy.)"
        .Cells(startRow, 4).Value = "ETP Mois (Moy.)"
        .Cells(startRow, 5).Value = "Jours de données"
        
        .Range(.Cells(startRow, 2), .Cells(startRow, 5)).Font.Bold = True
        .Range(.Cells(startRow, 2), .Cells(startRow, 5)).Interior.Color = RGB(248, 249, 250)
    End With
End Sub

' =============================================
' CRÉATION DES EN-TÊTES - PERFORMANCES
' =============================================
Sub CreatePerformanceHeaders(ws As Worksheet, startRow As Integer)
    With ws
        .Cells(startRow, 2).Value = "Indicateur"
        .Cells(startRow, 3).Value = "Jour"
        .Cells(startRow, 4).Value = "Semaine (Moy.)"
        .Cells(startRow, 5).Value = "Mois (Moy.)"
        
        .Range(.Cells(startRow, 2), .Cells(startRow, 5)).Font.Bold = True
        .Range(.Cells(startRow, 2), .Cells(startRow, 5)).Interior.Color = RGB(248, 249, 250)
        
        ' Bordures
        .Range(.Cells(startRow, 2), .Cells(startRow + 10, 5)).Borders.LineStyle = xlContinuous
    End With
End Sub

' =============================================
' CRÉATION DES EN-TÊTES - ETP
' =============================================
Sub CreateETPHeaders(ws As Worksheet, startRow As Integer)
    With ws
        .Cells(startRow, 2).Value = "Activité"
        .Cells(startRow, 3).Value = "ETP Jour"
        .Cells(startRow, 4).Value = "ETP Semaine (Moy.)"
        .Cells(startRow, 5).Value = "ETP Mois (Moy.)"
        
        .Range(.Cells(startRow, 2), .Cells(startRow, 5)).Font.Bold = True
        .Range(.Cells(startRow, 2), .Cells(startRow, 5)).Interior.Color = RGB(248, 249, 250)
        
        ' Bordures
        .Range(.Cells(startRow, 2), .Cells(startRow + 8, 5)).Borders.LineStyle = xlContinuous
    End With
End Sub

' =============================================
' FONCTION PRINCIPALE DE MISE À JOUR
' =============================================
Sub UpdateDashboard()
    Dim ws As Worksheet
    Dim selectedMonth As String
    Dim selectedDate As Date
    Dim hasSpecificDate As Boolean
    
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    selectedMonth = ws.Range("C4").Value
    
    ' Vérifier si une date spécifique est sélectionnée
    hasSpecificDate = IsDate(ws.Range("F4").Value)
    If hasSpecificDate Then
        selectedDate = ws.Range("F4").Value
    End If
    
    ' Effacer les données précédentes
    Call ClearDashboardData(ws)
    
    ' Charger et calculer les nouvelles données
    Call LoadDashboardData(ws, selectedMonth, selectedDate, hasSpecificDate)
    
    ' Message de confirmation
    ws.Range("B3").Value = "Dernière mise à jour : " & Format(Now, "dd/mm/yyyy hh:mm")
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").Font.Italic = True
End Sub

' =============================================
' EFFACEMENT DES DONNÉES PRÉCÉDENTES
' =============================================
Sub ClearDashboardData(ws As Worksheet)
    With ws
        .Range("B8:E10").ClearContents    ' Stats globales
        .Range("B14:E23").ClearContents   ' Performances
        .Range("B27:E34").ClearContents   ' ETP
    End With
End Sub

' =============================================
' CHARGEMENT ET CALCUL DES DONNÉES
' =============================================
Sub LoadDashboardData(ws As Worksheet, monthName As String, specificDate As Date, hasSpecificDate As Boolean)
    Dim sourceWs As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    Dim dayData As Range
    Dim weekData As Collection
    Dim monthData As Collection
    
    ' Récupérer la feuille source
    On Error Resume Next
    Set sourceWs = ThisWorkbook.Worksheets(monthName)
    On Error GoTo 0
    
    If sourceWs Is Nothing Then
        MsgBox "Feuille " & monthName & " introuvable!", vbExclamation
        Exit Sub
    End If
    
    ' Trouver la dernière ligne de données
    lastRow = sourceWs.Cells(sourceWs.Rows.count, "A").End(xlUp).Row
    
    If lastRow < 3 Then
        MsgBox "Aucune donnée trouvée dans la feuille " & monthName, vbExclamation
        Exit Sub
    End If
    
    ' Charger les données
    Set monthData = GetMonthData(sourceWs, lastRow)
    
    If hasSpecificDate Then
        Set dayData = GetDayData(sourceWs, specificDate, lastRow)
        Set weekData = GetWeekData(sourceWs, specificDate, lastRow)
    Else
        Debug.Print "Last Row: " & lastRow
        Set dayData = GetLatestDayData(sourceWs, lastRow)
        Set weekData = GetCurrentWeekData(sourceWs, lastRow)
    End If
    
    ' Calculer et afficher les statistiques
    Call DisplayGlobalStats(ws, dayData, weekData, monthData)
    Call DisplayPerformanceMetrics(ws, dayData, weekData, monthData)
    Call DisplayETPBreakdown(ws, dayData, weekData, monthData)
End Sub

' =============================================
' RÉCUPÉRATION DES DONNÉES DU MOIS
' =============================================
Function GetMonthData(sourceWs As Worksheet, lastRow As Long) As Collection
    Dim monthData As New Collection
    Dim i As Long
    Dim rowData As Object
    
    For i = 3 To lastRow ' Ignorer les en-têtes et la ligne Total
        If IsDate(sourceWs.Cells(i, 1).Value) Then
            Set rowData = CreateRowData(sourceWs, i)
            monthData.Add rowData
        End If
    Next i
    
    Set GetMonthData = monthData
End Function

' =============================================
' RÉCUPÉRATION DES DONNÉES D'UN JOUR SPÉCIFIQUE
' =============================================
Function GetDayData(sourceWs As Worksheet, targetDate As Date, lastRow As Long) As Object
    Dim i As Long
    
    For i = 3 To lastRow
        If sourceWs.Cells(i, 1).Value = targetDate Then
            Set GetDayData = CreateRowData(sourceWs, i)
            Exit Function
        End If
    Next i
    
    Set GetDayData = Nothing
End Function

' =============================================
' RÉCUPÉRATION DES DONNÉES DU DERNIER JOUR
' =============================================
Function GetLatestDayData(sourceWs As Worksheet, lastRow As Long) As Object
    Dim i As Long
    
    For i = lastRow To 3 Step -1
        If IsDate(sourceWs.Cells(i, 1).Value) Then
            Set GetLatestDayData = CreateRowData(sourceWs, i)
            Exit Function
        End If
    Next i
    
    Set GetLatestDayData = Nothing
End Function

' =============================================
' RÉCUPÉRATION DES DONNÉES DE LA SEMAINE
' =============================================
Function GetWeekData(sourceWs As Worksheet, targetDate As Date, lastRow As Long) As Collection
    Dim weekData As New Collection
    Dim weekStart As Date, weekEnd As Date
    Dim i As Long
    Dim currentDate As Date
    
    ' Calculer le début de la semaine (lundi)
    weekStart = targetDate - Weekday(targetDate, vbMonday) + 1
    weekEnd = weekStart + 6
    
    For i = 3 To lastRow
        If IsDate(sourceWs.Cells(i, 1).Value) Then
            currentDate = sourceWs.Cells(i, 1).Value
            If currentDate >= weekStart And currentDate <= weekEnd Then
                weekData.Add CreateRowData(sourceWs, i)
            End If
        End If
    Next i
    
    Set GetWeekData = weekData
End Function

' =============================================
' RÉCUPÉRATION DES DONNÉES DE LA SEMAINE COURANTE
' =============================================
Function GetCurrentWeekData(sourceWs As Worksheet, lastRow As Long) As Collection
    Dim latestDate As Date
    Dim i As Long
    
    ' Trouver la dernière date
    For i = lastRow To 3 Step -1
        If IsDate(sourceWs.Cells(i, 1).Value) Then
            latestDate = sourceWs.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    Set GetCurrentWeekData = GetWeekData(sourceWs, latestDate, lastRow)
End Function

' =============================================
' CRÉATION D'UN OBJET LIGNE DE DONNÉES
' =============================================
Function CreateRowData(sourceWs As Worksheet, rowIndex As Long) As Object
    Dim rowData As Object
    Set rowData = CreateObject("Scripting.Dictionary")
    
    With sourceWs
        rowData("Date") = .Cells(rowIndex, 1).Value
        rowData("ETP") = GetNumericValue(.Cells(rowIndex, 2).Value)
        rowData("OUVERTURES_MANUELLES") = GetNumericValue(.Cells(rowIndex, 3).Value)
        rowData("DOSSIERS_OUVERTURES_IA") = GetNumericValue(.Cells(rowIndex, 5).Value)
        rowData("ETP2") = GetNumericValue(.Cells(rowIndex, 8).Value)
        rowData("TICKETS_RETOUR") = GetNumericValue(.Cells(rowIndex, 9).Value)
        rowData("ETP3") = GetNumericValue(.Cells(rowIndex, 12).Value)
        rowData("TRANSFERTS") = GetNumericValue(.Cells(rowIndex, 13).Value)
        rowData("CONTROLE_TRANSFERTS") = GetNumericValue(.Cells(rowIndex, 17).Value)
        rowData("ETP6") = GetNumericValue(.Cells(rowIndex, 24).Value)
        rowData("BAL_SUCC") = GetNumericValue(.Cells(rowIndex, 25).Value)
        rowData("ETP7") = GetNumericValue(.Cells(rowIndex, 28).Value)
        rowData("UAN_COURRIERS") = GetNumericValue(.Cells(rowIndex, 29).Value)
        rowData("ETP12") = GetNumericValue(.Cells(rowIndex, 44).Value)
        rowData("PT_TEL") = GetNumericValue(.Cells(rowIndex, 45).Value)
        rowData("TOTAL_ETP") = GetNumericValue(.Cells(rowIndex, 50).Value)
    End With
    
    Set CreateRowData = rowData
End Function

' =============================================
' UTILITAIRE POUR RÉCUPÉRER UNE VALEUR NUMÉRIQUE
' =============================================
Function GetNumericValue(cellValue As Variant) As Double
    If IsNumeric(cellValue) And Not IsEmpty(cellValue) Then
        GetNumericValue = CDbl(cellValue)
    Else
        GetNumericValue = 0
    End If
End Function

' =============================================
' AFFICHAGE DES STATISTIQUES GLOBALES
' =============================================
Sub DisplayGlobalStats(ws As Worksheet, dayData As Object, weekData As Collection, monthData As Collection)
    Dim etpDay As Double, etpWeek As Double, etpMonth As Double
    
    ' Calcul ETP jour
    If Not dayData Is Nothing Then
        etpDay = dayData("TOTAL_ETP")
    End If
    
    ' Calcul ETP semaine (moyenne)
    etpWeek = CalculateAverageETP(weekData, "TOTAL_ETP")
    
    ' Calcul ETP mois (moyenne)
    etpMonth = CalculateAverageETP(monthData, "TOTAL_ETP")
    
    ' Affichage
    With ws
        .Range("B8").Value = Format(etpDay, "0.00")
        .Range("C8").Value = Format(etpWeek, "0.00")
        .Range("D8").Value = Format(etpMonth, "0.00")
        .Range("E8").Value = monthData.count
        
        ' Mise en forme
        .Range("B8:E8").Font.Bold = True
        .Range("B8:E8").HorizontalAlignment = xlCenter
    End With
End Sub

' =============================================
' AFFICHAGE DES MÉTRIQUES DE PERFORMANCE
' =============================================
Sub DisplayPerformanceMetrics(ws As Worksheet, dayData As Object, weekData As Collection, monthData As Collection)
    Dim metrics(1 To 7) As MetricInfo
    Dim i As Integer
    
    ' Définition des métriques
    metrics(1).Name = "Ouvertures Manuelles": metrics(1).dataColumn = "OUVERTURES_MANUELLES": metrics(1).etpColumn = "ETP": metrics(1).Unit = "dossiers"
    metrics(2).Name = "Ouvertures IA": metrics(2).dataColumn = "DOSSIERS_OUVERTURES_IA": metrics(2).etpColumn = "ETP": metrics(2).Unit = "dossiers"
    metrics(3).Name = "Tickets Retour": metrics(3).dataColumn = "TICKETS_RETOUR": metrics(3).etpColumn = "ETP2": metrics(3).Unit = "tickets"
    metrics(4).Name = "Transferts": metrics(4).dataColumn = "TRANSFERTS": metrics(4).etpColumn = "ETP3": metrics(4).Unit = "transferts"
    metrics(5).Name = "Contrôle Transferts": metrics(5).dataColumn = "CONTROLE_TRANSFERTS": metrics(5).etpColumn = "ETP3": metrics(5).Unit = "contrôles"
    metrics(6).Name = "BAL Succession": metrics(6).dataColumn = "BAL_SUCC": metrics(6).etpColumn = "ETP6": metrics(6).Unit = "mails"
    metrics(7).Name = "UAN Courriers": metrics(7).dataColumn = "UAN_COURRIERS": metrics(7).etpColumn = "ETP7": metrics(7).Unit = "courriers"
    
    For i = 1 To 7
        Dim dayValue As Double, weekAvg As Double, monthAvg As Double
        
        ' Valeur du jour
        If Not dayData Is Nothing Then
            dayValue = dayData(metrics(i).dataColumn)
        End If
        
        ' Moyenne semaine
        weekAvg = CalculateWeightedAverage(weekData, metrics(i).dataColumn, metrics(i).etpColumn)
        
        ' Moyenne mois
        monthAvg = CalculateWeightedAverage(monthData, metrics(i).dataColumn, metrics(i).etpColumn)
        
        ' Affichage
        With ws
            .Cells(13 + i, 2).Value = metrics(i).Name
            .Cells(13 + i, 3).Value = Format(dayValue, "0") & " " & metrics(i).Unit
            .Cells(13 + i, 4).Value = Format(weekAvg, "0.0") & " " & metrics(i).Unit & "/ETP"
            .Cells(13 + i, 5).Value = Format(monthAvg, "0.0") & " " & metrics(i).Unit & "/ETP"
        End With
    Next i
End Sub

' =============================================
' AFFICHAGE DE LA RÉPARTITION ETP
' =============================================
Sub DisplayETPBreakdown(ws As Worksheet, dayData As Object, weekData As Collection, monthData As Collection)
    Dim etpColumns As Variant
    Dim etpNames As Variant
    Dim i As Integer
    
    etpColumns = Array("ETP", "ETP2", "ETP3", "ETP6", "ETP7", "ETP12")
    etpNames = Array("Ouvertures", "Tickets Retour", "Transferts", "BAL Succession", "UAN Courriers", "Plateforme Tél")
    
    For i = 0 To UBound(etpColumns)
        Dim dayValue As Double, weekAvg As Double, monthAvg As Double
        
        ' Valeur du jour
        If Not dayData Is Nothing Then
            dayValue = dayData(etpColumns(i))
        End If
        
        ' Moyenne semaine
        weekAvg = CalculateAverageETP(weekData, etpColumns(i))
        
        ' Moyenne mois
        monthAvg = CalculateAverageETP(monthData, etpColumns(i))
        
        ' Affichage
        With ws
            .Cells(27 + i, 2).Value = etpNames(i)
            .Cells(27 + i, 3).Value = Format(dayValue, "0.0")
            .Cells(27 + i, 4).Value = Format(weekAvg, "0.0")
            .Cells(27 + i, 5).Value = Format(monthAvg, "0.0")
        End With
    Next i
End Sub

' =============================================
' CALCUL DE MOYENNE ETP
' =============================================
Function CalculateAverageETP(dataCollection As Collection, etpColumn As String) As Double
    Dim total As Double
    Dim count As Integer
    Dim item As Object
    
    total = 0
    count = 0
    
    For Each item In dataCollection
        If item(etpColumn) > 0 Then
            total = total + item(etpColumn)
            count = count + 1
        End If
    Next item
    
    If count > 0 Then
        CalculateAverageETP = total / count
    Else
        CalculateAverageETP = 0
    End If
End Function

' =============================================
' CALCUL DE MOYENNE PONDÉRÉE
' =============================================
Function CalculateWeightedAverage(dataCollection As Collection, dataColumn As String, etpColumn As String) As Double
    Dim totalProd As Double
    Dim totalETP As Double
    Dim item As Object
    
    totalProd = 0
    totalETP = 0
    
    For Each item In dataCollection
        If item(etpColumn) > 0 And item(dataColumn) > 0 Then
            totalProd = totalProd + item(dataColumn)
            totalETP = totalETP + item(etpColumn)
        End If
    Next item
    
    If totalETP > 0 Then
        CalculateWeightedAverage = totalProd / totalETP
    Else
        CalculateWeightedAverage = 0
    End If
End Function

