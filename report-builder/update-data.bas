Sub UpdateData()
    Dim wsSourceNew As Worksheet
    Dim wsSourceClosed As Worksheet
    Dim wsDestNew As Worksheet
    Dim wsDestClosed As Worksheet
    Dim wsDestReceivedClosed As Worksheet
    Dim wsMonthlyNewCases As Worksheet
    Dim wsSLA As Worksheet
    Dim i As Integer
    Dim lastCol As Integer
    Dim currentMonth As Integer
    Dim nextMonth As Integer
    Dim currentMonthName As String
    Dim currentMonthDate As Date
    Dim nextMonthDate As Date
    Dim nextMonthName As String
    Dim rowCount As Long
    Dim countSeverity1 As Long
    Dim countSeverity2 As Long
    Dim countSeverity3 As Long
    Dim countSeverity4 As Long
    Dim countSeverity1AndY As Long
    Dim countSeverity2AndY As Long
    Dim countSeverity3AndY As Long
    Dim countSeverity4AndY As Long
    Dim totalG As Double
    Dim totalF As Double
    Dim percentage As Double
    Dim totalE As Long
    Dim lastRow As Long

    Set wsSourceNew = Workbooks("report-data-extract.xlsx").Sheets("New")
    Set wsSourceClosed = Workbooks("report-data-extract.xlsx").Sheets("Closed")
    Set wsDestNew = Workbooks("report-builder.xlsm").Sheets("New")
    Set wsDestClosed = Workbooks("report-builder.xlsm").Sheets("Closed")
    Set wsDestReceivedClosed = Workbooks("report-builder.xlsm").Sheets("WorkedTickets")

    ' Data copy for New and Closed
    wsDestNew.Cells.ClearContents
    wsDestClosed.Cells.ClearContents

    wsSourceNew.Cells.Copy Destination:=wsDestNew.Cells(1, 1)
    wsSourceClosed.Cells.Copy Destination:=wsDestClosed.Cells(1, 1)

    ' Data update for WorkedTickets
    lastCol = wsDestReceivedClosed.Cells(3, wsDestReceivedClosed.Columns.Count).End(xlToLeft).Column

    For i = 2 To 6
        With wsDestReceivedClosed
            .Range(.Cells(i + 1, 1), .Cells(i + 1, lastCol)).Copy
            .Range(.Cells(i, 1), .Cells(i, lastCol)).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        End With
    Next i
    
    currentMonth = wsDestReceivedClosed.Cells(6, 1).Value

    If currentMonth = 12 Then
        nextMonth = 1
    Else
        nextMonth = currentMonth + 1
    End If

    wsDestReceivedClosed.Cells(7, 1).Value = nextMonth

    currentMonthName = wsDestReceivedClosed.Cells(6, 2).Value

    currentMonthDate = DateValue("1 " & currentMonthName & " " & Year(Date))

    nextMonthDate = DateAdd("m", 1, currentMonthDate)

    nextMonthName = Format(nextMonthDate, "mmmm")
    nextMonthName = UCase(Left(nextMonthName, 1)) & Mid(nextMonthName, 2)

    wsDestReceivedClosed.Cells(7, 2).Value = nextMonthName
    
    rowCount = Application.WorksheetFunction.CountA(wsDestNew.Range("A:A")) - 1

    wsDestReceivedClosed.Cells(7, 4).Value = rowCount
    
    rowCount = Application.WorksheetFunction.CountA(wsDestClosed.Range("A:A")) - 1
    
    wsDestReceivedClosed.Cells(7, 5).Value = rowCount
       
    'Data update for MonthlyNewCases
    
    Set wsMonthlyNewCases = Workbooks("report-builder.xlsm").Sheets("MonthlyNewCases")

    wsMonthlyNewCases.Cells(2, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "WO Error")
    wsMonthlyNewCases.Cells(3, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Latency")
    wsMonthlyNewCases.Cells(4, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "AC Error")
    wsMonthlyNewCases.Cells(5, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "User Error")
    wsMonthlyNewCases.Cells(6, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Preventive Maintenance Error")
    wsMonthlyNewCases.Cells(7, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Access Permissions")
    wsMonthlyNewCases.Cells(8, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "CSG Error")
    wsMonthlyNewCases.Cells(9, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Open Error")
    wsMonthlyNewCases.Cells(10, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "LISY Error")
    wsMonthlyNewCases.Cells(11, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Scannet Error")
    wsMonthlyNewCases.Cells(12, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Salesforce Error")
    wsMonthlyNewCases.Cells(13, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Outage Error")
    wsMonthlyNewCases.Cells(14, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Change Error")
    wsMonthlyNewCases.Cells(15, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Incident Error")
    wsMonthlyNewCases.Cells(16, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Impact Error")
    wsMonthlyNewCases.Cells(17, 2).Value = Application.WorksheetFunction.CountIf(wsDestNew.Range("O:O"), "Alarmed")
    
    totalE = wsMonthlyNewCases.Cells(2, 2).Value + wsMonthlyNewCases.Cells(3, 2).Value + _
                       wsMonthlyNewCases.Cells(4, 2).Value + wsMonthlyNewCases.Cells(5, 2).Value + _
                       wsMonthlyNewCases.Cells(6, 2).Value + wsMonthlyNewCases.Cells(7, 2).Value + _
                       wsMonthlyNewCases.Cells(8, 2).Value + wsMonthlyNewCases.Cells(9, 2).Value + _
                       wsMonthlyNewCases.Cells(10, 2).Value + wsMonthlyNewCases.Cells(11, 2).Value + _
                       wsMonthlyNewCases.Cells(12, 2).Value + wsMonthlyNewCases.Cells(13, 2).Value + _
                       wsMonthlyNewCases.Cells(14, 2).Value + wsMonthlyNewCases.Cells(15, 2).Value + _
                       wsMonthlyNewCases.Cells(16, 2).Value + wsMonthlyNewCases.Cells(17, 2).Value

    lastRow = wsDestNew.Cells(wsDestNew.Rows.Count, "P").End(xlUp).Row

    wsMonthlyNewCases.Cells(18, 2).Value = lastRow - totalE - 1
    
    With wsMonthlyNewCases
        .Cells(2, 5).Value = "Integration Error"
        .Cells(2, 6).Formula = "=SUM(B8,B9,B10,B11,B12,B13)"
        
        .Cells(3, 5).Value = "WO Error"
        .Cells(3, 6).Formula = "=B2"
        
        .Cells(4, 5).Value = "AC Error"
        .Cells(4, 6).Formula = "=B4"
        
        .Cells(5, 5).Value = "Preventive Maintenance Error"
        .Cells(5, 6).Formula = "=B6"
        
        .Cells(6, 5).Value = "Change/Incident/Impact Error"
        .Cells(6, 6).Formula = "=SUM(B14,B15,B16)"
        
        .Cells(7, 5).Value = "Latency"
        .Cells(7, 6).Formula = "=B3"
        
        .Cells(8, 5).Value = "Access Permissions"
        .Cells(8, 6).Formula = "=B7"
        
        .Cells(9, 5).Value = "Alarmed"
        .Cells(9, 6).Formula = "=B17"
        
        .Cells(10, 5).Value = "User Error"
        .Cells(10, 6).Formula = "=B5"
        
        .Cells(11, 5).Value = "Others"
        .Cells(11, 6).Formula = "=B18"
        
    End With
    
    ' Data update for SLA
    countSeverity1 = Application.WorksheetFunction.CountIf(wsDestClosed.Range("C:C"), "Severity - 1")
    countSeverity2 = Application.WorksheetFunction.CountIf(wsDestClosed.Range("C:C"), "Severity - 2")
    countSeverity3 = Application.WorksheetFunction.CountIf(wsDestClosed.Range("C:C"), "Severity - 3")
    countSeverity4 = Application.WorksheetFunction.CountIf(wsDestClosed.Range("C:C"), "Severity - 4")

    Set wsSLA = Workbooks("report-builder.xlsm").Sheets("SLA")
    wsSLA.Cells(6, 7).Value = countSeverity1
    wsSLA.Cells(7, 7).Value = countSeverity2
    wsSLA.Cells(8, 7).Value = countSeverity3
    wsSLA.Cells(9, 7).Value = countSeverity4
    
    wsSLA.Cells(13, 7).Value = countSeverity1
    wsSLA.Cells(14, 7).Value = countSeverity2
    wsSLA.Cells(15, 7).Value = countSeverity3
    wsSLA.Cells(16, 7).Value = countSeverity4

    countSeverity1AndY = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 1", wsDestClosed.Range("L:L"), "Y")
    countSeverity2AndY = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 2", wsDestClosed.Range("L:L"), "Y")
    countSeverity3AndY = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 3", wsDestClosed.Range("L:L"), "Y")
    countSeverity4AndY = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 4", wsDestClosed.Range("L:L"), "Y")
    
    countSeverity1AndY2 = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 1", wsDestClosed.Range("M:M"), "Y")
    countSeverity2AndY2 = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 2", wsDestClosed.Range("M:M"), "Y")
    countSeverity3AndY2 = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 3", wsDestClosed.Range("M:M"), "Y")
    countSeverity4AndY2 = Application.WorksheetFunction.CountIfs(wsDestClosed.Range("C:C"), "Severity - 4", wsDestClosed.Range("M:M"), "Y")

    wsSLA.Cells(6, 6).Value = countSeverity1AndY
    wsSLA.Cells(7, 6).Value = countSeverity2AndY
    wsSLA.Cells(8, 6).Value = countSeverity3AndY
    wsSLA.Cells(9, 6).Value = countSeverity4AndY
    
    wsSLA.Cells(13, 6).Value = countSeverity1AndY2
    wsSLA.Cells(14, 6).Value = countSeverity2AndY2
    wsSLA.Cells(15, 6).Value = countSeverity3AndY2
    wsSLA.Cells(16, 6).Value = countSeverity4AndY2

    For i = 6 To 9
        totalG = wsSLA.Cells(i, 7).Value
        totalF = wsSLA.Cells(i, 6).Value

        If totalG < 5 Or totalG = 0 Or totalF < 5 Or totalF = 0 Then
            wsSLA.Cells(i, 6).Value = "N/A"
            wsSLA.Cells(i, 7).Value = "N/A"
            wsSLA.Cells(i, 8).Value = "N/A"
        Else
            If totalG <> 0 Then
                percentage = Round((totalF / totalG) * 100) / 100
                wsSLA.Cells(i, 8).Value = percentage
                wsSLA.Cells(i, 8).NumberFormat = "0%"
            Else
                wsSLA.Cells(i, 8).Value = 0
            End If
        End If
    Next i
    
    For i = 13 To 16
        totalG = wsSLA.Cells(i, 7).Value
        totalF = wsSLA.Cells(i, 6).Value

        If totalG < 5 Or totalG = 0 Or totalF < 5 Or totalF = 0 Then
            wsSLA.Cells(i, 6).Value = "N/A"
            wsSLA.Cells(i, 7).Value = "N/A"
            wsSLA.Cells(i, 8).Value = "N/A"
        Else
            If totalG <> 0 Then
                percentage = Round((totalF / totalG) * 100) / 100
                wsSLA.Cells(i, 8).Value = percentage
                wsSLA.Cells(i, 8).NumberFormat = "0%"
            Else
                wsSLA.Cells(i, 8).Value = 0
            End If
        End If
    Next i

End Sub



