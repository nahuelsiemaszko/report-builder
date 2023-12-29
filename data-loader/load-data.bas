Sub LoadData()
    Dim wsSource As Worksheet
    Dim wsIntermediate As Worksheet
    Dim wsDestination As Worksheet
    Dim LastRow As Long
    Dim Title As String
    Dim LastRowSource As Long
    Dim LastRowDestination As Long
    Dim i As Long
    Dim Cell As Range
    Dim NewSheetName As String
    Dim DateTime As String
    Dim Range1 As Range
    Dim Range2 As Range
    
    Set wsSource = Nothing
    On Error Resume Next
    Set wsSource = Workbooks("modify-data-error.xlsx").Worksheets("Record Submission List")
    On Error Resume Next
    Set wsSource = Workbooks("modify-data-error.xlsx").Worksheets("Sheet1")
    
    If Not wsSource Is Nothing Then
    
        Title = wsSource.Cells(1, 1).Value
        
' Code when cell 1:1 of the source file = "Service ID"
        
        If wsSource.Name = "Record Submission List" And Title = "Service ID" Then
            
            Set wsSource = Workbooks("modify-data-error.xlsx").Worksheets("Record Submission List")
            Set wsIntermediate = Workbooks("data-loader.xlsm").Worksheets("BaseSheet")
            
            wsIntermediate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Set wsDestination = ActiveSheet
            
            NewSheetName = Format(Now, "dd-mm-yy")
            wsDestination.Name = "Record" & NewSheetName
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "P").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "P").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "A")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "O").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "B").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "O").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "B")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "X").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "C").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "X").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "C")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "Y").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "D").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "Y").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "D")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "S").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "E").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "S").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "E")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "T").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "F").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "T").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "F")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "U").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "G").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "U").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "G")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "V").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "H").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "V").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "H")
            Next i
            
          LastRowSource = wsSource.Cells(wsSource.Rows.Count, "N").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "J").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "N").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "J")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "K").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "A").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "K")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "H").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "L").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "H").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "L")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "I").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "M").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "I").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "M")
            Next i
            
           LastRowSource = wsSource.Cells(wsSource.Rows.Count, "J").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "N").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "J").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "N")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "K").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "O").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "K").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "O")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "L").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "P").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "L").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "P")
            Next i
            
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "E").End(xlUp).Row
            
            Set Range1 = wsDestination.Range("E1:H" & LastRowDestination)
            
            Range1.Replace What:="N", Replacement:=0, LookAt:=xlWhole, MatchCase:=False
            Range1.Replace What:="Y", Replacement:=1, LookAt:=xlWhole, MatchCase:=False
            
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row
            
            Set Range2 = wsDestination.Range("A3:P" & LastRowDestination)
            
            Range2.Replace What:="", Replacement:="~NULL~", LookAt:=xlWhole, MatchCase:=False
            
            DateTime = Format(Now, "yyyy-mm-dd\Thh:mm:ss-03:00")
            
            For i = 3 To LastRowDestination
              wsDestination.Cells(i, "I").Value = DateTime
            Next i
        
' Code when cell 1:1 of the source file = "subline"
    
       ElseIf wsSource.Name = "Sheet1" And Title = "subline" Then
            
            Set wsSource = Workbooks("modify-data-error.xlsx").Worksheets("Sheet1")
            Set wsIntermediate = Workbooks("data-loader.xlsm").Worksheets("BaseSheet")
            
            wsIntermediate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Set wsDestination = ActiveSheet
            
            NewSheetName = Format(Now, "dd-mm-yy")
            wsDestination.Name = "Record" & NewSheetName
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "U").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "U").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "A")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "K").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "B").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "K").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "B")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "L").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "C").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "L").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "C")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "M").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "D").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "M").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "D")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "H").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "E").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "H").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "E")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "G").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "F").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "G").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "F")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "J").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "G").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "J").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "G")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "I").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "H").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "I").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "H")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "T").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "J").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "T").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "J")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "K").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "A").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "K")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "O").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "L").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "O").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "L")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "P").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "M").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "P").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "M")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "Q").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "N").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "Q").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "N")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "R").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "O").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "R").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "O")
            Next i
            
            LastRowSource = wsSource.Cells(wsSource.Rows.Count, "S").End(xlUp).Row
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "P").End(xlUp).Row

            For i = 2 To LastRowSource
                wsSource.Cells(i, "S").Copy Destination:=wsDestination.Cells(LastRowDestination + i - 1, "P")
            Next i
            
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "E").End(xlUp).Row
            
            Set Range1 = wsDestination.Range("E1:H" & LastRowDestination)
            
            Range1.Replace What:="N", Replacement:=0, LookAt:=xlWhole, MatchCase:=False
            Range1.Replace What:="Y", Replacement:=1, LookAt:=xlWhole, MatchCase:=False
            
            LastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row
            
            Set Range2 = wsDestination.Range("A3:P" & LastRowDestination)
            
            Range2.Replace What:="", Replacement:="~NULL~", LookAt:=xlWhole, MatchCase:=False
            
            DateTime = Format(Now, "yyyy-mm-dd\Thh:mm:ss-03:00")
            
            For i = 3 To LastRowDestination
              wsDestination.Cells(i, "I").Value = DateTime
            Next i
            
' Message in case of unrecognized formats
    
        Else
            MsgBox "Unknown format in the source file sheet. No changes were made."
        End If
    Else
        MsgBox "Unknown source file sheet. No changes were made."
    End If
End Sub

