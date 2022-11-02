Private Sub CommandButton3_Click()
    Dim intNoOfRows
    Dim intNoOfColumns
    Dim objWord
    Dim objDoc
    Dim objSelection
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = True
    Dim addressLine1 As String
    Dim addressLine2 As String
    Dim addressLine3 As String
    Dim phoneNumberLine As String
    
    addressLine1 = "Office Address: "
    addressLine2 = "H.No. 21, Ichchapur, Gowalapara, Adityapur, P.O.: N.I.T. Jamshedpur, PIN -831014, Saraikela-Kharsawan"
    addressLine3 = "Landmark: Kamala Vastralaya, Sahara Garden City Road"
    phoneNumberLine = "Ph. No.: +91-7903400696/WhatsApp: +91-9835547589"
    
    With objDoc.PageSetup
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
        .LeftMargin = Application.InchesToPoints(0.6)
        .RightMargin = Application.InchesToPoints(0.6)
    End With
    
    Set objSelection = objWord.Selection
    objDoc.SaveAs ("C:\Users\dassh\Documents\Bill2")
    
    Dim j As Integer
    Dim l As Integer

    '// set your data range here
    '// Set dataRange = ActiveSheet.Range("A1:A400")
    Set dataRange = Sheet5.Range("A1:A114")
    
    addressColumnIndex = getBillColumnIndex("room")
    paperColumnIndex = getBillColumnIndex("paper")
    stopDateColumn = getBillColumnIndex("stop")
    startDateColumn = getBillColumnIndex("start")
    billColumn = getBillColumnIndex("bill")
    intNoOfRows = 2
    
    For i = 1 To dataRange.Rows.Count / 2
        With dataRange.Cells(intNoOfRows, addressColumnIndex)
        If .Value <> "" Then
            Set objRange = objSelection.Range
            objDoc.Tables.Add objRange, 10, 11
            Set objTable = objDoc.Tables(i)
            objTable.Borders.Enable = True
            objTable.Range.Font.Size = 9
            objTable.Range.Font.Bold = True
            j = 0
            l = 0
            For k = 1 To 2
                With objDoc.Tables(i)
                    '// For Heading
                    Set Rng = .Cell(1, j + 1).Range
                    Rng.End = .Cell(1, j + 5).Range.End
                    Rng.Cells.Merge
                    
                    '// For Address
                    Set Rng = .Cell(2, j + 1).Range
                    Rng.End = .Cell(4, j + 5).Range.End
                    Rng.Cells.Merge
                    
                    '// For Subscriber Address
                    Set Rng = .Cell(5, j + 1).Range
                    Rng.End = .Cell(5, j + 5).Range.End
                    Rng.Cells.Merge
                    
                    '// For Paper/Book column header
                    For Row = 6 To 9
                        Set Rng = .Cell(Row, l + 1).Range
                        Rng.End = .Cell(Row, l + 2).Range.End
                        Rng.Cells.Merge
                    Next
                    
                    '// For Signature row
                    Set Rng = .Cell(10, j + 1).Range
                    Rng.End = .Cell(10, j + 5).Range.End
                    Rng.Cells.Merge
                    
                    For Row = 6 To 9
                        objTable.Cell(Row, l + 1).Width = 65
                        objTable.Cell(Row, l + 2).Width = 64
                        objTable.Cell(Row, l + 3).Width = 64
                        objTable.Cell(Row, l + 4).Width = 46
                    Next
                    '// .PreferredWidth = 100
                End With
                objTable.Cell(1, j + 1).Range.Text = "                              TARUN KUMAR DAS" & Chr(10) & "  (Leading National Newspaper & Magazine Distributor)"
                objTable.Cell(2, j + 1).Range.Text = addressLine1 & addressLine2 & Chr(10) & addressLine3 & Chr(10) & phoneNumberLine
                objTable.Cell(5, j + 1).Range.Text = "Subscriber Address: " & dataRange.Cells(intNoOfRows, addressColumnIndex).Value & "                    Bill Month: AUG '17"
                objTable.Cell(6, l + 1).Range.Text = "Paper/Book"
                objTable.Cell(6, l + 2).Range.Text = "Start Dt"
                objTable.Cell(6, l + 3).Range.Text = "End Dt"
                objTable.Cell(6, l + 4).Range.Text = "Amt."
                objTable.Cell(7, l + 1).Range.Text = dataRange.Cells(intNoOfRows, paperColumnIndex).Value
                objTable.Cell(7, l + 2).Range.Text = Format(dataRange.Cells(intNoOfRows, startDateColumn), "dd/mm/yyyy")
                objTable.Cell(7, l + 3).Range.Text = Format(dataRange.Cells(intNoOfRows, stopDateColumn), "dd/mm/yyyy")
                '// Compute price
                objTable.Cell(7, l + 4).Range.Text = dataRange.Cells(intNoOfRows, 7)
                objTable.Cell(9, l + 3).Range.Text = "Total"
                objTable.Cell(10, j + 1).Range.Text = "                                 Signature: "
                '// If there is no book, then populate the total amount
                If dataRange.Cells(intNoOfRows, 3) = "" Then
                    objTable.Cell(9, l + 4).Range.Text = dataRange.Cells(intNoOfRows, 7)
                End If
                If dataRange.Cells(intNoOfRows, startDateColumn) = "" Then
                    objTable.Cell(7, l + 2).Range.Text = Format(dhFirstDayInMonth(CDate("8/3/2017")), "dd/mm/yyyy")
                End If
                If dataRange.Cells(intNoOfRows, stopDateColumn) = "" Then
                    objTable.Cell(7, l + 3).Range.Text = Format(GetNowLast(CDate("8/3/2017")), "dd/mm/yyyy")
                End If
                j = j + 2
                l = l + 5
                intNoOfRows = intNoOfRows + 1
            Next
            objSelection.EndKey 6
            objSelection.TypeParagraph
        End If
        '// intNoOfRows = intNoOfRows + 1
        End With
    Next
    Dim weekDay As Integer
    Dim dayOfWeek As Long

    objDoc.SaveAs ("C:\Users\dsourabh\Documents\Bill2")
End Sub

Function MonthWeekDays(dDate As Date, iWeekDay As Integer)
    Dim dLoop As Date
    If iWeekDay < 1 Or iWeekDay > 7 Then
        MonthWeekDays = CVErr(xlErrNum)
        Exit Function
    End If
    MonthWeekDays = 0
    dLoop = DateSerial(Year(dDate), Month(dDate), 1)
    Do While Month(dLoop) = Month(dDate)
        If weekDay(dLoop) = iWeekDay Then _
            MonthWeekDays = MonthWeekDays + 1
        dLoop = dLoop + 1
    Loop
End Function

Function WkDays(StartDate As Date, EndDate As Date, Days As Long) As Integer

    ' Returns the number of qualifying days between (and including)
    ' StartDate and EndDate. Qualifying days are whole numbers where
    ' each digit represents a day of the week that should be counted,
    ' with Monday=1, Tuesday=2, etc. For example, all Mondays, Tuesdays
    ' and Thursdays are to be counted between the two dates, set
    ' WkDays = 124 on your worksheet.
    '
    
    Dim iDate As Date
    Dim strQdays As String
    
    strQdays = CStr(Days)
    WkDays = 0
    
    For iDate = StartDate To EndDate
    If strQdays Like "*" & CStr(weekDay(iDate)) & "*" Then
        WkDays = WkDays + 1
    End If
    Next iDate

End Function
