Dim dataRange As Range

Private Sub CommandButton1_Click()
    numOfDays
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim thisBook As Workbook
    Dim newBook As Workbook
    Dim newRow As Integer
    Dim temp
    Dim stopData As Variant
    
    '// set your data range here
    Set dataRange = ActiveSheet.Range("A1:A400")

    '// create a new workbook
    '// Set newBook = Excel.Workbooks.Add
    Set dict = CreateObject("Scripting.Dictionary")
    Set bookDict = CreateObject("Scripting.Dictionary")
    Set areaDict = CreateObject("Scripting.Dictionary")
    '// loop through the data in book1, one column at a time
    Dim totalCount As Integer
    
    stopDateColumn = getBillColumnIndex("stop")
    startDateColumn = getBillColumnIndex("start")
    '// MsgBox (stopDateColumn)
    '// dataRange.Cells(7, stopDateColumn).Value = "Hello"
    '// MsgBox (dataRange.Cells(7, stopDateColumn).Value)
    For rowIndex = 2 To dataRange.Rows.Count
        Dim stopDate As Date
        With dataRange.Cells(rowIndex, 1)
        Dim LArray() As String
        Dim areaKey As String
        Dim paper As String

        '// ignore empty cells
        If .Value <> "" Then
            newRow = newRow + 1
            temp = doSomethingWith(.Value)
            '// newBook.ActiveSheet.Cells(newRow, 1).value = temp
            '// newBook.ActiveSheet.Cells(newRow, 2).value = dataRange.Cells(rowIndex, 2)
            With dict
            '// Debug.Print (dataRange.Cells(rowIndex, 2).Value)
            stopData = dataRange.Cells(rowIndex, stopDateColumn).Value
            If stopData <> "" And IsDate(stopData) = False Then
                stopData = ""
                dataRange.Cells(rowIndex, stopDateColumn).Value = ""
            End If
            startData = dataRange.Cells(rowIndex, startDateColumn).Value
            If startData <> "" And IsDate(startData) = False Then
                startData = ""
                dataRange.Cells(rowIndex, startDateColumn).Value = ""
            End If
            If shouldCount(stopData, startData) = True Then
                '// For newspapers only
                If dataRange.Cells(rowIndex, 2).Value <> "" Then
                    paper = dataRange.Cells(rowIndex, 2).Value
                    LArray = Split(dataRange.Cells(rowIndex, 1).Value, "/")
                    areaKey = LArray(0) & "/" & paper
                    If Not dict.Exists(dataRange.Cells(rowIndex, 2).Value) Then
                        '// Debug.Print (stopData)
                        dict.Add dataRange.Cells(rowIndex, 2).Value, 1
                        totalCount = totalCount + 1
                    Else
                        Dim val As Integer
                        '// stopDate = CDate(dataRange.Cells(rowIndex, stopDateColumn).Value)
                        '// If stopDate column contains something, then ignore it
                        '// Debug.Print stopData
                        val = dict.Item(dataRange.Cells(rowIndex, 2).Value) + 1
                        dict(dataRange.Cells(rowIndex, 2).Value) = val
                        totalCount = totalCount + 1
                    End If
                    
                    If Not areaDict.Exists(areaKey) Then
                        areaDict.Add areaKey, 1
                    Else
                        Dim areaVal As Integer
                        areaVal = areaDict.Item(areaKey) + 1
                        areaDict.Item(areaKey) = areaVal
                    End If
                End If
                '// For Books and magazines
                If dataRange.Cells(rowIndex, 3).Value <> "" Then
                    If Not bookDict.Exists(dataRange.Cells(rowIndex, 3).Value) Then
                        '// Debug.Print (dataRange.Cells(rowIndex, 3).Value)
                        '// Debug.Print (rowIndex)
                        bookDict.Add dataRange.Cells(rowIndex, 3).Value, 1
                    Else
                        Dim book As Integer
                        '// stopDate = CDate(dataRange.Cells(rowIndex, stopDateColumn).Value)
                        '// If stopDate column contains something, then ignore it
                        '// Debug.Print stopData
                        val = bookDict.Item(dataRange.Cells(rowIndex, 3).Value) + 1
                        bookDict(dataRange.Cells(rowIndex, 3).Value) = val
                    End If
                End If
                For colIndex = 1 To 6
                    dataRange.Cells(rowIndex, colIndex).Interior.ColorIndex = xlNone
                Next colIndex
            Else
                For colIndex = 1 To 6
                    dataRange.Cells(rowIndex, colIndex).Interior.Color = RGB(255, 0, 0)
                Next colIndex
            End If
            End With
        End If
        End With
    Next rowIndex
    Dim key As Variant
    Dim outRow As Integer
    outRow = 2
    '// SortDictionary dict
    ActiveSheet.Columns(13).ClearContents
    ActiveSheet.Columns(14).ClearContents
    
    dataRange.Cells(outRow, 13).Font.Bold = True
    dataRange.Cells(outRow, 14).Font.Bold = True
    dataRange.Cells(outRow, 13).Value = "PAPER"
    dataRange.Cells(outRow, 14).Value = "COUNT"
    outRow = outRow + 1
    For Each key In dict.Keys()
        dataRange.Cells(outRow, 13).Value = key
        If InStr(LCase(key), "paper") Then
            dataRange.Cells(outRow, 14).Value = "COUNT"
        Else
            dataRange.Cells(outRow, 14).Value = dict.Item(key)
        End If
        outRow = outRow + 1
    Next
    
    dataRange.Cells(outRow, 13).Font.Bold = True
    dataRange.Cells(outRow, 14).Font.Bold = True
    
    dataRange.Cells(outRow, 13).Value = "Total"
    dataRange.Cells(outRow, 14).Value = totalCount

    dataRange.Cells(outRow, 15).Value = localDate
    outRow = outRow + 2

    For Each key In bookDict.Keys()
        dataRange.Cells(outRow, 13).Value = key
        If InStr(LCase(key), "book") Then
            dataRange.Cells(outRow, 14).Value = "COUNT"
        Else
            dataRange.Cells(outRow, 14).Value = bookDict.Item(key)
        End If
        outRow = outRow + 1
    Next
    outRow = outRow + 2
    For Each key In areaDict.Keys()
        dataRange.Cells(outRow, 13).Value = key
        dataRange.Cells(outRow, 14).Value = areaDict.Item(key)
        outRow = outRow + 1
    Next
    
    Set dict = Nothing
    Set bookDict = Nothing
    Set areaDict = Nothing
    ActiveWorkbook.Save
End Sub

Private Function getBillColumnIndex(colName As String)
    Dim colIndex As Integer
    Dim pos As Integer
    Dim found As Boolean
    found = False
    Dim startDateColumnIndex As Integer
    Set headerRange = Sheet1.Range("A1:Z1")
    For colIndex = 1 To headerRange.Columns.Count
        With headerRange.Cells(1, colIndex)
            pos = InStr(LCase(.Value), colName)
            If pos <> 0 Then
                found = True
                getBillColumnIndex = colIndex
            End If
        End With
    Next
End Function

Private Function shouldCount(stopDate As Variant, StartDate As Variant) As Boolean
    Dim localDate As String
    localDate = Date
    localDate = DateAdd("d", 1, localDate)
    If stopDate = "" And StartDate = "" Then
        shouldCount = True
    ElseIf stopDate = "" And DateDiff("d", localDate, StartDate) > 0 Then
        shouldCount = False
    ElseIf stopDate <> "" And DateDiff("d", localDate, stopDate) <= 0 Then
        shouldCount = False
    ElseIf DateDiff("d", localDate, stopDate) >= 0 And DateDiff("d", localDate, StartDate) > 0 Then
        shouldCount = False
    Else
        shouldCount = True
    End If
End Function

Private Function doSomethingWith(aValue)

    '// This is where you would compute a different value
    '// for use in the new workbook
    '// In this example, I simply add one to it.
    aValue = aValue

    doSomethingWith = aValue
End Function

Sub SortDictionary(dict As Object)
    Dim i As Long
    Dim key As Variant

    With CreateObject("System.Collections.SortedList")
        For Each key In dict
            .Add key, dict(key)
        Next
        dict.RemoveAll
        For i = 0 To .Keys.Count - 1
            dict.Add .GetKey(i), .Item(.GetKey(i))
        Next
    End With
End Sub

Private Sub PaperWiseCount_Click()

End Sub

Private Sub numOfDays()
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim dataRange As Range
    Dim thisBook As Workbook
    Dim newBook As Workbook
    Dim newRow As Integer
    Dim temp
    Dim stopData As Variant
    Dim localDate As String
    localDate = Date
    '// set your data range here
    Set dataRange = Sheet1.Range("A2:A400")

    '// create a new workbook
    '// Set newBook = Excel.Workbooks.Add
    Set dict = CreateObject("Scripting.Dictionary")
    '// loop through the data in book1, one column at a time
    Dim totalCount As Integer
    
    stopDateColumn = getBillColumnIndex("stop")
    startDateColumn = getBillColumnIndex("start")
    billColumn = getBillColumnIndex("bill")
    amountColumn = getBillColumnIndex("paid")
    paperColumn = getBillColumnIndex("paper")
    
    Set priceDataRange = Sheet4.Range("A1:P10")
    Set priceDict = CreateObject("Scripting.Dictionary")
    For colIndex = 2 To priceDataRange.Columns.Count
        For rowIndex = 2 To priceDataRange.Rows.Count
            With priceDataRange.Cells(rowIndex, 1)
            If .Value <> "" Then
                priceDict(priceDataRange.Cells(1, colIndex) & "/" & priceDataRange.Cells(rowIndex, 1)) = priceDataRange.Cells(rowIndex, colIndex)
            End If
            End With
        Next
    Next
    
    Set monthDaysDict = CreateObject("Scripting.Dictionary")
    monthDaysDict("Monday") = MonthWeekDays(CDate("8/3/2017"), 1)
    monthDaysDict("Tuesday") = MonthWeekDays(CDate("8/3/2017"), 2)
    monthDaysDict("Wednesday") = MonthWeekDays(CDate("8/3/2017"), 3)
    monthDaysDict("Thursday") = MonthWeekDays(CDate("8/3/2017"), 4)
    monthDaysDict("Friday") = MonthWeekDays(CDate("8/3/2017"), 5)
    monthDaysDict("Saturday") = MonthWeekDays(CDate("8/3/2017"), 6)
    monthDaysDict("Sunday") = MonthWeekDays(CDate("8/3/2017"), 7)
    
    For rowIndex = 1 To dataRange.Rows.Count
    '// For rowIndex = 1 To 4
        Dim stopDate As Date
        With dataRange.Cells(rowIndex, 2)

        '// ignore empty cells
        If .Value <> "" Then
            newRow = newRow + 1
            '// MsgBox (priceDict.Item(.Value))
            stopData = dataRange.Cells(rowIndex, stopDateColumn).Value
            If stopData <> "" And IsDate(stopData) = False Then
                stopData = ""
                dataRange.Cells(rowIndex, stopDateColumn).Value = ""
            End If
            startData = dataRange.Cells(rowIndex, startDateColumn).Value
            If startData <> "" And IsDate(startData) = False Then
                startData = ""
                dataRange.Cells(rowIndex, startDateColumn).Value = ""
            End If

            If stopData = "" And startData = "" Then
                dataRange.Cells(rowIndex, billColumn).Value = nb_days_month(CDate("8/3/2017"))
                dataRange.Cells(rowIndex, amountColumn).Value = getBillAmount(dataRange.Cells(rowIndex, paperColumn).Value, priceDict, monthDaysDict)
            ElseIf stopData = "" And startData <> "" Then
                Set daysDictBetweenDates = CreateObject("Scripting.Dictionary")
                Dim dayOfWeek As Long
                For dayOfWeek = 1 To 7
                    daysDictBetweenDates(WeekdayName(dayOfWeek)) = WkDays(CDate(startData), GetNowLast(CDate(startData)), dayOfWeek)
                Next
                dataRange.Cells(rowIndex, amountColumn).Value = getBillAmount(dataRange.Cells(rowIndex, paperColumn).Value, priceDict, daysDictBetweenDates)
                dataRange.Cells(rowIndex, billColumn).Value = TestDates(startData, GetNowLast(CDate(startData)))
                Set daysDictBetweenDates = Nothing
            ElseIf stopData <> "" And startData = "" Then
                Set daysDictBetweenDates = CreateObject("Scripting.Dictionary")
                For dayOfWeek = 1 To 7
                    daysDictBetweenDates(WeekdayName(dayOfWeek)) = WkDays(dhFirstDayInMonth(CDate(stopData)), CDate(stopData), dayOfWeek)
                Next
                dataRange.Cells(rowIndex, amountColumn).Value = getBillAmount(dataRange.Cells(rowIndex, paperColumn).Value, priceDict, daysDictBetweenDates)
                dataRange.Cells(rowIndex, billColumn).Value = TestDates(dhFirstDayInMonth(CDate(stopData)), stopData)
                Set daysDictBetweenDates = Nothing
            ElseIf stopData < startData Then
                numberOfDays = TestDates(dhFirstDayInMonth(CDate(stopData)), stopData) + TestDates(startData, GetNowLast(CDate(startData)))
                dataRange.Cells(rowIndex, billColumn).Value = numberOfDays
            ElseIf stopData > startData Then
                Set daysDictBetweenDates = CreateObject("Scripting.Dictionary")
                For dayOfWeek = 1 To 7
                    daysDictBetweenDates(WeekdayName(dayOfWeek)) = WkDays(CDate(startData), CDate(stopData), dayOfWeek)
                Next
                dataRange.Cells(rowIndex, amountColumn).Value = getBillAmount(dataRange.Cells(rowIndex, paperColumn).Value, priceDict, daysDictBetweenDates)
                dataRange.Cells(rowIndex, billColumn).Value = TestDates(startData, stopData)
            End If
        End If
        End With
    Next
    Set priceDict = Nothing
    Set monthDaysDict = Nothing
End Sub

Private Function getBillAmount(paperName As String, priceDict As Variant, monthDaysDict As Variant)
    Dim sum As Integer
    Dim paperKey As String
    Dim day As Variant
    sum = 0
    For Each day In monthDaysDict.Keys()
        paperKey = paperName & "/" & day
        sum = sum + priceDict(paperKey) * monthDaysDict(day)
    Next
    getBillAmount = sum
End Function

Private Function nb_days_month(dateInMonth As Date)
    
    date_test = dateInMonth 'Any date will do for this example

    nb_days_month = day(DateSerial(Year(date_test), Month(date_test) + 1, 1) - 1)
   
End Function

Function TestDates(pDate1 As Variant, pDate2 As Variant) As Long

    TestDates = DateDiff("d", pDate1, pDate2) + 1

End Function

Function GetNowLast(inputDate As Date) As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Function dhFirstDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the first day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhFirstDayInMonth = DateSerial(Year(dtmDate), _
    Month(dtmDate), 1)
End Function

Private Function loadPaperPriceMap() As Variant
    '// set your data range here
    Set dataRange = Sheet2.Range("A1:A200")
    Set priceDict = CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To dataRange.Rows.Count
        With dataRange.Cells(rowIndex, 2)
        '// ignore empty cells
        If .Value <> "" Then
            priceDict.Add dataRange.Cells(rowIndex, 1).Value, .Value
        End If
        End With
    Next
    Set loadPaperPriceMap = priceDict
End Function

Public Function TotalDaysInMonth(pDate As Variant, pDay As Integer)
    'Update 20140210
    Dim xindex As Integer
    Dim EndDate As Integer
    EndDate = day(DateSerial(Year(pDate), Month(pDate) + 1, 0))
    For xindex = 1 To EndDate
        If weekDay(DateSerial(Year(pDate), Month(pDate), xindex)) = pDay Then
            TotalDaysInMonth = TotalDaysInMonth + 1
        End If
    Next
End Function

Public Function TotalDaysBetweenDates(StartDate As Variant, EndDate As Variant, pDay As Integer)
    'Update 20140210
    Dim xindex As Integer
    For xindex = day(StartDate) To day(EndDate)
        If weekDay(DateSerial(Year(StartDate), Month(StartDate), xindex)) = pDay Then
            TotalDaysBetweenDates = TotalDaysBetweenDates + 1
        End If
    Next
End Function
