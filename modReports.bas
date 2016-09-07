Attribute VB_Name = "modReports"
Option Explicit

Private Const intOffset As Integer = 1 ' How much space from previous table to next
Private Const intDataOffset As Integer = 3 ' Column offset of data from Period

Public Sub CopyDataToTemplates(strPeriod As String, strTicker As String)
    
    Dim pRowStart As Integer
    Dim pRowEnd As Integer
    Dim pRowDiff As Integer
    Dim grandTotal As Range
    Dim strBrand As String
    Dim rngPeriodStart As Range
    Dim rngPrevPeriodEnd As Range
    Dim rngNewPeriodStart As Range
    Dim rngTotalSpend As Range
    Dim intPeriodRowSize As Integer
    Dim wsPeriodWorksheet As Worksheet
    Dim channelColl As New Collection
    Dim channelCount As Integer
    Dim i As Integer
    Dim intDeleted As Integer
    
    On Error GoTo Error_Handler
    
    ' SPEEDUP
    Dim boolEnableEvents As Boolean
    Dim boolScreenUpdating As Boolean
    Dim eCalc As XlCalculation
    boolEnableEvents = Application.EnableEvents
    boolScreenUpdating = Application.ScreenUpdating
    eCalc = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
  
    ' Store reference to this workbook
    Set wsPeriodWorksheet = ThisWorkbook.Worksheets(strPeriod)
    
    ' Define the basic named ranges
    Set rngPeriodStart = wsPeriodWorksheet.Range("PeriodStart")
    Set rngTotalSpend = wsPeriodWorksheet.Range("TotalSpend50")
    
    ' Row size of period table
    intPeriodRowSize = wsPeriodWorksheet.Range("PeriodEnd").row - rngPeriodStart.row
    
    ' Set date
    wsPeriodWorksheet.Range("Date") = Date
    
    ' Find and write Company Name
    wsPeriodWorksheet.Range("company").Value2 = FindCompanyForTicker(FindTickerOnCoverSheet(strTicker))
        
    ' Find Company for TickerCoMatch
    strBrand = FindCompanyForTickerCoMatch(FindTickerCoMatchOnTemplate(strTicker))
    wsPeriodWorksheet.Range("Company").Value2 = strBrand
    
    ' FILTER BY BRAND
    Call ReportFiltering_SingleBrand(strTicker, strBrand)
    
    ' GROUP BY PERIOD
    Call PivotTableGrouping(strPeriod)
    
    'Clear everything below first table's headers
    Range(rngPeriodStart.Offset(1).EntireRow, rngPeriodStart.Offset(wsPeriodWorksheet.UsedRange.Rows.Count).EntireRow).ClearContents
    Range(rngPeriodStart.Offset(1).EntireRow, rngPeriodStart.Offset(wsPeriodWorksheet.UsedRange.Rows.Count).EntireRow).Delete xlShiftUp
    
    ' Find top and bottom rows of Pivot table on Pivot sheet
    pRowStart = wksPivot.Range("Row_Labels").row
    pRowEnd = wksPivot.Range("GrandTotal").row
    pRowDiff = pRowEnd - pRowStart

    ' Copy from Pivot table
    wksPivot.Range(wksPivot.Range("Row_Labels").Offset(1, 0), wksPivot.Cells(pRowEnd - 1, wksPivot.Range("Row_Labels").Column)).Copy
    rngPeriodStart.Offset(1, 0).PasteSpecial xlValues
    wksPivot.Range(wksPivot.Range("Row_Labels").Offset(1, 1), wksPivot.Cells(pRowEnd - 1, wksPivot.Range("PivotTotalSpend50").Column)).Copy
    rngPeriodStart.Offset(1, 3).PasteSpecial xlValues
    
    ' Format table
    intDeleted = FormatDates(rngPeriodStart, strPeriod)
    Call FormatReportData(Range(rngPeriodStart.Offset(1, intDataOffset), Application.Intersect(wsPeriodWorksheet.Range("PeriodEnd").EntireRow, rngTotalSpend.EntireColumn)))
    
    ' Add YoY table
    Set rngPrevPeriodEnd = AddYoYBelowTable(rngPeriodStart.Offset(1, 0), wsPeriodWorksheet.Range("PeriodEnd"), rngTotalSpend.Offset(1, 0), "Entire Ticker YoY")
    
    ' Find channels and count
    Set channelColl = FindChannels(strTicker)
    channelCount = FindChannelCount(strTicker)
    
    ' Loop through the channels
    For i = 1 To channelCount
        If channelColl(i) <> "ALL" Then ' should't be, but check anyway
            
            ' Filter data for channel
            Call ReportFiltering_SingleChannel(strBrand, channelColl(i))

            ' Copy Headers and update title
            wksTemplate.Range("ChannelTable").Copy
            rngPrevPeriodEnd.Offset(intOffset + 1).PasteSpecial xlPasteAll
            rngPrevPeriodEnd.Offset(intOffset + 1, 1).Value2 = channelColl(i)
            rngPrevPeriodEnd.Offset(intOffset + 1, 5).Value2 = strTicker
            
            ' Create reference to next table's periods
            Set rngNewPeriodStart = rngPrevPeriodEnd.Offset(intOffset + 1 + wksTemplate.Range("ChannelTable").Rows.Count)
            
            ' Find top and bottom rows of Pivot table on Pivot sheet
            pRowStart = wksPivot.Range("Row_Labels").row
            pRowEnd = wksPivot.Range("GrandTotal").row
            pRowDiff = pRowEnd - pRowStart - 1
        
            ' Copy from Pivot table
            wksPivot.Range(wksPivot.Range("Row_Labels").Offset(1, 0), wksPivot.Cells(pRowEnd - 1, wksPivot.Range("Row_Labels").Column)).Copy
            rngNewPeriodStart.PasteSpecial xlValues
            wksPivot.Range(wksPivot.Range("Row_Labels").Offset(1, 1), wksPivot.Cells(pRowEnd - 1, wksPivot.Range("PivotTotalSpend50").Column)).Copy
            rngNewPeriodStart.Offset(0, intDataOffset).PasteSpecial xlValues
            
            ' Format table
            intDeleted = FormatDates(rngNewPeriodStart.Offset(-1, 0), strPeriod)
            pRowDiff = pRowDiff - intDeleted ' Update size after deleting rows
            ' Recreate reference to next table's periods
            Set rngNewPeriodStart = rngPrevPeriodEnd.Offset(intOffset + 1 + wksTemplate.Range("ChannelTable").Rows.Count)
            
            Call FormatReportData(Range(rngNewPeriodStart.Offset(0, intDataOffset), Application.Intersect(rngNewPeriodStart.Offset(pRowDiff - 1, intDataOffset).EntireRow, rngTotalSpend.EntireColumn)))
            
            ' Add YoY table
            Set rngPrevPeriodEnd = AddYoYBelowTable(rngNewPeriodStart, rngNewPeriodStart.Offset(pRowDiff - 1), Application.Intersect(rngTotalSpend.EntireColumn, rngNewPeriodStart.EntireRow), channelColl(i) & " YoY")
            
        End If
    Next
    
    ' Final formatting of the sheet
    Range(rngPeriodStart, Application.Intersect(wsPeriodWorksheet.Range("PeriodEnd").EntireRow, rngTotalSpend.EntireColumn)).Columns.AutoFit
    Range("PeriodStart").ColumnWidth = 23
    Range(Range("Company"), Range("Company").Offset(0, 1)).ColumnWidth = 20
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    Worksheets(strPeriod).Range("A1").Select
    
CleanUp:
    Application.Calculation = eCalc
    Application.EnableEvents = boolEnableEvents
    Application.ScreenUpdating = boolScreenUpdating
    Application.CutCopyMode = False
    Exit Sub
    
Error_Handler:
    GoTo CleanUp
End Sub

' Add year on year below table passed in the arguments, returns bottom range of Periods list
Private Function AddYoYBelowTable(rngTableTopLeft As Range, rngTableBotLeft As Range, rngTableTopRight As Range, strTitle As String) As Range

    Dim intHeadersSize As Integer
    Dim intTableRowSize As Integer, intTableColSize As Integer
    Dim strFormula As String
    Dim rngYoYTopLeft As Range
    Dim rngYoYBotRight As Range
    Dim rngReturnValue As Range
    

    
    ' Size of table
    intTableRowSize = rngTableBotLeft.row - rngTableTopLeft.row + 1
    intTableColSize = rngTableTopRight.Column - rngTableTopLeft.Column + 1
    
    ' Find size of headers range to be pasted
    intHeadersSize = wksTemplate.Range("YoYTable").Rows.Count
    
    ' Copy/Paste headers
    wksTemplate.Range("YoYTable").Copy
    rngTableBotLeft.Offset(1 + intOffset).PasteSpecial xlPasteAll
    
    ' Write Title
    rngTableBotLeft.Offset(1 + intOffset).Value2 = strTitle
    
    ' Beginning and end of YoY data
    Set rngYoYTopLeft = rngTableBotLeft.Offset(intHeadersSize + intOffset + 1, intDataOffset)
    Set rngYoYBotRight = rngYoYTopLeft.Offset(intTableRowSize - 1, intTableColSize - intDataOffset - 1)
    
    ' Copy/Paste Periods
    Range(rngTableTopLeft, rngTableBotLeft.Offset(0, intDataOffset)).Copy
    rngYoYTopLeft.Offset(0, -intDataOffset).PasteSpecial xlPasteAll
    
    ' Build YoY formula relative to range
    strFormula = BuildYoYFormula(rngTableTopLeft, rngTableBotLeft, rngTableBotLeft.Offset(intHeadersSize + intOffset + 1, 0), intDataOffset)
    
    ' Assign Formula
    Range(rngYoYTopLeft, rngYoYBotRight).Formula = strFormula
    
    ' Format
    Range(rngYoYTopLeft, rngYoYBotRight).NumberFormat = "0.00%"
    Range(rngYoYTopLeft, rngYoYBotRight).Font.Name = "Calibri Light"
    Range(rngYoYTopLeft, rngYoYBotRight).Font.Size = 14
    
    Set rngReturnValue = RemoveYoYEmptyRows(rngYoYTopLeft, rngYoYBotRight).Offset(0, -intDataOffset)
    'Application.Intersect(rngTableTopLeft.EntireColumn, rngYoYBotRight.Offset(-intRowsRemoved, 0).EntireRow)
    
    Set AddYoYBelowTable = rngReturnValue

End Function

' Remove empty rows from YoY. Receives TopLeft and BotRight ranges, and the number of cols to ignore. Returns bottom left range
Private Function RemoveYoYEmptyRows(rngTopLeft As Range, rngBotRight As Range) As Range
    Dim rngToCheck As Range
    Dim i As Integer
    Dim intOutput As Integer
    Dim intBotRow As Integer
    Dim intTopRow As Integer
    Dim intRowSize As Integer
    Dim rngNewTopLeft As Range
    
    ' New range one up, in case the top row gets deleted and we lose rngTopLeft
    Set rngNewTopLeft = rngTopLeft.Offset(-1)
    
    
    intBotRow = rngBotRight.row
    intTopRow = rngTopLeft.row
    intRowSize = rngBotRight.row - rngTopLeft.row + 1
    
    intOutput = 0
    i = 0
    Do While intBotRow >= intTopRow + i
        Set rngToCheck = Range(rngNewTopLeft.Offset(i + 1, 0), Application.Intersect(rngNewTopLeft.Offset(i + 1, 0).EntireRow, rngBotRight.EntireColumn))
        If WorksheetFunction.CountBlank(rngToCheck) = rngToCheck.Cells.Count Then
            rngToCheck.EntireRow.Delete xlShiftUp
            intOutput = intOutput + 1
            intBotRow = intBotRow - 1
        Else
            i = i + 1
        End If
    Loop
    Set RemoveYoYEmptyRows = rngNewTopLeft.Offset(intRowSize - intOutput, 0)
    
End Function

Private Function BuildYoYFormula(rngTableTopLeft As Range, rngTableBotLeft As Range, rngFirstRow As Range, intDataOffset As Integer) As String
    Dim strVal As String, strTblCol As String, strLkUp As String
    Dim strFormula As String
    
    strVal = rngFirstRow.Address(False, True)
    strTblCol = Range(rngTableTopLeft, rngTableBotLeft).Offset(0, intDataOffset).Address(True, False)
    strLkUp = Range(rngTableTopLeft, rngTableBotLeft).Address(True, True)
    
    '"=IFERROR(IF(ISNUMBER(INDEX(E$18:E$35,MATCH(VALUE(LEFT($B41,4))-1&RIGHT($B41,3),$B$18:$B$35,0))), (INDEX(E$18:E$35,MATCH($B41,$B$18:$B$35,0))-INDEX(E$18:E$35,MATCH(VALUE(LEFT($B41,4))-1&RIGHT($B41,3),$B$18:$B$35,0)))/INDEX(E$18:E$35,MATCH(VALUE(LEFT($B41,4))-1&RIGHT($B41,3),$B$18:$B$35,0)),""""),"""")"
    strFormula = "=IFERROR(IF(ISNUMBER("
    strFormula = strFormula & "INDEX(@tblcol@,MATCH(VALUE(LEFT(@val@,4))-1&RIGHT(@val@,3),@lkup@,0))),"
    strFormula = strFormula & "(INDEX(@tblcol@,MATCH(@val@,@lkup@,0))-INDEX(@tblcol@,MATCH(VALUE(LEFT(@val@,4))-1&RIGHT(@val@,3),@lkup@,0)))"
    strFormula = strFormula & "/INDEX(@tblcol@,MATCH(VALUE(LEFT(@val@,4))-1&RIGHT(@val@,3),@lkup@,0)),""""),"""")"
    
    strFormula = Replace(strFormula, "@val@", strVal)
    strFormula = Replace(strFormula, "@tblcol@", strTblCol)
    strFormula = Replace(strFormula, "@lkup@", strLkUp)
    
    BuildYoYFormula = strFormula
    
End Function

Private Function FindTickerOnCoverSheet(strTicker As String) As Range
    Dim rngList As Range
    
    Set rngList = wksCover.Range("TickerList")
    
    Set FindTickerOnCoverSheet = rngList.Find(what:=strTicker, LookIn:=xlValues, _
    lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False)
    
End Function

Private Function FindCompanyForTicker(rngTicker As Range) As Range
    Set FindCompanyForTicker = Application.Intersect(rngTicker.EntireRow, wksCover.Range("Company").EntireColumn)
End Function

Private Function FindTickerCoMatchOnTemplate(strTicker As String) As Range
    Dim rngList As Range

    ' Check list not empty
    If wksTemplate.Range("TickerCoMatch").Address <> wksTemplate.Range("TickerCoMatchEnd").Address Then
        Set rngList = Range(wksTemplate.Range("TickerCoMatch"), wksTemplate.Range("TickerCoMatchEnd"))
    End If
    Set FindTickerCoMatchOnTemplate = rngList.Find(what:=strTicker, LookIn:=xlValues, _
    lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False)
    
End Function

Private Function FindCompanyForTickerCoMatch(rngTickerCM As Range) As Range
    Set FindCompanyForTickerCoMatch = Application.Intersect(rngTickerCM.EntireRow, wksTemplate.Range("DropDownList").EntireColumn)
End Function


Function FormatDates(rStart As Range, period As String) As Integer
    Dim rng As Range
    Dim str As String
    Dim pRowStart As Integer
    Dim pRowEnd As Integer
    Dim i As Integer
    Dim Flag As Boolean
    Dim intDelCnt As Integer
    
    On Error GoTo Error_Handler
    
    ' SPEEDUP
    Dim boolEnableEvents As Boolean
    Dim boolScreenUpdating As Boolean
    Dim eCalc As XlCalculation
    boolEnableEvents = Application.EnableEvents
    boolScreenUpdating = Application.ScreenUpdating
    eCalc = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    intDelCnt = 0
    i = 1
    Flag = True
    While Flag = True
    If rStart.Offset(i, 0) <> "" Then
        'if there is still data go on
        i = i + 1
    Else
        'if there is no more data left stop the loop
        Flag = False
    End If
    Wend

    pRowEnd = i - 1
    
    If period = "Quarterly" Then
        For Each rng In Range(rStart, rStart.Offset(i, 0))
            If InStr(rng.Value2, "Qtr") = 0 Then
                str = rng.Value2
            ElseIf str <> "" Then
                rng.Value2 = str & " " & rng.Value2
                rng.Value2 = Replace(rng.Value2, "Qtr", "Q")
                If InStr(rng.Value2, "Q1") > 0 Then
                    rng.Offset(0, 1) = CDate("1/1/" & str)
                    rng.Offset(0, 2) = CDate("3/31/" & str)
                ElseIf InStr(rng.Value2, "Q2") > 0 Then
                    rng.Offset(0, 1) = CDate("4/1/" & str)
                    rng.Offset(0, 2) = CDate("6/30/" & str)
                ElseIf InStr(rng.Value2, "Q3") > 0 Then
                    rng.Offset(0, 1) = CDate("7/1/" & str)
                    rng.Offset(0, 2) = CDate("9/30/" & str)
                ElseIf InStr(rng.Value2, "Q4") > 0 Then
                    rng.Offset(0, 1) = CDate("10/1/" & str)
                    rng.Offset(0, 2) = CDate("12/31/" & str)
                End If
            End If
        Next
        For i = pRowEnd To 1 Step -1
            If InStr(rStart.Offset(i, 0).Value2, "Q") = 0 And rStart.Offset(i, 0).Value2 <> "" Then
            rStart.Offset(i, 0).EntireRow.Delete
            intDelCnt = intDelCnt + 1
            Else
            End If
        Next
    ElseIf period = "Monthly" Then
        For Each rng In Range(rStart, rStart.Offset(pRowEnd, 0))
            If InStr(rng.Value2, "20") > 0 Then
                str = rng.Value2
            ElseIf str <> "" Then
                If rng.Value2 = "Jan" Then
                    'rng.Text = "01"
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/31/" & str)
                    rng.Value2 = str & "-01"
                ElseIf rng.Value2 = "Feb" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/28/" & str)
                    rng.value = str & "-02"
                ElseIf rng.Value2 = "Mar" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/31/" & str)
                    rng.Value2 = str & "-03"
                ElseIf rng.Value2 = "Apr" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/30/" & str)
                    rng.Value2 = str & "-04"
                ElseIf rng.Value2 = "May" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/31/" & str)
                    rng.Value2 = str & "-05"
                ElseIf rng.Value2 = "Jun" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/30/" & str)
                    rng.Value2 = str & "-06"
                ElseIf rng.Value2 = "Jul" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/31/" & str)
                    rng.Value2 = str & "-07"
                ElseIf rng.Value2 = "Aug" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/31/" & str)
                    rng.Value2 = str & "-08"
                ElseIf rng.Value2 = "Sep" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/30/" & str)
                    rng.Value2 = str & "-09"
                ElseIf rng.Value2 = "Oct" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/31/" & str)
                    rng.Value2 = str & "-10"
                ElseIf rng.Value2 = "Nov" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/30/" & str)
                    rng.Value2 = str & "-11"
                ElseIf rng.Value2 = "Dec" Then
                    
                    rng.Offset(0, 1) = CDate(rng.Value2 & "/1/" & str)
                    rng.Offset(0, 2) = CDate(rng.Value2 & "/31/" & str)
                    rng.Value2 = str & "-12"
                End If
                
                
            End If
        Next
            For i = pRowEnd To 1 Step -1
            If InStr(rStart.Offset(i, 0).Value2, "-") = 0 And rStart.Offset(i, 0).Value2 <> "" Then
            rStart.Offset(i, 0).EntireRow.Delete
            intDelCnt = intDelCnt + 1
            Else
            End If
        Next
    ElseIf period = "Weekly" Then
        For Each rng In Range(rStart.Offset(1, 0), rStart.Offset(pRowEnd, 0))
            If rng.Value2 <> "" Then
                rng.Offset(0, 1) = CDate(Split(rng.Value2, " -")(0))
                rng.Offset(0, 2) = CDate(Split(rng.Value2, "- ")(1))
                rng.Value2 = CStr(Right(rng.Offset(0, 1), 4)) & "-W" & CStr(Int(((rng.Offset(0, 1) - DateSerial(Year(rng.Offset(0, 1)), 1, 0)) + 6) / 7))
            End If
        Next
        For i = pRowEnd To 1 Step -1
            If InStr(rStart.Offset(i, 0).Value2, "W") = 0 And rStart.Offset(i, 0).Value2 <> "" Then
            rStart.Offset(i, 0).EntireRow.Delete
            Else
            End If
        Next
    ElseIf period = "Daily" Then
        For Each rng In Range(rStart.Offset(1, 0), rStart.Offset(pRowEnd, 0))
            If InStr(rng.Value2, "-") = 0 And rng.Value2 <> "" Then
                str = rng.Value2
            ElseIf str <> "" Then
                rng = CStr(Left(CStr(CDate(rng.Value2)), Len(CStr(CDate(rng.Value2))) - 4)) & str
            End If
        Next
        For i = pRowEnd To 1 Step -1
            If InStr(rStart.Offset(i, 0), "/") = 0 And rStart.Offset(i, 0).Value2 <> "" Then
            rStart.Offset(i, 0).EntireRow.Delete
            intDelCnt = intDelCnt + 1
            Else
            End If
        Next
    End If
    
    Range(rStart.Offset(1, 0), rStart.Offset(pRowEnd, 2)).Font.Name = "Calibri Light"
    Range(rStart.Offset(1, 0), rStart.Offset(pRowEnd, 2)).Font.Size = 14
    
    FormatDates = intDelCnt

CleanUp:

    Application.Calculation = eCalc
    Application.EnableEvents = boolEnableEvents
    Application.ScreenUpdating = boolScreenUpdating

    Exit Function
    
Error_Handler:
    GoTo CleanUp

End Function
Sub FormatReportData(rngTable As Range)
    
    wksTemplate.Range("RowForFormatting").Copy
    rngTable.PasteSpecial xlPasteFormats
    
    Application.CutCopyMode = False
End Sub


Public Sub PivotTableGrouping(period As String)
    'parameters are "Quarterly", "Monthly", "Weekly", and "Daily"
    Dim pt As PivotTable
    Dim ptField As PivotField
    Dim rPTRange As Range
    Dim groupBy As Integer
    Dim arr(1 To 7) As Boolean ' Array initializes as False
    'Grouping in Pivot Tables uses an array of booleans to determine which
    'periods are on or off
    
    '''''''''''''''''''''''
    'Period Settings
    '1 = Second
    '2 = Minute
    '3 = Hour
    '4 = Day
    '5 = Month
    '6 = Quarter
    '7 = Year
    '''''''''''''''''''''''
    groupBy = 1 ' default for all except Week
    
    arr(7) = True ' Turn on Yearly Grouping
    
    If period = "Quarterly" Then
        arr(6) = True
    ElseIf period = "Monthly" Then
        arr(5) = True
    ElseIf period = "Weekly" Then ' weekly doesn't have a period, so we use day and group by 7
        arr(4) = True
        arr(7) = False
        groupBy = 7
    ElseIf period = "Daily" Then
        arr(4) = True
    Else
        arr(5) = True ' If a mistake is made, default to Month
    End If
    
    Set pt = wksPivot.PivotTables("PivotTable2")
    Set ptField = pt.RowFields("begin_date")
    Set rPTRange = ptField.DataRange.Cells(1, 1)
    
    
    rPTRange.Group Start:=True, End:=True, Periods:=arr, By:=groupBy
    
    
End Sub

Public Function PivotCount(str As String) As Integer
    'counts the number of unique Brands
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim Count As Integer
    
    Set pt = wksPivot.PivotTables("PivotTable2")
    Set pf = pt.PivotFields(str)
    
    PivotCount = pf.PivotItems.Count
    
End Function


Public Sub ReportFiltering_Single(filter As String)
    'PURPOSE: Filter on a single item with the Report Filter field
    'SOURCE: www.TheSpreadsheetGuru.com
    'argument added by William Fehringer
    
    Dim pf As PivotField
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Brand")
    
    pf.ClearAllFilters
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Ticker")
    
    'Clear Out Any Previous Filtering
    pf.ClearAllFilters
    
    'Filter on argument
    pf.CurrentPage = filter
    
End Sub

    'PURPOSE: Filter on a single item with the Report Filter field
    'SOURCE: www.TheSpreadsheetGuru.com
    'argument added by William Fehringer
Public Sub ReportFiltering_SingleBrand(ticker As String, Optional brand As String = "ALL", Optional channel As String = "ALL")

    
    Dim pf As PivotField
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Ticker")
    
    pf.ClearAllFilters
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Brand")
    
    'Clear Out Any Previous Filtering
    pf.ClearAllFilters
    
    'Filter on argument
    pf.CurrentPage = brand
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("channel")
    If channel <> "NONE" Then
        pf.CurrentPage = channel
    Else
        pf.ClearAllFilters
    End If
End Sub
Public Sub ReportFiltering_SingleTicker(ticker As String, Optional brand As String = "ALL", Optional channel As String = "ALL")
    'PURPOSE: Filter on a single item with the Report Filter field
    'SOURCE: www.TheSpreadsheetGuru.com
    'argument added by William Fehringer
    
    Dim pf As PivotField
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Brand")
    
    pf.ClearAllFilters
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Ticker")
    
    'Clear Out Any Previous Filtering
    pf.ClearAllFilters
    
    'Filter on argument
    pf.CurrentPage = ticker
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("channel")
    If channel <> "NONE" Then
        pf.CurrentPage = channel
    Else
        pf.ClearAllFilters
    End If
End Sub

Public Sub ReportFiltering_SingleChannel(brand As String, channel As String)
    'PURPOSE: Filter on a single item with the Report Filter field
    'SOURCE: www.TheSpreadsheetGuru.com
    'argument added by William Fehringer
    
    Dim pf As PivotField
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Ticker")
    
    pf.ClearAllFilters
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Brand")
    
    'Clear Out Any Previous Filtering
    pf.ClearAllFilters
    
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("channel")
    pf.ClearAllFilters
    
    If channel <> "NONE" Then
        pf.CurrentPage = channel
    Else
        pf.ClearAllFilters
    End If
    
    'Filter on argument
    Set pf = wksPivot.PivotTables("PivotTable2").PivotFields("Brand")
    pf.CurrentPage = brand
    

End Sub
Public Sub RefreshingPivotTables()
    'PURPOSE: Shows various ways to refresh Pivot Table Data
    'SOURCE: www.TheSpreadsheetGuru.com
    
    'Refresh A Single Pivot Table
    wksPivot.PivotTables("PivotTable2").PivotCache.Refresh
    wksPivot.PivotTables("PivotTable2").RefreshTable
    
End Sub

Public Sub RefreshDropDownLists()
    Dim LastRow1 As Integer
    Dim LastRow2 As Integer
    Dim countTicker As Integer
    Dim countBrand As Integer
    Dim i As Integer
    Dim rng As Range
    
    On Error GoTo Error_Handler
    
    ' DISABLE EVENTS AND SCREENUPDATE !
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call RefreshingPivotTables
    
    countTicker = PivotCount("Ticker")
    countBrand = PivotCount("Brand")
    
    ' Clear and Populate DropDownTicker list
    If wksTemplate.Range("DropDownTickerEnd").row > wksTemplate.Range("DropDownTicker").row Then ' List is not empty
        Range(wksTemplate.Range("DropDownTicker").Offset(1, 0), wksTemplate.Range("DropDownTickerEnd")).ClearContents
    End If
    wksTemplate.Activate
    For i = 1 To countTicker
        wksTemplate.Range("DropDownTicker").Offset(i, 0).value = wksPivot.PivotTables("PivotTable2").PivotFields("Ticker").PivotItems(i).value
    Next
    
    ' Clear DropDownList and TickerCoMatch
    If wksTemplate.Range("DropDownListEnd").row > wksTemplate.Range("DropDownList").row Then ' List is not empty
        Range(wksTemplate.Range("DropDownList").Offset(1, 0), wksTemplate.Range("DropDownListEnd")).ClearContents
    End If
    If wksTemplate.Range("TickerCoMatchEnd").row > wksTemplate.Range("TickerCoMatch").row Then ' List is not empty
        Range(wksTemplate.Range("TickerCoMatch").Offset(1, 0), wksTemplate.Range("TickerCoMatchEnd")).ClearContents
    End If
    
    ' Populate DropDownList and TickerCoMatch
    For i = 1 To countBrand
        If wksPivot.PivotTables("PivotTable2").PivotFields("Brand").PivotItems(i) <> "ALL" Then
        wksTemplate.Range("DropDownList").Offset(i, 0).value = wksPivot.PivotTables("PivotTable2").PivotFields("Brand").PivotItems(i).value
        
        Set rng = wksUserData.Columns(7).Find(what:=wksPivot.PivotTables("PivotTable2").PivotFields("Brand").PivotItems(i).value, _
            LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:=False)
            wksTemplate.Range("TickerCoMatch").Offset(i, 0).value = rng.Offset(0, -1).value
        End If
    Next
    
    
    For i = countBrand To 1 Step -1
        If wksTemplate.Range("DropDownList").Offset(i, 0) = "" Then
            wksTemplate.Range("DropDownList").Offset(i, 0).Delete Shift:=xlUp
            wksTemplate.Range("TickerCoMatch").Offset(i, 0).Delete Shift:=xlUp
        End If
    Next
    Call Update_DataValidation(countTicker)
CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    Exit Sub
    
Error_Handler:
    GoTo CleanUp
End Sub

'This function will return the total count of rows in the
'drop down list(data validation) source
Private Function Get_Count() As Integer
    'counter
    Dim i As Integer
    'determines if the we have reached the end
    Dim Flag As Boolean
    
    i = 1
    Flag = True
    While Flag = True
        If wksTemplate.Range("DropDownTicker").Offset(1, 0).Cells(i, 1) <> "" Then
            'if there is still data go on
            i = i + 1
        Else
            'if there is no more data left stop the loop
            Flag = False
        End If
    Wend

    'return the total row count
    Get_Count = i - 1
End Function

'the function below updates the source range for the data validation
'based on the number of rows provided by the input
Private Sub Update_DataValidation(ByVal intRow As Integer)
    'the reference string to the source range
    Dim strSourceRange As String
    Dim StartRow As Integer
    StartRow = wksTemplate.Range("DropDownTicker").row
    strSourceRange = "=Template!B31:B" + Strings.Trim(str(StartRow + intRow))
    
    With wksQuarterly.Range("TickerDropdown").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=strSourceRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With wksMonthly.Range("TickerDropdown").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=strSourceRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With wksWeekly.Range("TickerDropdown").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=strSourceRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With wksDaily.Range("TickerDropdown").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=strSourceRange
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub
'Test sub, likely not needed
Private Sub GenerateReportSheets()
    'creates individual sheets with report for each Brand
    'Code in NewWorkSheet event adds "XSHOW_" to name of sheets generated by pivot table
    
    Application.DisplayAlerts = False
    
    Application.CommandBars(13).Controls("&PivotTable").Controls("Show Report Filter &Pages...").Execute
    'Application.SendKeys ("~")
    Application.DisplayAlerts = True
End Sub

Private Sub DeleteGeneratedSheets()
    'Removes any sheet generated by 'parseTickers' so they can be replaces with updated data
    
    Dim ws As Worksheet
    Dim PW As String
    Dim rng As Range
    Dim Test As String
    
    Dim Resp As Long
    Dim ShowCount As Long
    
    ShowCount = 0
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, 5)) = "XSHOW" Then
            ShowCount = ShowCount + 1
        End If
    Next ws
    
    If ShowCount > 0 Then
        
        Application.DisplayAlerts = False
        For Each ws In ThisWorkbook.Worksheets
            If UCase(Left(ws.Name, 5)) = "XSHOW" Then
                ws.Delete
            End If
        Next ws
        
    End If
    Set ws = Nothing
End Sub



Sub parseTickers()
'This sub splits the sheet Ticker Brand into separate sheets for each ticker

    Dim lr As Long
    Dim ws As Worksheet
    Dim vcol, i As Integer
    Dim icol As Long
    Dim myarr() As Variant
    Dim title As String
    Dim titlerow As Integer
    
    Call DeleteGeneratedSheets ' clear out the old sheets
    
    
    vcol = 6 ' splitting on the second column
    Set ws = wksTickerBrand
    lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).row
    title = "A1:AM1"
    titlerow = ws.Range(title).Cells(1).row
    icol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column + 1
    ws.Cells(1, icol) = "Unique"
    For i = 2 To lr
        On Error Resume Next
        If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
            ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
        End If
    Next
    
    myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    ws.Columns(icol).Clear
    
    For i = 2 To UBound(myarr)
        ws.Range(title).AutoFilter field:=vcol, Criteria1:=myarr(i) & ""
        If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
            Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "XSHOW_Ticker" & myarr(i) & ""
        Else
            Sheets(myarr(i) & "").Move After:=Worksheets(Worksheets.Count)
        End If
        ws.Range("A" & titlerow & ":A" & lr).EntireRow.Copy Sheets("XSHOW_Ticker" & myarr(i) & "").Range("A1")
        Sheets("XSHOW_Ticker" & myarr(i) & "").Columns.AutoFit
    Next
    Erase myarr
    On Error GoTo 0
    Application.CutCopyMode = False
    ws.AutoFilterMode = False
    ws.Activate
End Sub

Sub tickerToPivotSource(strUser As String)

    Dim ws As Worksheet
    Dim wsDest As Worksheet
    Dim wsTicker As Worksheet
    Dim sheetnames As String
    Dim shtArr() As String
    Dim i As Integer
    Dim strPW As String
    
    sheetnames = ""
    
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "XSHOW_" Then
            strPW = UserDecryptSheetPassword(strUser, Replace(Right(ws.Name, Len(ws.Name) - 6), "Ticker", ""))
            If strPW <> "" Then
                sheetnames = sheetnames & "," & ws.Name
            End If
        End If
    Next

    'sheetnames = wksAccess.Range("D7")
    ' Populate with list allowed to
''''''''''''''''''''''''''''''''''''''''''''''
    sheetnames = Right(sheetnames, Len(sheetnames) - 1)


    shtArr = Split(Trim(sheetnames), ",", -1, vbTextCompare)
    
    Set wsDest = wksUserData
    wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(wsDest.Cells(Rows.Count, "A").End(xlUp).row, wsDest.Range("A1").End(xlToRight).Column)).ClearContents
    
    
    For i = LBound(shtArr) To UBound(shtArr)
        If WorksheetExists(shtArr(i)) Then ' Not all tickers present in data
            Set wsTicker = Worksheets(shtArr(i))
            wsTicker.Range(wsTicker.Cells(2, 1), wsTicker.Cells(wsTicker.Cells(Rows.Count, "A").End(xlUp).row, wsTicker.Range("A1").End(xlToRight).Column)).Copy
            wsDest.Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial xlPasteValues
        End If
    Next
    wsDest.Columns(2).NumberFormat = "mm/dd/yyyy"
    wsDest.Columns(3).NumberFormat = "mm/dd/yyyy"
    wsDest.Columns(4).NumberFormat = "mm/dd/yyyy"

    Erase shtArr

End Sub
Sub PivotSourceRefresh()
    Dim SrcData As String
    Dim PivTbl As PivotTable

   SrcData = "UserData!" & Range(wksUserData.Cells(1, 1), wksUserData.Cells(wksUserData.Cells(Rows.Count, "A").End(xlUp).row, wksUserData.Range("A1").End(xlToRight).Column)).Address(ReferenceStyle:=xlR1C1)

    On Error Resume Next
        Set PivTbl = wksPivot.PivotTables("PivotTable2")
    On Error GoTo 0
        If PivTbl Is Nothing Then
            'create pivot
        Else
            PivTbl.ChangePivotCache ActiveWorkbook. _
            PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData, _
            Version:=xlPivotTableVersion14)

            PivTbl.RefreshTable
        End If

End Sub

Sub CoverReport()
    
    Dim Last As Integer
    Dim First As Integer
    Dim i As Integer
    Dim ticker As String
    Dim company As String
    Dim rngCover As Range
    Dim oldReportDate As Date
    Dim newReportDate As Date
    Dim oRange As Range
    Dim tempVal As String
    Dim ListLRow As Integer
    Dim ListRange As Range
    
    On Error GoTo Error_Handler
    
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    First = wksCover.Range("CoverHeaders").row
    Last = wksCover.Cells(Rows.Count, Range("CoverHeaders").Cells(1, 1).Column).End(xlUp).row
    ListLRow = wksCoverData.Cells(Rows.Count, Range("A1").Cells(1, 1).Column).End(xlUp).row
    Set ListRange = Range(wksCoverData.Range("A2"), wksCoverData.Range("A" & ListLRow))
    'Clear Old Report
    If Last > First Then
        Set rngCover = Range(Range("CoverHeaders").Offset(1, 0), Range("CoverHeaders").Offset(Last - First, 0))
        rngCover.ClearContents
    Else
    End If
    
    oldReportDate = wksCover.Range("CoverReportDate")
    'newReportDate = wksUserData.Range("D2")
    wksCover.Range("CoverDateRange") = getDateRange ' Get the range of dates from UserData
    wksCover.Range("CoverReportDate") = Date ' Today's Date, perhaps updated by DB data during workbook update
    
    'If newReportDate > oldReportDate Then
    
    '    wksCover.Range("CoverReportID").Value = wksCover.Range("CoverReportID").Value + 1 'Report ID - Updated/Incremented by workbook update?
    
    'End If
    
    First = wksTemplate.Range("DropDownTicker").row
    Last = wksTemplate.Cells(Rows.Count, wksTemplate.Range("DropDownTicker").Column).End(xlUp).row
    
    If Last > First Then
        For i = 1 To Last - First
            'copy ticker
     
    
            'match ticker to company
        ticker = wksTemplate.Range("DropDownTicker").Offset(i, 0).Value2
                Set oRange = ListRange.Find(what:=ticker, LookIn:=xlValues, _
        lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False)
    
        If Not oRange Is Nothing Then
            Range(oRange, oRange.Offset(0, 7)).Copy
            Range(wksCover.Range("CoverHeaders").Cells(1, 1).Offset(i, 0), wksCover.Range("CoverHeaders").Cells(1, 1).Offset(i, 0)).PasteSpecial xlPasteValues
    '        wksCover.Range("CoverHeaders").Cells(1, 2).Offset(i, 0).Value2 = wksTemplate.Columns(3).Find(what:=ticker, _
    '                            LookIn:=xlValues, _
    '                lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    '                MatchCase:=False, SearchFormat:=False).Offset(0, 1)
    '
    '        ' hyperlink company to Summary sheet ''Incomplete
    '        company = wksCover.Range("CoverHeaders").Cells(1, 2).Offset(i, 0).Value2
    '
    '        'Initiation Date - If same as report date, easy to update
    '        'wksCover.Range("CoverHeaders").Cells(1, 8).Offset(i, 0).Value2
    '
    '        'count brands in ticker
            wksCover.Range("CoverHeaders").Cells(1, 9).Offset(i, 0).Value2 = countTickerBrand(wksTemplate.Range("DropDownTicker").Offset(i, 0))
            
            'count channels
            wksCover.Range("CoverHeaders").Cells(1, 10).Offset(i, 0).Value2 = FindChannelCount(wksCover.Range("CoverHeaders").Cells(1, 1).Offset(i, 0).Value2)
        End If
        Next
    Else
    End If
    

    
CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    Exit Sub
    
Error_Handler:
    GoTo CleanUp
End Sub

Sub SummaryReport()

    
    Dim minSeparation As Integer
    Dim brandCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim firstCell As Range
    Dim lastCell As Range
    Dim Last As Integer
    Dim pRowStart As Integer
    Dim pRowEnd As Integer
    Dim grandTotal As Range
    Dim delay As String
    Dim Flag2 As Boolean
    Dim rng As Range
    Dim str As String
    Dim chtObj As ChartObject
    Dim chtTAS As ChartObject
    Dim chtYoY As ChartObject
    
    On Error GoTo Error_Handler
    
    Dim boolEnableEvents As Boolean
    Dim boolScreenUpdating As Boolean
    Dim eCalc As XlCalculation
    boolEnableEvents = Application.EnableEvents
    boolScreenUpdating = Application.ScreenUpdating
    eCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    

    
    wksSummary.Activate
    minSeparation = 23 ' number of rows between each Company
    delay = "0:00:03" ' value can be adjusted down for better performance, at risk of causing procedure to fail
    
    brandCount = countBrand
    wksSummary.Range("SummaryDateRange") = getDateRange
    wksSummary.Range("SummaryReportDate") = Date
    
    For Each chtObj In wksSummary.ChartObjects
        chtObj.Delete
    Next
    
    wksSummary.Rows(wksSummary.Range("SummaryStart").Offset(1, 0).row & ":" & wksSummary.UsedRange.Count).Delete
    Last = wksTemplate.Cells(Rows.Count, wksTemplate.Range("DropDownList").Column).End(xlUp).row
    
    'In case Template is not sorted properly
    With wksTemplate.Range(wksTemplate.Range("TickerCoMatch"), wksTemplate.Cells(Last, wksTemplate.Range("DropDownList").Column))
        .Sort Key1:=.Cells(1, 1), Order1:=xlAscending, _
               key2:=.Cells(1, 2), order2:=xlAscending, _
                Header:=xlGuess
    End With
    
    Set chtTAS = wksTemplate.ChartObjects("TotalAdjustedSpend")
    DoEvents
    Set chtYoY = wksTemplate.ChartObjects("YOYGrowth")
    DoEvents
    For i = 1 To brandCount
        'copy summary header for each Company
        wksTemplate.Range("SummaryTable").Copy
        wksSummary.Range("SummaryStart").Offset(1 + ((i - 1) * minSeparation), 0).PasteSpecial
        wksSummary.Range("SummaryStart").Offset(1 + ((i - 1) * minSeparation), 1) = wksTemplate.Range("TickerCoMatch").Offset(i, 0)
        wksSummary.Range("SummaryStart").Offset(1 + ((i - 1) * minSeparation), 3) = wksTemplate.Range("DropDownList").Offset(i, 0)
        
        '''''''''''''''''''''''
        ' Section for copying
        ' Report data
        '''''''''''''''''''''''
        
        Call ReportFiltering_SingleChannel(wksTemplate.Range("DropDownList").Offset(i, 0), "ALL")
        PivotTableGrouping ("Quarterly")
        pRowStart = wksPivot.Range("Row_Labels").row
    
        On Error Resume Next
    
        With wksPivot
        
            Set grandTotal = .Columns(1).Find(what:="Grand Total", After:=.Cells(1, 1), LookIn:=xlValues, lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        
            On Error GoTo 0
        
        End With
    
        pRowEnd = grandTotal.row
        
        wksPivot.Range(wksPivot.Range("Row_Labels").Offset(1, 0), wksPivot.Cells(pRowEnd - 1, wksPivot.Range("Row_Labels").Column)).Copy
        wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).PasteSpecial xlValues
        
        wksPivot.Range(wksPivot.Range("Row_Labels").Offset(1, 6), wksPivot.Cells(pRowEnd - 1, wksPivot.Range("Row_Labels").Offset(0, 6).Column)).Copy
        wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 1).PasteSpecial xlValues
        
            j = 1
    Flag2 = True
    While Flag2 = True
    If wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).Cells(j, 1) <> "" Then
        'if there is still data go on
        j = j + 1
    Else
        'if there is no more data left stop the loop
        Flag2 = False
    End If
    Wend
    
    pRowEnd = j - 1
        
        
            For Each rng In Range(wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0), wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).Cells(pRowEnd, 1))
                If InStr(rng.Value2, "Qtr") = 0 Then
                    str = rng.Value2
                ElseIf str <> "" Then
                    rng.Value2 = str & " " & rng.Value2
                    rng.Value2 = Replace(rng.Value2, "Qtr", "Q")
                    
                    Dim strFormula As String
                    Dim strVal As String, strVals As String, strQtr As String, strQtrs As String
                    strVal = rng.Offset(0, 1).Address(False, True)
                    strVals = Range(wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0), wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).Cells(pRowEnd, 1)).Offset(0, 1).Address(True, True)
                    strQtr = rng.Address(False, True)
                    strQtrs = Range(wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0), wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).Cells(pRowEnd, 1)).Address(True, True)
                    strFormula = "=IFERROR((@val@-INDEX(@vals@,MATCH(VALUE(LEFT(@qtr@,4))-1&RIGHT(@qtr@,3),@qtrs@,0)))/INDEX(@vals@,MATCH(VALUE(LEFT(@qtr@,4))-1&RIGHT(@qtr@,3),@qtrs@,0)),"""")"
                    
                    strFormula = Replace(strFormula, "@val@", strVal)
                    strFormula = Replace(strFormula, "@vals@", strVals)
                    strFormula = Replace(strFormula, "@qtr@", strQtr)
                    strFormula = Replace(strFormula, "@qtrs@", strQtrs)
                    
                    rng.Offset(0, 1).NumberFormat = "$#,##0"
                    rng.Offset(0, 2).Formula = strFormula
                    rng.Offset(0, 2).NumberFormat = "0.00%"
                        
                    
                End If
            Next
            
            For j = pRowEnd To 0 Step -1
                If InStr(wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).Cells(j, 1).Value2, "Q") = 0 And wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).Cells(j, 1).Value2 <> "" Then
                wksSummary.Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).Cells(j, 1).EntireRow.Delete
                Else
                End If
            Next
        ''''''''''''''''''''
        ' Copying Charts
        ''''''''''''''''''''
        
        'wksTemplate.ChartObjects("TotalAdjustedSpend").Activate
        chtTAS.Activate
        chtTAS.Copy 'Destination:=wksSummary.Range("SummaryStart").Offset(3 + (i - 1) * minSeparation, 4)
        DoEvents
        Application.Wait (Now + TimeValue(delay))
        wksSummary.Activate
        wksSummary.Range("SummaryStart").Offset(3 + (i - 1) * minSeparation, 4).Select
        'Application.Wait (Now + TimeValue(delay))
        wksSummary.Paste
        'Application.Wait (Now + TimeValue(delay))
        ' Update Data source for chart
        ActiveChart.SetSourceData Source:=wksSummary.Range(Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0), _
                            Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 1).End(xlDown)), PlotBy:=xlColumns
        ' Rename chart with unique name
        wksSummary.ChartObjects("TotalAdjustedSpend").Name = wksTemplate.Range("DropDownList").Offset(i, 0) & " TAS"
        Application.CutCopyMode = False
        ' Repeat for YOY Chart
        chtYoY.Activate
        chtYoY.Copy
        DoEvents
        Application.Wait (Now + TimeValue(delay))
        wksSummary.Activate
        wksSummary.Range("SummaryStart").Offset(3 + (i - 1) * minSeparation, 8).Select
        wksSummary.Paste
        ActiveChart.SetSourceData Source:=Union(wksSummary.Range(Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0), _
                            Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 0).End(xlDown)), wksSummary.Range(Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 2), _
                            Range("SummaryStart").Offset(5 + ((i - 1) * minSeparation), 2).End(xlDown))), PlotBy:=xlColumns
        wksSummary.ChartObjects("YOYGrowth").Name = wksTemplate.Range("DropDownList").Offset(i, 0) & " YOY"
        
        ' Clear and give system time to empty clipboard
        Application.CutCopyMode = False
        'Application.Wait (Now + TimeValue(delay))
    Next
    
    
    
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    wksSummary.Range("A1").Select

CleanUp:
    Application.Calculation = eCalc
    Application.EnableEvents = boolEnableEvents
    Application.ScreenUpdating = boolScreenUpdating
    Application.CutCopyMode = False
    Exit Sub
    
Error_Handler:
    GoTo CleanUp
End Sub
Function countTickerBrand(ticker As String) As Integer
    'checks how many brands are under a certain ticker
    'May be superseded by Database-provided values
    
    Dim Last As Integer
    
    Last = wksTemplate.Cells(Rows.Count, wksTemplate.Range("TickerCoMatch").Column).End(xlUp).row
    countTickerBrand = Application.WorksheetFunction.CountIf(wksTemplate.Range(wksTemplate.Range("TickerCoMatch"), wksTemplate.Range("C" & Last)), ticker)
End Function

Function countBrand() As Integer
    Dim First As Integer
    Dim Last As Integer
    First = wksTemplate.Range("DropDownList").row
    Last = wksTemplate.Cells(Rows.Count, wksTemplate.Range("DropDownList").Column).End(xlUp).row
    countBrand = Last - First
End Function

Function getDateRange() As String
    Dim First As Integer
    Dim Last As Integer
    Dim Earliest As Date
    Dim Latest As Date
    
    First = wksUserData.Range("A1").row
    Last = wksUserData.Cells(Rows.Count, wksUserData.Range("A1").Column).End(xlUp).row
    
    If Last > First Then
        Earliest = Application.WorksheetFunction.Min(wksUserData.Range("B" & First + 1 & ":B" & Last))
        Latest = Application.WorksheetFunction.max(wksUserData.Range("C" & First + 1 & ":C" & Last)) ' May want to convert to last day of month
        getDateRange = CStr(Earliest) & " - " & CStr(Latest)
    Else
        getDateRange = "None"
    End If

End Function

Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    Dim ret As Boolean
    ret = False
    wsName = UCase(wsName)
    For Each ws In ActiveWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret
End Function

Function FindChannelCount(FirstRange As String) As Integer
    'FirstRange is the Company Name
    Dim aCell As Range, bCell As Range, oRange As Range
    Dim var As Variant
    Dim coll As New Collection
    Dim Flag As Integer
    Dim i As Integer
    Dim temp As Integer
    Dim ListRange As Range
    
    Flag = 0
    Set ListRange = wksUserData.Range("F:F")
    Set oRange = ListRange.Find(what:=FirstRange, LookIn:=xlValues, _
    lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False)

    If Not oRange Is Nothing Then
        Set bCell = oRange
        
        coll.Add (oRange.Offset(0, 2).Value2)
        
        Do
            Set oRange = ListRange.Find(what:=FirstRange, After:=oRange, LookIn:=xlValues, _
            lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False)
            If Not oRange Is Nothing Then
                If oRange.Address = bCell.Address Then Exit Do
                        
                For Each var In coll
                    If var = oRange.Offset(0, 2).Value2 Then
                        Flag = 0
                        Exit For
                    Else
                        Flag = 1
                    End If
                    
                Next
                If Flag > 0 Then
                    coll.Add (oRange.Offset(0, 2).Value2)
                    Flag = 0
                End If
                        
                
            Else
                Exit Do
            End If
        Loop
        temp = 0
        For i = 1 To coll.Count
            If coll(i) = "ALL" Then
                temp = i
                Exit For
            End If
        Next
        If temp > 0 Then
            coll.Remove (temp)
        End If
        FindChannelCount = coll.Count
        
    Else
        FindChannelCount = 0
        
    End If
    
    Set coll = Nothing
End Function

Function FindChannels(FirstRange As String) As Collection
    'FirstRange is the Company Name
    Dim aCell As Range, bCell As Range, oRange As Range
    Dim var As Variant
    Dim coll As New Collection
    Dim Flag As Integer
    Dim i As Integer
    Dim temp As Integer
    Dim ListRange As Range
    
    Flag = 0
    Set ListRange = wksUserData.Range("F:F")
    Set oRange = ListRange.Find(what:=FirstRange, LookIn:=xlValues, _
    lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False)

    If Not oRange Is Nothing Then
        Set bCell = oRange
        
        coll.Add (oRange.Offset(0, 2).Value2)
        
        Do
            Set oRange = ListRange.Find(what:=FirstRange, After:=oRange, LookIn:=xlValues, _
            lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False)
            If Not oRange Is Nothing Then
                If oRange.Address = bCell.Address Then Exit Do
                        
                    For Each var In coll
                        If var = oRange.Offset(0, 2).Value2 Then
                            Flag = 0
                            Exit For
                        Else
                            Flag = 1
                        End If
                        
                    Next
                    If Flag > 0 Then
                        coll.Add (oRange.Offset(0, 2).Value2)
                        Flag = 0
                    End If

            Else
                Exit Do
            End If
        Loop
        temp = 0
        For i = 1 To coll.Count
            If coll(i) = "ALL" Then
                temp = i
                Exit For
            End If
        Next
        If temp > 0 Then
            coll.Remove (temp)
        End If
        Set FindChannels = coll
        
    Else
        Set FindChannels = coll
        
    End If
    
    Set coll = Nothing
End Function
'--------------------------------------
'Code below for Testing. To be deleted.
'--------------------------------------

'~~> 07th May 2006 # Siddharth Rout #
'~~> The below code retrieves the control id's of the VBA Editor
Private Sub GetMeIDs()
    Dim ws As Worksheet
    Dim Ctl As CommandBarControl
    Dim cCtl As CommandBarControl
    Dim i As Long, j As Long, k As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    On Error Resume Next
    
    i = 1: j = 1: k = 1
    
    '~~> Loop through the Top Level menus like File, Edit, View etc
    For Each Ctl In Application.CommandBars(13).Controls(30253).Controls
        '~~> Write the ID and caption to Col A
        ws.Range("A" & i).value = Ctl.ID & " - " & Ctl.Caption
        i = i + 1
        
        '~~> Loop through Sub Level menus like File~~>New Project, File~~>Open Project etc
        For Each cCtl In Application.CommandBars(13).Controls(Ctl.Caption).Controls
            '~~> Write the ID and caption to Col B
            ws.Range("B" & j).value = cCtl.ID & " - " & cCtl.Caption
            j = j + 1
        Next
        i = j
    Next
End Sub
Private Sub GetMeMoreIDs()
    Dim ws As Worksheet
    Dim Ctl As CommandBarControl
    Dim cCtl As CommandBarControl
    Dim i As Long, j As Long, k As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    On Error Resume Next
    
    i = 1: j = 1: k = 1
    For k = 1 To 166
        '~~> Loop through the Top Level menus like File, Edit, View etc
        For Each Ctl In Application.CommandBars(k).Controls
            '~~> Write the ID and caption to Col A
            ws.Range("A" & i).value = Ctl.ID & " - " & Ctl.Caption
            i = i + 1
            
            ws.Range("B" & i).value = k
            '~~> Loop through Sub Level menus like File~~>New Project, File~~>Open Project etc
            
        Next
    Next k
End Sub






