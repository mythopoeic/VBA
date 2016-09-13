Attribute VB_Name = "modYahooAPI"
Public Sub GetYahooStocks()
    Dim w As Worksheet: Set w = Worksheets("Summary")
    Dim i As Integer
    Dim Last As Integer
    Last = w.Range("B6000").End(xlUp).Row
    If Last = 1 Then Exit Sub
    Dim Symbols As String
    For i = 7 To Last
        Symbols = Symbols & w.Range("B" & i).Value & "+"
    Next i
    Symbols = Left(Symbols, Len(Symbols) - 1)
    
    Dim URL As String: URL = "http://finance.yahoo.com/d/quotes.csv?s=" & Symbols & "&f=l1kk5jj1"
    
    Dim x As New WinHttpRequest
    x.Open "GET", URL, False
    x.send
    
    Dim Resp As String: Resp = x.responseText
    Dim Lines As Variant: Lines = Split(Resp, vbLf)
    Dim sLine As String
    Dim Values As Variant
    For i = 0 To UBound(Lines)
        sLine = Lines(i)
        If InStr(sLine, ",") > 0 Then
            Values = Split(sLine, ",")
            w.Cells(i + 7, 4).Value = Values(UBound(Values) - 4)
            w.Cells(i + 7, 5).Value = Values(UBound(Values) - 3)
            w.Cells(i + 7, 6).Value = Values(UBound(Values) - 2)
            w.Cells(i + 7, 7).Value = Values(UBound(Values) - 1)
            w.Cells(i + 7, 8).Value = Values(UBound(Values))
            
        End If
    Next i
    
    
    Call convertMarketCap
    w.Range("D2").Value = Date
    w.Range("D7:H9").Copy
    Worksheets("Summary - full valuation").Range("E12").Select
    Worksheets("Summary - full valuation").Range("E12").PasteSpecial xlPasteValues
    Worksheets("Summary - full valuation").Range("E7").Value = Date
    Worksheets("Summary - full valuation").Range("A1").Select
    Application.CutCopyMode = False
    'W.Cells.Columns.AutoFit
    'W.Columns("C").EntireColumn.Hidden = True
End Sub

Private Sub convertMarketCap()
    
    'Remove B or M from Market Cap data. Convert Billions to Millions
    
    Dim w As Worksheet: Set w = Worksheets("Summary")
    Dim i As Integer
    Dim n As Single
    Dim Last As Integer
    Dim myString As String
    Last = w.Range("H6000").End(xlUp).Row
    
    For i = 7 To Last
        myString = w.Range("H" & i).Value
        If InStr(myString, "B") > 0 Then
            myString = Left(myString, Len(myString) - 1)
            n = CSng(myString)
            w.Range("H" & i).Value = n * 1000
            ElseIf InStr(myString, "M") > 0 Then
            myString = Left(myString, Len(myString) - 1)
            n = CSng(myString)
            w.Range("H" & i).Value = n
        Else
        End If
    Next i
    
    
End Sub
