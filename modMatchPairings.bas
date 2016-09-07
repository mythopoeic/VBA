Attribute VB_Name = "modMatchPairings"
''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Golf Pairings Spreadsheet
''   created by William Fehringer
''   for Mark Anderson of Batavia Mfg. c. 2016
''''''''''''''''''''''''''''''''''''''''''''''''''''

''Deliverable: Excel Workbook with VBA
''
''See attached spreadsheet to see how the tabs should be arranged and the data formatted.
''
''The purpose of this spreadsheet is to automatically create matches for a golf group that pairs
''players together based upon whom they have not been paired with in the past and whom they
''have not yet faced as an opponent and finally random distribution if no history is available to
''determine the pairings and opponents.
''
Option Explicit
Public pList() As String
Public IntList() As Integer
Public finArr() As Integer
Public matchDate As Date
Public bCancel As Boolean

Sub ParticipantsList()
    
    ''1.  Participants
    ''
    ''a.  There will be a Participants tab that lists all possible players up to a possible 25
    ''
    ''b.  Not all players listed in the participants tab will participate in a match
    ''
    ''c.  There will be an entry area for who will participate in the match
    ''
    ''d.  A minimum of four participants are required to have a match
    ''
    Dim n As Integer
    Dim strDate As String
    bCancel = False
    Dim row As Integer
    
    On Error GoTo ErrHandler
    
    
    strDate = InputBox("Enter the date of the match")
    
    If IsDate(strDate) = False Then
        MsgBox ("Please enter a date in the correct format to continue.")
        bCancel = True
        Exit Sub
    End If
    
    matchDate = CDate(strDate)
    
    ' populate participant list in selection box
    uSelect.lActivities.RowSource = ""
    For row = 2 To 26
        If shParticipants.Cells(row, 1) <> "" Then
            uSelect.lActivities.AddItem shParticipants.Cells(row, 1)
        End If
    Next row
    uSelect.Show
    
    
    
    If bCancel = True Then
        Exit Sub
    End If
    
    n = UBound(pList) - LBound(pList) + 1
    
    If n < 4 Then
        MsgBox ("Minimum of 4 Participants required for match.")
        Exit Sub
    Else
        
        Call MatchPairings
    End If
    
ExitSub:
    ' clean up before exiting
    Exit Sub
    
ErrHandler:
    MsgBox ("Minimum of 4 Participants required for match.")
    Resume ExitSub
    
End Sub



Private Sub MatchPairings()
    ''3.  Match Pairings
    ''
    ''a.  Once the participants are selected, the groups will be created in the match pairings tabs
    ''
    ''b.  Each group consists of 4 or 5 players.
    ''
    ''c.  Each group will have a four-person match that consists of two partner players vs. two partner players
    ''
    ''d.  If there is group of five, three players will participate in two matches
    ''
    ''- Two of the three players will remain partners from the other match
    ''- The fifth player will be paired with one of the partners to form a match
    ''
    ''e.  There can never be groups less than four
    ''
    ''f.  There can never be groups greater than five
    ''
    ''g.  Participants will be eliminated if groups of four and five are not achievable
    ''
    ''h.  The Match history tab will be display the created groups with the following sorting selection (in order of priority)
    ''
    ''- Match players together that they have not yet been a partner in match history
    ''- Match players that they have not yet faced as an opponent in match history
    ''- Random if no decisions can be based upon history
    ''- Anyone that has been the fifth in a fivesome will be recorded and will not be a fifth until the rest of the players
    ''    available that week have played as a fifth. Added 3/11/2016
    
    Dim n As Integer
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    Dim Last As Integer
    Dim Last2 As Integer
    Dim Min1 As Integer
    Dim Min2 As Integer
    Dim exclusions As Integer
    Dim fourGroups As Integer
    Dim fiveGroups As Integer
    Dim rowCount As Integer
    Dim randSample As Integer
    Dim fiftyFifty As Integer
    Dim p1 As String
    Dim p2 As String
    Dim p3 As String
    Dim p4 As String
    Dim p5 As String
    Dim pExcluded As String
    Dim rng As Range
    Dim str As String
    Dim strPair As String
    Dim firstTimeFifth As Boolean
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    'firstTimeFifth = True
    
    Call printCombinations(pList, 2)
    
    ' Create matched pairs
    
    Last = shScratch.Cells(Rows.Count, 1).End(xlUp).row
    shScratch.Range("C1").Formula = "=CONCATENATE(A1,B1)"
    shScratch.Range("C1:C" & Last).FillDown
    shScratch.Calculate
    
    
    shMatchHistory.Activate
    For i = 1 To Last
        On Error Resume Next
        strPair = shScratch.Range("C" & i).Text
        With shMatchHistory
            
            Set rng = .Columns(14).Find(What:=strPair, After:=.Cells(1, 14), LookIn:=xlValues, LookAt:= _
            xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
            , SearchFormat:=False)
            
            
            On Error GoTo ErrHandler
            
            If Not rng Is Nothing Then Application.Goto rng, True
            
        End With
        
        shScratch.Range("D" & i).Value = rng.Offset(0, 1).Value
    Next
    
    n = UBound(pList) - LBound(pList) + 1
    
    ' algorithm to determine best mix of four- and five-groups to minimize exclusions
    
    i = n \ 4 'integer division
    r = n Mod 4 'remainder
    exclusions = 0
    
    If r > i Then
        exclusions = r - i
        fiveGroups = i
    Else
        fourGroups = i - r
        fiveGroups = r
    End If
    
    
    shMatchPairings.Rows("2:" & Rows.Count).ClearContents
    
    'players who have not been partners before
    'Min = Application.Min(Range(matchupCount))
    
    'players who have not been opponents before
    
    
    Randomize
    If exclusions > 0 Then
        shScratch2.UsedRange.ClearContents
        shMatchPairings.Activate
        Last = Cells(Rows.Count, 1).End(xlUp).row - 1
        Last2 = shMatchHistory.Cells(Rows.Count, 1).End(xlUp).row
        For i = 1 To exclusions
            shMatchHistory.Activate
            For j = 0 To UBound(pList)
                On Error Resume Next
                
                str = pList(j)
                With shMatchHistory
                    
                    Set rng = .Columns(11).Find(What:=str, After:=.Cells(1, 11), LookIn:=xlValues, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                    , SearchFormat:=False)
                    
                    
                    On Error GoTo ErrHandler
                    
                    If Not rng Is Nothing Then Application.Goto rng, True
                    
                End With
                
                shScratch.Range("G" & j + 1).Value = rng.Value
                shScratch.Range("H" & j + 1).Value = rng.Offset(0, 2).Value
            Next
            rowCount = shScratch.Cells(Rows.Count, 8).End(xlUp).row
            Min1 = Application.Min(shScratch.Range("H1:H" & rowCount))
            Min2 = -1
            rowCount = shScratch.Cells(Rows.Count, 8).End(xlUp).row
            
            shScratch.Activate
            Randomize
            Do While Min2 <> Min1
                
                randSample = Int(rowCount * Rnd + 1)
                Min2 = Cells(randSample, 8).Value
            Loop
            
            pExcluded = Cells(randSample, 7).Text
            
            For j = 0 To UBound(pList)
                If pList(j) = pExcluded Then
                    k = j
                End If
            Next
            
            Call RemItem(pList, k)
            
            shMatchHistory.Activate
            shMatchHistory.Cells(Last2 + i, 9).Value = pExcluded
            shScratch2.Activate
            
            shScratch2.Cells(i, 1).Value = pExcluded
            shScratch.Activate
            rowCount = shScratch.Cells(Rows.Count, 1).End(xlUp).row
            For j = rowCount To 1 Step -1
                If (Cells(j, 1).Value) = pExcluded Or (Cells(j, 2).Value) = pExcluded Then
                    Cells(j, 1).EntireRow.Delete
                End If
                
            Next j
            shMatchHistory.Calculate
        Next i
    End If
    'forming the five-groups
    
    
    If fiveGroups > 0 Then
        'Choose the fifth player
        For i = 1 To fiveGroups
            shMatchPairings.Activate
            Last = Cells(Rows.Count, 1).End(xlUp).row
            
            Last2 = shMatchHistory.Cells(Rows.Count, 1).End(xlUp).row
            shMatchHistory.Activate
            
            If i = 1 Then
                'For j = 1 To rowCount
                For j = 0 To UBound(pList)
                    On Error Resume Next
                    
                    str = pList(j)
                    With shMatchHistory
                        
                        Set rng = .Columns(11).Find(What:=str, After:=.Cells(1, 11), LookIn:=xlValues, LookAt:= _
                        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                        , SearchFormat:=False)
                        
                        
                        On Error GoTo ErrHandler
                        
                        If Not rng Is Nothing Then Application.Goto rng, True
                        
                    End With
                    
                    shScratch.Range("E" & j + 1).Value = rng.Value
                    shScratch.Range("F" & j + 1).Value = rng.Offset(0, 1).Value
                Next
            Else
                shScratch.Activate
                rowCount = shScratch.Cells(Rows.Count, 6).End(xlUp).row
                For j = rowCount To 1 Step -1
                    If (Cells(j, 5).Value) = p5 Or (Cells(j, 5).Value) = p4 Or (Cells(j, 5).Value) = p3 Or (Cells(j, 5).Value) = p2 Or (Cells(j, 5).Value) = p1 Then
                        Range("E" & j & ":F" & j).Delete Shift:=xlUp
                    End If
                    
                Next j
            End If
            
            rowCount = shScratch.Cells(Rows.Count, 6).End(xlUp).row
            Min1 = Application.Min(shScratch.Range("F1:F" & rowCount))
            Min2 = -1
            
            
            shScratch.Activate
            Randomize
            Do While Min2 <> Min1
                
                randSample = Int(rowCount * Rnd + 1)
                Min2 = Cells(randSample, 6).Value
            Loop
            
            p5 = Cells(randSample, 5).Text
            
            '        fiftyFifty = Int(2 * Rnd)
            '        shMatchPairings.Activate
            '        shMatchPairings.Cells(Last + i, 8 + fiftyFifty).PasteSpecial xlPasteValues
            '        shMatchHistory.Activate
            '        shMatchHistory.Cells(Last2 + i, 8).PasteSpecial xlPasteValues
            '        shMatchHistory.Cells(Last2 + i + 1, 2 + fiftyFifty).PasteSpecial xlPasteValues
            '        shMatchHistory.Calculate
            shScratch.Activate
            rowCount = Cells(Rows.Count, 1).End(xlUp).row
            For j = rowCount To 1 Step -1
                If (Cells(j, 1).Value) = p5 Or (Cells(j, 2).Value) = p5 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
                
            Next j
            
            'find the other 4
            rowCount = shScratch.Cells(Rows.Count, 1).End(xlUp).row
            Min1 = Application.Min(shScratch.Range("D1:D" & rowCount))
            shMatchPairings.Activate
            Last = Cells(Rows.Count, 1).End(xlUp).row
            
            shMatchPairings.Range("A" & Last + 1) = "Match " & Last
            
            rowCount = shScratch.Cells(Rows.Count, 1).End(xlUp).row
            Min2 = -1
            
            
            Do While Min2 <> Min1
                randSample = Int(rowCount * Rnd + 1)
                Min2 = Cells(randSample, 4).Value
            Loop
            
            p1 = shScratch.Range("A" & randSample).Text
            p2 = shScratch.Range("B" & randSample).Text
            shScratch.Activate
            shScratch.Range(Cells(randSample, 1), Cells(randSample, 2)).Copy
            shMatchPairings.Activate
            shMatchPairings.Range(Cells(Last + 1, 2), Cells(Last + 1, 3)).PasteSpecial xlPasteValues
            shMatchPairings.Range(Cells(Last + 1, 8), Cells(Last + 1, 9)).PasteSpecial xlPasteValues
            shMatchPairings.Range("D" & Last + 1) = "vs."
            shMatchPairings.Range("J" & Last + 1) = "vs."
            ' Add to Match History Tab
            shMatchHistory.Activate
            Last2 = shMatchHistory.Cells(Rows.Count, 1).End(xlUp).row
            shMatchHistory.Range("A" & Last2 + 1) = matchDate
            shMatchHistory.Range("A" & Last2 + 2) = matchDate
            shMatchHistory.Range(Cells(Last2 + 1, 2), Cells(Last2 + 1, 3)).PasteSpecial xlPasteValues
            shMatchHistory.Range(Cells(Last2 + 2, 2), Cells(Last2 + 2, 3)).PasteSpecial xlPasteValues
            shMatchHistory.Range("D" & Last2 + 1) = "vs."
            shMatchHistory.Range("D" & Last2 + 2) = "vs."
            shMatchHistory.Range("P1").Formula = "=CONCATENATE(B1,C1)"
            shMatchHistory.Range("P1:P" & Last2 + 2).FillDown
            shMatchHistory.Range("Q1").Formula = "=CONCATENATE(C1,B1)"
            shMatchHistory.Range("Q1:Q" & Last2 + 2).FillDown
            shMatchHistory.Range("R1").Formula = "=CONCATENATE(E1,F1)"
            shMatchHistory.Range("R1:R" & Last2 + 2).FillDown
            shMatchHistory.Range("S1").Formula = "=CONCATENATE(F1,E1)"
            shMatchHistory.Range("S1:S" & Last2 + 2).FillDown
            shMatchHistory.Calculate
            shScratch.Activate
            rowCount = Cells(Rows.Count, 1).End(xlUp).row
            For j = rowCount To 1 Step -1
                If (Cells(j, 1).Value) = p1 Or (Cells(j, 2).Value) = p1 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
                
                If (Cells(j, 1).Value) = p2 Or (Cells(j, 2).Value) = p2 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
            Next j
            
            
            
            
            rowCount = shScratch.Cells(Rows.Count, 1).End(xlUp).row
            Min1 = Application.Min(shScratch.Range("D1:D" & rowCount))
            Min2 = -1
            'rowCount = rng.Rows.Count
            
            shScratch.Activate
            Do While Min2 <> Min1
                randSample = Int(rowCount * Rnd + 1)
                Min2 = Cells(randSample, 4).Value
            Loop
            p3 = Range("A" & randSample).Text
            p4 = Range("B" & randSample).Text
            shScratch.Range(Cells(randSample, 1), Cells(randSample, 2)).Copy
            shMatchPairings.Activate
            shMatchPairings.Range(Cells(Last + 1, 5), Cells(Last + 1, 6)).PasteSpecial xlPasteValues
            shMatchPairings.Range(Cells(Last + 1, 11), Cells(Last + 1, 12)).PasteSpecial xlPasteValues
            shMatchPairings.Range("G" & Last + 1) = "Round 2"
            shMatchHistory.Activate
            shMatchHistory.Range(Cells(Last2 + 1, 5), Cells(Last2 + 1, 6)).PasteSpecial xlPasteValues
            shMatchHistory.Range(Cells(Last2 + 2, 5), Cells(Last2 + 2, 6)).PasteSpecial xlPasteValues
            fiftyFifty = Int(2 * Rnd)
            shMatchPairings.Activate
            shMatchPairings.Cells(Last + 1, 8 + fiftyFifty).Value = p5
            shMatchHistory.Activate
            shMatchHistory.Cells(Last2 + 1, 8).Value = p5
            shMatchHistory.Cells(Last2 + 2, 2 + fiftyFifty).Value = p5
            shMatchHistory.Calculate
            
            
            shScratch.Activate
            rowCount = Cells(Rows.Count, 1).End(xlUp).row
            For j = rowCount To 1 Step -1
                If (Cells(j, 1).Value) = p3 Or (Cells(j, 2).Value) = p3 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
                
                If (Cells(j, 1).Value) = p4 Or (Cells(j, 2).Value) = p4 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
            Next j
            
            
            
        Next
    End If
    
    
    'forming the four-groups
    If fourGroups > 0 Then
        
        
        For i = 1 To fourGroups
            'choose random set from pList
            shMatchPairings.Activate
            Last = Cells(Rows.Count, 1).End(xlUp).row
            If Last = -1 Then
                Last = 0
            End If
            Last2 = shMatchHistory.Cells(Rows.Count, 1).End(xlUp).row
            shMatchPairings.Range("A" & Last + 1) = "Match " & Last
            
            rowCount = shScratch.Cells(Rows.Count, 1).End(xlUp).row
            Min1 = Application.Min(shScratch.Range("D1:D" & rowCount))
            Min2 = -1
            'rowCount = rng.Rows.Count
            
            shScratch.Activate
            Do While Min2 <> Min1
                randSample = Int(rowCount * Rnd + 1)
                Min2 = Cells(randSample, 4).Value
            Loop
            
            shScratch.Activate
            p1 = Range("A" & randSample).Text
            p2 = Range("B" & randSample).Text
            shScratch.Range(Cells(randSample, 1), Cells(randSample, 2)).Copy
            shMatchPairings.Activate
            shMatchPairings.Range(Cells(Last + 1, 2), Cells(Last + 1, 3)).PasteSpecial xlPasteValues
            shMatchPairings.Range("D" & Last + 1) = "vs."
            shMatchHistory.Activate
            
            shMatchHistory.Range("A" & Last2 + 1) = matchDate
            
            shMatchHistory.Range(Cells(Last2 + 1, 2), Cells(Last2 + 1, 3)).PasteSpecial xlPasteValues
            
            shMatchHistory.Range("D" & Last2 + 1) = "vs."
            
            shMatchHistory.Range("P1").Formula = "=CONCATENATE(B1,C1)"
            shMatchHistory.Range("P1:P" & Last2 + 2).FillDown
            shMatchHistory.Range("Q1").Formula = "=CONCATENATE(C1,B1)"
            shMatchHistory.Range("Q1:Q" & Last2 + 2).FillDown
            shMatchHistory.Range("R1").Formula = "=CONCATENATE(E1,F1)"
            shMatchHistory.Range("R1:R" & Last2 + 2).FillDown
            shMatchHistory.Range("S1").Formula = "=CONCATENATE(F1,E1)"
            shMatchHistory.Range("S1:S" & Last2 + 2).FillDown
            shMatchHistory.Calculate
            shScratch.Activate
            rowCount = Cells(Rows.Count, 1).End(xlUp).row
            For j = rowCount To 1 Step -1
                If (Cells(j, 1).Value) = p1 Or (Cells(j, 2).Value) = p1 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
                
                If (Cells(j, 1).Value) = p2 Or (Cells(j, 2).Value) = p2 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
            Next j
            
            
            
            
            rowCount = shScratch.Cells(Rows.Count, 1).End(xlUp).row
            Min1 = Application.Min(shScratch.Range("D1:D" & rowCount))
            Min2 = -1
            Do While Min2 <> Min1
                randSample = Int(rowCount * Rnd + 1)
                Min2 = Cells(randSample, 4).Value
            Loop
            shScratch.Activate
            p3 = Range("A" & randSample).Text
            p4 = Range("B" & randSample).Text
            shScratch.Range(Cells(randSample, 1), Cells(randSample, 2)).Copy
            shMatchPairings.Activate
            shMatchPairings.Range(Cells(Last + 1, 5), Cells(Last + 1, 6)).PasteSpecial xlPasteValues
            shMatchHistory.Activate
            shMatchHistory.Range(Cells(Last2 + 1, 5), Cells(Last2 + 1, 6)).PasteSpecial xlPasteValues
            shScratch.Activate
            rowCount = Cells(Rows.Count, 1).End(xlUp).row
            For j = rowCount To 1 Step -1
                If (Cells(j, 1).Value) = p3 Or (Cells(j, 2).Value) = p3 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
                
                If (Cells(j, 1).Value) = p4 Or (Cells(j, 2).Value) = p4 Then
                    Range("A" & j & ":D" & j).Delete Shift:=xlUp
                End If
            Next j
            
        Next i
    End If
    
    If exclusions > 0 Then
        shScratch2.Activate
        rowCount = Cells(Rows.Count, 1).End(xlUp).row
        Range("A1:A" & rowCount).Copy
        
        shMatchPairings.Activate
        Last = Cells(Rows.Count, 1).End(xlUp).row
        
        Range("A" & Last + 2) = "Excluded:"
        
        Range("A" & Last + 3).PasteSpecial xlPasteValues
        
        
    End If
    shMatchHistory.Activate
    Application.Goto shMatchHistory.Range("A1"), True
    shMatchPairings.Activate
    Application.Goto shMatchPairings.Range("A1"), True
    
ExitSub:
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    Exit Sub
    
ErrHandler:
    
    Resume ExitSub
    
End Sub



Public Sub printCombinations(ByRef pool() As String, ByVal r As Integer)
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    
    shScratch.UsedRange.ClearContents
    Application.Calculation = xlCalculationManual
    
    'Application.ScreenUpdating = False
    n = UBound(pool) - LBound(pool)
    
    ' Please do add error handling for when r>n
    
    Dim idx() As Integer
    ReDim idx(1 To r)
    For i = 1 To r
        idx(i) = i - 1
    Next i
    
    
    k = 1
    
    Do
        'Write current combination
        For j = 1 To r
            shScratch.Cells(k, j) = pool(idx(j))
            'Debug.Print pool(idx(j));
            'or whatever you want to do with the numbers
        Next j
        'Debug.Print
        k = k + 1
        ' Locate last non-max index
        i = r
        While (idx(i) = n - r + i)
        i = i - 1
        If i = 0 Then
            'All indexes have reached their max, so we're done
            Exit Sub
        End If
    Wend
    
    'Increase it and populate the following indexes accordingly
    idx(i) = idx(i) + 1
    For j = i + 1 To r
        idx(j) = idx(i) + j - i
    Next j
    
Loop
Application.Calculation = xlCalculationAutomatic
End Sub

Sub playerList()
    Dim blDimensioned As Boolean
    Dim jcounter As Integer
    Dim n As Integer
    blDimensioned = False
    
    n = shParticipants.UsedRange.Rows.Count
    
    For jcounter = 2 To n
        'If lActivities.Selected(jcounter) = True Then
        
        If blDimensioned = True Then
            
            ReDim Preserve pList(0 To UBound(pList) + 1) As String
            'ReDim Preserve IntList(0 To UBound(IntList) + 1) As Integer
            
        Else
            ReDim pList(0 To 0) As String
            'ReDim IntList(0 To 0) As Integer
            blDimensioned = True
        End If
        
        pList(UBound(pList)) = shParticipants.Range("A" & jcounter).Value
        'IntList(UBound(IntList)) = shParticipants.Range("A" & jcounter).Row - 1
        
    Next
End Sub

Sub test()
    Call playerList
    Call printCombinations(pList, 2)
End Sub

Sub selectParticipants()
Dim row As Integer
uSelect.lActivities.RowSource = ""
For row = 2 To 26
    If shParticipants.Cells(row, 1) <> "" Then
        uSelect.lActivities.AddItem shParticipants.Cells(row, 1)
    End If
Next row
    uSelect.Show
End Sub

Sub RemItem(v As Variant, n As Long)
    Dim i As Long
    For i = n + 1 To UBound(v) - LBound(v)
        v(i - 1) = v(i)
    Next i
    ReDim Preserve v(LBound(v) To UBound(v) - 1)
    
End Sub

Sub manualHistory()
Dim Last As Integer
    'Enter items manually into Match History and then click button to update matchup count
    Last = shMatchHistory.Cells(Rows.Count, 1).End(xlUp).row

    shMatchHistory.Range("P1").Formula = "=CONCATENATE(B1,C1)"
    shMatchHistory.Range("P1:P" & Last + 2).FillDown
    shMatchHistory.Range("Q1").Formula = "=CONCATENATE(C1,B1)"
    shMatchHistory.Range("Q1:Q" & Last + 2).FillDown
    shMatchHistory.Range("R1").Formula = "=CONCATENATE(E1,F1)"
    shMatchHistory.Range("R1:R" & Last + 2).FillDown
    shMatchHistory.Range("S1").Formula = "=CONCATENATE(F1,E1)"
    shMatchHistory.Range("S1:S" & Last + 2).FillDown
    shMatchHistory.Calculate
End Sub
