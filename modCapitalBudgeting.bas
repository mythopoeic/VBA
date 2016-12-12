Attribute VB_Name = "modCapitalBudgeting"
'---------------------------
'Capital Budgeting Simulator
'---------------------------

Option Base 1
Public rI As Long ' index for records
Public rCycle As Long
Public nwRanking() As String
Public budgetArr() As Double
Public rMethod As String
Public rRow As Long
Public relOrder As Long
Public arrKey As Long
Public progress As Long
Public tProgress As Long


Function simulateFirstCost(ByVal minCost As Long, ByVal maxCost As Long)
'Function to make random first cost
    simulateFirstCost = Int((maxCost + 1 - minCost) * Rnd + minCost)
End Function

Function simulateProjectLife(ByVal minLife As Long, ByVal maxLife As Long)
    simulateProjectLife = Int((maxLife + 1 - minLife) * Rnd + minLife)
End Function

Function simulateIRR(ByVal minIRR As Single, ByVal maxIRR As Single)
    simulateIRR = (maxIRR - minIRR) * Rnd + minIRR
End Function

Function simulateRisk(ByVal minRisk As Single, ByVal maxRisk As Single)
    simulateRisk = (maxRisk - minRisk) * Rnd + minRisk
End Function
Function yearlyCashFlow(cost As Long, irr As Single, life As Long)
    yearlyCashFlow = cost * (((1 + irr) ^ life) * irr / ((1 + irr) ^ life - 1))
End Function

Function payback(ByVal firstCost As Long, ByVal cashFlow As Long)
    payback = firstCost / cashFlow
End Function

Function presentValue(firstCost As Long, ycf As Double, i As Single, N As Long)
Dim PV As Double
Dim k As Long
PV = 0
For k = 0 To N
    If k <> 0 Then
        PV = PV + ycf * (1 / (1 + i)) ^ k
    Else
        PV = PV - firstCost
    End If
Next
presentValue = PV
End Function

Function annualWorth(firstCost As Long, ycf As Double, i As Single, N As Long)
    annualWorth = presentValue(firstCost, ycf, i, N) * ((i * (1 + i) ^ N) / ((1 + i) ^ N - 1))
End Function

Function debtPayback(amount As Long, i As Single, N As Long)
    debtPayback = i * amount / (1 - (1 + i) ^ -N)
End Function



Sub projectGenerator()
'''
'1' Generate projects
'''
Dim cost As Long
Dim life As Long
Dim irr As Single
Dim ycf As Double
Dim pb As Double
Dim npv As Double
Dim aw As Double
Dim risk As Single
Dim cmplx As Single
Dim N As Long
Dim i As Long
Dim rank As Range
Dim order As String
Dim Arr() As Variant
Dim cplxArr() As Variant
Dim sumNPV As Double
Dim sumPB As Double
Dim sumFC As Double
Dim sumIRR As Double
Dim sumAW As Double
Dim sumRisk As Double
Dim wtNPV As Double
Dim wtPB As Double
Dim wtFC As Double
Dim wtIRR As Double
Dim wtAW As Double
Dim wtRisk As Double
Dim cplxNPV As Double
Dim cplxPB As Double
Dim cplxFC As Double
Dim cplxIRR As Double
Dim cplxAW As Double
Dim cplxRisk As Double
Dim invNPV As Double
Dim invPB As Double
Dim invFC As Double
Dim invIRR As Double
Dim invAW As Double
Dim invRisk As Double


N = Range("numCurrentYear")
ReDim Arr(N, 15)
Worksheets("Simulations").Activate

For i = 1 To N

    cost = simulateFirstCost(Range("minFirstCost"), Range("maxFirstCost"))
    life = simulateProjectLife(Range("minProjLife"), Range("maxProjLife"))
    irr = simulateIRR(Range("minIRR"), Range("maxIRR"))
    ycf = yearlyCashFlow(cost, irr, life)
    npv = calcNPV(life, cost, ycf)
    aw = annualWorth(cost, ycf, Range("discountRate"), life)
    risk = simulateRisk(Range("minRisk"), Range("maxRisk"))
    
    '''
    '2' Payback with ranking below
    '''
    pb = cost / ycf
    
    Arr(i, 1) = i
    Arr(i, 2) = cost
    Arr(i, 3) = life
    Arr(i, 4) = irr
    Arr(i, 5) = ycf
    Arr(i, 6) = pb
    Arr(i, 7) = npv
    Arr(i, 8) = aw
    Arr(i, 9) = aw / cost
    Arr(i, 10) = npv / cost
    Arr(i, 11) = Rnd
    Arr(i, 12) = risk
    Arr(i, 13) = 0
    Arr(i, 14) = 0
Next

Worksheets("Simulations").Range("A2:N" & N + 1) = Arr

'Creating Complex Variable
ReDim cplxArr(N, 1)

sumNPV = Application.Sum(Range(Range("simNPV").Offset(1, 0), Range("simNPV").Offset(N, 0)))
sumPB = Application.Sum(Range(Range("simPB").Offset(1, 0), Range("simPB").Offset(N, 0)))
sumFC = Application.Sum(Range(Range("simFirstCost").Offset(1, 0), Range("simFirstCost").Offset(N, 0)))
sumIRR = Application.Sum(Range(Range("simIRR").Offset(1, 0), Range("simIRR").Offset(N, 0)))
sumAW = Application.Sum(Range(Range("simAW").Offset(1, 0), Range("simAW").Offset(N, 0)))
sumRisk = Application.Sum(Range(Range("simRisk").Offset(1, 0), Range("simRisk").Offset(N, 0)))

wtNPV = Range("wtNPV").Value
wtPB = Range("wtPB").Value
wtFC = Range("wtFC").Value
wtIRR = Range("wtIRR").Value
wtAW = Range("wtAW").Value
wtRisk = Range("wtRisk").Value

invPB = 0
invFC = 0
invRisk = 0

For i = 1 To N
invPB = invPB + (sumPB / Arr(i, 6))

invFC = invFC + (sumFC / Arr(i, 2))


invRisk = invRisk + (sumRisk / Arr(i, 12))

Next

For i = 1 To N
    cplxPB = wtPB * (sumPB / Arr(i, 6)) / invPB
    cplxNPV = wtNPV * (Arr(i, 7) / sumNPV)
    cplxFC = wtFC * (sumFC / Arr(i, 2)) / invFC
    
    cplxIRR = wtIRR * (Arr(i, 4) / sumIRR)
    cplxAW = wtAW * (Arr(i, 8) / sumAW)
    
    If Range("riskTolerance") = "Risky" Then
        cplxRisk = wtRisk * (Arr(i, 12) / sumRisk)
    Else
        cplxRisk = wtRisk * (sumRisk / Arr(i, 12)) / invRisk
    End If
    
    cplxArr(i, 1) = (cplxPB + cplxNPV + cplxFC + cplxIRR + cplxAW + cplxRisk) * 100
Next

Range(Worksheets("Simulations").Range("simComplex").Offset(1, 0), Worksheets("Simulations").Range("simComplex").Offset(N, 0)) = cplxArr

'Call getNPV

'Set rank = Range(Range("A1:L1"), Range("A1:L1").Offset(N, 0))
'
'If Range(sortRng) = "LP" Then
'    If Range("rankLP") = "Yes" Then
'        Worksheets("Simulations").Calculate
'        Call RunSolverNPV
'    End If
'End If
'
'If Range(sortRng) = "Payback" Or Range(sortRng) = "Random" Then
'    order = xlAscending
'ElseIf Range(sortRng) = "IRR" Or Range(sortRng) = "NPV" Then
'    order = xlDescending
'Else
'    order = xlDescending
'End If
'
'rank.Sort key1:=Range(sortRng), order1:=order, Header:=xlYes

End Sub

Sub multitimesProjects(numYears As Long)
'Generates a number of projects per year for the given number of years

Dim i As Long
Dim j As Long
Dim k As Long
Dim growth As Long
Dim capital As Double
Dim projectRange As Range
Dim sTemp As String
Dim sortRng As String

k = Range("numYear1") ' starts at 25 and increase by growth
growth = Range("projGrowthAmt")
capital = Range("initialCapBudget") ' starts at $600,000

Worksheets("Simulations").Activate
Set projectRange = Range(Worksheets("Simulations").Range("A1:N1").Offset(1, 0), _
                    Worksheets("Simulations").Range("A1:N1").Offset(k, 0))

For i = 1 To numYears
    Range("currentYear") = i
    Range("numCurrentYear") = k
    Range("currentCapBudget") = capital
    projectRange.ClearContents
    Set projectRange = Range(Worksheets("Simulations").Range("A1:N1").Offset(1, 0), _
                        Worksheets("Simulations").Range("A1:N1").Offset(k, 0))
    
    Call projectGenerator
    Call buyProjects(k, i)
    
    For j = LBound(budgetArr) To UBound(budgetArr)
    sortRng = nwRanking(j, 1)
    Select Case sortRng
    Case "F1"
        sTemp = "PB"
    Case "D1"
        sTemp = "IRR"
    Case "G1"
        sTemp = "NPV"
    Case "H1"
        sTemp = "AW"
    Case "I1"
        sTemp = "AWFC"
    Case "J1"
        sTemp = "NPVFC"
    Case "K1"
        sTemp = "Random"
    Case "L1"
        sTemp = "Risk"
    Case "M1"
        sTemp = "Complex"
    Case "N1"
        sTemp = "LP"
        
    End Select
    
    Worksheets("Cash Flows").Range("cfRM") = sTemp
    
    Worksheets("Cash Flows").Range("cfRep") = rCycle
    Call CashFlows
    
    budgetArr(j) = Worksheets("Income Statement").Range("ISAvailableCapital").Offset(0, i)
    If i = numYears Then
        budgetArr(j) = budgetArr(j) + netWealth(sTemp, numYears)
    End If
    Next
    k = k + growth
Next

'Range("currentCapBudget") = capital + netWealth(numYears)

End Sub

Sub buyProjects(cNumProjects As Long, year As Long)
'''
'3' buy as many projects as possible
'''
Dim i As Long
Dim j As Long
Dim budget As Double

Dim lRow As Long
Dim lRow2 As Long
Dim sortRng As String
Dim recArr() As Variant
Dim recArrT() As Variant
Dim recRow As Long
Dim bTranspose As Boolean


ReDim recArr(12, 1000)
recRow = 1
For j = LBound(budgetArr) To UBound(budgetArr)
If year = 1 Then
budget = Range("InitialCapBudget").Value
Else
budget = budgetArr(j)
End If
relOrder = 1
sortRng = nwRanking(j, 1)
rMethod = Right(nwRanking(j, 2), Len(nwRanking(j, 2)) - 2)
'lRow = Worksheets("Simulations").Cells(Worksheets("Simulations").Rows.count, Range(rMethod & "buyYear").Column).End(xlUp).Row
rRow = Worksheets("Records").Cells(Worksheets("Records").Rows.count, Range("recIndex").Column).End(xlUp).Row
lRow2 = Worksheets("Simulations").Cells(Worksheets("Simulations").Rows.count, "A").End(xlUp).Row

Set rank = Range(Worksheets("Simulations").Range("A1:N1"), Worksheets("Simulations").Range("A1:N1").Offset(lRow2 - 1, 0))

If Range(sortRng) = "LP" Then
    If Range("rankLP") = "Yes" Then
        Range("Capital_Available") = budget
        Worksheets("Simulations").Calculate
        Call RunSolverNPV
    End If
End If

If Range(sortRng) = "Payback" Or Range(sortRng) = "Random" Then
    order = xlAscending
ElseIf Range(sortRng) = "IRR" Or Range(sortRng) = "NPV" Then
    order = xlDescending
ElseIf Range(sortRng) = "Risk" Then
    If Range("riskTolerance") = "Risky" Then
        order = xlDescending
    Else
        order = xlAscending
    End If
Else
    order = xlDescending
End If

rank.Sort key1:=Range(sortRng), order1:=order, Header:=xlYes

Randomize

For i = 1 To cNumProjects
   If Range(sortRng) <> "LP" Then
    If Range("simFirstCost").Offset(i, 0) <= budget And budget >= Range("minFirstCost") Then
    'If it meets criteria, "buy" by moving values to
    'purchased columns
    
        
        'Section for keeping records
        
        recArr(1, recRow) = rI
        recArr(2, recRow) = Range("simFirstCost").Offset(i, 0) * (1 + (2 * Range("simRisk").Offset(i, 0) * Rnd - Range("simRisk").Offset(i, 0))) ' First Cost Risk
        If Rnd < Range("simRisk").Offset(i, 0) Then ' 'Duration Risk
            recArr(3, recRow) = Range("simLife").Offset(i, 0) - 1
        Else
            recArr(3, recRow) = Range("simLife").Offset(i, 0)
        End If
        recArr(4, recRow) = Range("simIRR").Offset(i, 0)
        recArr(5, recRow) = rCycle
        
        recArr(6, recRow) = year
        recArr(7, recRow) = budget
        recArr(8, recRow) = Range("discountRate")
        recArr(9, recRow) = rMethod
        recArr(10, recRow) = Range(nwRanking(j, 3)).Offset(i, 0)
        recArr(11, recRow) = Range("simRisk").Offset(i, 0)
        recArr(12, recRow) = relOrder
        budget = budget - recArr(2, recRow)
        


        recRow = recRow + 1
        rI = rI + 1
        relOrder = relOrder + 1
    End If
   ElseIf Range("simLP").Offset(i, 0) = 1 And budget >= Range("minFirstCost") Then
    'If it meets criteria, "buy" by moving values to
    'purchased columns
    

        'Section for keeping records
        
        recArr(1, recRow) = rI
        recArr(2, recRow) = Range("simFirstCost").Offset(i, 0) * (1 + (2 * Range("simRisk").Offset(i, 0) * Rnd - Range("simRisk").Offset(i, 0))) ' First Cost Risk
        If Rnd < Range("simRisk").Offset(i, 0) Then ' 'Duration Risk
            recArr(3, recRow) = Range("simLife").Offset(i, 0) - 1
        Else
            recArr(3, recRow) = Range("simLife").Offset(i, 0)
        End If
        recArr(4, recRow) = Range("simIRR").Offset(i, 0)
        recArr(5, recRow) = rCycle
        
        recArr(6, recRow) = year
        recArr(7, recRow) = budget
        recArr(8, recRow) = Range("discountRate")
        recArr(9, recRow) = rMethod
        recArr(10, recRow) = Range(nwRanking(j, 3)).Offset(i, 0)
        recArr(11, recRow) = Range("simRisk").Offset(i, 0)
        recArr(12, recRow) = relOrder
        budget = budget - recArr(2, recRow)
        

        recRow = recRow + 1
        rI = rI + 1
        relOrder = relOrder + 1
    End If
Next
budgetArr(j) = budget
progress = progress + 1
Application.StatusBar = "Progress: " & progress & " of " & tProgress & ": " & Format(progress / tProgress, "Percent")
Application.CutCopyMode = False
ActiveWorkbook.Save
Next
If recRow > 1 Then
recRow = recRow - 1
End If
ReDim Preserve recArr(12, recRow) ' You can only change the last dimension of an array
bTranspose = TransposeArray(recArr, recArrT) ' so you have to transpose it.
Range(Range("recIndex").Offset(rRow, 0), Range("recRelOrder").Offset(rRow + recRow - 1, 0)) = recArrT



End Sub

Function getProfits(year As Long)
Dim i As Long
Dim budget As Double
Dim lRow As Long

budget = Range("currentCapBudget")
lRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "N").End(xlUp).Row

For i = 1 To lRow - 1
    If (year - Range("buyYear").Offset(i, 0) + 1) <= Range("buyLife").Offset(i, 0) Then
        budget = budget + Range("buyYCF").Offset(i, 0)
    End If
Next

getProfits = budget

End Function

Function netWealth(method As String, year As Long)
'''
'4' Calculate profits from future payments
'''
Dim i As Long
Dim j As Long
Dim limit As Long
Dim profit As Double
Dim lRow As Long

    Worksheets("Cash Flows").Range("cfRM") = method
    
    Worksheets("Cash Flows").Range("cfRep") = rCycle
    Call CashFlows

profit = 0
limit = 0
'lRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "N").End(xlUp).Row


profit = Application.Sum(Range(Range("NetCashFlow0").Offset(0, year), Range("NetCashFlow0").Offset(0, 20 + year)))

'For i = 1 To lRow - 1
'    If (Range("buyYear").Offset(i, 0) + Range("buyLife").Offset(i, 0) - 1) > limit Then
'        limit = Range("buyYear").Offset(i, 0) + Range("buyLife").Offset(i, 0) - 1
'    End If
'Next
'
'For j = year + 1 To limit
'    For i = 1 To lRow - 1
'        If j <= Range("buyYear").Offset(i, 0) + Range("buyLife").Offset(i, 0) - 1 Then
'            profit = profit + Range("buyYCF").Offset(i, 0)
'        End If
'    Next
'Next

netWealth = profit

End Function

Sub main()
' Main procedure that starts the program
Dim cycles As Long
Dim numYears As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim N As Long
Dim lRow As Long
Dim rCount As Long
Dim rate As Single
'Worksheets("Main").Range("B1").Value = Time
'Worksheets("Main").Range("B1").NumberFormat = "h:mm:ss AM/PM"


If Range("rankLP") = "No" Then
    ReDim budgetArr(9)
    ReDim nwRanking(9, 3)
    rCount = 9
Else
    ReDim budgetArr(10)
    ReDim nwRanking(10, 3)
    nwRanking(10, 1) = "N1"
    nwRanking(10, 2) = "nwLP"
    nwRanking(10, 3) = "simLP"
    rCount = 10
End If

nwRanking(1, 1) = "F1"
nwRanking(1, 2) = "nwPB"
nwRanking(1, 3) = "simPB"
nwRanking(2, 1) = "G1"
nwRanking(2, 2) = "nwNPV"
nwRanking(2, 3) = "simNPV"
nwRanking(3, 1) = "D1"
nwRanking(3, 2) = "nwIRR"
nwRanking(3, 3) = "simIRR"
nwRanking(4, 1) = "H1"
nwRanking(4, 2) = "nwAW"
nwRanking(4, 3) = "simAW"
nwRanking(5, 1) = "I1"
nwRanking(5, 2) = "nwAWFC"
nwRanking(5, 3) = "simAWFC"
nwRanking(6, 1) = "J1"
nwRanking(6, 2) = "nwNPVFC"
nwRanking(6, 3) = "simNPVFC"
nwRanking(7, 1) = "K1"
nwRanking(7, 2) = "nwRandom"
nwRanking(7, 3) = "simRandom"
nwRanking(8, 1) = "L1"
nwRanking(8, 2) = "nwRisk"
nwRanking(8, 3) = "simRisk"
nwRanking(9, 1) = "M1"
nwRanking(9, 2) = "nwComplex"
nwRanking(9, 3) = "simComplex"


For i = LBound(budgetArr) To UBound(budgetArr)
    budgetArr(i) = Range("initialCapBudget")
Next

lRow = Worksheets("Main").UsedRange.Rows(Worksheets("Main").UsedRange.Rows.count).Row
Worksheets("Main").Range("D2:Z" & lRow).ClearContents

Application.ScreenUpdating = False

lRow = Worksheets("Records").UsedRange.Rows(Worksheets("Records").UsedRange.Rows.count).Row
Worksheets("Records").Range("A2:L" & lRow).ClearContents

lRow = Worksheets("ANOVA").UsedRange.Rows(Worksheets("ANOVA").UsedRange.Rows.count).Row
Worksheets("ANOVA").Range("A1:G" & lRow).ClearContents

cycles = Range("numReplications")
numYears = Range("numYears")
rI = 1
progress = 0
tProgress = cycles * numYears * rCount

Application.StatusBar = "Progress: " & progress & " of " & tProgress & ": " & (progress / tProgress) * 100 & "%"
'''
'5' Runs process for all eight ranking methods
'''


'Calculate Net Wealth 30 times for each ranking
'Removed 100 lines of code with additional For Loop
'For j = 1 To N
'    arrKey = j
'    rMethod = Right(nwRanking(j, 2), Len(nwRanking(j, 2)) - 2)
    
'    If nwRanking(j, 1) = "L1" And Range("rankLP") = "No" Then Exit For

'rate = Range("initialCapBudget").Value
    For i = 1 To cycles
        'If i <> 1 Then
        'Range("initialCapBudget") = rate + 100000
        'rate = Range("initialCapBudget").Value
        'End If
        
        rCycle = i
'        If rCycle > 1 Then
'        Stop
'        End If
        Range("numCurrentYear") = Range("numYear1")
        'Range("currentCapBudget") = Range("initialCapBudget")
        lRow = Worksheets("Simulations").UsedRange.Rows(Worksheets("Simulations").UsedRange.Rows.count).Row
        Worksheets("Simulations").Range("A2:N" & lRow).ClearContents
    
        Call multitimesProjects(numYears)

        Range("nwCycle").Offset(i, 0) = i
        Application.ScreenUpdating = True
        Range(Range("nwPB"), Range("nwLP")).Offset(i, 0) = budgetArr
        Application.ScreenUpdating = False
    Next
    
    'Worksheets("EPS").Range("epsTable").Offset(j, 0).Value = Worksheets("Income Statement").Range("eps").Value
    
    Worksheets("Main").Activate
    
    
        Range("nwCycle").Offset(cycles + 2, 0) = "Average:"
        Range("nwCycle").Offset(cycles + 3, 0) = "Standard"
        Range("nwCycle").Offset(cycles + 4, 0) = "Deviation:"
    
    For j = LBound(budgetArr) To UBound(budgetArr)
    Range(nwRanking(j, 2)).Offset(cycles + 2, 0) = Excel.WorksheetFunction.Average(Range _
        (Range(nwRanking(j, 2)).Offset(1, 0), Range(nwRanking(j, 2)).Offset(cycles, 0)))
    Range(nwRanking(j, 2)).Offset(cycles + 4, 0) = Excel.WorksheetFunction.StDev(Range _
        (Range(nwRanking(j, 2)).Offset(1, 0), Range(nwRanking(j, 2)).Offset(cycles, 0)))
    Next
    Worksheets("Simulations").Activate
'Next

Erase nwRanking

If Range("rankLP") = "No" Then
Call [atpvbaen.xls].Anova1(Worksheets("Main").Range("$E$1:$M" & cycles + 1), Worksheets("ANOVA").Range("$A$1"), "C", True, 0.05)
Else
Call [atpvbaen.xls].Anova1(Worksheets("Main").Range("$E$1:$N" & cycles + 1), Worksheets("ANOVA").Range("$A$1"), "C", True, 0.05)
End If


Application.StatusBar = False

'Worksheets("Main").Range("B2").Value = Time
'Worksheets("Main").Range("B2").NumberFormat = "h:mm:ss AM/PM"
ActiveWorkbook.Save
Worksheets("Main").Activate
Application.ScreenUpdating = True
End Sub

Sub getNPV()
Dim npv As Double
Dim r As Single
Dim i As Long
Dim lRow As Long


r = Range("discountRate")
lRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "D").End(xlUp).Row

For i = 2 To lRow
    npv = 0
    For j = 0 To Range("F" & i)
        If j <> 0 Then
            npv = npv + Range("H" & i) / ((1 + r) ^ j)
        Else
            npv = -Range("H" & i)
        End If
    Next
    
    Range("J" & i) = npv
Next

End Sub

Function calcNPV(life As Long, fc As Long, ycf As Double)
Dim npv As Double
Dim r As Single
Dim i As Long
Dim lRow As Long


r = Range("discountRate")
'lRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "D").End(xlUp).Row


    npv = 0
    For j = 0 To life
        If j <> 0 Then
            npv = npv + ycf / ((1 + r) ^ j)
        Else
            npv = -fc
        End If
    Next
    
    calcNPV = npv


End Function

Sub CashFlows()
Dim rMethod As String
Dim replication As Long
Dim methodFirst As Long
Dim methodLast As Long
Dim repCount As Long
Dim repFirst As Long
Dim repLast As Long
Dim i As Long
Dim j As Long
Dim risk As Single
Dim cfArray() As Double
Dim depArray() As Double
Dim debtArray() As Double
Dim cf As Long
Dim dep As Long
Dim debt As Long
Dim year0 As Long
Dim years As Long
Dim lRow As Long
Dim lRow2 As Long
Dim rank As Range
lRow = Worksheets("Cash Flows").UsedRange.Rows(Worksheets("Cash Flows").UsedRange.Rows.count).Row
Worksheets("Cash Flows").Range("C9:ZZ" & lRow).ClearContents
Worksheets("Depreciation Table").Range("C2:ZZ" & lRow).ClearContents
Worksheets("Debt").Range("C8:ZZ" & lRow).ClearContents
If Range("cfRM") = "" Or Range("cfRep") = "" Then
    Exit Sub
End If

lRow2 = Worksheets("Records").Cells(Worksheets("Records").Rows.count, "A").End(xlUp).Row

Set rank = Range(Worksheets("Records").Range("A1:L1"), Worksheets("Records").Range("A1:L1").Offset(lRow2 - 1, 0))

rank.Sort key1:=Range("recRankMethod"), order1:=xlAscending, key2:=Range("recReplication"), order2:=xlAscending, key3:=Range("recYearPurchased"), order3:=xlAscending, Header:=xlYes

rMethod = Range("cfRM")
replication = Range("cfRep")

methodFirst = GetFirstCell(Range("cfRM"), Range("recRankMethod"))
methodLast = GetLastCell(Range("cfRM"), Range("recRankMethod"), methodFirst)

repFirst = Application.WorksheetFunction.Match(Range("cfRep").Value, Range(Range("recReplication").Offset _
            (methodFirst - 1, 0), Range("recReplication").Offset(methodLast - 1, 0)), 0) + methodFirst - 1

repCount = Application.WorksheetFunction.CountIf(Range(Range("recReplication").Offset _
            (methodFirst - 1, 0), Range("recReplication").Offset(methodLast - 1, 0)), replication)

repLast = repFirst + repCount - 1

years = Range("numYears") + Range("maxProjLife") - 1

ReDim cfArray(repCount, 3 + years)
ReDim depArray(repCount, 3 + years)
ReDim debtArray(repCount, 3 + years)

Randomize
For i = 0 To repCount - 1
    cfArray(i + 1, 1) = Range("recIndex").Offset(repFirst + i - 1, 0)
    depArray(i + 1, 1) = Range("recIndex").Offset(repFirst + i - 1, 0)
    debtArray(i + 1, 1) = Range("recIndex").Offset(repFirst + i - 1, 0)
    cfArray(i + 1, 2) = Range("recRankVal").Offset(repFirst + i - 1, 0)
    cfArray(i + 1, 3) = -Range("recFC").Offset(repFirst + i - 1, 0)
    debtArray(i + 1, 2) = -Range("recFC").Offset(repFirst + i - 1, 0) * Range("debtRatio")
    risk = Range("recRisk").Offset(repFirst + i - 1, 0)
    'year0 = -Range("recFC").Offset(repFirst + i - 1, 0)
    cf = yearlyCashFlow(Range("recFC").Offset(repFirst + i - 1, 0), _
                        Range("recIRR").Offset(repFirst + i - 1, 0), Range("recLife").Offset(repFirst + i - 1, 0))
    dep = -Range("recFC").Offset(repFirst + i - 1, 0) / Range("recLife").Offset(repFirst + i - 1, 0)
    
    debt = debtPayback(-Range("recFC").Offset(repFirst + i - 1, 0) * Range("debtRatio"), Range("interestRate"), Range("recLife").Offset(repFirst + i - 1, 0))
    
    For j = 0 To years
        'Add risk to debt interest rate
        debt = debtPayback(-Range("recFC").Offset(repFirst + i - 1, 0) * Range("debtRatio"), Range("interestRate") * (1 + ((2 * risk) * Rnd - risk)), Range("recLife").Offset(repFirst + i - 1, 0))
        If j >= Range("recYearPurchased").Offset(repFirst + i - 1, 0) And j <= Range("recYearPurchased").Offset(repFirst + i - 1, 0) + Range("recLife").Offset(repFirst + i - 1, 0) - 1 Then
            cfArray(i + 1, j + 3) = cf * (1 + ((2 * risk) * Rnd - risk)) ' Apply risk to cash flows
            depArray(i + 1, j + 1) = dep
            debtArray(i + 1, j + 2) = debt
        'ElseIf j = Range("recYearPurchased").Offset(repFirst + i - 1, 0) - 1 Then
        '    cfArray(i + 1, j + 3) = year0
        'Else
         '   cfArray(i + 1, j + 3) = 0
        End If
    Next
Next

Range(Range("cfIndex").Offset(1, 0), Range("cfIndex").Offset(repCount, 2 + years)) = cfArray
Range(Range("depIndex").Offset(1, 0), Range("depIndex").Offset(repCount, 2 + years)) = depArray
Range(Range("debtIndex").Offset(1, 0), Range("debtIndex").Offset(repCount, 2 + years)) = debtArray


Worksheets("Cash Flows").Activate
Worksheets("Cash Flows").Range("F2:AE2").Formula = "=cashOutlay()"
Worksheets("Cash Flows").Calculate
Worksheets("Debt").Activate
Worksheets("Debt").Range("E2:AD2").Formula = "=debtBalance()"
Worksheets("Debt").Calculate
Worksheets("Cash Flows").Activate


Erase cfArray
Erase depArray

End Sub

Function GetFirstCell(CellRef As Range, col As Range) As Long
    Dim l As Long
    Dim lRow As Long
    lRow = Worksheets("Records").Cells(Worksheets("Records").Rows.count, "A").End(xlUp).Row
    l = Application.WorksheetFunction.Match(CellRef.Value, Worksheets("Records").Range(col, Worksheets("Records").Cells(lRow, col.Column)), 0)
    GetFirstCell = l
End Function

Function GetLastCell(CellRef As Range, col As Range, lFirstCell As Long)
    Dim l As Long
    Dim lRow As Long
    lRow = Worksheets("Records").Cells(Worksheets("Records").Rows.count, "A").End(xlUp).Row
    l = Application.WorksheetFunction.CountIf(Range(col, Worksheets("Records").Cells(lRow, col.Column)), CellRef.Value)
    GetLastCell = lFirstCell + l - 1
End Function

Function debtBalance() As Double
Dim rng As Range
Dim i As Long
Dim cell As Range
Dim fRow As Long
Dim lRow As Long
Dim lRng As Range
Dim xRng As Range
Dim col As Long
Dim count As Long
Dim Temp As Double
Dim Result As Double
'Application.Volatile
    Set rng = Application.Caller
    fRow = Worksheets("Debt").Range("debtIndex").Row
    
    col = rng.Column
    lRow = Worksheets("Debt").Cells(Worksheets("Debt").Rows.count, col).End(xlUp).Row
    Set lRng = Worksheets("Debt").Cells(Worksheets("Debt").Rows.count, col).End(xlUp)
    Set xRng = Worksheets("Debt").Range(Cells(fRow + 1, col), lRng)
    count = Application.CountIf(xRng, 0)
    'debtBalance = rng.Address
If Worksheets("Debt").Cells(fRow, col) <> 1 Then
    Temp = 0
    For Each cell In xRng
        If cell <> 0 And cell.Offset(0, -1) = 0 Then
            Temp = Temp + Worksheets("Debt").Cells(cell.Row, Worksheets("Debt").Range("debtIndex").Column + 1)
        End If
    Next
    Result = Temp + rng.Offset(0, -1) - rng.Offset(3, -1)
Else
    Result = Application.Sum(Worksheets("Debt").Range(Cells(fRow + 1, col - 1), Cells(lRow - count, col - 1)))
End If

If Result < 0 Then
    debtBalance = Result
Else
    debtBalance = 0
End If
End Function

Function cashOutlay() As Double
Dim rng As Range
Dim i As Long
Dim cell As Range
Dim fRow As Long
Dim lRow As Long
Dim lRng As Range
Dim xRng As Range
Dim col As Long
Dim count As Long
Dim Temp As Double
Dim Result As Double
'Application.Volatile
    Set rng = Application.Caller
    fRow = Worksheets("Cash Flows").Range("cfIndex").Row
    
    col = rng.Column
    lRow = Worksheets("Cash Flows").Cells(Worksheets("Cash Flows").Rows.count, col).End(xlUp).Row
    Set lRng = Worksheets("Cash Flows").Cells(Worksheets("Cash Flows").Rows.count, col).End(xlUp)
    Set xRng = Worksheets("Cash Flows").Range(Cells(fRow + 1, col), lRng)
    count = Application.CountIf(xRng, 0)
    'debtBalance = rng.Address
If Worksheets("Cash Flows").Cells(fRow, col) <> 1 Then
    Temp = 0
    For Each cell In xRng
        If cell <> 0 And cell.Offset(0, -1) = 0 Then
            Temp = Temp + Worksheets("Cash Flows").Cells(cell.Row, Worksheets("Cash Flows").Range("cfIndex").Column + 2)
        End If
    Next
    Result = Temp
Else
    Result = Application.Sum(Worksheets("Cash Flows").Range(Cells(fRow + 1, col - 1), Cells(lRow - count, col - 1)))
End If

cashOutlay = Result

End Function



