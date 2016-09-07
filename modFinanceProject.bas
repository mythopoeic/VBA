Attribute VB_Name = "modFinanceProject"
Option Explicit 'Forces you to declare all variables
'Private variables are not limited to within Subs, can be used throughout module

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Project runs by calling sub "SortingMethods"
' Assign SortingMethods to a button on Sheet1
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private pCount As Integer
Private startRow As Integer
Private sortKey As Range
Private sortOrder As String
Private o As Integer

Private Sub GenerateProjects()
'Declare variables

Dim minCost As Long
Dim maxCost As Long
Dim pCost As Long
Dim minLife As Integer
Dim maxLife As Integer
Dim pLife As Integer
Dim minIRR As Single
Dim maxIRR As Single
Dim pIRR As Single
Dim i As Integer
Dim j As Integer
Dim pPayment As Long
Dim sortRange As Range
Dim r As Single
Dim npv As Single

'''''''''''''''''''''''''''''
'1. Generate 25 projects
'''''''''''''''''''''''''''''

minCost = 50000
maxCost = 300000
minLife = 2
maxLife = 10
minIRR = 0.05
maxIRR = 0.45

' Clear old values and set heading labels
Range("B" & startRow) = "Project Cost"
Range("C" & startRow) = "Project Life"
Range("D" & startRow) = "Project IRR"
Range("E" & startRow) = "Yearly Cash Flow"

Randomize ' Sets the Rnd function seed value

For i = 1 To pCount
    pCost = Int((maxCost + 1 - minCost) * Rnd + minCost) ' Use Int for whole dollar values, add 1 because Int cuts off decimal values
    pLife = Int((maxLife + 1 - minLife) * Rnd + minLife) ' Use Int to get discrete integer values
    pIRR = (maxIRR - minIRR) * Rnd + minIRR
    pPayment = Pmt(pIRR, pLife, pCost)
    
    Range("A" & startRow + i) = "Project " & i
    Range("B" & startRow + i) = pCost
    Range("C" & startRow + i) = pLife
    Range("D" & startRow + i) = pIRR
    Range("D" & startRow + i).Value = Round(Range("D" & startRow + i), 3) ' Round the IRR to 3 decimal places (0.xxx or xx.x%)
    Range("E" & startRow + i) = -pPayment ' negative to get positive cash flow
    Range("E" & startRow + i).Value = Round(Range("E" & startRow + i), 0)
Next

'''''''''''''''''''''''''''''
'2. Compute payback
'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''
'5. Compute NPV
'''''''''''''''''''''''''''''
Range("F" & startRow) = "Payback"
Range("G" & startRow) = "NPV"

r = 0.15 ' discount rate

' Payback equals Project Cost divided by Yearly Cash Flow
For i = 1 To pCount

    Range("F" & startRow + i) = Range("B" & startRow + i) / Range("E" & startRow + i)
    
    ' NPV is the sum of the discounted cash flows
    npv = 0
    
    For j = 0 To Range("C" & startRow + i)
        If j = 0 Then
            npv = -Range("B" & startRow + i)
        Else
            npv = npv + Range("E" & startRow + i) / ((1 + r) ^ j) ' equation for one year return
        End If
    Next
    
    Range("G" & startRow + i) = npv
Next

' Sort the projects by various ranking methods
' sortKey is the column sorted by, sortOrder is either ascending or descending
Set sortRange = Range("A" & startRow & ":G" & startRow + pCount)

sortRange.Sort Key1:=sortKey, Order1:=sortOrder, Header:=xlYes

End Sub

Private Sub FiveYearProjectBuy()
'''''''''''''''''''''''''''''
'3. Buy projects
'''''''''''''''''''''''''''''
Dim totalYears As Integer
Dim CapBudget As Long
Dim totalSpent As Long
Dim lastCost As Long
Dim yearlyProfit As Long
Dim i As Integer
Dim j As Integer
Dim k As Integer


Worksheets("Sheet1").Range("A36:X219").Clear ' remove any old data from sheet
totalYears = 5
CapBudget = 600000

startRow = 36

For i = 1 To totalYears

    If i <> 1 Then
        startRow = startRow + pCount + 2 ' Add two to create space between years
    End If
    
    Range("A" & startRow) = "Year " & i
    
    pCount = 20 + 5 * i ' increase the number of projects generated per year
    
    Call GenerateProjects
    
    Range("I" & startRow) = "Projects Purchased"
    Range("J" & startRow) = "Purchase Amount"
    For j = 1 To 14
       Cells(startRow, 10 + j) = "Year " & j
    Next
    
    CapBudget = CapBudget - totalSpent + yearlyProfit ' determine budget to spend on new projects
    
    totalSpent = 0
    yearlyProfit = 0
    
    ' select projects if they do not exceed the budget amount, copy to area on right
    
    For j = startRow + 1 To startRow + pCount
        If totalSpent + Range("B" & j) > CapBudget Then
        Else
            Range("I" & j) = Range("A" & j)
            Range("J" & j) = Range("B" & j)
            For k = i To Range("C" & j) + i - 1 ' copy yearly cash flows
                Cells(j, 10 + k) = Range("E" & j)
            Next
            
            totalSpent = totalSpent + Range("B" & j)
            
        End If
    Next
    
    'total cash flows from all projects for given year, using Excel's sum function
    yearlyProfit = Application.Sum(Range(Cells(11, 10 + i), Cells(11 + startRow + pCount, 10 + i)))
    
Next

End Sub

Private Sub CalcNetWealth30Cycles()
'''''''''''''''''''''''''''''
'4. Calculate Net Wealth
'''''''''''''''''''''''''''''
Dim netWealth As Single
Dim totalWealth As Single
Dim average As Single
Dim variance As Single
Dim stDev As Single
Dim N As Integer
Dim i As Integer

Worksheets("Sheet1").Range("A36:X219").Clear
N = 30
netWealth = 0
totalWealth = 0

Range("D2").Offset(0, o) = "Cycle" ' Use Offset function to shift columns over for different sorting methods
Range("E2").Offset(0, o) = "Net Wealth"

For i = 1 To N
    Call FiveYearProjectBuy
    
    ' Use Excel function to add all values from Year 5 and later
    netWealth = Application.Sum(Range("O37:X219"))
    totalWealth = totalWealth + netWealth
    Range("D" & i + 2).Offset(0, o) = "# " & i
    Range("E" & i + 2).Offset(0, o) = netWealth
Next

average = totalWealth / N

Range("D33").Offset(0, o) = "Average:"
Range("E33").Offset(0, o) = average

'calculate variance to ge the standard deviation
variance = 0

For i = 1 To N
    variance = variance + (average - Range("E" & i + 2).Offset(0, o)) ^ 2
Next
variance = variance / N ' divide total by number of samples
stDev = variance ^ 0.5

Range("D34").Offset(0, o) = "St. Dev.:"
Range("E34").Offset(0, o) = stDev

End Sub

Sub SortingMethods() ' Not private so it can be assigned to the button
'''''''''''''''''''''''''''''
'5. Calculate Net Wealth for
'   different sorting methods
'''''''''''''''''''''''''''''
Worksheets("Sheet1").UsedRange.Clear

'Rank projects by sorting from lowest to highest payback value
Range("D1") = "Payback"
Set sortKey = Range("F36")
sortOrder = xlAscending
o = 0 ' initially no offset
Call CalcNetWealth30Cycles

'Rank projects by sorting from highest to lowest NPV
Range("F1") = "NPV"
Set sortKey = Range("G36")
sortOrder = xlDescending
o = 2
Call CalcNetWealth30Cycles

'Rank projects by sorting from highest to lowest IRR
Range("H1") = "IRR"
Set sortKey = Range("D36")
sortOrder = xlDescending
o = 4
Call CalcNetWealth30Cycles

End Sub

