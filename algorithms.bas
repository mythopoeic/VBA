Attribute VB_Name = "Module1"
Function binSearch(rng As Range, targetValue As Integer)
Dim min As Integer
Dim max As Integer
Dim guess As Integer
Dim count As Integer
min = 1
max = rng.count
count = 0
Do While (max >= min)
    count = count + 1
    guess = (max + min) \ 2
    If rng.Cells(guess) = targetValue Then
        binSearch = CStr(rng.Cells(guess).Address) & ", " & count
        Exit Function
    ElseIf rng.Cells(guess) < targetValue Then
        min = guess + 1
    Else
        max = guess - 1
    End If
Loop
binSearch = "-1, " & count
End Function

Sub swap(rng As Range, idx1 As Integer, idx2 As Integer)
Dim temp As Variant

    temp = rng.Cells(idx1)
    rng.Cells(idx1) = rng.Cells(idx2)
    rng.Cells(idx2) = temp
End Sub

Function indexOfMinimum(rng As Range, startIndex As Integer)
Dim minValue As Integer
Dim minIndex As Integer
Dim i As Integer

minValue = rng.Cells(startIndex)
minIndex = startIndex

For i = minIndex + 1 To rng.count
    If rng.Cells(i) < minValue Then
        minIndex = i
        minValue = rng.Cells(i)
    End If
Next

indexOfMinimum = minIndex

End Function

Sub selectionSort(rng As Range)
Dim idx As Integer
Dim i As Integer
For i = 1 To rng.count
    idx = indexOfMinimum(rng, i)
    Call swap(rng, i, idx)
Next
End Sub

Sub insert(rng As Range, rightIndex As Integer, value As Integer)
Dim idx As Integer

For idx = rightIndex To idx = 1 Step -1
    If idx >= 1 Then
        If rng.Cells(idx) > value Then
            rng.Cells(idx + 1) = rng.Cells(idx)
        Else
            Exit For
        End If
    Else
        Exit For
    End If
Next
rng.Cells(idx + 1) = value
End Sub

Sub insertionSort(rng As Range)
Dim i As Integer
For i = 1 To rng.count - 1
    Call insert(rng, i, rng.Cells(i + 1))
Next
End Sub

Function factorial(n As Integer)
Dim i As Integer
Dim result As Double
result = 1

For i = 1 To n
    result = result * i
Next
factorial = result
End Function

Function rFactorial(n As Integer)
'Factorial by recursion
If n = 0 Then
    rFactorial = 1
Else
    rFactorial = n * rFactorial(n - 1)
End If
End Function

Function firstCharacter(str As String)
    firstCharacter = Left(str, 1)
End Function

Function lastCharacter(str As String)
    lastCharacter = Right(str, 1)
End Function

Function middleCharacters(str As String)
    middleCharacters = (Left(Right(str, Len(str) - 1), Len(str) - 2))
End Function

Function isPalindrome(str As String)
    If Len(str) <= 1 Then 'Base case 1
        isPalindrome = True
        Exit Function
    End If
    
    If firstCharacter(str) = lastCharacter(str) Then
        isPalindrome = isPalindrome(middleCharacters(str)) 'Recursion
    Else
        isPalindrome = False 'Base Case 2
    End If
End Function

Function risEven(n As Integer) As Boolean
    If n Mod 2 = 0 Then
        risEven = True
    Else
        risEven = False
    End If
End Function

Function risOdd(n As Integer) As Boolean
    risOdd = Not (risEven(n))
End Function

Function rPower(x As Integer, n As Integer)
Dim y As Integer
    'Base Case
    If n = 0 Then
        rPower = 1
        Exit Function
    End If
    'recursive case: n is negative
    If n < 0 Then
        rPower = 1 / rPower(x, -n)
        Exit Function
    End If
    'recursive case: n is odd
    If risOdd(n) Then
        rPower = rPower(x, n - 1) * x
        Exit Function
    End If
    'recursive case: n is even
    If risEven(n) Then
        y = rPower(x, n / 2)
        rPower = y * y
        Exit Function
    End If
End Function
