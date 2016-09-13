Attribute VB_Name = "modExporttoPPT"
Option Explicit
Public ppReport As Boolean
Public workStreamlist() As String
' Declaring Public objects for use through out project

Public ppApp As Object, ppPres As Object
Public pSlide As Object, sld As Object, Shape As Object
Public xlApp As Excel.Application, xlsheet As Excel.Worksheet, aRange As Excel.Range, bRange As Excel.Range, wkbk As Excel.Workbook
Public strPath As String, strFileName As String, strRange As String
Public sType As String, strWkSh As String, SheetName As String, RangeName As String, OldSheetName As String
Public T As Integer, l As Integer, H As Integer, w As Integer, s As Integer


#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If



'For each item in workStreamlist
''change values according to workstream list

'projteam = uSelect.lActivities.List(jcounter)
'wksSetup.Range("ProjectTeam") = projteam
'ActiveWorkbook.Save

'Call XLExport

'Next




'------------ Used for Production -----------------

Sub GetPPTFile2()
    On Error GoTo Err_GetPPTFile
    
    ' Update All Outputs
    '   Call GraphRowsUpdate
    '   Call UpdateUA
    '   Call UpdateAAR
    '   Call UpdateMD
    
    
    ' Update All Outputs
    Dim strPathFile
    Dim nday
    Dim brandName
    Dim j
    Dim S_Old
    Dim S_New
    Dim SlideSkip As Integer
    Dim spTmark As Integer
    Dim k As Integer
    Dim StSlidesRange As Range
    Dim StSlidesCtRange As Range
    Dim STSlideCt As Integer
    Dim SlideMax As Integer
    Dim exportBool As Boolean
    Dim multi As Boolean
    Dim d As Integer
    Dim cell As Range
    Dim scount As Integer
    Dim lRow As Integer
    Dim lCol As Integer
    'Open Template PPT File
    'strFileName = "Adynovate Launch Report Template.pptx"
    'strPath = Application.ActiveWorkbook.Path & "\"
  
    strPathFile = ActiveWorkbook.Path
    nday = Application.WorksheetFunction.Text(Date, "YYYY-MMM-DD")
    brandName = "Sample"
    
    strPath = Application.GetOpenFilename("PowerPoint Files (*.ppt*), *.ppt*", , "Please select a PowerPoint Template...", , False)
    
    Set ppApp = CreateObject("PowerPoint.Application")
    ppApp.Visible = True
    
    DoEvents
    Application.Wait (Now + TimeValue("0:00:10"))
    
    Set ppPres = ppApp.Presentations.Open(strPath)
    Set pSlide = ppPres.Slides
    
    
    
    
    Dim i As Integer
    For i = 1 To 6
        DoEvents
        Sleep 500 'milliseconds
    Next i
    
    
    
    ppApp.Visible = msoTrue
    
    
    ppApp.ActivePresentation.SaveAs Filename:=strPathFile & "\" & brandName & " (" & "Slide Output" & ") - " & nday & ".pptx"
    
    Set aRange = Application.Worksheets("Setup").Range("Setup_Range")
    'Set StSlidesCtRange = Application.Worksheets("Setup").Range("StSlidesCtRange")
    exportBool = True
    d = 0
    ' Delete Previous Charts and Tables Search the Setup Range to get slide information.
    For Each bRange In aRange.Rows
        
        SheetName = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 2, False)  'Sheet Name
        RangeName = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 4, False) ' Range Name or Chart Name
        T = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 5, False) ' Top
        l = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 6, False) ' Left
        H = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 7, False) ' Height
        w = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 8, False) ' Width
        s = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 3, False) ' Slide index
        sType = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 9, False) ' Type
        S_Old = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range"), 3, False) ' Slide index
        
        For i = LBound(workStreamlist) To UBound(workStreamlist)
            If workStreamlist(i) = "All Charts" Then
                exportBool = True
                Exit For
            ElseIf workStreamlist(i) = SheetName Then
                exportBool = True
                Exit For
            Else
                exportBool = False
            End If
            
        Next i
        
        
        s = s - d
        
        pSlide(s).Select
        If exportBool = False Then
            pSlide(s).Delete
            d = d + 1
        Else
            Set sld = pSlide(s)
            Set xlsheet = ThisWorkbook.Sheets(SheetName)
            xlsheet.Select
            
            For Each Shape In sld.Shapes
                
                Select Case Shape.Type
                    Case msoTable
                    If Shape.Name = RangeName Then
                        Shape.Delete                        ' deletes existing table
                        xlsheet.Range(RangeName).Select     ' selects range in worksheet
                        xlsheet.Range(RangeName).Copy
                        'Application.Selection.Copy          ' Copies range
                        
                        
                        ppPres.Windows(1).Activate
                        ppApp.ActiveWindow.View.GotoSlide s
                        ppApp.ActiveWindow.View.PasteSpecial (10)      ' pastes range to ppt file
                        
                        'Formats table in slide
                        ppApp.ActiveWindow.Selection.ShapeRange.Name = RangeName
                        ppApp.ActiveWindow.Selection.ShapeRange.Top = T
                        ppApp.ActiveWindow.Selection.ShapeRange.Left = l
                        ppApp.ActiveWindow.Selection.ShapeRange.Height = H
                        ppApp.ActiveWindow.Selection.ShapeRange.Width = w

                        DoEvents
                        DoEvents
                        Application.CutCopyMode = False
                        Application.Wait (Now + TimeValue("0:00:01"))
                        DoEvents
                        DoEvents
                    Else
                    End If
                    
                    Case msoChart
                    If Shape.Name = RangeName Then
                        Shape.Delete
                        xlsheet.ChartObjects(RangeName).Activate
                        'Application.ActiveSheet.ChartObjects(RangeName).Activate
                        
                        xlsheet.ChartObjects(RangeName).Copy
                        'Application.ActiveChart.ChartArea.Copy
                        
                        ppPres.Windows(1).Activate
                        'ppApp.ActiveWindow.View.GotoSlide S
                        sld.Shapes.PasteSpecial (10)
                        sld.Select
                        sld.Shapes(sld.Shapes.Count).Name = RangeName
                        sld.Shapes(RangeName).Top = T
                        sld.Shapes(RangeName).Left = l
                        sld.Shapes(RangeName).Height = H
                        sld.Shapes(RangeName).Width = w
                        DoEvents
                        DoEvents
                        Application.CutCopyMode = False
                        Application.Wait (Now + TimeValue("0:00:01"))
                        DoEvents
                        DoEvents
                    End If
                    Case msoPicture
                    
                    If Shape.Name = RangeName Then
                        Shape.Delete
                        xlsheet.Range(RangeName).Select     ' selects range in worksheet
                        xlsheet.Range(RangeName).Copy
                        For i = 1 To 6
                            DoEvents
                            Sleep 500 'milliseconds
                        Next i
                        
                        ppPres.Windows(1).Panes(2).Activate
                        ppApp.ActiveWindow.View.GotoSlide s
                        ppApp.ActiveWindow.View.PasteSpecial (2)       ' pastes range to ppt file
                        
                        'Formats table in slide
                        ppApp.ActiveWindow.Selection.ShapeRange.Name = RangeName
                        ppApp.ActiveWindow.Selection.ShapeRange.Top = T
                        ppApp.ActiveWindow.Selection.ShapeRange.Left = l
                        ppApp.ActiveWindow.Selection.ShapeRange.Height = H
                        ppApp.ActiveWindow.Selection.ShapeRange.Width = w
                        DoEvents
                        DoEvents
                        Application.CutCopyMode = False
                        Application.Wait (Now + TimeValue("0:00:01"))
                        DoEvents
                        DoEvents
                    Else
                    End If
                    
                    '                    Shape.Top = 521.4286
                    '                    Shape.Left = 522
                    '                    Shape.Height = 18
                    '                    Shape.Width = 162
                    '                    Shape.Name = "Picture 19"
                    
                End Select
                
                Application.CutCopyMode = False
                
            Next
        End If
    Next
    '----------------------------------------------------------
    'multi-slide section
    '----------------------------------------------------------
    
    
    
    Set aRange = Application.Worksheets("Setup").Range("Setup_Range2")
    Set StSlidesRange = Application.Worksheets("Setup").Range("StSlidesRange")
    'Set StSlidesCtRange = Application.Worksheets("Setup").Range("StSlidesCtRange")
    Application.Worksheets("Setup").Range("Setup2EStart") = "=Setup2StartSlide"
    SlideSkip = Application.Worksheets("Setup").Range("SlideSkip").Value
    
    'For section = 1 To 10
    
    'Application.Worksheets("Upcoming Activities - Starting").DropDowns("Level1Sort").ListIndex = section
    
    'Call AllReports 'Generates reports for the current workstream
    
    'Application.Calculate
    
    spTmark = 0
    d = 2
    exportBool = True
    
            For i = LBound(workStreamlist) To UBound(workStreamlist)
                If workStreamlist(i) = "All Charts" Then
                    d = d - 2
                    Exit For
                ElseIf workStreamlist(i) = "Summary - full valuation" Then
                    
                    d = d - 1
                    
                    
                ElseIf workStreamlist(i) = "Summary" Then
                    d = d - 1
                    
                End If
                
            Next i
    
    
    ' Delete Previous Charts and Tables Search the Setup Range to get slide information.
    ' If multi = True Then
    
    For Each bRange In aRange.Rows
        
        S_Old = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 3, False)
        If S_Old = 0 Or S_Old > 11 Then
            
            
            'Exit Sub
        Else
            
            SheetName = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 2, False)  'Sheet Name
            RangeName = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 4, False) ' Range Name or Chart Name
            T = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 5, False) ' Top
            l = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 6, False) ' Left
            H = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 7, False) ' Height
            w = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 8, False) ' Width
            
            S_New = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 10, False) ' Slide index
            sType = Application.WorksheetFunction.VLookup(bRange.Cells(1, 1), Application.Worksheets("Setup").Range("Setup_Range2"), 9, False) ' Type
            ' Additional variables for split tables
            j = bRange.Cells(1, 11) ' Hide Row
            
            
            
            For i = LBound(workStreamlist) To UBound(workStreamlist)
                If workStreamlist(i) = "All Charts" Then
                    exportBool = True
                    Exit For
                    ElseIf workStreamlist(i) = SheetName Then
                    exportBool = True
                    'd = d + 1
                    Exit For
                    
                Else
                    exportBool = False
                    
                End If
                
            Next i
            
            
            
            
            'Save file
            
            
            'ppApp.ActivePresentation.Save
            
            
            For k = 1 To 10
                DoEvents
                Sleep 500 'milliseconds
            Next k
            
            
            ' Application.Wait (Now + TimeValue("0:00:05"))
            
            
            scount = ppPres.Slides(ppPres.Slides.Count).SlideNumber
            
            'Select slide for positioning
            If S_New - d <= scount Then
            Set sld = pSlide(S_New - d)
            
            If exportBool = False Then
            If SheetName <> OldSheetName Then
                If SheetName = "Division Presidents" Or SheetName = "CEO Report" Or SheetName = "Assets Passed On" Then
                 sld.Delete
                 d = d + 1

                End If
            Else
            d = d + 1
            End If

            Else
                
                ' Make copies of template slides for split tables - i.e., report tables requiring >1 slide
                ' If count of "SplitChart" >0
                
                
                
                
                For Each cell In StSlidesRange
                    'Make copies of template slides for report tables that are too large
                    If S_Old = cell.Value Then 'And S_New <> 27
                        If S_Old <> spTmark Then
                            STSlideCt = cell.Offset(0, 4).Value
                            'Set pptLayout = ActivePresentation.Slides(S_Old).CustomLayout
                            SlideMax = Application.Min(STSlideCt, 5)
                            If STSlideCt > 0 Then
                                For i = 1 To SlideMax - 1
                                    sld.Duplicate
                                Next i
                            End If
                            spTmark = S_Old
                        End If
                    End If
                Next cell
            
            
            
            'Switch slide selection to new position
            Set sld = pSlide(S_New - d)
            
            'Remove prior iteration of the object if it is present
            For Each Shape In sld.Shapes
                If Shape.Name = RangeName Then
                    Shape.Delete
                End If
            Next
            
            'Add new content
            '====================================================================
            If sType = "Chart" Then
                Worksheets(SheetName).Activate
                Worksheets(SheetName).ChartObjects(RangeName).Activate
                Worksheets(SheetName).ChartObjects(RangeName).Chart.ChartArea.Copy
                'Spend time to ensure proper clipboard loading/unloading
                
                
                For k = 1 To 10
                    DoEvents
                    Sleep 500 'milliseconds
                Next k
                Application.Wait (Now + TimeValue("0:00:05"))
                
                'DoEvents
                'DoEvents
                'Application.Wait (Now + TimeValue("0:00:05"))
                'DoEvents
                'Paste into slide
                ppPres.Windows(1).Activate
                ppApp.ActiveWindow.View.GotoSlide S_New - d
                ppApp.ActiveWindow.View.Paste       'pastes chart to ppt file
                
                'Format object in slide - name, size & position
                ppApp.ActiveWindow.Selection.ShapeRange.Name = RangeName
                ppApp.ActiveWindow.Selection.ShapeRange.Top = T
                ppApp.ActiveWindow.Selection.ShapeRange.Left = l
                ppApp.ActiveWindow.Selection.ShapeRange.Height = H
                ppApp.ActiveWindow.Selection.ShapeRange.Width = w
                
                
                Application.CutCopyMode = False
                For k = 1 To 20
                    DoEvents
                    Sleep 500 'milliseconds
                Next k
                
                
                
            Else
                
                If sType = "Table" Then
                    Worksheets(SheetName).Activate
                    Worksheets(SheetName).Range(Range(RangeName).Value).Select     ' selects range in worksheet
                    Worksheets(SheetName).Range(Range(RangeName).Value).Copy
                    'Application.Selection.Copy          ' Copies range
                    
                    For k = 1 To 6
                        DoEvents
                        Sleep 500 'milliseconds
                    Next k
                    
                    ppPres.Windows(1).Activate
                    ppApp.ActiveWindow.View.GotoSlide S_New - d
                    ppApp.ActiveWindow.View.PasteSpecial (0)       ' pastes range to ppt file
                    
                    'Formats table in slide
                    ppApp.ActiveWindow.Selection.ShapeRange.Name = RangeName
                    ppApp.ActiveWindow.Selection.ShapeRange.Top = T
                    ppApp.ActiveWindow.Selection.ShapeRange.Left = l
                    ppApp.ActiveWindow.Selection.ShapeRange.Height = H
                    ppApp.ActiveWindow.Selection.ShapeRange.Width = w
                            sld.Shapes(RangeName).Select
        
                        With sld.Shapes(RangeName).Table
                        'sld.Shapes(RangeName).TextRange.Font.Size = 10
                        For lRow = 1 To .Rows.Count
                        For lCol = 1 To .Columns.Count
                        With .cell(lRow, lCol).Shape
        
                            .TextFrame.TextRange.Font.Size = 10
                        End With
                        Next
                        Next
                        End With
                    DoEvents
                    DoEvents
                    Application.CutCopyMode = False
                    For k = 1 To 6
                        DoEvents
                        Sleep 500 'milliseconds
                    Next k
                    
                    Application.Wait (Now + TimeValue("0:00:02"))
                Else
                End If
                
                
                ' If RangeName = "TableY" Then
                '    ActiveWorkbook.Names("TableY").Delete
                '  'Added SH
                '   Dim tylastRow As Integer
                '  tylastRow = ActiveWorkbook.Worksheets("Missing Dates").Cells(Rows.Count, 2).End(xlUp).Row
                ' ActiveWorkbook.Names.Add "TableY", RefersTo:="='Missing Dates'!$B$7:$L$" & tylastRow
                '   ActiveWorkbook.Names.Add "Table7", RefersTo:="='Missing Dates'!$B$4:$K$18"
                
                'End If
                
                '     If RangeName = "TableX" Then
                '        ActiveWorkbook.Names("TableX").Delete
                '       'Added SH
                '       Dim t8lastRow As Integer
                '       t8lastRow = ActiveWorkbook.Worksheets("Activities at Risk").Cells(Rows.Count, 2).End(xlUp).Row
                '       ActiveWorkbook.Names.Add "Table8", RefersTo:="='Activities at Risk'!$B$4:$K$" & t8lastRow
                '  '        ActiveWorkbook.Names.Add "Table8", RefersTo:="='Activities At Risk'!$B$4:$K$4"
                
            End If
            
            
            Application.CutCopyMode = False
            
            End If
        End If
        End If
        OldSheetName = SheetName
    Next
    
    'Else
    
    'End If
    
    Application.Worksheets("Setup").Range("Setup2EStart") = S_New + SlideSkip
    
    
                For i = LBound(workStreamlist) To UBound(workStreamlist)
                If workStreamlist(i) = "All Charts" Then
                    exportBool = True
                    Exit For
                    ElseIf workStreamlist(i) = "Executive Summary" Then
                    exportBool = True
                    'd = d + 1
                    Exit For
                    
                Else
                    exportBool = False
                    
                End If
                
            Next i

    If exportBool = False Then
        Set sld = pSlide(2)
        sld.Delete
    End If
    
    
    ' Delete Previous Charts and Tables Search the Setup Range to get slide information.
    ' If multi = True Then
    

    ' Goto first page in PowerPoint, select the Dashboard sheet, verify that the process is completed
    'ppPres.Save
    ppApp.ActivePresentation.Slides(1).Select
    Application.ActiveWorkbook.Sheets("Summary - full valuation").Select
    Application.Worksheets("Summary - full valuation").Select
    'MsgBox "Export Completed", vbInformation, "Comprehensive Reporting"
    
    
    Set ppApp = Nothing
    Set ppPres = Nothing
    SheetName = ""
    Set pSlide = Nothing
    Set sld = Nothing
    strPath = ""
    strFileName = ""
    Erase workStreamlist
Exit_GetPPTFile:
    Exit Sub
Err_GetPPTFile:
    MsgBox "Error Line: " & Erl & " Description: " & Err.Description
    Resume Exit_GetPPTFile
    
End Sub


