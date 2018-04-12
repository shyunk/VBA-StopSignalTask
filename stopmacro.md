# VBA-StopSignalTask

This code analyzes [Logan's Stop Signal Task](https://www.psytoolkit.org/experiment-library/stopsignal.html) that is modified for an fMRI experiment
```
Sub StopV13()

Dim sh As Worksheet
    Dim Newsh As Worksheet
    Dim RwNum As Long
    Dim Basebook As Workbook
    Dim StopLC As Integer, LastRow As Integer
    Dim NextCol As Long
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim myarray As Variant, newarray As Variant
    Dim firstrown As Long, Orown As Long, frown As Long, arown As Long
    Dim pStopFail As Double, GoRT As Double
    Dim soaRng As Range, soaResult As Double, quantRe As Double, ssrtRe As Double

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    'Delete the sheet "Summary-Sheet" if it exist
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Summary-Sheet").Delete
    'ThisWorkbook.Worksheets("Post-QC-Sheet").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    'Add a worksheet with the name "Summary-Sheet"
    Set Basebook = ThisWorkbook
    Set Newsh = Basebook.Worksheets.Add
    Newsh.Name = "Summary-Sheet"
    'Set NewshT = Basebook.Worksheets.Add
    'NewshT.Name = "Quantile_Check"


    'The links to the first sheet will start in row 2
    RwNum = 1
    'count participants
    participant = 0
    
    myarray = Array("Subject", "Session", "Block", "Arrow", "FixJitter.Duration", "FixJitter.FinishTime", "FixJitter.OnsetTime", "FixJitter.RESP", "FixJitter.RT", "GoImage.ACC", "GoImage.ActionDelay", "GoImage.ActionTime", "GoImage.Duration", "GoImage.DurationError", "GoImage.OffsetDelay", "GoImage.OnsetDelay", "GoImage.OnsetTime", "GoImage.RESP", "GoImage.RT", "GoImage.RTTime", "GoImage1.ACC", "GoImage1.ActionDelay", "GoImage1.ActionTime", "GoImage1.Duration", "GoImage1.DurationError", "GoImage1.OffsetDelay", "GoImage1.OnsetDelay", "GoImage1.OnsetTime", "GoImage1.RESP", "GoImage1.RT", "GoImage1.RTTime", "StopImage.ACC", "StopImage.ActionDelay", "StopImage.ActionTime", "StopImage.Duration", "StopImage.DurationError", "StopImage.OffsetDelay", "StopImage.OnsetDelay", "StopImage.OnsetTime", "StopImage.RESP", "StopImage.RT", "StopImage.RTTime")

    firstrown = 10
            For Each Element In myarray
            Newsh.Cells(1, firstrown).Value = Element
            firstrown = firstrown + 1
            Next Element
            
    newarray = Array("total_Participant", "TCR", "BN", "StopBlock", "StopFail", "PStopFail", "SOA+SSRT", "SOA", "SSRT", "RR")
    'BN Block Number, TCR total correct response rate, SOA stimulus-onset asynchrony, RR Response Rate for stop
    frown = 1
    For Each Element In newarray
    Newsh.Cells(1, frown).Value = Element
    frown = frown + 1
    Next Element
    
    For Each sh In Basebook.Worksheets
    With Newsh
               LastRow = .Cells(.Rows.Count, "K").End(xlUp).Row
    End With

        If sh.Name <> Newsh.Name And sh.Visible Then
            participant = participant + 1
            sh.Select
            'StopLC is the number of rows that exist so far in the participant sheet
                    StopLC = sh.Cells(sh.Rows.Count, "K").End(xlUp).Row
                    If StopLC > 1 Then
                    StopLC = StopLC + 1
                    End If
                ' copy paste important values from each sheet use Orown as the space
                Orown = 10
                i = None
                
                
                For Each Element In myarray
                Rows(1).Select
                Selection.Find(What:=Element, After:=ActiveCell, LookIn:=xlFormulas, Lookat _
                        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                        False, SearchFormat:=False).Activate
                ActiveCell.Offset(1, 0).Resize(StopLC - 2).Copy
                Newsh.Cells(LastRow + 1, Orown).PasteSpecial Paste:=xlPasteValues
                'store action time for TCR calc
                If Element = "GoImage.ActionTime" Then
                i = Orown
                End If
                'store col number for Go Image Response
                If Element = "GoImage.RESP" Then
                k = Orown
                End If
                'store col number for stop
                If Element = "StopImage.ActionTime" Then
                StopCol = Orown
                End If
                'store col number for acc
                If Element = "StopImage.ACC" Then
                StopAccCol = Orown
                End If
                'store col number for GoImage.RT this is used for quantile calc
                If Element = "GoImage.RT" Then
                GoRT = Orown
                End If
                'storing values for SOA
                'SOA part 1: GoImage1.OnsetTime store as GoOnset
                
                If Element = "GoImage1.OnsetTime" Then
                GoOnset = Orown
                'MsgBox (GoOnset)
                End If
                
                'SOA part 2: StopImage.ActionTime store as StopOnset
                If Element = "StopImage.ActionTime" Then
                StopOnset = Orown
                End If
                
                
                ActiveCell.Offset(1, 0).Resize(StopLC - 3).Copy
                Newsh.Cells(LastRow + 1, Orown).PasteSpecial Paste:=xlPasteValues
                

                Orown = Orown + 1
                Next Element
                
                'calculation and analysis step
                
            'calculate new last row
            With Newsh
            NewLastRow = .Cells(.Rows.Count, "K").End(xlUp).Row
            End With
            
            'Derive BN from StopLC
            arown = 3 ' see above for indexing
            Newsh.Cells(LastRow + 1, arown).Value = StopLC - 3
            
            'count total number of stop trials
            StopBlock = 0
            For l = LastRow + 1 To NewLastRow - 1
            TestBlock = Newsh.Cells(l, StopCol)
            If (IsNumeric(TestBlock) = True) And TestBlock > 0 Then
            StopBlock = StopBlock + 1
            'testcolumn
            'Newsh.Cells(StopBlock + 1, arown + 5).Value = Newsh.Cells(l, StopCol)
            End If
            Newsh.Cells(LastRow + 1, arown + 1).Value = StopBlock
            'TCRresp * 100 / TCRTotal
            Next l
            
            StopAcc = 0
            For i = LastRow + 1 To NewLastRow - 1
            TestACCBl = Newsh.Cells(i, StopAccCol)
            If (IsNumeric(TestACCBl) = True) And TestACCBl > 0 Then
            StopAcc = StopAcc + 1
            'Newsh.Cells(i, arown + 5).Value = Newsh.Cells(i, StopAccCol)
            End If
            Next i
            Newsh.Cells(LastRow + 1, arown + 2).Value = StopBlock - StopAcc
            'derive percentage of stop failure
            pStopFail = (StopBlock - StopAcc) / (StopBlock)
            Newsh.Cells(LastRow + 1, arown + 3).Value = pStopFail
            
            'quantile start
            'first copy response time to some other column
            Newsh.Activate
            Newsh.Range(Cells(LastRow, GoRT), Cells(NewLastRow - 1, GoRT)).Copy Destination:=Newsh.Range(Cells(LastRow, Orown + 1), Cells(NewLastRow, Orown + 1))
            If Cells(1, Orown + 1) = "GoImage.RT" Then
            Newsh.Cells(1, Orown + 1).Value = 0
            End If
            Newsh.Sort.SortFields.Clear
            Newsh.Sort.SortFields.Add Key:=Range(Cells(LastRow + 1, Orown + 1), Cells(StopLC - 1, Orown + 1)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

  With Newsh.Sort
        .SetRange Range(Cells(LastRow, Orown + 1), Cells(NewLastRow - 1, Orown + 1))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Dim strArray As Variant
Dim ia As Integer, nonZrow As Integer

If LastRow = 1 Then
nonZrow = LastRow - 1
Else
nonZrow = LastRow
End If

For ia = LastRow To NewLastRow:
If IsNumeric(Cells(ia, Orown + 1)) = True And Cells(ia, Orown + 1) <> 0 Then
nonZrow = nonZrow + 1
End If
Next ia

strArray = Range(Cells(LastRow, Orown + 1), Cells(nonZrow, Orown + 1)).Value

quantRe = WorksheetFunction.Percentile_Exc(strArray, pStopFail)
Newsh.Cells(LastRow + 1, arown + 4).Value = quantRe
            'qauntile end
            
            
            'calcuate soa
        For iSOA = LastRow + 1 To NewLastRow - 1
            If Newsh.Cells(iSOA, GoOnset).Value <> 0 And Newsh.Cells(iSOA, StopOnset).Value <> 0 Then
            Newsh.Cells(iSOA, arown + 5).Value = Newsh.Cells(iSOA, GoOnset).Value - Newsh.Cells(iSOA, StopOnset).Value
            Else
            Newsh.Cells(iSOA, arown + 5).Value = ""
            End If
            Next iSOA
        'calculate ssrt
        'set soaRng for average
        
        Set soaRng = Range(Newsh.Cells(LastRow + 1, arown + 5), Newsh.Cells(NewLastRow - 1, arown + 5))
        soaRng = WorksheetFunction.Average(soaRng)
        soaResult = Newsh.Cells(LastRow + 1, arown + 5).Value
        
        ssrtRe = soaResult + quantRe
        Newsh.Cells(LastRow + 1, arown + 6).Value = ssrtRe

            
        End If
        
        
        
Newsh.Cells(LastRow + 1, 1).Value = participant
    Next sh
    'show number of participants


    Newsh.UsedRange.Columns.AutoFit
    


    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With


End Sub

```




