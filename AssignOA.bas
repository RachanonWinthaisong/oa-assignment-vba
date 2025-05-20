Attribute VB_Name = "Module3"
Sub AssignOA_ByPercentage_AvoidDuplicates()

    Dim wsData As Worksheet, wsMaster As Worksheet
    Dim lastRow As Long, i As Long
    Dim province As String, assignedOA As String
    Dim pastOAs(1 To 4) As String
    Dim possibleOAs As Collection
    Dim totalPct As Double
    Dim randVal As Double, cumPct As Double
    Dim j As Long
    
    Set wsData = ThisWorkbook.Sheets("Sheet1")
    Set wsMaster = ThisWorkbook.Sheets("OA_Master")
    
    lastRow = wsData.Cells(wsData.Rows.Count, "AI").End(xlUp).Row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 2 To lastRow
        province = Trim(wsData.Cells(i, "AI").Value)
        
        ' ???? OA 4 ?????????????
        pastOAs(1) = Trim(wsData.Cells(i, "S").Value) ' May
        pastOAs(2) = Trim(wsData.Cells(i, "T").Value) ' Apr
        pastOAs(3) = Trim(wsData.Cells(i, "U").Value) ' Mar
        pastOAs(4) = Trim(wsData.Cells(i, "V").Value) ' Feb
        
        Set possibleOAs = New Collection
        
        ' ????????? OA ??? OA_Master
        Dim mLastRow As Long
        mLastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
        
        totalPct = 0
        For j = 2 To mLastRow
            If Trim(wsMaster.Cells(j, "A").Value) = province Then
                Dim thisOA As String
                Dim thisPct As Double
                
                thisOA = Trim(wsMaster.Cells(j, "B").Value)
                thisPct = wsMaster.Cells(j, "C").Value
                
                If Not IsInArray(thisOA, pastOAs) Then
                    possibleOAs.Add Array(thisOA, thisPct)
                    totalPct = totalPct + thisPct
                End If
            End If
        Next j
        
        ' ???????? OA ??????????????????? ?????? OA ??????????????????
        If possibleOAs.Count = 0 Then
            Dim fallbackOA As String
            For j = 4 To 1 Step -1 ' ???????? Feb ? Mar ? Apr ? May
                If pastOAs(j) <> "" Then
                    fallbackOA = pastOAs(j)
                    Exit For
                End If
            Next j
            assignedOA = fallbackOA
        Else
            ' ????????? OA ??? possibleOAs ??????????
            randVal = Rnd * totalPct
            cumPct = 0
            
            For j = 1 To possibleOAs.Count
                cumPct = cumPct + possibleOAs(j)(1)
                If randVal <= cumPct Then
                    assignedOA = possibleOAs(j)(0)
                    Exit For
                End If
            Next j
        End If
        
        wsData.Cells(i, "R").Value = assignedOA
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Assignment Complete!", vbInformation

End Sub

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim k As Variant
    For Each k In arr
        If val = k Then
            IsInArray = True
            Exit Function
        End If
    Next k
    IsInArray = False
End Function
