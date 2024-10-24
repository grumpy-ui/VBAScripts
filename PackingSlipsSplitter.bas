Attribute VB_Name = "PackingSlipsSplitter"
Function GetFullNumber(baseNumber As Long, num As String) As Long
    ' Given a base number and a short form number, get the full number
    Dim numVal As Long
    Dim baseLength As Integer
    Dim numLength As Integer

    numVal = Val(num)
    baseLength = Len(CStr(baseNumber))
    numLength = Len(num)

    If numLength < baseLength Then
        ' Adjust logic for continuation of numbers
        GetFullNumber = (baseNumber \ (10 ^ numLength)) * (10 ^ numLength) + numVal
        If GetFullNumber < baseNumber Then
            GetFullNumber = GetFullNumber + (10 ^ numLength)
        End If
    Else
        ' If the partial number is the same length or longer, it's a full number
        GetFullNumber = numVal
    End If
End Function

Sub ProcessPackingSlipsV2()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim packSlips As String, shipId As String
    Dim slips() As String, subSlips() As String
    Dim baseNumber As Long, fullNumber As Long, startRange As Long, endRange As Long
    Dim lastFullNumber As Long
    Dim slipSet As Object
    Set slipSet = CreateObject("Scripting.Dictionary")
    Dim wsSourceRowCnt As Long


    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set wsSource = ActiveWorkbook.ActiveSheet
    Set wsTarget = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.ActiveSheet, Count:=ActiveWorkbook.Worksheets.Count)
    wsTarget.Name = "ProcessedPackingSlips"
    wsTarget.Cells(1, 1).Value = "Packing Slip"
    wsTarget.Cells(1, 2).Value = "Shipment ID"

    Dim rowTarget As Long
    rowTarget = 2

    For i = 2 To wsSource.Cells(wsSource.Rows.Count, 8).End(xlUp).Row
        packSlips = wsSource.Cells(i, 8).Value
        shipId = wsSource.Cells(i, 2).Value
        slips = Split(packSlips, ",") ' Split entries by comma
        lastFullNumber = 0 ' Initialize the last full number for each row

        For Each slip In slips
            subSlips = Split(slip, ";") ' Further split by semicolon if necessary

            For j = LBound(subSlips) To UBound(subSlips)
                If InStr(subSlips(j), "-") > 0 Then
                    ' It's a range
                    Dim rangeParts() As String
                    rangeParts = Split(subSlips(j), "-")
                    baseNumber = GetFullNumber(lastFullNumber, rangeParts(0))
                    lastFullNumber = baseNumber ' Update last full number

                    ' Calculate the start and end of the range
                    startRange = baseNumber
                    endRange = GetFullNumber(startRange, rangeParts(1))

                    ' Add the range of numbers to the worksheet, ensuring each is only added once
                    For k = startRange To endRange
                        If Not slipSet.Exists(k) Then
                            slipSet.Add k, True
                            wsTarget.Cells(rowTarget, 1).Value = k
                            wsTarget.Cells(rowTarget, 2).Value = shipId
                            wsTarget.Cells(rowTarget, 3).Value = "YES"
                            rowTarget = rowTarget + 1
                        End If
                    Next k
                Else
                    ' It's an individual number, possibly after a range
                    fullNumber = GetFullNumber(lastFullNumber, subSlips(j))
                    lastFullNumber = fullNumber ' Update last full number

                    If Not slipSet.Exists(fullNumber) Then
                        slipSet.Add fullNumber, True
                        wsTarget.Cells(rowTarget, 1).Value = fullNumber
                        wsTarget.Cells(rowTarget, 2).Value = shipId
                        wsTarget.Cells(rowTarget, 3).Value = "YES"
                        rowTarget = rowTarget + 1
                    End If
                End If
            Next j
        Next slip
    Next i
    
    wsSourceRowCnt = Application.WorksheetFunction.CountA(Range("B:B"))
    
    
        wsSource.Range("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    wsSource.Range("C1").Value = "Processed"
    wsSource.Range("C2").FormulaR1C1 = _
        "=XLOOKUP(@C[-1],ProcessedPackingSlips!C[-1],ProcessedPackingSlips!C,""NO"")"
    wsSource.Range("C2").AutoFill Destination:=wsSource.Range("C2:C" & wsSourceRowCnt)
    Calculate
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalc
End Sub





