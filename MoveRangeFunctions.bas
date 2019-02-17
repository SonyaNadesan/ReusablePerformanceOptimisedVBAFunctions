Attribute VB_Name = "MoveRangeFunctions"
'Sonya - Move Ranges

Sub shiftData_keepDataWithinRanges(ByVal sheetname As String, ByVal columnShift As Integer, ByRef tableRangesAsString() As String, Optional ByVal showConfirm = True)
  With Worksheets(sheetname)
    Dim warnUser As Boolean
    Dim allowShiftData As Boolean
    Dim subAddress As String
    Dim subTable As Range
    
    'set default values
    warnUser = False
    allowShiftData = False
    Dim allowedShift As Integer
    Dim abs_columnShift As Integer
    
    'check if confirmation is to be shown
    If showConfirm = True Then
        'check if data will be lost; assign the result to variable which will be used to determine whether we will display warning message
        warnUser = willDataBeLost(tableRangesAsString, sheetname, columnShift)
    End If
    
    'check whether the shift is to the left(-) or right(+) and store the number of shifts in a varible
    If columnShift < 0 Then
        abs_columnShift = columnShift * -1
    Else
        abs_columnShift = columnShift
    End If
    
    'check if user is to be warned
    If warnUser = False Then
            Dim aRangeAsStr As Variant
            'loop through the ranges
            For Each aRangeAsStr In tableRangesAsString
                If aRangeAsStr <> vbNullString Then
                    'shifting to the left
                    If columnShift < 0 Then
                        allowedShift = getNumberOfEmptyCellsAtTheStartOfRow(aRangeAsStr, sheetname)
                    Else
                        'shifting to the right
                        If columnShift > 0 Then
                            allowedShift = getNumberOfEmptyCellsAtTheEndOfRow(aRangeAsStr, sheetname)
                        End If
                    End If
                    subAddress = getSubTableWithinRange(aRangeAsStr, sheetname)
                    If subAddress <> vbNullString Then
                        Set subTable = .Range(subAddress)
                        If allowedShift < abs_columnShift And columnShift > 0 Then
                            Call emptyLastXcolumnsInRow(abs_columnShift - allowedShift, subAddress, sheetname)
                        Else
                            If allowedShift < abs_columnShift And columnShift < 0 Then
                                Call emptyFirstXcolumnsInRow(abs_columnShift - allowedShift, subAddress, sheetname)
                            End If
                        End If
                        subAddress = getSubTableWithinRange(aRangeAsStr, sheetname)
                        Call shiftData(sheetname, subTable, columnShift, 0)
                    End If
                End If
            Next aRangeAsStr
    Else
        If showConfirm = True Then
                If MsgBox("Warning: You may lose data. Would you like to proceed?", vbExclamation + vbYesNo) = vbYes Then
                    allowShiftData = True
                Else
                    allowShiftData = False
                End If
        Else
            allowShiftData = True
        End If
        If allowShiftData = True Then
            Dim aRangeAsStr2 As Variant
            For Each aRangeAsStr2 In tableRangesAsString
                If aRangeAsStr2 <> vbNullString Then
                    If columnShift < 0 Then
                        allowedShift = getNumberOfEmptyCellsAtTheStartOfRow(aRangeAsStr2, sheetname)
                    Else
                        If columnShift > 0 Then
                            allowedShift = getNumberOfEmptyCellsAtTheEndOfRow(aRangeAsStr2, sheetname)
                        End If
                    End If
                    subAddress = getSubTableWithinRange(aRangeAsStr2, sheetname)
                    If subAddress <> vbNullString Then
                        Set subTable = .Range(subAddress)
                        If allowedShift < abs_columnShift And columnShift > 0 Then
                            Call emptyLastXcolumnsInRow(abs_columnShift - allowedShift, subAddress, sheetname)
                        Else
                            If allowedShift < abs_columnShift And columnShift < 0 Then
                                Call emptyFirstXcolumnsInRow(abs_columnShift - allowedShift, subAddress, sheetname)
                            End If
                        End If
                        subAddress = getSubTableWithinRange(aRangeAsStr2, sheetname)
                        If subAddress <> vbNullString Then
                            Set subTable = .Range(subAddress)
                            Call shiftData(sheetname, subTable, columnShift, 0)
                        End If
                    End If
                End If
            Next aRangeAsStr2
        End If
    End If
  End With
End Sub
Sub shiftData(ByVal sheetname As String, ByRef rng As Range, ByVal columnShift As Integer, ByVal rowShift As Integer)
    Worksheets(sheetname).Activate
    Dim noOfCells As Integer
    noOfCells = rng.Count
    Dim valuesInRng() As String
    ReDim valuesInRng(noOfCells)
    Dim cell As Range
    Dim i As Integer
    i = 0
    For Each cell In rng.Cells
        valuesInRng(i) = cell.Value
        cell.Value = ""
        i = i + 1
    Next cell
    Dim c As Range
    Dim j As Integer
    j = 0
    For Each c In rng
        Dim c2 As Integer
        Dim r2 As Integer
        c2 = c.column + columnShift
        r2 = c.row + rowShift
        If c2 > 0 And r2 > 0 Then
            If Worksheets(sheetname).Range(Cells(r2, c2), Cells(r2, c2)).AllowEdit = True Then
                Worksheets(sheetname).Range(Cells(r2, c2), Cells(r2, c2)).Value = valuesInRng(j)
            End If
        End If
        j = j + 1
    Next c
End Sub

Function getSubTableWithinRange(ByVal rangeAsString As String, ByVal sheetname As String)
    Dim firstCellAddress As String
    Dim lastCellAddress As String
    With Worksheets(sheetname)
        Dim r As Range
        Set r = .Range(rangeAsString)
        Dim c As Range
        Dim isFirstSet As Boolean
        isFirstSet = False
        For Each c In r
            If c.Value <> vbNullString And isFirstSet = False Then
                firstCellAddress = c.address
                isFirstSet = True
            Else
                If c.Value <> vbNullString And isFirstSet = True Then
                    lastCellAddress = c.address
                End If
            End If
        Next c
    End With
    Dim result As String
    result = firstCellAddress & ":" & lastCellAddress
    If result = ":" Then
        result = ""
    Else
        If firstCellAddress <> vbNullString And lastCellAddress = vbNullString Then
            result = firstCellAddress & ":" & firstCellAddress
        End If
    End If
    getSubTableWithinRange = result
End Function
Function willDataBeLost(ByRef tableRangesAsString() As String, ByVal sheetname As String, ByVal columnShift As Integer)
    Dim result As Boolean
    result = False
    Dim str_range As Variant
    If columnShift > 0 Then
        For Each str_range In tableRangesAsString
            If str_range <> vbNullString Then
                If getNumberOfEmptyCellsAtTheEndOfRow(str_range, sheetname) < columnShift Then
                    result = True
                    Exit For
                End If
            End If
        Next str_range
    Else
        If columnShift < 0 Then
            Dim abs_columnShift As Integer
            abs_columnShift = columnShift * -1
            For Each str_range In tableRangesAsString
                If str_range <> vbNullString Then
                    If getNumberOfEmptyCellsAtTheStartOfRow(str_range, sheetname) < abs_columnShift Then
                        result = True
                        Exit For
                    End If
                End If
            Next str_range
        End If
    End If
    willDataBeLost = result
End Function
