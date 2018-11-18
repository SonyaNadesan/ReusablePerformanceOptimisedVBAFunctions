Attribute VB_Name = "MoveRangeFunctions"
'Sonya - Move Ranges

Sub moveRanges_KeepWithinTableRanges(ByVal sheetname As String, ByVal columnShift As Integer, ByRef tableRangesAsString() As String, Optional ByVal showConfirm = True)
  With Worksheets(sheetname)
    Dim warnUser As Boolean
    Dim allowShiftData As Boolean
    Dim subAddress As String
    Dim subTable As Range
    warnUser = False
    allowShiftData = False
    Dim allowedShift As Integer
    
    If showConfirm = True Then
        For Each str_range In tableRangesAsString
            If str_range <> vbNullString Then
                If getNumberOfEmptyCellsAtTheEndOfRow(str_range, sheetname) < columnShift Then
                    warnUser = True
                    Exit For
                End If
            End If
        Next str_range
    End If
    
    
    If warnUser = False Then
            Dim aRangeAsStr As Variant
            For Each aRangeAsStr In tableRangesAsString
                If aRangeAsStr <> vbNullString Then
                    allowedShift = getNumberOfEmptyCellsAtTheEndOfRow(aRangeAsStr, sheetname)
                    subAddress = getSubTableWithinRange(aRangeAsStr, sheetname)
                    If subAddress <> vbNullString Then
                        Set subTable = .Range(subAddress)
                        If allowedShift < columnShift Then
                            Call emptyLastXcolumnsInRow(columnShift - allowedShift, subAddress, sheetname)
                        End If
                        subAddress = getSubTableWithinRange(aRangeAsStr, sheetname)
                        Call moveRange(sheetname, subTable, columnShift, 0)
                    End If
                End If
            Next aRangeAsStr
    Else
        If showConfirm = True Then
                If MsgBox("Warning: You may lose data. Would you like to proceed?", vbExclamation + vbYesNo) = vbYes Then
                    allowShiftData = True
                End If
        Else
            allowShiftData = True
        End If
        If allowShiftData = True Then
            Dim aRangeAsStr2 As Variant
            For Each aRangeAsStr2 In tableRangesAsString
                If aRangeAsStr2 <> vbNullString Then
                    allowedShift = getNumberOfEmptyCellsAtTheEndOfRow(aRangeAsStr2, sheetname)
                    subAddress = getSubTableWithinRange(aRangeAsStr2, sheetname)
                    If subAddress <> vbNullString Then
                        Set subTable = .Range(subAddress)
                        If allowedShift < columnShift Then
                            Call emptyLastXcolumnsInRow(columnShift - allowedShift, subAddress, sheetname)
                        End If
                        subAddress = getSubTableWithinRange(aRangeAsStr2, sheetname)
                        If subAddress <> vbNullString Then
                            Set subTable = .Range(subAddress)
                            Call moveRange(sheetname, subTable, columnShift, 0)
                        End If
                    End If
                End If
            Next aRangeAsStr2
        End If
    End If
  End With
End Sub
Sub moveRange(ByVal sheetname As String, ByRef rng As Range, ByVal columnShift As Integer, ByVal rowShift As Integer)
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
        If Worksheets(sheetname).Range(Cells(r2, c2), Cells(r2, c2)).AllowEdit = True Then
            Worksheets(sheetname).Range(Cells(r2, c2), Cells(r2, c2)).Value = valuesInRng(j)
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
            If c <> vbNullString And isFirstSet = False Then
                firstCellAddress = c.address
                isFirstSet = True
            Else
                If c <> vbNullString And isFirstSet = True Then
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
