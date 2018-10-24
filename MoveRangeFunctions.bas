Attribute VB_Name = "MoveRangeFunctions"
'Sonya - Move Ranges

Sub moveRanges_KeepWithinTableRanges(ByVal sheetname As String, ByVal columnShift As Integer, ByRef tableRangesAsString() As String, Optional ByVal showConfirm = True)
  With Worksheets(sheetname)
    For Each str_tblRange In tableRangesAsString
        Dim table As Range
        Set table = .Range(str_tblRange)
        Dim tableAddressSplitByColon() As String
        tableAddressSplitByColon = Split(table.address, ":")
        Dim lastCellAddress As String
        lastCellAddress = tableAddressSplitByColon(1)
        Dim firstCellAddress As String
        firstCellAddress = tableAddressSplitByColon(0)
        Dim lastCellAddres_split() As String
        lastCellAddres_split = Split(lastCellAddress, "$")
        Dim firstCellAddress_split() As String
        firstCellAddress_split = Split(firstCellAddress, "$")
        
        Dim lastColIndex As Integer
        lastColIndex = getColNum(lastCellAddres_split(1))
        Dim lastRowIndex As Integer
        lastRowIndex = CInt(lastCellAddres_split(2))
        Dim firstColIndex As Integer
        firstColIndex = getColNum(firstCellAddress_split(1))
        Dim firstRowIndex As Integer
        firstRowIndex = CInt(firstCellAddress_split(2))
    
        Dim subTable As Range
        Dim lastNonEmptyCell As Range
    
        Set lastNonEmptyCell = .Range(lastNonEmptyCellAddressInTableRange(sheetname, str_tblRange))
        Set subTable = .Range(firstCellAddress & ":" & lastNonEmptyCell.address)
        Dim lastColIndexSub As Integer
        lastColIndexSub = getColNum(Split(Split(subTable.address, ":")(1), "$")(1))
    
        Dim allowedShift As Integer
        allowedShift = lastColIndex - lastNonEmptyCell.column
    
        If columnShift > allowedShift Then
            Dim colIndex As Integer
            colIndex = lastColIndexSub - (columnShift - allowedShift)
            Set subTable = .Range(Cells(firstRowIndex, firstColIndex), Cells(lastRowIndex, colIndex))
            If showConfirm = False Then
                    Call emptyColumnAfterThisColumn(colIndex, table)
                    Call moveRange(sheetname, subTable, columnShift, 0)
            Else
                If MsgBox("Warning: You may lose data. Would you like to continue?", vbYesNo) = vbYes Then
                    Call emptyColumnAfterThisColumn(colIndex, table)
                    Call moveRange(sheetname, subTable, columnShift, 0)
                    showConfirm = False
                End If
            End If
        Else
            Call moveRange(sheetname, subTable, columnShift, 0)
        End If
    Next str_tblRange
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

