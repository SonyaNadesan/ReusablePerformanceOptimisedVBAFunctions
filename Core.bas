Attribute VB_Name = "Core"
'Sonya - General Functions
Function getSubTableWithinRange(ByVal rangeAsString As String, ByVal sheetname As String)
    Dim firstCellAddress As String
    Dim lastCellAddress As String
    With Worksheets(sheetname)
        Dim r As range
        Set r = .range(rangeAsString)
        Dim c As range
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
Function emptyLastXcolumnsInRow(ByVal numberOfColumnsToEmpty As Integer, ByVal rangeAsString As String, ByVal sheetname As String)
    With Worksheets(sheetname)
        Dim r As range
        Dim c As range
        Dim i As Integer
        Dim counter As Integer
        counter = 0
        Set r = .range(rangeAsString)
        For i = r.Cells.Count To 1 Step -1
            If counter < numberOfColumnsToEmpty Then
                r.Cells(i).Value = vbNullString
                counter = counter + 1
            Else
                Exit For
            End If
        Next i
    End With
End Function
Function getWholeCol(ByVal first As Integer, ByVal sheetname As String, ByVal column As String) As range
    Dim i As Integer
    i = first
    Dim max As Integer
    With Worksheets(sheetname)
        Do Until i < 0
            If .range(column & i) = vbNullString Then
                max = i - 1
                i = -1
            Else
                i = i + 1
            End If
        Loop
        Set getWholeCol = .range(column & first & ":" & column & max)
    End With
End Function
Function getNumberOfEmptyCellsAtTheEndOfRow(ByVal rngAsString As String, ByVal sheetname As String) As Integer
    With Worksheets(sheetname)
        Dim cellCount As Integer
        Dim rng As range
        Set rng = .range(rngAsString)
        cellCount = rng.Cells.Count
        Dim i As Integer
        Dim counter As Integer
        counter = 0
        For i = cellCount To 1 Step -1
            If rng.Cells(i) = vbNullString Then
                counter = counter + 1
            Else
                i = 0
            End If
        Next i
    End With
    getNumberOfEmptyCellsAtTheEndOfRow = counter
End Function
Function mergeCsvWithoutRepetition(ByVal csv1 As String, ByVal csv2) As String
    Dim result As String
    If csv1 = vbNullString Or csv2 = vbNullString Then
        result = csv1 & csv2
    Else
        Dim csvSplit1() As String
        csvSplit1 = Split(csv1, ",")
        Dim csvSplit2() As String
        csvSplit2 = Split(csv2, ",")
        Dim item As Variant
    
        'get smallest csv to minimise additions to the other csv
        Dim csv1Count As Integer
        Dim csv2Count As Integer
        csv1Count = UBound(csvSplit1) - LBound(csvSplit1) + 1
        csv2Count = UBound(csvSplit2) - LBound(csvSplit2) + 1
        If csv1Count > 0 And csv1Count < csv2Count Then
            'loop through smallest, non-empty csv
            For Each item In csvSplit1
                csv2 = addValueToCsvIfAbsent(csv2, item)
            Next item
            result = csv2
        Else
            'loop through smallest, non-empty csv
            For Each item In csvSplit2
                csv1 = addValueToCsvIfAbsent(csv1, item)
            Next item
            result = csv1
        End If
    End If
    mergeCsvWithoutRepetition = Replace(result, ",,", ",")
End Function
Function addValueToCsvIfAbsent(ByVal csv_indicateLastValueWithoutCommaAtTheEnd As String, ByVal val As String) As String
    Dim result As String
    Dim arr() As String
    If csv_indicateLastValueWithoutCommaAtTheEnd = vbNullString Then
        result = val
    Else
        arr = Split(csv_indicateLastValueWithoutCommaAtTheEnd, ",")
        Dim s As Variant
        result = vbNullString
        For Each s In arr
            If result = vbNullString Then
                result = s
            Else
                If isDataInCSV(result, s) = False Then
                    result = result & "," & s
                End If
            End If
        Next s

        If result = vbNullString Then
            result = val
        Else
            If isDataInCSV(result, val) = False Then
                result = result & "," & val
            End If
        End If
    End If
    addValueToCsvIfAbsent = result
End Function
Function isDataInCSV(ByVal csv_indicateLastValueWithoutCommaAtTheEnd As String, ByVal data As String) As Boolean
    Dim result As Boolean
    If csv_indicateLastValueWithoutCommaAtTheEnd = vbNullString Then
        result = False
    Else
        Dim csvSplit() As String
        csvSplit = Split(UCase(csv_indicateLastValueWithoutCommaAtTheEnd), ",")
        Dim v As Variant
        result = False
        data = UCase(data)
        For Each v In csvSplit
            If v = data Then
                result = True
                Exit For
            End If
        Next v
    End If
    isDataInCSV = result
End Function
Sub emptyThisRange(ByVal rngAsString As String, ByVal sheetname As String)
    Worksheets(sheetname).range(rngAsString).ClearContents
End Sub
Function getColumnAsLetter(ByVal cellAddress As String) As String
    Dim addressSplit() As String
    If cellAddress <> vbNullString Then
        addressSplit = Split(cellAddress, "$")
        getColumnAsLetter = addressSplit(1)
    Else
        getColumnAsLetter = ""
    End If
End Function
Function lastRowNumOfNonEmptyCellInCol(ByVal first As Integer, ByVal sheetname As String, ByVal column As String) As Integer
    Dim i As Integer
    i = first
    Dim max As Integer
    With Worksheets(sheetname)
        Do Until i < 0
            If .range(column & i) = vbNullString Then
                max = i - 1
                If max < first Then
                    max = first
                End If
                i = -1
            Else
                i = i + 1
            End If
        Loop
    End With
    lastRowNumOfNonEmptyCellInCol = max
End Function
Function firstNonEmptyCell(ByVal sheetname As String, ByVal rangeAsString As String) As range
    Dim cell As range
    Dim r As range
    Set r = Worksheets(sheetname).range(rangeAsString)
    Dim i As Integer
    i = 1
    Dim result As range
    For Each cell In r
        If cell.Value <> vbNullString And result Is Nothing Then
            Set result = cell
            Exit For
        End If
    Next cell
    Set firstNonEmptyCell = result
End Function
Function getColNum(ByVal letter As String) As Integer
    Dim r As range
    Set r = range(letter & "1")
    getColNum = r.column
End Function
Function lastNonEmptyCellAddressInTableRange(ByVal sheetname As String, ByVal tableRangeAsString As String) As String
    Dim table As range
    Set table = Worksheets(sheetname).range(tableRangeAsString)
    Dim cell As range
    Dim result As range
    For Each cell In table
        If cell.Value <> vbNullString Then
            Set result = cell
        End If
    Next cell
    lastNonEmptyCellAddressInTableRange = result.address
End Function
Sub emptyColumnAfterThisColumn(ByVal columnNum As Integer, ByVal r As range)
    Dim c As range
    For Each c In r
        If c.column > columnNum Then
            c.Value = vbNullString
        End If
    Next c
End Sub
Function arraySize(ByRef arr() As String) As Integer
    arraySize = UBound(arr) - LBound(arr) + 1
End Function
Function intArraySize(ByRef arr() As Integer) As Integer
    intArraySize = UBound(arr) - LBound(arr) + 1
End Function
Function lastNonEmptyCellAddressInRow(ByVal rangeAsString As String, ByVal sheetname As String) As String
  With Worksheets(sheetname)
    .Activate
    Dim firstCell As range
    Set firstCell = .range(rangeAsString)
    Dim firstColNum As Integer
    firstColNum = firstCell.column
    Dim lastCellAddress As String
    lastCellAddress = firstCell.address
    If InStr(firstCell.address, ":") Then
        lastCellAddress = Split(firstCell.address, ":")(0)
        Dim r As range
        Set r = .range(rangeAsString)
        Dim cell As range
        For Each cell In r
            If cell.Value <> vbNullString Then
                lastCellAddress = cell.address
            End If
        Next cell
    Else
        Dim i As Integer
        i = 1
        Do Until i < 0
            Dim c As range
            Set c = .range(Cells(firstCell.row, firstColNum + i), Cells(firstCell.row, firstColNum + i))
            If c.Value <> vbNullString Then
                lastCellAddress = c.address
                i = i + 1
            Else
                i = -1
            End If
        Loop
    End If
    End With
    lastNonEmptyCellAddressInRow = lastCellAddress
End Function
Function doesRowExistInRange(ByVal tbl As range, ByVal rowAsRange As range) As Boolean
    Dim rowInTbl As range
    Dim finalResult As Boolean
    finalResult = False
    Dim lastRowNum As Integer
    Dim colCount As Integer
    colCount = tbl.columns.Count
    For Each rowInTbl In tbl.Rows
        Dim counter As Integer
        For counter = 1 To colCount
            If rowInTbl.Cells(counter) = rowAsRange.Cells(counter) Then
                If counter = colCount Then
                    finalResult = True
                End If
            Else
                counter = colCount
            End If
        Next counter
    Next rowInTbl
    doesRowExistInRange = finalResult
End Function
Function doesRowExistInRange_whereRowISsProvidedAsACSV(ByVal tbl As range, ByVal rowAsCSV_indicateLastValueWithoutComma As String) As Boolean
    Dim rowInTbl As range
    Dim finalResult As Boolean
    finalResult = False
    Dim lastRowNum As Integer
    Dim csvSplit() As String
    csv = Split(rowAsCSV_indicateLastValueWithoutComma, ",")
    Dim colCount As Integer
    colCount = tbl.columns.Count
    For Each rowInTbl In tbl.Rows
        Dim counter As Integer
        For counter = 1 To colCount
            If rowInTbl.Cells(counter) = csv(counter - 1) Then
                If counter = colCount Then
                    finalResult = True
                End If
            Else
                counter = colCount
            End If
        Next counter
        If finalResult = True Then
            Exit For
        End If
    Next rowInTbl
    doesRowExistInRange_whereRowISsProvidedAsACSV = finalResult
End Function
Function rowToCSV(ByVal rowRange As range) As String
    Dim result As String
    result = vbNullString
    Dim c As range
    For Each c In rowRange.Cells
        If result <> vbNullString Then
            result = result & "," & c.Value
        Else
            result = c.Value
        End If
    Next c
    rowToCSV = result
End Function
