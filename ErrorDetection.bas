Attribute VB_Name = "ErrorDetection"
'Sonya - Error Detection & Suggestion Feature

Function getTyposAndSuggestions(ByVal userInputs As Range, ByVal options As Range, ByVal issue As String, ByVal sheetname As String, Optional ByVal nameColumn = "n/a")
    Dim anInput As String
    Dim anOption As String
    Dim i As Integer
    Dim j As Integer
    Dim iLimit As Integer
    Dim jLimit As Integer
    Dim results As String
    
    iLimit = userInputs.Count
    jLimit = options.Count
    With Worksheets(userInputs.Worksheet.name)
    For i = 1 To iLimit
        Dim selectedItem As Range
        Set selectedItem = userInputs(i)
        anInput = selectedItem.Value
        For j = 1 To jLimit
            anOption = options(j).Value
            'if a mtach is found, the input is valid, move onto the next input
            If anInput = anOption Then
                j = jLimit + 1
            Else
                'if input has been compared to all options, add it to the validation sheet
                If j = options.Count And anInput <> Empty Then
                        results = getSuggestedValues(anInput, options)
                        Dim issue2 As String
                        issue2 = issue
                        If nameColumn <> "n/a" Then
                            issue2 = issue2 & " (" & .Range(nameColumn & selectedItem.row) & ")"
                        End If
                        Call updateValidationSheet(issue2, sheetname, selectedItem.address, anInput, results)
                        results = vbNullString
                End If
            End If
        Next j
    Next i
    End With
End Function
Function getTyposAndSuggestions_multiSelect(ByVal userInputs As Range, ByVal options As Range, ByVal issue As String, ByVal sheetname As String, Optional ByVal nameColumn = "n/a", Optional ByVal delimeter = ";")
    Dim anInput As String
    Dim anOption As String
    Dim i As Integer
    Dim j As Integer
    Dim iLimit As Integer
    Dim jLimit As Integer
    Dim results As String
    
    iLimit = userInputs.Count
    jLimit = options.Count
    With Worksheets(userInputs.Worksheet.name)
    For i = 1 To iLimit
        Dim selectedItem As Range
        Set selectedItem = userInputs(i)
        anInput = selectedItem.Value
        Dim inputArray() As String
        inputArray = Split(anInput, delimeter)
        Dim str As Variant
        Dim position As Integer
        position = 0
        For Each str In inputArray
            position = position + 1
            For j = 1 To jLimit
           anOption = options(j).Value
            'if a mtach is found, the input is valid, move onto the next input
            If str = anOption Then
                j = jLimit + 1
            Else
                'if input has been compared to all options, add it to the validation sheet
                If j = options.Count And anInput <> Empty Then
                        results = getSuggestedValues(str, options)
                        Dim issue2 As String
                        issue2 = issue
                        If nameColumn <> "n/a" Then
                            issue2 = issue2 & " - Entry at position " & position & " (" & .Range(nameColumn & selectedItem.row) & ")"
                        End If
                        Call updateValidationSheet(issue2, sheetname, selectedItem.address, anInput, results)
                        results = vbNullString
                End If
            End If
        Next j
        Next str
    Next i
    End With
End Function
Function getSuggestedValues(ByVal key As String, ByVal rangeOfData As Range) As String
    Dim isAdvancedSearchOn As Boolean
    isAdvancedSearchOn = True
    Dim result As String
    result = vbNullString
    Dim cell As Range
    'prevent case-sensitivity
    key = UCase(key)
    'key length
    Dim keyLen As Integer
    keyLen = Len(key)
    'get abbreviated key
    Dim abbr As String
    abbr = abbreviate(key)
    Dim abbrLen As Integer
    abbrLen = Len(abbr)
    For Each cell In rangeOfData
        'prevent case-sesitivity
        word = UCase(cell.Value)
        Dim abbr_cell As String
        abbr_cell = abbreviate(word)
        If abbr = word Or abbr_cell = key Then
            result = result & word & ", "
        Else
            'check if the key contains the word or vice versa; check if smaller string is contained in the bigger string as the reverse is impossible
            Dim big As String
            big = word
            Dim small As String
            small = key
            If keyLen > Len(word) Then
                big = key
                small = word
            End If
            If InStr(big, small) > 0 Then
                result = word & ", "
                isAdvancedSearchOn = False
            Else
                If isAdvancedSearchOn = True Then
                    'check if the key is an anagram of the word
                    If is_Anagram(key, word, False) <> False Then
                        result = result & word & ", "
                    Else
                        'check if abbreviated key is an anagram of abbreviated word or whether abbreviated key is an anagram of the word
                        If Len(abbr_cell) > 1 Then
                            If is_Anagram(abbr, abbr_cell, False) <> False Then
                                result = result & word & ", "
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next cell
    getSuggestedValues = result
End Function
Function abbreviate(ByVal key As String) As String
    Dim abbr As String
    abbr = vbNullString
    If InStr(key, " ") = False Then
        abbr = Mid(key, 1, 1)
    Else
        'split the keyword by space
        Dim keySplit() As String
        keySplit = Split(UCase(key), " ")
        Dim i As Integer
        Dim lenKeySplit As Integer
        lenKeySplit = UBound(keySplit)
        'iterate through each item
        For i = 0 To lenKeySplit
            'concatenate the first letter of each item
            abbr = abbr & Mid(keySplit(i), 1, 1)
        Next i
    End If
    abbreviate = abbr
End Function
Function is_Anagram(ByVal key As String, ByVal word As String, Optional ByVal exactMatch As Boolean = True) As Boolean
    Dim i As Integer
    Dim noOfRemainingChars As Integer
    Dim big As String
    Dim small As String
    Dim lenSmall As Integer
    Dim lenBig As Integer
    Dim lenKey As Integer
    Dim lenWord As Integer
    is_Anagram = False
    
    lenKey = Len(key)
    lenWord = Len(word)

    'find out which has the least number of characters
    If lenKey > lenWord Then
        big = key
        small = word
        lenSmall = lenWord
        lenBig = lenKey
    Else
        big = word
        small = key
        lenSmall = lenKey
        lenBig = lenWord
    End If
    
    If lenBig = lenSmall Or lenBig = (lenSmall + 1) Then
        If exactMatch = True Then
            noOfRemainingChars = 0
        Else
            noOfRemainingChars = 1
        End If
        'reducs number of loops by iterating through the shorter value
        For i = 1 To lenSmall
            Dim letter As String
            letter = Mid(small, i, 1)
            big = Replace(big, letter, "", Count:=1)
        Next i
    
        If Len(big) <= noOfRemainingChars Then
            is_Anagram = True
        Else
            is_Anagram = False
        End If
    End If
End Function
Function findExtremeValues(ByVal sheetname As String, ByVal rangeAsString As String, ByVal minimumVal As Integer, ByVal maxVal As Integer, Optional ByVal min_messageBeforeValue = "Extreme Value - Minimum value is ", Optional ByVal min_messageAfterValue = vbNullString, Optional ByVal max_messageBeforeValue = "Extreme Value - Maximum value is ", Optional ByVal max_messageAfterValue = vbNullString, Optional ByVal nameColumn = "n/a") As Integer
    Dim r As Range
    Dim cell As Range
    With Worksheets(sheetname)
        Set r = .Range(rangeAsString)
        Dim counter As Integer
        counter = 0
        For Each cell In r
            Dim name As String
            If cell.Value <> vbNullString Then
                If cell.Value < minimumVal Then
                    If nameColumn <> "n/a" Then
                        name = " (" & .Range(nameColumn & cell.row) & ")"
                    End If
                    Call updateValidationSheet(min_messageBeforeValue & minimumVal & min_messageAfterValue & " " & name, sheetname, cell.address, cell.Value, "")
                    counter = counter + 1
                Else
                    If cell.Value > maxVal Then
                        If nameColumn <> "n/a" Then
                            name = " (" & .Range(nameColumn & cell.row) & ")"
                        End If
                        Call updateValidationSheet(max_messageBeforeValue & maxVal & max_messageAfterValue & " " & name, sheetname, cell.address, cell.Value, "")
                        counter = counter + 1
                    End If
                End If
            End If
        Next cell
    End With
    findExtremeValues = counter
End Function
Function findExtremeValues_EntryDifferentToLookUp(ByVal lookupSheetname As String, ByVal lookupRange As String, ByVal minimumVal As Integer, ByVal maxVal As Integer, ByVal dataEntry_sheetname As String, ByVal dataEntry_range As String, Optional ByVal min_messageBeforeValue = "Extreme Value - Minimum value is ", Optional ByVal min_messageAfterValue = vbNullString, Optional ByVal max_messageBeforeValue = "Extreme Value - Maximum value is ", Optional ByVal max_messageAfterValue = vbNullString, Optional ByVal nameColumn = "n/a") As Integer
    Dim r As Range
    Dim dataEntryRange As Range
    Dim cell As Range
    With Worksheets(lookupSheetname)
        Set r = .Range(lookupRange)
        Set dataEntryRange = Worksheets(dataEntry_sheetname).Range(dataEntry_range)
        Dim counter As Integer
        Dim entryCell As Range
        counter = 0
        Dim cellCounter As Integer
        cellCounter = 1
        For Each cell In r
            Dim name As String
            If cell.Value <> vbNullString Then
                If cell.Value < minimumVal Then
                    Set entryCell = dataEntryRange.Cells(cellCounter)
                    If nameColumn <> "n/a" Then
                        name = " (" & .Range(nameColumn & cell.row) & ")"
                    End If
                    Call updateValidationSheet(min_messageBeforeValue & minimumVal & min_messageAfterValue & " " & name, lookupSheetname, entryCell.address, entryCell.Value, "")
                    counter = counter + 1
                Else
                    If cell.Value > maxVal Then
                        Set entryCell = dataEntryRange.Cells(cellCounter)
                        If nameColumn <> "n/a" Then
                            name = " (" & .Range(nameColumn & cell.row) & ")"
                        End If
                        Call updateValidationSheet(max_messageBeforeValue & maxVal & max_messageAfterValue & " " & name, lookupSheetname, entryCell.address, entryCell.Value, "")
                        counter = counter + 1
                    End If
                End If
            End If
            cellCounter = cellCounter + 1
        Next cell
    End With
    findExtremeValues_EntryDifferentToLookUp = counter
End Function
Function mandatoryChecksForCorrespondingCells(ByRef ranges() As Range, ByRef tblNames() As String, Optional ByVal nameColumn = "n/a")
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim rangCount As Integer
    rangCount = (UBound(ranges) - LBound(ranges)) - 1
    
    For i = 0 To rangCount
        Dim table As Range
        Set table = ranges(i)
        Dim cellCount As Integer
        cellCount = table.Cells.Count
        For j = 1 To cellCount
            Dim c As Range
            Set c = table.Cells(j)
            If c = vbNullString Then
                    For k = 0 To rangCount
                        If k <> i Then
                            Dim table2 As Range
                            Set table2 = ranges(k)
                            Dim cell As Range
                            Set cell = table2.Cells(j)
                            If cell <> vbNullString Then
                                Dim name As String
                                If nameColumn <> "n/a" Then
                                    name = "(" & Worksheets(c.Worksheet.name).Range(nameColumn & c.row) & ")"
                                End If
                                Call updateValidationSheet("Missing Value in Correpsonding Cell - " & tblNames(i) & " " & name, table.Worksheet.name, c.address, c.Value, "", True)
                            End If
                        End If
                    Next k
            End If
        Next j
    Next i
End Function
Function mandatoryChecksForCorrespondingCells_MsgBox(ByRef arr_ranges() As Range, ByRef tblNames() As String, ByVal sheetname As String, Optional ByVal colLetter_names As String = "n/a", Optional ByVal rowNum_headers As String = "n/a", Optional ByVal displayMsgBox = True)
    Dim i As Integer
    Dim j As Integer
    Dim table As Range
    Dim lboundArrRanges As Integer
    lboundArrRanges = LBound(arr_ranges)
    Dim uboundArrRanges As Integer
    uboundArrRanges = UBound(arr_ranges) - 1
    Dim result As String
    Dim messageCount As Integer
    messageCount = 0
    
    'loop though ranges
    For i = lboundArrRanges To uboundArrRanges
        'get table at position i
        Set table = arr_ranges(i)
        Dim col As Range
        'track the column that is being looking at
        Dim column As Integer
        column = 0
        result = vbNullString
        'loop though columns in table
        For Each col In table.columns
            'increment column by 1
            column = column + 1
            Dim cell As Range
            'keep track of the row that is being looked at
            Dim row As Integer
            row = 0
            'loo through cells in column
            For Each cell In col.Cells
                'increment row by 1
                row = row + 1
                'check if the cell is blank
                If cell = vbNullString Then
                    Dim table2 As Range
                    'loop through ranges
                    For j = lboundArrRanges To uboundArrRanges
                        'get table at position j
                        Set table2 = arr_ranges(j)
                        'prevent a table being compared to itself
                        If table.address <> table2.address Then
                            Dim cell2 As Range
                            'get cell at corresponding position
                            Set cell2 = table2.Cells(row, column)
                            'check if the corresponding cell is not blank
                            If cell2 <> vbNullString Then
                                Dim label1 As String
                                Dim label2 As String
                                With Worksheets(sheetname)
                                    If colLetter_names <> "n/a" Then
                                        label1 = .Range(colLetter_names & cell.row)
                                    End If
                                    If rowNum_headers <> "n/a" Then
                                        label2 = .Range(getColumnAsLetter(cell.address) & rowNum_headers)
                                    End If
                                End With
                                'add result to array
                                If result = vbNullString Then
                                    result = "Missing information for " & tblNames(i) & " for " & ":"
                                End If
                                result = result & vbNewLine & label1 & " for " & label2
                                'once a non-empty corresponding cell is found, exit the loop to move onto the next row to prevent duplicate results
                                j = uboundArrRanges
                            End If
                        End If
                    Next j
                End If
            Next cell
        Next col
        If result <> vbNullString Then
            If displayMsgBox = True Then
                Call MsgBox(result)
            End If
            messageCount = messageCount + 1
        End If
    Next i
    mandatoryChecksForCorrespondingCells_MsgBox = messageCount
End Function
