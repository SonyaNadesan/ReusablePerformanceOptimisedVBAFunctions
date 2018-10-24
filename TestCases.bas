Attribute VB_Name = "TestCases"
Sub runAllTests()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Call clearOutput
    Dim rowNum As Integer
    rowNum = 5
    Dim result As String
    
    With Worksheets("Version")
        
        'is_anagram
        result = is_Anagram("NADII", "INDIA", True)
        Call addOutput(rowNum, "is_anagram() [exact, exactMatch]", result, "True")
        rowNum = rowNum + 1
        result = is_Anagram("INDA", "INDIA", True)
        Call addOutput(rowNum, "is_anagram()  [missing char, exactMatch]", result, "False")
        rowNum = rowNum + 1
        result = is_Anagram("INDISA", "INDIA", True)
        Call addOutput(rowNum, "is_anagram()  [additional char, exactMatch]", result, "False")
        rowNum = rowNum + 1
        result = is_Anagram("INDIA", "INDIA", False)
        Call addOutput(rowNum, "is_anagram() [exact, exactMatch=False]", result, "True")
        rowNum = rowNum + 1
        result = is_Anagram("INDA", "INDIA", False)
        Call addOutput(rowNum, "is_anagram() [missing char, exactMatch=False]", result, "True")
        rowNum = rowNum + 1
        result = is_Anagram("INDISA", "INDIA", False)
        Call addOutput(rowNum, "is_anagram() [additional char, exactMatch=False]", result, "True")
        rowNum = rowNum + 1
        
        'invalidInput()
        Call ClearIssues
        result = invalidInput(.Range("F13:F21"), .Range("I15:I21"), "", "Version")
        result = Worksheets("Validation").Range("B1")
        Call addOutput(rowNum, "invalidInput()", result, "6")
        Call ClearIssues
        rowNum = rowNum + 1
        
        'suggestions()
        Dim r As Range
        Set r = .Range("I15:I21")
        result = suggestions("inda", r)
        Call addOutput(rowNum, "suggestions() [missing char]", result, "India", False)
        rowNum = rowNum + 1
        result = suggestions("indisa", r)
        Call addOutput(rowNum, "suggestions() [extra char]", result, "India", False)
        rowNum = rowNum + 1
        result = suggestions("UK", r)
        Call addOutput(rowNum, "suggestions() [abbreviation]", result, "United Kingdom", False)
        rowNum = rowNum + 1
        result = suggestions("US", r)
        Call addOutput(rowNum, "suggestions() [abbreviation, missing char]", result, "USA", False)
        rowNum = rowNum + 1
        result = suggestions("UKS", r)
        Call addOutput(rowNum, "suggestions() [abbreviation, extra char - UKS]", result, "United Kingdom", False)
        rowNum = rowNum + 1
        result = suggestions("USK", r)
        Call addOutput(rowNum, "suggestions() [abbreviation, extra char - USK]", result, "United Kingdom", False)
        rowNum = rowNum + 1
        result = suggestions("badui", r)
        Call addOutput(rowNum, "suggestions() [anagram]", result, "Dubai", False)
        rowNum = rowNum + 1
        result = suggestions("badu", r)
        Call addOutput(rowNum, "suggestions() [anagram, missing char]", result, "Dubai", False)
        rowNum = rowNum + 1
        result = suggestions("baduis", r)
        Call addOutput(rowNum, "suggestions() [anagram, extra char]", result, "Dubai", False)
        rowNum = rowNum + 1
        result = suggestions("zxcv", r)
        Call addOutput(rowNum, "suggestions() [zxcv]", result, "", True)
        rowNum = rowNum + 1
        
        
        'abbreviation()
        result = abbreviation("United Kingdom")
        Call addOutput(rowNum, "abbreviation() [United Kingdom]", result, "UK")
        rowNum = rowNum + 1
        
        'getDayWithSuffix()
        Dim d As Date
        d = DateValue("November 29, 1995")
        result = getDayWithSuffix(d)
        Call addOutput(rowNum, "getDayWithSuffix() [November 29 1995]", result, "29th")
        rowNum = rowNum + 1
        d = DateValue("November 21, 2995")
        result = getDayWithSuffix(d)
        Call addOutput(rowNum, "getDayWithSuffix() [November 21 1995]", result, "21st")
        rowNum = rowNum + 1
        d = DateValue("November 22, 2995")
        result = getDayWithSuffix(d)
        Call addOutput(rowNum, "getDayWithSuffix() [November 22 1995]", result, "22nd")
        rowNum = rowNum + 1
        d = DateValue("November 11, 1995")
        result = getDayWithSuffix(d)
        Call addOutput(rowNum, "getDayWithSuffix() [November 11 1995]", result, "11th")
        rowNum = rowNum + 1
        d = DateValue("November 12, 1995")
        result = getDayWithSuffix(d)
        Call addOutput(rowNum, "getDayWithSuffix() [November 12 1995]", result, "12th")
        rowNum = rowNum + 1
        d = DateValue("November 13, 1995")
        result = getDayWithSuffix(d)
        Call addOutput(rowNum, "getDayWithSuffix() [November 13 1995]", result, "13th")
        rowNum = rowNum + 1
        
        'addZeroToHrOrMin()
        result = addZeroToHrOrMin(0)
        Call addOutput(rowNum, "addZeroToHrOrMin() [0]", result, "00")
        rowNum = rowNum + 1
        result = addZeroToHrOrMin(9)
        Call addOutput(rowNum, "addZeroToHrOrMin() [9]", result, "09")
        rowNum = rowNum + 1
        result = addZeroToHrOrMin(10)
        Call addOutput(rowNum, "addZeroToHrOrMin() [10]", result, "10")
        rowNum = rowNum + 1
        
        'moveRange()
        Call constructTestTables
        Dim rng As Range
        Set rng = .Range("F4:H7")
        Call moveRange("Version", rng, 1, 0)
        result = .Range("F8") & "," & .Range("G8") & "," & .Range("H8")
        Call addOutput(rowNum, "moveRange() [1 column, 0 rows]", result, "0,10,26")
        rowNum = rowNum + 1
        Call constructTestTables
        Dim rngArr(1) As String
        rngArr(0) = "F4:H7"
        rngArr(1) = "J4:L7"
        Call moveRanges_KeepWithinTableRanges("Version", 2, rngArr, False)
        result = .Range("F8") & "," & .Range("G8") & "," & .Range("H8")
        Call addOutput(rowNum, "moveRanges_KeepWithinTableRanges() [2 column, 0 rows]", result, "0,0,10")
        rowNum = rowNum + 1
        
        'findExtremeValues()
        Call constructTestTables
        result = findExtremeValues("Version", "F4:H7", 4, 7)
        Call addOutput(rowNum, "findExtremeValues() [1 to 8][min: 4, max: 7]", result, "4")
        Call ClearIssues
        rowNum = rowNum + 1
        
        'isDataInCSV()
        result = isDataInCSV("abc,def,ghi", "abc")
        Call addOutput(rowNum, "isDataInCSV() [first value]", result, "True")
        rowNum = rowNum + 1
        result = isDataInCSV("abc,def,ghi", "ghi")
        Call addOutput(rowNum, "isDataInCSV() [absent value]", result, "True")
        rowNum = rowNum + 1
        result = isDataInCSV("abc,def,ghi", "jkl")
        Call addOutput(rowNum, "isDataInCSV() [last value]", result, "False")
        rowNum = rowNum + 1
        
        'addValueToCsvIfAbsent
        result = addValueToCsvIfAbsent("Croydon,Sutton", "Kingston")
        Call addOutput(rowNum, "addValueToCsvIfAbsent() [add Kingston to Croydon,Sutton]", result, "Croydon,Sutton,Kingston")
        rowNum = rowNum + 1
        result = addValueToCsvIfAbsent("Croydon,Sutton,Kingston", "Kingston")
        Call addOutput(rowNum, "addValueToCsvIfAbsent() [reject Kingston]", result, "Croydon,Sutton,Kingston")
        rowNum = rowNum + 1
        
        'mergeCsvWithoutRepetition
        result = mergeCsvWithoutRepetition("a,b,c", "d,e,f")
        Call addOutput(rowNum, "mergeCsvWithoutRepetition() [add def to abc]", result, "a,b,c,d,e,f")
        rowNum = rowNum + 1
        result = mergeCsvWithoutRepetition("a,b,c", "a")
        Call addOutput(rowNum, "mergeCsvWithoutRepetition() [reject a,b]", result, "a,b,c")
        rowNum = rowNum + 1
        result = mergeCsvWithoutRepetition("a,b,c", "c")
        Call addOutput(rowNum, "mergeCsvWithoutRepetition() [reject c,b]", result, "a,b,c")
        rowNum = rowNum + 1
        
        'getWholeCol
        result = getWholeCol(4, "Version", "F").address
        Call addOutput(rowNum, "getWholeCol() [F4]", result, "$F$4:$F$8")
        rowNum = rowNum + 1
        
        'getColumnAsLetter()
        result = getColumnAsLetter("$B$5")
        Call addOutput(rowNum, "getColumnAsLetter() [$B$5]", result, "B")
        rowNum = rowNum + 1
        
        'getColNum()
        result = getColNum("J")
        Call addOutput(rowNum, "getColNum() [J]", result, "10")
        rowNum = rowNum + 1
        
        'arraySize()
        Dim arr() As String
        arr = Split("A,B,C", ",")
        result = arraySize(arr)
        Call addOutput(rowNum, "arraySize() [['A','B','C']]", result, "3")
        rowNum = rowNum + 1
        
        'intArraySize()
        Dim arr2(3) As Integer
        arr2(0) = 1
        arr2(1) = 2
        arr2(2) = 3
        arr2(3) = 4
        result = intArraySize(arr2)
        Call addOutput(rowNum, "intArraySize() [[1,2,3,4]]", result, "4")
        rowNum = rowNum + 1
        
        'emptyThisRange()
        Call constructTestTables
        Call emptyThisRange("F5:H7", "Version")
        result = .Range("F8") & "," & .Range("G8") & "," & .Range("H8")
        Call addOutput(rowNum, "emptyThisRange() [F5:H7]", result, "1,5,0")
        rowNum = rowNum + 1
        
        'lastRowNumOfNonEmptyCellInCol()
        Call constructTestTables
        result = lastRowNumOfNonEmptyCellInCol(4, "Version", "F")
        Call addOutput(rowNum, "emptyThisRange() [F4]", result, "8")
        rowNum = rowNum + 1
        
        'emptyColumnAfterThisColumn
        Call constructTestTables
        Call emptyColumnAfterThisColumn(6, .Range("F4:G7"))
        result = .Range("F8") & "," & .Range("G8") & "," & .Range("H8")
        Call addOutput(rowNum, "emptyThisRange() [F]", result, "10,0,0")
        rowNum = rowNum + 1
        
        'firstNonEmptyCell()
        Call constructTestTables
        result = firstNonEmptyCell("Version", "G1:G7").address
        Call addOutput(rowNum, "firstNonEmptyCell() [G1:G7]", result, "$G$4")
        rowNum = rowNum + 1
        
        'lastNonEmptyCellInTableRange()
        Call constructTestTables
        result = lastNonEmptyCellAddressInTableRange("Version", "F4:H7")
        Call addOutput(rowNum, "lastNonEmptyCellInTableRange() [F4:H7]", result, "$G$7")
        rowNum = rowNum + 1
        
        
        'lastNonEmptyCellAddressInRow()
        Call constructTestTables
        result = lastNonEmptyCellAddressInRow("F4:H4", "Version")
        Call addOutput(rowNum, "lastNonEmptyCellAddressInRow() [F4:H4]", result, "$G$4")
        rowNum = rowNum + 1
        
        'rowToCSV()
        Call constructTestTables
        result = rowToCSV(.Range("F4:H4"))
        Call addOutput(rowNum, "rowToCSV() [F4:H4]", result, "1,5,")
        rowNum = rowNum + 1
        
        'doesRowExistInRange()
        Call constructTestTables
        result = doesRowExistInRange(.Range("F4:H4"), .Range("F5:H5"))
        Call addOutput(rowNum, "doesRowExistInRange() [F4:H4, F5:H5]", result, "False")
        rowNum = rowNum + 1
        .Range("F5") = .Range("F4")
        .Range("G5") = .Range("G4")
        .Range("H5") = .Range("H4")
        result = doesRowExistInRange(.Range("F4:H4"), .Range("F5:H5"))
        Call addOutput(rowNum, "doesRowExistInRange() [F4:H4, F5:H5]", result, "True")
        
        'getEarliestDate
        result = getEarliestDate("Version", "N10:N14")
        Call addOutput(rowNum, "getEarliestDate() [N10:N14]", result, "09/06/1966")
        rowNum = rowNum + 1
        
        'getLatestDate
        result = getLatestDate("Version", "N10:N14")
        Call addOutput(rowNum, "getLatestDate() [N10:N14]", result, "27/07/2017")
        rowNum = rowNum + 1
        
        'quickValidateRequiredFieldsInCorrespondingTables
        Call ClearIssues
        Call constructTestTables
        .Range("F4") = ""
        .Range("J5") = ""
        Dim rangesArr(2) As Range
        Set rangesArr(0) = .Range("F4:H7")
        Set rangesArr(1) = .Range("J4:L7")
        Dim tblNames(2) As String
        tblNames(0) = "yellow"
        tblNames(1) = "green"
        
        Call quickValidateRequiredFieldsInCorrespondingTables(rangesArr, tblNames)
        Call addOutput(rowNum, "quickValidateRequiredFieldsInCorrespondingTables() [2 missing values]", "2", Worksheets("Validation").Range("B1"))
        rowNum = rowNum + 1
        Call ClearIssues
        
        .Range("F6") = ""
        Dim numberOfErrors As Integer
        numberOfErrors = quickValidateRequiredFieldsInCorrespondingTables_MsgBox(rangesArr, tblNames, "Version", displayMsgBox:=False)
        Call addOutput(rowNum, "quickValidateRequiredFieldsInCorrespondingTables_MsgBox() [2 tables with missing value(s)]", "2", numberOfErrors)
        rowNum = rowNum + 1
        Call ClearIssues
        
        Call constructTestTables
        .Range("J6") = 30
        Call findExtremeValues_where_dataEntryRangeIsDifferentToInputRange("Version", "J4:K7", 0, 9, "Version", "F4:G7")
        result = Worksheets("Validation").Range("C3")
        Call addOutput(rowNum, "findExtremeValues_where_dataEntryRangeIsDifferentToInputRange [expect: F6]", result, "$F$6")
        rowNum = rowNum + 1
        Call ClearIssues
        
        .Range("A1" & ":C" & rowNum).columns.AutoFit
    End With
End Sub
Sub clearOutput()
    Dim lastEntryRowNum As Integer
    lastEntryRowNum = lastRowNumOfNonEmptyCellInCol(5, "Version", "A")
    Worksheets("Version").Range("A5" & ":C" & lastEntryRowNum).ClearContents
End Sub
Sub addOutput(ByVal rowNum As Integer, ByVal test As String, ByVal actual As String, ByVal expectedOutput As String, Optional ByVal exactMatch = True)
    With Worksheets("Version")
        .Range("A" & rowNum) = test
        If exactMatch = True Then
            If UCase(actual) = UCase(expectedOutput) Then
                .Range("B" & rowNum) = "PASSED"
            Else
                .Range("B" & rowNum) = "FAILED"
                .Range("C" & rowNum) = "Expected: " & expectedOutput & "   Actual Output: " & actual
            End If
        Else
            If InStr(UCase(actual), UCase(expectedOutput)) Then
                .Range("B" & rowNum) = "PASSED"
            Else
                .Range("B" & rowNum) = "FAILED"
                .Range("C" & rowNum) = "Expected: " & expectedOutput & "   in Actual Output: " & actual
            End If
        End If
    End With
End Sub
Sub constructTestTables()
  With Worksheets("Version")
    .Range("F4:L7").ClearContents
    .Range("F4") = 1
    .Range("F5") = 2
    .Range("F6") = 3
    .Range("F7") = 4
    .Range("G4") = 5
    .Range("G5") = 6
    .Range("G6") = 7
    .Range("G7") = 8
    
    .Range("J4") = 1
    .Range("J5") = 2
    .Range("J6") = 3
    .Range("J7") = 4
    .Range("K4") = 5
    .Range("K5") = 6
    .Range("K6") = 7
    .Range("K7") = 8
  End With
End Sub
