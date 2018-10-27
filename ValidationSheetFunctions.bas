Attribute VB_Name = "ValidationSheetFunctions"
'Sonya - Validation Sheet

Function updateValidationSheet(ByVal issue As String, ByVal sheetname As String, ByVal address As String, ByVal currentVal As String, ByVal suggestedvalues As String, Optional ByVal checkForDuplicatesBeforeUpdate = False)
    Dim row As Integer
    With Worksheets("Validation")
        row = .Range("B1").Value + 3
        Dim newRowAsCSV As String
        newRowAsCSV = issue & "," & sheetname & "," & address & "," & currentVal & "," & suggestedvalues
        If checkForDuplicatesBeforeUpdate = False Then
            Call addIssue(row, issue, sheetname, address, currentVal, suggestedvalues)
        Else
            If doesRowExistInRange_whereRowISsProvidedAsACSV(.Range("A2:E" & (row - 1)), newRowAsCSV) = False Then
                Call addIssue(row, issue, sheetname, address, currentVal, suggestedvalues)
            End If
        End If
        'createNavigateButtons (Worksheets("Validation").Range("B1").Value)
    End With
End Function
Sub addIssue(ByVal row As Integer, ByVal issue As String, ByVal sheetname As String, ByVal address As String, ByVal currentVal As String, ByVal suggestedvalues As String)
    With Worksheets("Validation")
        .Range("B1").Value = (row - 2)
        .Range("A" & row) = issue
        .Range("B" & row) = sheetname
        .Range("C" & row) = address
        If currentVal <> vbNullString Then
            .Range("D" & row) = currentVal
        End If
        If suggestedvalues <> vbNullString Then
            .Range("E" & row) = suggestedvalues
        End If
    End With
End Sub
Sub ClearIssues()
    With Worksheets("Validation")
        .Cells.ClearContents
        .Range("B1") = 0
        .Range("A1") = "No. Of Issues"
        .Range("A2") = "Validation Check Type"
        .Range("B2") = "Sheet"
        .Range("C2") = "Address"
        .Range("D2") = "Current Value"
        .Range("E2") = "Suggested Values"
        .Range("A2:D2").Interior.ColorIndex = 48
    End With
End Sub
Sub createNavigateButtons(ByVal noOfIssues As Integer)
        Application.ScreenUpdating = False
        Worksheets("Validation").Buttons.Delete
        Dim btn As Button
        For i = 3 To noOfIssues + 2
            Dim t As Range
            Set t = Range(Cells(i, 3), Cells(i, 3))
            Set t = Worksheets("Validation").Range(t.address)
            If t.Value <> vbNullString Then
             Dim sheetname As String
             sheetname = Worksheets("Validation").Range("B" & i).Value
             Set btn = Worksheets("Validation").Buttons.add(t.Left, t.Top, t.Width, t.Height)
             With btn
                .OnAction = "NavigateToCell"
                .Caption = Replace(t.Value, "$", "")
                .name = sheetname & "!" & t.Value
             End With
            End If
        Next i
End Sub
Sub NavigateToCell()
    Dim cellRange As String
    cellRange = Application.Caller
    Dim data() As String
    data = Split(cellRange, "!")
    ThisWorkbook.Sheets(data(0)).Activate
    Range(data(1)).Select
End Sub
