Attribute VB_Name = "DateFunctions"
'Sonya - Date Functions

Function getEarliestDate(ByVal sheetname As String, ByVal rangeAsString As String) As Date
    Dim rng As Range
    Set rng = Worksheets(sheetname).Range(rangeAsString)
    Dim cell As Range
    Dim result As Date
    result = CDate(rng(1).Value)
    For Each cell In rng
        If CDate(cell.Value) < result Then
            result = CDate(cell.Value)
        End If
    Next cell
    getEarliestDate = result
End Function
Function getLatestDate(ByVal sheetname As String, ByVal rangeAsString As String) As Date
    Dim rng As Range
    Set rng = Worksheets(sheetname).Range(rangeAsString)
    Dim cell As Range
    Dim result As Date
    result = CDate(rng(1).Value)
    For Each cell In rng
        If CDate(cell.Value) > result Then
            result = CDate(cell.Value)
        End If
    Next cell
    getLatestDate = result
End Function
Function getDayWithSuffix(ByVal d As Date) As String
    Dim dayStr As String
    dayStr = day(d)
    Dim found As Boolean
    found = False
    If CInt(dayStr) = "11" Or CInt(dayStr) = "12" Or CInt(dayStr) = "13" Then
        dayStr = dayStr & "th"
        found = True
    Else
        Dim number As String
        number = Mid(dayStr, Len(dayStr), 1)
        If number = "1" Then
            dayStr = dayStr & "st"
            found = True
        Else
            If number = "2" Then
                dayStr = dayStr & "nd"
                found = True
            Else
                If number = "3" Then
                    dayStr = dayStr & "rd"
                    found = True
                End If
            End If
        End If
    End If
    If found = False Then
        dayStr = dayStr & "th"
    End If
    getDayWithSuffix = dayStr
End Function
Function addZeroToHrOrMin(ByVal hourOrMin As Integer) As String
    addZeroToHrOrMin = hourOrMin
    If hourOrMin < 10 Then
        addZeroToHrOrMin = "0" & hourOrMin
    End If
End Function
