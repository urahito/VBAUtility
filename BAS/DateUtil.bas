Attribute VB_Name = "DateUtil"
Private Const START_YEAR = 2017

Public Function DateCompare(ByVal dateA As Date, ByVal dateB As Date) As Integer
    If dateA > dateB Then
        IsDateEqual = 1
    ElseIf dateA = dateB Then
        IsDateEqual = 0
    Else
        IsDateEqual = -1
    End If
End Function

Public Function IsDateStr(ByRef dateStr As String, Optional askConv As Boolean = True) As Boolean
    On Error GoTo Err
    Dim formatStr As String
    Dim HasColon As Boolean, HasSlash As Boolean
    IsDateStr = False
    formatStr = ""
    
    HasColon = InStr(1, dateStr, DATE_COLON) > 0
    HasSlash = InStr(1, dateStr, DATE_SLASH) > 0
    
    dateStr = Replace(dateStr, DATE_SLASH, "")
    dateStr = Replace(dateStr, DATE_COLON, "")
    dateStr = Replace(dateStr, " ", "")
    
    If IsNumeric(dateStr) Then
        Select Case Len(dateStr)
            Case 8
                formatStr = "@@@@/@@/@@"
            Case 13
                formatStr = "@@@@/@@/@@ @:@@:@@"
            Case 14
                formatStr = "@@@@/@@/@@ @@:@@:@@"
            Case 5
                If HasColon Then
                    formatStr = "@:@@:@@"
                Else
                    GoTo Err
                End If
            Case 6
                If HasSlash Then
                    formatStr = "@@@@/@@"
                ElseIf HasColon Then
                    formatStr = "@@:@@:@@"
                Else
                    formatStr = "@@@@/@@"
                End If
            Case Else
                GoTo Err
        End Select
        
        Debug.Print "formatStr: " & formatStr
        If askConv = False Then
            dateStr = Format(dateStr, formatStr)
            Debug.Print "dateStr: " & dateStr
            IsDateStr = True
            
            Exit Function
        ElseIf MsgBox("ì˙ïtå^Ç…ïœä∑ÇµÇ‹Ç∑Ç©ÅH", vbYesNo) = vbYes Then
            dateStr = Format(dateStr, formatStr)
            Debug.Print "dateStr: " & dateStr
            IsDateStr = True
            
            Exit Function
        End If
    End If
    Exit Function
Err:

End Function

Public Function GetMonthFirst(ByVal valDate As Date, Optional ByVal excludeHoliday As Boolean = False, Optional ByVal excludeShukujitsu As Boolean = False) As Date
    Dim shukujitsuRng As Range, pivotRng As Range, shukujitsuDate As Date
    Dim monthLastDate As Date
    Set shukujitsuRng = Sheets("èjì˙àÍóó").Range("A2:A367")
    monthLastDate = WorksheetFunction.EoMonth(valDate, 0)
    GetMonthFirst = DateSerial(Year(valDate), Month(valDate), 1)
    
    If excludeHoliday = False Then Exit Function
    
    Do
        If GetMonthFirst > monthLastDate Then Exit Do
        
        If IsHoliday(GetMonthFirst) = False Then Exit Do
        
        GetMonthFirst = DateAdd("d", 1, GetMonthFirst)
    Loop While GetMonthFirst <= monthLastDate
    
    If excludeShukujitsu = False Then Exit Function
    
    Set pivotRng = shukujitsuRng.Cells(1, 1)
    Do
        shukujitsuDate = pivotRng.Value
        
        If IsCorrectYoubi(shukujitsuDate, "ì˙") Then
            shukujitsuDate = DateAdd("d", 1, shukujitsuDate)
        End If
        
        If GetMonthFirst <> shukujitsuDate Then Exit Do
        GetMonthFirst = DateAdd("d", 1, GetMonthFirst)
    Loop While GetDownRowValue(pivotRng) And (GetMonthFirst <= monthLastDate)
    
End Function

Public Function GetWorkdayBetween(ByVal startDate As Date, ByVal endDate As Date) As Integer
    Dim tempDate As Date
    
    Set nHolidayRng = GetShukujitsuRng
    
    If startDate > endDate Then
        Call clsPubDebug.MsgWarning("GetWorkdayBetween", "startDate>", "endDate", "startDate: ", startDate, "endDate: ", endDate)
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    GetWorkdayBetween = WorksheetFunction.NetworkDays(startDate, endDate, nHolidayRng)
End Function

Public Function GetHoursDbl(ByVal dateVal As Date) As Double
    GetHoursDbl = CDbl(dateVal / TimeSerial(1, 0, 0))
End Function

Public Function SetHoursDate(ByVal daysVal As Double, ByVal hoursVal As Double, ByVal minutesVal As Double) As Date
    SetHoursDate = CDate((daysVal * 24 + hoursVal + minutesVal / 60) * TimeSerial(1, 0, 0))
End Function

Public Function GetToday() As Date
    GetToday = DateSerial(Year(Now), Month(Now), Day(Now))
End Function

Public Function GetDateOnly(ByVal dateVal As Date, Optional ByVal addDay As Integer = 0) As Date
    dateVal = Fix(dateVal)
    GetDateOnly = DateAdd("d", addDay, dateVal)
End Function

Public Function GetNextDay(ByVal dateVal As Date) As Date
    GetNextDay = GetDateOnly(dateVal, 1)
End Function

Public Function IsNextMonth(ByRef dateVal As Date) As Boolean
    Dim nowDate As Date, nextDate As Date
    nowDate = dateVal
    nextDate = GetNextDay(dateVal)
    IsNextMonth = (Month(nextDate) > Month(nowDate))
    If IsNextMonth = False Then dateVal = nextDate
End Function

Public Function GetNowTimeStr(ByVal timeFormat As String) As String
    GetNowTimeStr = Format(Now, timeFormat)
End Function

Public Sub SetYearHolidayMain()
    Dim yearInt As Integer
    
    Focus = True
    Call ClearList
    For yearInt = START_YEAR To Year(Now) + 1
        Call SetYearHoliday(yearInt)
        Call CopyToList(yearInt - START_YEAR)
    Next
    Call GetShukujitsuRng
    Focus = False
End Sub

Public Function GetShukujitsuRng() As Range
    Dim nHolidaySht As Worksheet
    Dim maxRow As Long, rowInt As Integer
    Dim tempDate As Date
    Dim pivotRng As Range
    Dim yearInt As Integer, renkyuCount As Integer, prevRowInt As Integer
    
    Set nHolidaySht = Sheets("èjì˙àÍóó")
    maxRow = GetMaxRow(nHolidaySht)
    
    With nHolidaySht
        Set GetShukujitsuRng = .Range(.Cells(2, 1), .Cells(maxRow, 1))
    End With
    
    Set pivotRng = GetShukujitsuRng.Cells(1, 1)
    rowInt = 1
    nHolidaySht.Cells(2, 13).Value = 0
    
    Do
        If IsCorrectYoubi(pivotRng, "ì˙") Then
            With pivotRng
                .Value = DateAdd("d", 1, .Value)
                
                .Offset(0, 2).Value = .Offset(0, 2).Value & "ÅiêUë÷ãxì˙Åj"
            End With
        End If
        
        pivotRng.Offset(0, 1).Value = pivotRng.Value
        yearInt = Year(pivotRng.Value)
        rowInt = yearInt - START_YEAR + 2
        
        If rowInt <> prevRowInt Then
            Call RecordRenkyuCount(nHolidaySht, prevRowInt, renkyuCount)
            prevRowInt = rowInt
        ElseIf renkyuCount > 0 Then
            Call RecordRenkyuCount(nHolidaySht, rowInt, renkyuCount)
        End If
        
        If IsRenkyu(pivotRng.Value) = False Then
            renkyuCount = renkyuCount + 1
        Else
            renkyuCount = 0
        End If
        
        nHolidaySht.Cells(rowInt, 12).Value = yearInt
    Loop While GetDownRowValue(pivotRng) = True
    
    Call RecordRenkyuCount(nHolidaySht, rowInt - 1, renkyuCount)
End Function

Private Sub RecordRenkyuCount(ByRef sht As Worksheet, ByVal rowInt As Integer, ByRef renkyuCount As Integer)
    If renkyuCount > 0 Then
        With sht.Cells(rowInt, 13)
            .Value = .Value + 1
        End With
        renkyuCount = 0
    End If
End Sub

Private Function IsRenkyu(ByVal valDate As Date) As Boolean
    Dim holidayCount As Integer
    Dim startDate As Date
    IsRenkyu = False
    holidayCount = -1
    
    startDate = valDate
    
    Do
        holidayCount = holidayCount + 1
        valDate = DateAdd("d", -1, valDate)
    Loop Until IsHoliday(valDate) = False
    
    valDate = startDate
    
    holidayCount = holidayCount - 1
    
    Do
        holidayCount = holidayCount + 1
        valDate = DateAdd("d", 1, valDate)
    Loop Until IsHoliday(valDate) = False
    
    IsRenkyu = holidayCount > 0
End Function

Public Function IsHoliday(ByVal valDate As Date) As Boolean
    IsHoliday = (WorksheetFunction.Weekday(valDate, 2) >= 6)
End Function

Public Function IsCorrectYoubi(ByVal valDate As Date, ByRef weekDayVar As Variant) As Boolean
    On Error GoTo ProcessFalse
    Dim youbiInt As Integer
    IsCorrectYoubi = True

    If VarType(weekDayVar) = vbString Then
        If IncludeStrs(CStr(weekDayVar), "åé", "Mon") Then
            weekDayVar = 1
        ElseIf IncludeStrs(CStr(weekDayVar), "âŒ", "Tue") Then
            weekDayVar = 2
        ElseIf IncludeStrs(CStr(weekDayVar), "êÖ", "Wed") Then
            weekDayVar = 3
        ElseIf IncludeStrs(CStr(weekDayVar), "ñÿ", "Thu") Then
            weekDayVar = 4
        ElseIf IncludeStrs(CStr(weekDayVar), "ã‡", "Fri") Then
            weekDayVar = 5
        ElseIf IncludeStrs(CStr(weekDayVar), "ìy", "Sat") Then
            weekDayVar = 6
        ElseIf IncludeStrs(CStr(weekDayVar), "ì˙", "Sun") Then
            weekDayVar = 7
        Else
            IsCorrectYoubi = False
        End If
    End If
    
    If IsCorrectYoubi = False Then GoTo ProcessFalse
    
    If VarType(weekDayVar) = vbInteger Or VarType(weekDayVar) = vbLong Then
        youbiInt = CInt(weekDayVar)
        
        If youbiInt < 1 And youbiInt > 7 Then
            IsCorrectYoubi = False
            Exit Function
        End If
    Else
        IsCorrectYoubi = False
    End If
    
    If IsCorrectYoubi = False Then GoTo ProcessFalse
    
    weekDayVar = youbiInt
    IsCorrectYoubi = (WorksheetFunction.Weekday(valDate, 2) = youbiInt)
    Exit Function
ProcessFalse:
    IsCorrectYoubi = False
    weekDayVar = 0
End Function

Private Sub ClearList()
    Dim shukujitsuSht As Worksheet, toRng As Range, countRng As Range
    Dim maxRow As Long
    
    Set shukujitsuSht = Sheets("èjì˙àÍóó")
    maxRow = GetMaxRow(shukujitsuSht)
    
    With shukujitsuSht
        Set toRng = .Range(.Cells(2, 1), .Cells(maxRow, 3))
        Set countRng = .Range(.Cells(2, 12), .Cells(maxRow, 13))
    End With
    
    toRng.Clear
    countRng.Clear
End Sub

Private Sub CopyToList(ByVal offsetInt As Integer)
    Dim shukujitsuSht As Worksheet, fromRng As Range, toRng As Range
    Dim maxRow As Long, recordCount As Integer, offsetRow As Integer, toRow As Integer
    
    Set shukujitsuSht = Sheets("èjì˙àÍóó")
    maxRow = GetMaxRow(shukujitsuSht, 5)
    recordCount = maxRow - 1
    toRow = 2 + recordCount * offsetInt
    
    With shukujitsuSht
        Set fromRng = .Range(.Cells(2, 8), .Cells(maxRow, 10))
        Set toRng = .Range(.Cells(toRow, 1), .Cells(toRow + recordCount - 1, 3))
    End With
    
    Call fromRng.Copy(toRng)
End Sub

Private Sub SetYearHoliday(ByVal yearInt As Integer)
    Dim shukujitsuSht As Worksheet, pivotRng As Range
    Dim monthInt As Integer, weekNum As Integer, dayInt As Integer
    Dim isBlankWeek As Boolean, isBlankDay As Boolean
    
    Set shukujitsuSht = Sheets("èjì˙àÍóó")
    Set pivotRng = shukujitsuSht.Cells(2, 5)
    
    Do
        With pivotRng
            isBlankWeek = (.Offset(0, 1).Value = "")
            isBlankDay = (.Offset(0, 2).Value = "")
            
            monthInt = .Value
            weekNum = IIf(isBlankWeek, 0, .Offset(0, 1).Value)
            dayInt = IIf(isBlankDay, 0, .Offset(0, 2).Value)
        End With
        
        With pivotRng.Offset(0, 3)
            .NumberFormat = "yyyy/M/d"
            If isBlankWeek And isBlankDay Then
                Select Case monthInt
                    Case 3
                        .Value = GetSpringDay(yearInt)
                    Case 9
                        .Value = GetAutumnDay(yearInt)
                    Case Else
                        .Value = ""
                End Select
            ElseIf isBlankWeek = False Then
                .Value = GetDateFromWeekNum(yearInt, monthInt, weekNum)
            ElseIf isBlankDay = False Then
                .Value = DateSerial(yearInt, monthInt, dayInt)
            Else
                .Value = ""
            End If
        End With
        
        With pivotRng.Offset(0, 4)
            .NumberFormat = "(aaa)"
            .Value = pivotRng.Offset(0, 3).Value
        End With
    Loop While GetDownRowValue(pivotRng) = True
End Sub

Public Function GetSpringDay(ByVal yearInt As Integer) As Date
    Dim dayInt As Integer
    dayInt = Int(20.8431 + 0.242194 * (yearInt - 1980) - Int((yearInt - 1980) / 4))
    GetSpringDay = DateSerial(yearInt, 3, dayInt)
End Function

Public Function GetAutumnDay(ByVal yearInt As Integer) As Date
    Dim dayInt As Integer
    dayInt = Int(20.8431 + 0.242194 * (yearInt - 1980) - Int((yearInt - 1980) / 4))
    GetAutumnDay = DateSerial(yearInt, 9, dayInt)
End Function

Public Function GetDateFromWeekNum(ByVal yearInt As Integer, ByVal monthInt As Integer, ByVal weekNumInt As Integer) As Date
    Dim pivotDate As Date
    pivotDate = GetMonthFirst(DateSerial(yearInt, monthInt, 1), True, False)
    
    Do Until IsCorrectYoubi(pivotDate, "åé")    'åéójì˙(1)éwíËÅFç≈èâÇÃåéójì˙ÇéÊìæÇ∑ÇÈ
        pivotDate = DateAdd("d", 1, pivotDate)
    Loop
    
    GetDateFromWeekNum = DateAdd("ww", (weekNumInt - 1), pivotDate)     'ëÊnèTÇÃåéójì˙ÇéÊìæÇ∑ÇÈ
End Function

