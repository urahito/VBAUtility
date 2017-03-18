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

Public Function GetMonthFirst(ByVal valDate As Date, Optional ByVal excludeHoliday As Boolean = False) As Date
    GetMonthFirst = WorksheetFunction.EoMonth(valDate, -1) + 1
    
    If excludeHoliday = False Then Exit Function
    
    GetMonthFirst = GetNextMonday(GetMonthFirst)
End Function

Public Function GetNextMonday(ByVal dateVal As Date) As Date
    GetNextMonday = dateVal + 8 - WorksheetFunction.Weekday(dateVal, 2)
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
    GetToday = GetDateOnly(Now)
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
    Dim nHolidaySht As Worksheet, nHolidayCols As Range
    Dim rng As Range, findRng As Range
    Dim yearInt As Integer
    Dim pivotDate As Date, endDate As Date, startDate As Date
    Dim renkyuCount As Long
    
    Focus = True
    Call FindSheet("èjì˙àÍóó", nHolidaySht, False, False, ThisWorkbook)
    pivotDate = DateSerial(START_YEAR, 1, 1)
    endDate = DateAdd("yyyy", 2, pivotDate) - 1
    
    Set rng = nHolidaySht.Cells(5, 5)
    With nHolidaySht
        Set nHolidayCols = .Range(.Columns(1), .Columns(3))
    End With
    
    'ìyì˙èjì˙éÊìæ
    Do
        With nHolidaySht
            Set findRng = nHolidayCols.Find(CDate(pivotDate))
        End With
        If Not findRng Is Nothing Then
            If findRng.Value = pivotDate Then
                With rng
                    Call SetValueWithFormat(rng, pivotDate, "yyyy/MM/dd")
                    Call SetValueWithFormat(.Offset(0, 1), pivotDate, "aaa")
                    Call SetValueWithFormat(.Offset(0, 2), findRng.Offset(0, 1).Value, "@")
                End With
                Set rng = rng.Offset(1, 0)
            End If
        Else
            If IsHoliday(pivotDate) = True Then
                With rng
                    Call SetValueWithFormat(rng, pivotDate, "yyyy/MM/dd")
                    Call SetValueWithFormat(.Offset(0, 1), pivotDate, "aaa")
                    Call SetValueWithFormat(.Offset(0, 2), "ãxì˙", "@")
                End With
                Set rng = rng.Offset(1, 0)
            End If
        End If
        pivotDate = pivotDate + 1
        Set findRng = Nothing
    Loop While pivotDate <= endDate
    
    'òAãxéÊìæ
    Set rng = nHolidaySht.Cells(5, 5)
    Set findRng = nHolidaySht.Cells(5, 9)
    renkyuCount = 1
    Do
        If rng.Row > 5 Then
            If rng.Value - 1 = rng.Offset(-1, 0).Value Then
                If renkyuCount = 1 Then
                    startDate = rng.Value - 1
                End If
                renkyuCount = renkyuCount + 1
            Else
                endDate = rng.Offset(-1, 0).Value
                If renkyuCount > 2 Then
                    Call SetValueWithFormat(findRng, startDate, "yyyy/MM/dd")
                    Call SetValueWithFormat(findRng.Offset(0, 1), endDate, "yyyy/MM/dd")
                    Call SetValueWithFormat(findRng.Offset(0, 2), renkyuCount, "0")
                    Set findRng = findRng.Offset(1, 0)
                End If
                renkyuCount = 1
            End If
        End If
    Loop While GetDownRowValue(rng)
    
    Focus = False
End Sub

Public Function IsHoliday(ByVal valDate As Date) As Boolean
    If valDate = 0 Then Exit Function
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
    pivotDate = GetMonthFirst(DateSerial(yearInt, monthInt, 1), True)
    
    Do Until IsCorrectYoubi(pivotDate, "åé")    'åéójì˙(1)éwíËÅFç≈èâÇÃåéójì˙ÇéÊìæÇ∑ÇÈ
        pivotDate = DateAdd("d", 1, pivotDate)
    Loop
    
    GetDateFromWeekNum = DateAdd("ww", (weekNumInt - 1), pivotDate)     'ëÊnèTÇÃåéójì˙ÇéÊìæÇ∑ÇÈ
End Function

