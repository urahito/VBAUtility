Attribute VB_Name = "Common"
Public Const DELIMITER_COMMA = ","
Public Const DELIMITER_TAB = vbTab
Public Const DATA_DBLQUOTE = """"
Public Const DATA_4_DBLQUOTE = DATA_DBLQUOTE & DATA_DBLQUOTE & DATA_DBLQUOTE & DATA_DBLQUOTE
Public Const DATA_EMPTY = ""
Public Const FILE_EXT = "."
Public Const FILE_CSV = "csv"
Public Const FILE_TSV = "tsv"
Public Const FILE_XLSX = "xlsx"
Public Const DATE_COLON = ":"
Public Const DATE_SLASH = "/"
Public Const ISNOT_BLANK = "<>"""""
Public Const IS_BLANK = """"""
Private Const PRINT_VALUE_INFO = False

Property Let Focus(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Property

Public Function GetRangeInfo(ByVal rng As Range) As String
    GetRangeInfo = "[no value or nothing]"
    If IsNoValue(rng, "", False) = True Then Exit Function
    
    GetRangeInfo = "(Address: " & rng.Address & " [" & rng.Worksheet.Name & "] = " & rng.Value & ")"
End Function

Public Sub CopyPasteSpecial(ByVal rngFrom As Range, ByVal rngDest As Range)
    DoEvents
    rngFrom.Copy
    rngDest.PasteSpecial (xlPasteFormats)
    Application.CutCopyMode = False
End Sub

Public Sub CopyPasteValues(ByVal rngFrom As Range, ByVal rngDest As Range)
    DoEvents
    rngFrom.Copy
    rngDest.PasteSpecial (xlPasteValuesAndNumberFormats)
    Application.CutCopyMode = False
End Sub

Public Function LastActivateSheet(ParamArray shts()) As Boolean
    On Error GoTo Err
    Dim item As Variant
    Dim sht As Worksheet, shtName As String
    LastActivateSheet = False
    shtName = "(None)"
    DoEvents
    Call clsPubDebug.MsgDebugBegin("LastActivateSheet")
    
    For Each item In shts
        If IsNothing(item) Then
            Call clsPubDebug.MsgDebug("(Nothing)")
        ElseIf TypeName(item) = "Worksheet" Then
            Set sht = item
            sht.Activate
            Call clsPubDebug.MsgDebug("sheetName: ", sht.Name)
            LastActivateSheet = True
            Exit For
        End If
        DoEvents
    Next
    
    If LastActivateSheet = False And FindSheet("メイン画面", sht, False, False) = True Then
        sht.Activate
        Call clsPubDebug.MsgWarning("引数が無効だったため、メイン画面を表示します")
    End If
    Call clsPubDebug.MsgDebugEnd("Result: ", LastActivateSheet, "Activated: ", shtName)
    If Not IsNothing(sht) Then shtName = sht.Name
    Exit Function
Err:
    Call clsPubDebug.MsgError(Err.Number, Err.Description, Err.Source)
    
End Function

Public Sub FindAndDeleteSheet(ByVal sheetName As String, Optional ByVal likeSearch As Boolean = False, Optional wkBook As Workbook = Nothing)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim tempShtName As String
    DoEvents
    
    Set wb = IIf(wkBook Is Nothing, ActiveWorkbook, wkBook)
    wb.Activate
    tempShtName = sheetName
    
    Do While likeSearch = True And FindSheet(sheetName, ws, likeSearch, False, wb)
        ws.Delete
        sheetName = tempShtName
        DoEvents
    Loop
End Sub

Public Function FindSheet(ByRef sheetName As String, ByRef sheetObj As Worksheet, Optional ByVal likeSearch As Boolean = False, Optional ByVal makeNewSheet As Boolean = True, Optional wkBook As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    Dim wb As Workbook
    FindSheet = False
    DoEvents
    
    Set wb = IIf(wkBook Is Nothing, ActiveWorkbook, wkBook)
    
    For Each ws In wb.Worksheets
        If likeSearch = True And ws.Name Like sheetName & "*" Then
            sheetName = ws.Name
            FindSheet = True
        ElseIf ws.Name = sheetName Then
            sheetName = ws.Name
            FindSheet = True
        End If
        
        If FindSheet = True Then
            Set sheetObj = Worksheets(sheetName)
            Exit Function
        End If
    Next ws
    
    If makeNewSheet = True Then
        Set sheetObj = Worksheets.Add
        sheetObj.Name = sheetName
        FindSheet = True
    End If
End Function

Public Sub SetValueWithFormat(ByRef rng As Range, ByVal val As Variant, ByVal formatStr As String)
    DoEvents
    With rng
        .NumberFormat = formatStr
        .Value = val
    End With
End Sub

Public Sub SetValueAsStr(ByRef rng As Range, ByVal val As Variant, ByVal withoutQuot As Boolean)
    DoEvents
    With rng
        .NumberFormat = "@"
        .Value = RemoveDblQuote(val, withoutQuot)
    End With
End Sub

Public Function RemoveDblQuote(ByVal val As Variant, ByVal withoutQuot As Boolean)
    Dim quotStr As String
    DoEvents
    quotStr = IIf(withoutQuot, DATA_DBLQUOTE, DATA_EMPTY)
    
    RemoveDblQuote = Replace(CStr(val), quotStr, DATA_EMPTY)
End Function

Public Function GetInputValue(ByRef val As Variant, ByVal promptStr As String, ByVal errMsgStr As String, ByVal defaultValue As Variant) As Boolean
    val = InputBox(promptStr)
    DoEvents
    GetInputValue = Not IsNoValue(val, errMsgStr)
    
    If GetInputValue = False Then
        val = defaultValue
    End If
End Function

Public Function IsNoValue(ByRef val As Variant, ByVal errMsgStr As String, Optional ByVal viewMsg As Boolean = True) As Boolean
    On Error GoTo Err
    Dim dateStr As String
    DoEvents
    Call PrintValueInfo(val)
    
    IsNoValue = False
    
    If errMsgStr = "" Then
        errMsgStr = "無効な値です"
    End If
    
    If VarType(val) = vbString Then
        IsNoValue = False
        
        If val = "" Or Len(val) = 0 Then    'InputBox空欄OK時, str = ""
            IsNoValue = True
        End If
    ElseIf IsNothing(val) Then
        IsNoValue = True
    ElseIf IsEmpty(val) Then    '初期化されていない値
        IsNoValue = True
    ElseIf IsMissing(val) Then  'Optional変数で値がないとき（空文字を入れる）
        val = ""
        IsNoValue = True
    ElseIf IsError(val) Then    'エラー値のとき（空文字を入れる）
        val = ""
        IsNoValue = True
    ElseIf IsNumeric(val) Then
        IsNoValue = False
    ElseIf IsDate(val) Then     '日付データのとき
        val = CDate(val)
        IsNoValue = False
    ElseIf StrPtr(val) = 0 Then 'InputBoxキャンセル時
        IsNoValue = True
    ElseIf IsNull(val) Then     'オブジェクトが空
        IsNoValue = True
    ElseIf val = "" Then        'InputBox空欄OK時, str = ""
        IsNoValue = True
    ElseIf Len(val) = 0 Then    '空文字列
        IsNoValue = True
    End If
    
    If IsNoValue = False Then
        If IsDate(val) = False Then
            dateStr = CStr(val)
            If IsDateStr(dateStr) = True Then
                val = CDate(dateStr)
            End If
        End If
    ElseIf IsArray(val) Then
        Debug.Print "Array: Len(" & UBound(val) & ")"
    ElseIf viewMsg = True Then
        MsgBox errMsgStr
    End If
    Call PrintValueInfo(val, IsNoValue)
    Exit Function
Err:
    IsNoValue = True
    Call PrintValueInfo(val, IsNoValue)
    MsgBox errMsgStr & vbCrLf & Err.Number & ": " & Err.Description
End Function

Private Sub PrintValueInfo(ByVal val As Variant, Optional ByVal result As String = "")
    Dim valStr As String
    
    DoEvents
    If PRINT_VALUE_INFO = False Then Exit Sub
    
    If IsNothing(val) Then
        valStr = ""
    Else
        valStr = CStr(val)
    End If
    
    If Len(valStr) > 29 Then
        valStr = Left(valStr, 29) & "..."
    End If
        
    Debug.Print "TypeName(val): " & TypeName(val) & " = " & _
                    IIf(TypeName(val) = "Nothing", "Nothing", valStr) & _
                    IIf(result = "", "", "(" & result & ")")
End Sub

Public Function IncludeStrs(ByVal val As String, ParamArray strArr() As Variant) As Boolean
    Dim itemVar As Variant
    IncludeStrs = True
    
    For Each itemVar In strArr
        If VarType(itemVar) = vbString Then
            If InStr(1, val, itemVar) > 0 Then
                Exit Function
            End If
        End If
        DoEvents
    Next
    
    IncludeStrs = False
End Function

Public Function GetDownRowValue(ByRef rng As Range) As Boolean
    GetDownRowValue = False
    DoEvents
    If rng.Row = Rows.Count Then Exit Function
    Set rng = rng.Offset(1, 0)
    GetDownRowValue = Not IsNoValue(rng.Value, "", False)
End Function

Public Function GetUpRowValue(ByRef rng As Range) As Boolean
    GetUpRowValue = False
    DoEvents
    If rng.Row = 1 Then Exit Function
    Set rng = rng.Offset(-1, 0)
    GetUpRowValue = Not IsNoValue(rng.Value, "", False)
End Function

Public Function GetRightColValue(ByRef rng As Range) As Boolean
    GetRightColValue = False
    DoEvents
    If rng.Column = Columns.Count Then Exit Function
    Set rng = rng.Offset(0, 1)
    GetRightColValue = Not IsNoValue(rng.Value, "", False)
End Function

Public Function GetLeftColValue(ByRef rng As Range) As Boolean
    GetLeftColValue = False
    DoEvents
    If rng.Column = 1 Then Exit Function
    Set rng = rng.Offset(0, 1)
    GetLeftColValue = Not IsNoValue(rng.Value, "", False)
End Function

Public Function GetMaxRow(ByRef sht As Worksheet, Optional ByVal colNum As Long = 1) As Long
    DoEvents
    GetMaxRow = sht.Cells(Rows.Count, colNum).End(xlUp).Row
End Function

Public Function GetMaxCol(ByRef sht As Worksheet, Optional ByVal rowNum As Long = 1) As Long
    DoEvents
    GetMaxCol = sht.Cells(rowNum, Columns.Count).End(xlUp).Column
End Function

Public Function IsNothing(ByVal objvar As Variant) As Boolean
    DoEvents
    IsNothing = (TypeName(objvar) = "Nothing")
End Function


