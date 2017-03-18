VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDebugUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const LOG_DEBUG = "Debug"
Private Const LOG_WARN = "Warning"
Private Const LOG_ERR = "Error"

Private filePath As String
Private dirPath As String
Private fileName As String
Private fileNo As Integer

Private countErr As Long
Private countWarn As Long
Private countNoMethodName As Long
Private callCount As Long
Private methodName As String
Private tempMethodName As String

Private Sub Class_Initialize()
    On Error GoTo Err
    dirPath = ThisWorkbook.Path & IIf(Right(ThisWorkbook.Path, 1) = "\", "", "\")
    fileName = "log_" & GetNowTimeStr("yyyyMMddHHmmss") & ".txt"
    filePath = dirPath & fileName
    fileNo = FreeFile
    countErr = 0
    countWarn = 0
    callCount = 0
    countNoMethodName = 0
    
    Open filePath For Output As #fileNo
    Print #fileNo, ThisWorkbook.Name
    Print #fileNo, "Start at " & GetNowTimeStr("yyyy/MM/dd HH:mm:ss")
    Exit Sub
Err:
    Close #fileNo
End Sub

Private Sub Class_Terminate()
    Dim msgStr As String
    msgStr = GetMsg()

    If IsNoValue(msgStr, "", False) = False Then
        MsgBox msgStr
    End If
    Close
End Sub

Public Function MsgDebug(ParamArray disStrArr()) As Boolean
    On Error GoTo LogErr
    Dim errStr As String, disStr As String
    Dim arrIdx As Integer, paramCount As Integer
    
    If methodName = "" Then
        Debug.Print "methodName: Empty, Maybe?: " & tempMethodName
        methodName = tempMethodName
        countNoMethodName = countNoMethodName + 1
    End If
    
    disStr = ""
    paramCount = 0
    For arrIdx = LBound(disStrArr) To UBound(disStrArr)
        paramCount = paramCount + 1
        disStr = disStr & disStrArr(arrIdx) & AddComma(paramCount, LBound(disStrArr), UBound(disStrArr))
    Next
    
    errStr = GetLogStr(LOG_DEBUG, methodName, errDisStr:=disStr)
    Print #fileNo, errStr
    
Err:
    MsgDebug = True
    Exit Function
LogErr:
    errStr = GetLogStr(LOG_ERR, "MsgDebug(" & methodName & ")", Err.Number, Err.Description, Err.Source)
    Debug.Print errStr
    
    MsgDebug = False
End Function

Public Function MsgDebugBegin(ByVal methodNm As String, ParamArray inputVal()) As Boolean
    On Error GoTo LogErr
    Dim errStr As String, disStr As String
    Dim arrIdx As Integer, paramCount As Integer
    Dim commaPos As Integer
    
    If methodNm = "" Then
        Call MsgWarning("MsgDebugBegin", "methodName: ", "Empty")
    Else
        methodName = methodNm
        tempMethodName = methodNm
    End If
    
    callCount = callCount + 1
    
    disStr = ""
    paramCount = 0
    For arrIdx = LBound(inputVal) To UBound(inputVal)
        paramCount = paramCount + 1
        disStr = disStr & inputVal(arrIdx) & AddComma(paramCount, LBound(inputVal), UBound(inputVal))
    Next
    
    If Len(disStr) > 0 Then
        disStr = "--BEGIN(" & disStr & ")"
    Else
        disStr = "--BEGIN"
    End If
    
    errStr = GetLogStr(LOG_DEBUG, methodName, errDisStr:=disStr)
    Print #fileNo, "[" & callCount & "] ------------------------------------------------------------------------"
    Print #fileNo, errStr
    
Err:
    MsgDebugBegin = True
    Exit Function
LogErr:
    errStr = GetLogStr(LOG_ERR, "MsgDebugBegin(" & methodName & ")", Err.Number, Err.Description, Err.Source)
    Debug.Print errStr
    
    MsgDebugBegin = False
End Function

Public Function MsgDebugEnd(ParamArray outputVal()) As Boolean
    On Error GoTo LogErr
    Dim errStr As String, disStr As String
    Dim arrIdx As Long, paramCount As Long
    
    If methodName = "" Then methodName = tempMethodName
    
    disStr = ""
    paramCount = 0
    For arrIdx = LBound(outputVal) To UBound(outputVal)
        paramCount = paramCount + 1
        disStr = disStr & outputVal(arrIdx) & AddComma(paramCount, LBound(outputVal), UBound(outputVal))
    Next
    
    If Len(disStr) > 0 Then
        disStr = "--END(" & disStr & ")"
    Else
        disStr = "--END"
    End If
    
    errStr = GetLogStr(LOG_DEBUG, methodName, errDisStr:=disStr)
    Print #fileNo, errStr
    Print #fileNo, "[" & callCount & "] ------------------------------------------------------------------------"
    
    If callCount = 1 And countWarn > 0 Then
        Print #fileNo, GetLogStr(LOG_DEBUG, methodName, errDisStr:="Warning: " & countWarn)
    End If
    
    callCount = callCount - 1
Err:
    MsgDebugEnd = True
    methodName = ""
    Exit Function
LogErr:
    errStr = GetLogStr(LOG_ERR, "MsgDebugEnd(" & methodName & ")", Err.Number, Err.Description, Err.Source)
    Debug.Print errStr
    
    methodName = ""
    MsgDebugEnd = False
End Function

Public Function MsgWarning(ParamArray disStrArray()) As Boolean
    On Error GoTo LogErr
    Dim errStr As String, disStr As String
    Dim arrIdx As Integer
        
    countWarn = countWarn + 1
    
    disStr = ""
    For arrIdx = LBound(disStrArray) To UBound(disStrArray)
        disStr = disStr & disStrArray(arrIdx)
    Next
    
    errStr = GetLogStr(LOG_WARN, methodName, errDisStr:=disStr)
    Print #fileNo, errStr
    
Err:
    MsgWarning = True
    Exit Function
LogErr:
    countErr = countErr + 1
    
    errStr = GetLogStr(LOG_ERR, "MsgWarning(" & methodName & ")", Err.Number, Err.Description, Err.Source)
    Debug.Print errStr
    
    MsgWarning = False
End Function

Public Function MsgError(ByVal errNo As Integer, ByVal errDisStr As String, errSrcStr As String) As Boolean
    On Error GoTo LogErr
    Dim errStr As String
    
    If errNo = 0 Then Exit Function
    
    countErr = countErr + 1
    
    errStr = GetLogStr(LOG_ERR, methodName, errNo, errDisStr, errSrcStr)
    Print #fileNo, errStr
    
Err:
    MsgError = True
    Exit Function
LogErr:
    countErr = countErr + 1
    errStr = GetLogStr(LOG_ERR, "MsgError(" & methodName & ")", Err.Number, Err.Description, Err.Source)
    MsgBox errStr
    
    methodName = ""
    MsgError = False
End Function

Private Function GetLogStr(ByVal LogType As String, ByVal methodName As String, Optional ByVal errNo As Integer = 0, Optional ByVal errDisStr As String = "", Optional errSrcStr As String = "") As String
    GetLogStr = LogType & " in [" & methodName & "], " & _
                GetNowTimeStr("HH:mm:ss") & ", "
                
    If errNo > 0 Then
        GetLogStr = GetLogStr & "<" & errNo & "> "
    End If
    If IsNoValue(errDisStr, "", False) = False Then
        GetLogStr = GetLogStr & errDisStr & IIf(LogType <> LOG_ERR, "", DELIMITER_COMMA & " ")
    End If
    If IsNoValue(errSrcStr, "", False) = False Then
        GetLogStr = GetLogStr & "at " & errSrcStr
    End If
End Function

Private Function GetMsg() As String
    GetMsg = ""
    
    If (countErr + CDbl(countWarn) / 10) >= 1 Then
        GetMsg = "[" & LOG_ERR & "]: " & countErr & vbCrLf & _
                "[" & LOG_WARN & "]: " & countWarn & vbCrLf & _
                "[No method name]: " & countNoMethodName
    End If
End Function

Private Function AddComma(ByVal countInt As Integer, ByVal firstIdx As Integer, ByVal lastIdx As Integer) As String
    AddComma = ""
    lastIdx = lastIdx - firstIdx + 1
    If countInt = lastIdx Then Exit Function
    
    If countInt Mod 2 = 0 Then
        AddComma = DELIMITER_COMMA & " "
    End If
End Function

Property Let TEMP_METHOD_NAME(methodNm As String)
    tempMethodName = methodNm
End Property


