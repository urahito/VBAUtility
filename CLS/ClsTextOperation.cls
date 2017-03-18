VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsTextOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim fileInfo As ClsFileInfo
Dim fileNo As Integer
Dim fileMode As OPEN_MODE
Dim lineStr As String

Public Enum OPEN_MODE
    INPUT_EXIST
    OUTPUT_NEW
    OUTPUT_APPEND
End Enum

Private Sub Class_Initialize()
    Set fileInfo = New ClsFileInfo
End Sub

Private Sub Class_Terminate()
    Close #fileNo
    Set fileInfo = Nothing
End Sub

Public Function OpenFile(ByVal filePath As String, ByVal mode As OPEN_MODE) As Boolean
    On Error GoTo Err
    OpenFile = False
    fileMode = mode
    
    If fileInfo.SetPaths(filePath) = False Then Exit Function
    If IncludeStrs(fileInfo.FILE_EXT, ".txt", ".csv", ".tsv", ".log") = True Then
    Else
        GoTo Err
    End If
    
    fileNo = FreeFile
    
    Select Case fileMode
        Case INPUT_EXIST
            If fileInfo.CAN_READ = False Then GoTo Err
            Open fileInfo.GET_FULL_PATH For Input As #fileNo
        Case OUTPUT_APPEND
            Open fileInfo.GET_FULL_PATH For Append As #fileNo
        Case OUTPUT_NEW
            Open fileInfo.GET_FULL_PATH For Output As #fileNo
    End Select
    
    OpenFile = True
    Exit Function
Err:
    Close #fileNo
End Function

Public Function ReadAll() As String
    On Error GoTo Err
    ReadAll = ""
    
    If ReadFile = True Then
        Call OverwriteToCrLf
        If ReadFile = False Then Exit Function
    End If
    
    Do
        ReadAll = ReadAll & lineStr & vbCrLf
    Loop While ReadFile = True
Err:
    Close #fileNo
End Function

Public Function ReadFile() As Boolean
    On Error GoTo Err
    ReadFile = False
    lineStr = ""
    
    If IsInput = False Then Exit Function
    If EOF(fileNo) = True Then Exit Function
    
    Line Input #fileNo, lineStr
    ReadFile = True
    Exit Function
Err:
    Close #fileNo
End Function

Private Sub OverwriteToCrLf()
    If InStr(1, lineStr, vbCr) > 0 Then
    ElseIf InStr(1, lineStr, vbLf) > 0 Then
    Else
        Exit Sub
    End If
    If IsInput = False Then Exit Sub
    
    Close #fileNo
    fileNo = FreeFile
    Open fileInfo.GET_FULL_PATH For Output As #fileNo
    
    lineStr = Replace(lineStr, vbCr, vbLf)
    lineStr = Replace(lineStr, vbLf, vbCrLf)
    
    Print #fileNo, lineStr
    Close #fileNo
    Debug.Print "è„èëÇ´äÆóπ"
    
    Call OpenFile(fileInfo.GET_FULL_PATH, INPUT_EXIST)
End Sub

Public Function WriteFile() As Boolean
    WriteFile = False
    lineStr = ""
    
    If IsInput = False Then Exit Function
End Function

Private Function IsInput() As Boolean
    Select Case fileMode
        Case INPUT_EXIST
            IsInput = True
        Case Else
            IsInput = False
    End Select
End Function

Property Get GET_LINE() As String
    GET_LINE = lineStr
End Property

