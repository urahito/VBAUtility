VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private filePath As String
Private dirPath As String
Private fileNameWithExt As String
Private fileNameNoExt As String
Private extentionStr As String

Private strAttr As String

Private Sub Class_Initialize()
    Call InitPaths
End Sub

Public Sub InitPaths()
    Call SetPaths("")
End Sub

Public Function GetFileSuccess() As Boolean
    Dim OpenFileName As Variant
    OpenFileName = Application.GetOpenFilename("csvファイル,*.csv")
    OpenFileName = IIf(OpenFileName = False, "", OpenFileName)
    GetFileSuccess = SetPathsSuccess(CStr(OpenFileName))
End Function

Public Function SetPathsSuccess(ByVal pathStr As String) As Boolean
    SetPathsSuccess = SetPaths(pathStr) And CAN_READ
End Function

Public Function SetPaths(ByVal pathStr As String) As Boolean
    On Error GoTo Init
    Dim dirPos As Integer
    Dim extPos As Integer
    SetPaths = False
    
    dirPos = InStrRev(pathStr, "\")

    filePath = pathStr
    dirPath = Left(pathStr, dirPos)
    fileNameWithExt = Mid(filePath, dirPos + 1)
    
    extPos = InStrRev(fileNameWithExt, ".")
    extentionStr = Mid(fileNameWithExt, extPos)
    fileNameNoExt = Left(fileNameWithExt, extPos - 1)
    
    SetPaths = EXIST_FILE
    Exit Function
Init:
    If pathStr = "" Then Exit Function
    Call SetPaths("")
End Function

Private Function GetFileAttr() As Boolean
    Dim intAttr As Integer
    strAttr = ""
    
    ' ﾌｧｲﾙ属性の取得
    intAttr = GetAttr(filePath)                      ' ①
    
    ' ﾌｧｲﾙ属性の判定
    If (intAttr And vbReadOnly) <> 0 Then               ' ②
        strAttr = strAttr & vbCr & "読み取り専用"
    End If
    If (intAttr And vbHidden) <> 0 Then
        strAttr = strAttr & vbCr & "隠しファイル"
    End If
    If (intAttr And vbSystem) <> 0 Then
        strAttr = strAttr & vbCr & "システムファイル"
    End If
    If (intAttr And vbDirectory) <> 0 Then
        strAttr = strAttr & vbCr & "フォルダ"
    End If
    If (intAttr And vbArchive) <> 0 Then
        strAttr = strAttr & vbCr & "アーカイブ"
    End If
    If (intAttr And vbNormal) <> 0 Then
        strAttr = strAttr & vbCr & "通常ファイル"
    End If
    
    Debug.Print strAttr
    
    GetFileAttr = (Len(strAttr) > 0)
End Function

Property Get CAN_READ() As Boolean
    CAN_READ = IncludeStrs(strAttr, "アーカイブ", "通常ファイル")
End Property

Property Get GET_FULL_PATH() As String
    GET_FULL_PATH = filePath
End Property

Property Get GET_FULL_PATH_NO_EXT() As String
    GET_FULL_PATH_NO_EXT = dirPath & fileNameNoExt
End Property

Property Get DIR_PATH() As String
    DIR_PATH = dirPath
End Property

Property Get FILE_NAME() As String
    FILE_NAME = fileNameWithExt
End Property

Property Get FILE_NAME_NO_EXT() As String
    FILE_NAME_NO_EXT = fileNameNoExt
End Property

Property Get FILE_EXT() As String
    FILE_EXT = extentionStr
End Property

Property Get EXIST_FILE() As Boolean
    EXIST_FILE = False
    
    If Dir(filePath) = "" Then Exit Property
    
    EXIST_FILE = GetFileAttr()
End Property

Property Get EXIST_DIR(Optional ByVal subDir As String = "") As Boolean
    EXIST_DIR = False
    
    If Dir(dirPath & subDir) = "" Then Exit Property
    
    EXIST_DIR = True
End Property


