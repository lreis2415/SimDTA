VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Const TABChar = Chr(9)
Const SPACEChar = " "
Const COMMAChar = ","

Public FileName As String

' Reference: Microsoft Scripting Runtime
Private fs As New FileSystemObject
Private ts As TextStream

Private sLine As String
Private lLineIndex As Long

Private colText As New Collection
Private lLineCount As Long
Private sSplitChar As String
Private arrCols() As String


Public Property Let SplitChar(ByVal sChar As String)
    If Len(sChar) = 1 Then
        sSplitChar = sChar
    Else
        MsgBox "Wrong split char: " & sChar
    End If
End Property
Public Property Get SplitChar() As String
    SplitChar = sSplitChar
End Property

Public Property Get LastReadLine() As String
    LastReadLine = sLine
End Property

Public Function GetCols(Cols() As String) As Boolean
On Error GoTo ErrHandler
    Cols = Split(sLine, sSplitChar)
    GetCols = True
    Exit Function
ErrHandler:
    GetCols = False
End Function

Public Function GetCol(Col As String, ByVal ColIndex As Integer) As Boolean
On Error GoTo ErrHandler
    Col = arrCols(ColIndex)
    GetCol = True
    Exit Function
ErrHandler:
    GetCol = False
End Function

Public Function ReadLine(line As String) As Boolean
On Error GoTo ErrHandler
    If lLineIndex = -1 Then
        line = ""
        ReadLine = False
        Exit Function
    End If
    If ts.AtEndOfStream Then
        line = ""
        ReadLine = False
    Else
        line = ts.ReadLine
        lLineIndex = lLineIndex + 1
        lLineCount = lLineCount + 1
        sLine = line
        arrCols = Split(sLine, sSplitChar)
        ReadLine = True
    End If
    Exit Function
ErrHandler:
    MsgBox "Error: readline function"
    line = ""
    ReadLine = False
End Function

Public Function OpenFile() As Boolean
On Error GoTo ErrHandler
    If FileName = "" Then
        MsgBox "Set file name firstly."
        OpenFile = False
        Exit Function
    End If
    Set ts = fs.OpenTextFile(FileName, ForReading, False, TristateUseDefault)
    lLineCount = 0
    lLineIndex = 0
    OpenFile = True
    Exit Function
ErrHandler:
    MsgBox "Error: open file " & FileName
    On Error Resume Next
    'ts.Close
    OpenFile = False
End Function

Public Function CloseFile() As Boolean
On Error GoTo ErrHandler
    If lLineCount = -1 Then
        MsgBox "Not file open."
        CloseFile = False
        Exit Function
    End If
    ts.Close
    lLineCount = -1
    lLineIndex = -1
    sLine = ""
    Set colText = Nothing
    CloseFile = True
    Exit Function
ErrHandler:
    MsgBox "Error: close file " & FileName
    CloseFile = False
End Function

Private Sub Class_Initialize()
    FileName = ""
    sLine = ""
    lLineIndex = -1
    lLineCount = -1
    sSplitChar = SPACEChar
End Sub

'close log file
Private Sub Class_Terminate()
   
    If lLineCount > -1 Then ts.Close
    Set fs = Nothing
End Sub

