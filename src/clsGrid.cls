VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Const TABChar = Chr(9)
'Const SPACEChar = " "
'Const COMMAChar = ","
Const C_FILEHEAD_NCOLS = "ncols"
Const C_FILEHEAD_NROWS = "nrows"
Const C_FILEHEAD_XLLCORNER = "xllcorner"
Const C_FILEHEAD_YLLCORNER = "yllcorner"
Const C_FILEHEAD_CELLSIZE = "cellsize"
Const C_FILEHEAD_NODATA = "NODATA_value"

' Asc file data: nrows (y) * ncols (x);
'  xllcorner: x left-low , to right
'  yllcorner: y left-low, to up
'  Data (icol=0, irow=0) -> left-up corner
'  (xllcorner, yllcorner) -> data(iCol=0, iRow=nrows-1)

Public sAscGridFileName As String

Public nCols As Integer, nRows As Integer
Public xllcorner As Double, yllcorner As Double
Public CellSize As Double
Public NoData_Value As Double
Public vData As Variant

Private m_dMax As Double, m_dMin As Double
Private m_blnDataTypeIsInt As Boolean

Public Function NearNoDataCell(iCol As Integer, iRow As Integer) As Boolean
   Dim iCol1 As Integer, iRow1 As Integer
   NearNoDataCell = False
   If iCol = 0 Or iCol = nCols - 1 Or iRow = 0 Or iRow = nRows - 1 Then
      NearNoDataCell = True
      Exit Function
   Else
      For iCol1 = iCol - 1 To iCol + 1
         For iRow1 = iRow - 1 To iRow + 1
            If vData(iCol, iRow) = NoData_Value Then
               NearNoDataCell = True
               Exit Function
            End If
         Next
      Next
   End If
End Function

Public Property Let Cell(iCol As Integer, iRow As Integer, dValue As Double)
   If m_blnDataTypeIsInt Then
      vData(iCol, iRow) = Int(dValue)
   Else
      vData(iCol, iRow) = dValue
   End If
End Property

Public Property Get Cell(iCol As Integer, iRow As Integer) As Double
   Cell = vData(iCol, iRow)
End Property

Public Property Get Maximum() As Double
   Dim iCol As Integer, iRow As Integer
   
   If m_dMax = NoData_Value Then
      m_dMax = MIN_SINGLE
      For iCol = 0 To nCols - 1
         For iRow = 0 To nRows - 1
            If vData(iCol, iRow) <> NoData_Value Then
               If m_dMax < vData(iCol, iRow) Then m_dMax = vData(iCol, iRow)
            End If
         Next
      Next
      If m_dMax = MIN_SINGLE Then m_dMax = NoData_Value
   End If
   Maximum = m_dMax
End Property

Public Property Get Minimum() As Double
   Dim iCol As Integer, iRow As Integer
   
   If m_dMin = NoData_Value Then
      m_dMin = MAX_SINGLE
      For iCol = 0 To nCols - 1
         For iRow = 0 To nRows - 1
            If vData(iCol, iRow) <> NoData_Value Then
               If m_dMin > vData(iCol, iRow) Then m_dMin = vData(iCol, iRow)
            End If
         Next
      Next
      If m_dMin = MAX_SINGLE Then m_dMin = NoData_Value
   End If
   Minimum = m_dMin
End Property

Public Function IsValidCell(iCol As Integer, iRow As Integer) As Boolean
   If iCol >= 0 And iCol < nCols And iRow >= 0 And iRow < nRows Then
      IsValidCell = True
   Else
      IsValidCell = False
   End If
End Function

Public Function IsValidCellValue(iCol As Integer, iRow As Integer, Optional dValue As Double) As Boolean
   If iCol >= 0 And iCol < nCols And iRow >= 0 And iRow < nRows Then
      dValue = vData(iCol, iRow)
      IsValidCellValue = IIf((dValue = NoData_Value), False, True)
   Else
      IsValidCellValue = False
   End If
End Function

Private Sub Class_Initialize()
    '
    sAscGridFileName = ""
    nCols = 0: nRows = 0
    xllcorner = 0#: yllcorner = 0#: CellSize = 1#: NoData_Value = -9999
    
    InitGRIDStatistics
End Sub

Public Function NewGrid(Optional iCols As Integer = 0, Optional iRows As Integer = 0, _
                        Optional xll As Double = MIN_SINGLE, Optional yll As Double = MIN_SINGLE, _
                        Optional dCell As Double = 0#, Optional dNoData As Double = MIN_SINGLE, _
                        Optional dInitValue As Double = MIN_SINGLE, Optional blnDataTypeIsInt As Boolean = False) As Boolean
   Dim iRow As Integer, iCol As Integer
On Error GoTo ErrH
   NewGrid = False
   If iCols < 0 Or iRows < 0 Or dCell < 0 Then
      MsgBox "wrong parameters for a new grid!"
      Exit Function
   End If
   If iCols > 0 Then nCols = iCols
   If iRows > 0 Then nRows = iRows
   If nCols <= 0 Or nRows <= 0 Then
      MsgBox "need set parameters for a new grid!"
      Exit Function
   End If
   If xll > MIN_SINGLE Then xllcorner = xll
   If yll > MIN_SINGLE Then yllcorner = yll
   If dCell > 0 Then CellSize = dCell
   If dNoData > MIN_SINGLE Then NoData_Value = dNoData
   
   m_blnDataTypeIsInt = blnDataTypeIsInt
   If m_blnDataTypeIsInt Then
      ReDim vData(0 To nCols - 1, 0 To nRows - 1) As Integer
   Else
      ReDim vData(0 To nCols - 1, 0 To nRows - 1) As Double
   End If
   
   If dInitValue > MIN_SINGLE Then
      If m_blnDataTypeIsInt Then
         For iCol = 0 To nCols - 1
            For iRow = 0 To nRows - 1
               vData(iCol, iRow) = Int(dInitValue)
            Next
         Next
      Else
         For iCol = 0 To nCols - 1
            For iRow = 0 To nRows - 1
               vData(iCol, iRow) = dInitValue
            Next
         Next
      End If
   End If
   
   InitGRIDStatistics
   NewGrid = True
   Exit Function
ErrH:
   'MsgBox "NewGrid Error!"
End Function

'
'When bOpenDialog = TRUE, use file dialog to create folder and overwrite file;
'if else, without query for creating folder and overwriting file
'
Public Function SaveAscGrid(Optional sAscFile As String = "", Optional bOpenDialog As Boolean = False, Optional iKeepDotNum As Integer = 3) As Boolean
On Error GoTo ErrH
   ' Reference: Microsoft Scripting Runtime
   Dim fs As New FileSystemObject
   Dim ts As TextStream
   Dim sTemp As String, strFormat As String, i As Integer
   
   SaveAscGrid = False
   If bOpenDialog Then      'open with FileDialog
      'sAscFile
   End If
   If sAscFile <> "" Then sAscGridFileName = sAscFile
   If sAscGridFileName = "" Then
        MsgBox "Set file name firstly."
        Exit Function
   End If
   If Not bOpenDialog Then   'create folder (if it doesn't exist) and overwrite file (if it exists), without query
      sTemp = fs.GetFileName(sAscGridFileName)
      sTemp = Left(sAscGridFileName, Len(sAscGridFileName) - Len(sTemp))
      If Not fs.FolderExists(sTemp) Then fs.CreateFolder sTemp
   Else
      If fs.FileExists(sAscGridFileName) Then
         If MsgBox(sAscGridFileName & vbCrLf & "exists. Overwrite it?", vbQuestion + vbYesNo + vbDefaultButton2, "Overwrite file?") = vbNo Then Exit Function
      End If
   End If
   Set ts = fs.OpenTextFile(sAscGridFileName, ForWriting, True, TristateUseDefault)
   
   Dim iRow As Integer '-> y
   Dim iCol As Integer '-> x
   
   If iKeepDotNum < 0 Then
      strFormat = ""
   ElseIf iKeepDotNum = 0 Then
      strFormat = "0"
   Else
      strFormat = "0." & String(iKeepDotNum, "0")
   End If
   
   ts.WriteLine C_FILEHEAD_NCOLS & " " & nCols
   ts.WriteLine C_FILEHEAD_NROWS & " " & nRows
   ts.WriteLine C_FILEHEAD_XLLCORNER & " " & xllcorner
   ts.WriteLine C_FILEHEAD_YLLCORNER & " " & yllcorner
   ts.WriteLine C_FILEHEAD_CELLSIZE & " " & CellSize
   ts.WriteLine C_FILEHEAD_NODATA & " " & NoData_Value
   If m_blnDataTypeIsInt Then
      For iRow = 0 To nRows - 1
         For iCol = 0 To nCols - 2
            ts.Write CStr(vData(iCol, iRow)) & " "
         Next
         ts.WriteLine CStr(vData(nCols - 1, iRow))
         DoEvents
      Next
   Else
'      If iKeepDotNum < 0 Then
'         For iRow = 0 To nRows - 1
'            For iCol = 0 To nCols - 2
'               ts.Write vData(iCol, iRow) & " "
'            Next
'            ts.WriteLine vData(nCols - 1, iRow)
'            DoEvents
'         Next
'      Else
         For iRow = 0 To nRows - 1
            For iCol = 0 To nCols - 2
               ts.Write Format(vData(iCol, iRow), strFormat) & " "
            Next
            ts.WriteLine Format(vData(nCols - 1, iRow), strFormat)
            DoEvents
         Next
'      End If
   End If
      
   ts.Close
   Set fs = Nothing
   SaveAscGrid = True
   Exit Function
ErrH:
   On Error Resume Next
   ts.Close
   Set fs = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


Public Function LoadAscGrid(Optional sAscFile As String = "", Optional blnDataTypeIsInt As Boolean = False, _
                           Optional bOpenDialog As Boolean = False) As Boolean
On Error GoTo ErrH
   LoadAscGrid = False
   If bOpenDialog Then      'open with FileDialog
      'sAscFile
   End If
   
   Dim pReadFile As New clsReadFile
   Dim iRow As Integer '-> y
   Dim iCol As Integer '-> x
   Dim s As String
   Dim bool As Boolean
   Dim arrCol() As String
   
   With pReadFile
      If sAscFile = "" Then
         .FileName = sAscGridFileName
      Else
         .FileName = sAscFile
      End If
      .OpenFile
      
      'Asc GRID file head: ncols, ...
      For iRow = 1 To 6
         If .ReadLine(s) Then
            If .GetCols(arrCol) Then
               Select Case arrCol(0)
               Case C_FILEHEAD_NCOLS
                  nCols = CInt(arrCol(UBound(arrCol)))
               Case C_FILEHEAD_NROWS
                  nRows = CInt(arrCol(UBound(arrCol)))
               Case C_FILEHEAD_XLLCORNER
                  xllcorner = CDbl(arrCol(UBound(arrCol)))
               Case C_FILEHEAD_YLLCORNER
                  yllcorner = CDbl(arrCol(UBound(arrCol)))
               Case C_FILEHEAD_CELLSIZE
                  CellSize = CDbl(arrCol(UBound(arrCol)))
               Case C_FILEHEAD_NODATA
                  NoData_Value = CDbl(arrCol(UBound(arrCol)))
               End Select
            End If
         End If
      Next
      
      m_blnDataTypeIsInt = blnDataTypeIsInt
      If m_blnDataTypeIsInt Then
         ReDim vData(0 To nCols - 1, 0 To nRows - 1) As Integer
      Else
         ReDim vData(0 To nCols - 1, 0 To nRows - 1) As Double
      End If
      If m_blnDataTypeIsInt Then
         For iRow = 0 To nRows - 1
            .ReadLine (s)
            If .GetCols(arrCol) Then
               For iCol = 0 To nCols - 1
                  vData(iCol, iRow) = CInt(arrCol(iCol))
               Next
            End If
            DoEvents
         Next
      Else
         For iRow = 0 To nRows - 1
            .ReadLine (s)
            If .GetCols(arrCol) Then
               For iCol = 0 To nCols - 1
                  vData(iCol, iRow) = CDbl(arrCol(iCol))
               Next
            End If
            DoEvents
         Next
      End If
      .CloseFile
   End With
   
   Set pReadFile = Nothing
   If sAscFile <> "" Then sAscGridFileName = sAscFile
   InitGRIDStatistics
   LoadAscGrid = True
   Exit Function
ErrH:
   '数据可能已经覆盖，可通过reload sAscGridFileName文件来重新获取原来的可用数据
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   On Error Resume Next
   pReadFile.CloseFile
   Set pReadFile = Nothing
End Function

Private Sub Class_Terminate()
    '
    vData = Empty
    'If lLineCount > -1 Then ts.Close
    'Set fs = Nothing
End Sub

Private Sub InitGRIDStatistics()
   m_dMax = NoData_Value
   m_dMin = NoData_Value
End Sub
