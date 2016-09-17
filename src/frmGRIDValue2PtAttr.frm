VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmGRIDValue2PtAttr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GRID value -> Point attribute"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame frameSrc 
      Caption         =   "Source"
      Height          =   3375
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   11055
      Begin VB.ComboBox cboSplit 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   2940
         Width           =   2235
      End
      Begin VB.TextBox txtSrcPtFileHead 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   25
         Top             =   2520
         Width           =   9615
      End
      Begin VB.CommandButton cmdSrcPtFile 
         Caption         =   "Src Point File"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSrcPtFile 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   22
         Top             =   1800
         Width           =   9615
      End
      Begin VB.CommandButton cmdSrcGRID 
         Caption         =   "Src GRID"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSrcGRID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   9615
      End
      Begin VB.Frame frameFileHead 
         Caption         =   "File Head"
         Enabled         =   0   'False
         Height          =   675
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   10815
         Begin VB.TextBox txtNoData 
            Height          =   315
            Left            =   10020
            TabIndex        =   12
            Text            =   "-9999"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCellSize 
            Height          =   315
            Left            =   8100
            TabIndex        =   11
            Text            =   "1"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtYll 
            Height          =   315
            Left            =   5880
            TabIndex        =   10
            Text            =   "0"
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox txtXll 
            Height          =   315
            Left            =   3660
            TabIndex        =   9
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtRows 
            Height          =   315
            Left            =   2040
            TabIndex        =   8
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtCols 
            Height          =   315
            Left            =   600
            TabIndex        =   7
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "NoData_Value"
            Height          =   315
            Index           =   84
            Left            =   8880
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "CellSize"
            Height          =   315
            Index           =   85
            Left            =   7380
            TabIndex        =   17
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "YllCorner"
            Height          =   315
            Index           =   86
            Left            =   5160
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "XllCorner"
            Height          =   315
            Index           =   87
            Left            =   2880
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "nRows"
            Height          =   315
            Index           =   88
            Left            =   1500
            TabIndex        =   14
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "nCols"
            Height          =   315
            Index           =   89
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Split Char:"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   2940
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "File head of ""Src Point File"": ""[X]""  ""[Y]""  ""[Attr]""  ..."
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   2220
         Width           =   7035
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   9060
      TabIndex        =   4
      Top             =   5040
      Width           =   1875
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output Point File"
      Height          =   1095
      Left            =   60
      TabIndex        =   1
      Top             =   3420
      Width           =   11055
      Begin VB.TextBox txtSaveGRID 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   9615
      End
      Begin VB.CommandButton cmdSaveGRID 
         Caption         =   "Save File..."
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Add Point Attribute with GRID Value"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   5040
      Width           =   5235
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   420
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   60
      TabIndex        =   21
      Top             =   4500
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmGRIDValue2PtAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_Split_Comma = "Comma"
Const C_Split_Space = "Space"
Const C_Split_TAB = "TAB"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_bRunning As Boolean
Dim m_strBasePath As String
Dim m_strFilePre As String
Dim m_pBaseGRID As clsGrid

Private Sub cmdRun_Click()
On Error GoTo ErrH
   Dim strSaveFile As String, strPtFile As String
   Dim pReadFile As clsReadFile
   Dim arrCol() As String, strLine As String
   Dim fs As FileSystemObject
   Dim ts As TextStream
   Dim strSplit As String
   Dim pSrcGRID As clsGrid
   Dim dX As Double, dY As Double
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, dCol As Double, dRow As Double
   Dim dValue As Double
   Dim lCount As Long, i As Integer, j As Integer
      
   If m_bRunning Then Exit Sub
         
   ' get parameters
   If m_pBaseGRID Is Nothing Then
      MsgBox "Assign the INPUT file name first", vbInformation, APP_TITLE
      Exit Sub
   End If
   strPtFile = Trim(txtSrcPtFile.Text)
   If strPtFile = "" Then
      MsgBox "Assign the Src Point file name first", vbInformation, APP_TITLE
      Exit Sub
   End If
   strSaveFile = Trim(txtSaveGRID.Text)
   If strSaveFile = "" Then
      MsgBox "Assign the OUTPUT file name first", vbInformation, APP_TITLE
      Exit Sub
   End If
   ' prepare for writing output file
   Set fs = New FileSystemObject
   If fs.FileExists(strSaveFile) Then
      If MsgBox(strPtFile & vbCrLf & "exists. Overwrite it?", vbQuestion + vbYesNo + vbDefaultButton2, "Overwrite file?") = vbNo Then
         Set fs = Nothing
         Exit Sub
      End If
   End If
   Set ts = fs.OpenTextFile(strSaveFile, ForWriting, True, TristateUseDefault)
   Select Case cboSplit.Text
   Case C_Split_Comma
      strSplit = ","
   Case C_Split_Space
      strSplit = " "
   Case C_Split_TAB
      strSplit = Chr(9)
   End Select
     
   m_bRunning = True
   Me.MousePointer = 11
   '
   With m_pBaseGRID
      iCols = .nCols: iRows = .nRows
      dXll = .xllcorner: dYll = .yllcorner
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   Set pSrcGRID = m_pBaseGRID
         
   Set pReadFile = New clsReadFile
   With pReadFile
      .FileName = strPtFile
      .OpenFile
      .SplitChar = strSplit
      
      .ReadLine strLine
      lCount = 0
      strLine = strLine & strSplit & """"
      strLine = strLine & m_strFilePre & """"
      ts.WriteLine strLine
      While .ReadLine(strLine)
         lCount = lCount + 1
         If .GetCols(arrCol) Then
            For i = LBound(arrCol) To UBound(arrCol)
               If IsNumeric(arrCol(i)) Then
                  dX = CDbl(arrCol(i))
                  For j = i + 1 To UBound(arrCol)
                     If IsNumeric(arrCol(j)) Then
                        dY = CDbl(arrCol(j))
                        dCol = (dX - dXll) / dCellSize
                        dRow = iRows - (dY - dYll) / dCellSize
                        If dCol < 0 Or dCol > iCols Then
                           iCol = -1
                        Else
                           iCol = Int(dCol)
                        End If
                        If dRow < 0 Or dRow > iRows Then
                           iRow = -1
                        Else
                           iRow = Int(dRow)
                        End If
                        If pSrcGRID.IsValidCellValue(iCol, iRow, dValue) Then
                           ts.WriteLine strLine & strSplit & dValue
                        Else
                           ts.WriteLine strLine & strSplit & dNoData
                        End If
                        Exit For
                     End If
                  Next
                  Exit For
               End If
            Next
         End If
         SetProgressBarValue (lCount Mod 50) * 2
      Wend
      .CloseFile
   End With
   
   SetProgressBarValue 100
   MsgBox "Completed. Save result file: " & vbCrLf & strSaveFile, vbInformation, APP_TITLE
   
ErrH:
   Me.MousePointer = 0
   m_bRunning = False
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   On Error Resume Next
   ts.Close
   Set fs = Nothing
   SetProgressBarValue 0
End Sub

Private Sub cmdQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub

Private Sub cmdSaveGRID_Click()
   comdlg.DialogTitle = "Save GRID"
   comdlg.FileName = ""
   txtSaveGRID.Text = GetFileName(comdlg, False, , ".asc")
End Sub

Private Sub cmdSrcGRID_Click()
   Dim strBaseDEM As String
On Error GoTo ErrH
   If m_bRunning Then Exit Sub
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   strBaseDEM = GetFileName(comdlg, True, , ".asc")
   If strBaseDEM = "" Then Exit Sub
   
   Dim strPath As String, strName As String, strSuffix As String
   Dim i As Integer
    
   Me.MousePointer = 11
   m_bRunning = True
   
   i = InStrRev(strBaseDEM, "\")
   m_strBasePath = Left(strBaseDEM, i)
   strName = Right(strBaseDEM, Len(strBaseDEM) - i)
   i = InStrRev(strName, ".")
   If i = 0 Then
      m_strFilePre = strName
   Else
      m_strFilePre = Left(strName, i - 1)
   End If
   
   txtSrcGRID.Text = strBaseDEM
   
   ' load BaseDEM, read parameters in file head
   If Not (m_pBaseGRID Is Nothing) Then Set m_pBaseGRID = Nothing
   Set m_pBaseGRID = New clsGrid
   With m_pBaseGRID
      .LoadAscGrid strBaseDEM
      txtCols.Text = .nCols
      txtRows.Text = .nRows
      txtXll.Text = .xllcorner
      txtYll.Text = .yllcorner
      txtCellSize.Text = .CellSize
      txtNoData.Text = .NoData_Value
   End With
   
   'txtSaveGRID.Text = m_strBasePath & m_strFilePre & "_FillNone.asc"
ErrH:
   Me.MousePointer = 0
   m_bRunning = False
End Sub

Private Sub cmdSrcPtFile_Click()
   Dim strPtFile As String
   
On Error GoTo ErrH
   If m_bRunning Then Exit Sub
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   strPtFile = GetFileName(comdlg, True, , ".txt")
   If strPtFile = "" Then Exit Sub
   
   Dim strPath As String, strName As String, strSuffix As String, strFilePre As String
   Dim i As Integer, str As String
    
   Me.MousePointer = 11
   m_bRunning = True
   
   i = InStrRev(strPtFile, "\")
   strPath = Left(strPtFile, i)
   strName = Right(strPtFile, Len(strPtFile) - i)
   i = InStrRev(strName, ".")
   If i = 0 Then
      strFilePre = strName
      strSuffix = ""
   Else
      strFilePre = Left(strName, i - 1)
      strSuffix = Right(strName, Len(strName) - i + 1)
   End If
   
   txtSrcPtFile.Text = strPtFile
   
   ' read file head
   Dim pReadFile As New clsReadFile
   With pReadFile
      .FileName = strPtFile
      .OpenFile
      .ReadLine str
      .CloseFile
   End With
   
   txtSrcPtFileHead.Text = str
   txtSaveGRID.Text = strPath & strFilePre & "_Attr" & strSuffix
ErrH:
   Me.MousePointer = 0
   m_bRunning = False
   Set pReadFile = Nothing
End Sub

Private Sub Form_Load()
   'initialize var
   m_bRunning = False
   Set m_pBaseGRID = Nothing
   SetProgressBarValue 0
   txtSrcPtFileHead.Text = ""
   With cboSplit
      .Clear
      .AddItem C_Split_Comma
      .AddItem C_Split_Space
      .AddItem C_Split_TAB
      .ListIndex = 0
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bRunning Then
      Cancel = 1
      Exit Sub
   End If
   If Not (m_pBaseGRID Is Nothing) Then Set m_pBaseGRID = Nothing
End Sub

Private Sub SetProgressBarValue(iValue As Integer)
   If iValue > 100 Or iValue < 0 Then Exit Sub
   With progbar
      .Value = iValue
      .Refresh
   End With
End Sub


