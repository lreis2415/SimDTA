VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSlopeLength 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Slope Length"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cdmQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   6420
      Width           =   1935
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   6420
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTabFunc 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Slope Length"
      TabPicture(0)   =   "frmSlopeLength.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "framePara(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameOutput(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frameInput(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2535
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   780
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   0
               Left            =   600
               TabIndex        =   24
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   0
               Left            =   2160
               TabIndex        =   23
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   0
               Left            =   4140
               TabIndex        =   22
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   0
               Left            =   6180
               TabIndex        =   21
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   0
               Left            =   8100
               TabIndex        =   20
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   0
               Left            =   10020
               TabIndex        =   19
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   29
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   2
               Left            =   3240
               TabIndex        =   28
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   3
               Left            =   5280
               TabIndex        =   27
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   4
               Left            =   7320
               TabIndex        =   26
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   5
               Left            =   8880
               TabIndex        =   25
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID0 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1380
            TabIndex        =   17
            Top             =   360
            Width           =   9555
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "DEM"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1275
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "ArcInfo FlowDir"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   1500
            Width           =   1275
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "Valley"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   1860
            Width           =   1275
         End
         Begin VB.TextBox txtSrcGRID0 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1380
            TabIndex        =   13
            Top             =   1500
            Width           =   9555
         End
         Begin VB.TextBox txtSrcGRID0 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1380
            TabIndex        =   12
            Top             =   1860
            Width           =   9555
         End
         Begin VB.Label Label2 
            Caption         =   "(Default: No Valley Grid Assigned)"
            Height          =   255
            Index           =   1
            Left            =   1380
            TabIndex        =   33
            Top             =   2220
            Width           =   4995
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1395
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   4260
         Width           =   11055
         Begin VB.TextBox txtSaveGRID0 
            Height          =   375
            Index           =   0
            Left            =   1380
            TabIndex        =   10
            Top             =   360
            Width           =   9555
         End
         Begin VB.CommandButton cmdSaveGRID0 
            Caption         =   "Slope Length"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtSaveGRID0 
            Height          =   375
            Index           =   1
            Left            =   1380
            TabIndex        =   8
            Top             =   780
            Width           =   9555
         End
         Begin VB.CommandButton cmdSaveGRID0 
            Caption         =   "Upslope Cells"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   780
            Width           =   1275
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   3300
         Width           =   11055
         Begin VB.TextBox txtValleyTag 
            Height          =   315
            Index           =   0
            Left            =   1380
            TabIndex        =   4
            Text            =   "1"
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Valley Tag"
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   "Calculate Slope Length by reversed single flow direction"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Top             =   420
         Width           =   10935
      End
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   960
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   0
      TabIndex        =   32
      Top             =   5760
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSlopeLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_VALLEY = 1

Dim m_bRunning As Boolean

Dim mpDEM As clsGrid, mpFlowDir As clsGrid, mpValley As clsGrid
Dim mpUpslpCells As clsGrid, mpSlpLen As clsGrid   ', mpUpslpDir As clsGrid

Dim miRows As Integer, miCols As Integer, mdCell As Double
Dim mpTag As clsGrid

Dim miVlyTag As Integer
Dim msDEM As String, msFlowDir As String, msVly As String
Dim msUpslpCells As String, msSlpLen As String, msUpslpDir As String

Private Function SlpLen_SearchUpslope_bySFD() As Boolean
   Dim iRow As Integer, iCol As Integer
On Error GoTo ErrH
   SlpLen_SearchUpslope_bySFD = False
   
   Set mpDEM = New clsGrid
   mpDEM.LoadAscGrid msDEM
   Set mpFlowDir = New clsGrid
   mpFlowDir.LoadAscGrid msFlowDir, True
   
   With mpDEM
      miRows = .nRows: miCols = .nCols: mdCell = .CellSize
      
      Set mpValley = New clsGrid
      If msVly = "" Then
         mpValley.NewGrid miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True
      Else
         mpValley.LoadAscGrid msVly, True
      End If
      
      If miRows <> mpFlowDir.nRows Or miCols <> mpFlowDir.nCols _
            Or miRows <> mpValley.nRows Or miCols <> mpValley.nCols Then
         Err.Raise Number:=vbObjectError + 513, Description:="INPUT GRID are not with same size"
      End If
      
      Set mpUpslpCells = New clsGrid
      If Not mpUpslpCells.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpSlpLen = New clsGrid
      If Not mpSlpLen.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
'      Set mpUpslpDir = New clsGrid
'      If Not mpUpslpDir.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
'         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
'      End If
      
      Set mpTag = New clsGrid
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, 0, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   
   '''''''''''''''''''''''''
   ' search upslope
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = 0 Then SearchUpslope_bySFD iCol, iRow
         DoEvents
      Next
      SetProgressBarValue Int((iRow + 1) * 100# / miRows)
      DoEvents
   Next
   
   ''''''''''''''''''''
   ' output para
   If msUpslpCells <> "" Then
      If mpUpslpCells.SaveAscGrid(msUpslpCells) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpCells
      End If
   End If
   If msSlpLen <> "" Then
      If mpSlpLen.SaveAscGrid(msSlpLen) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msSlpLen
      End If
   End If
'   If msUpslpDir <> "" Then
'      If mpUpslpDir.SaveAscGrid(msUpslpDir) Then
'         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
'      Else
'         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpDir
'      End If
'   End If
   
   SlpLen_SearchUpslope_bySFD = True
ErrH:
   If Not SlpLen_SearchUpslope_bySFD Then
      If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   End If
   On Error Resume Next
   CleanUpMemory
End Function

'
Private Function SearchUpslope_bySFD(iProcCol As Integer, iProcRow As Integer) As Boolean
   Dim dElev As Double, dTemp As Double
   Dim iTempRow As Integer, iTempCol As Integer
   Dim iDir As Integer, iDirTemp As Integer, iFlowDir As Integer
   Dim dVal As Double
   Dim iUpslpCount As Integer
   
   SearchUpslope_bySFD = False
   
   dElev = mpDEM.Cell(iProcCol, iProcRow)
   If dElev = mpDEM.NoData_Value Then
      mpUpslpCells.Cell(iProcCol, iProcRow) = mpUpslpCells.NoData_Value
      mpSlpLen.Cell(iProcCol, iProcRow) = mpSlpLen.NoData_Value
'      mpUpslpDir.Cell(iProcCol, iProcRow) = mpUpslpDir.NoData_Value
   Else
      iUpslpCount = 0
      For iDir = 1 To DIRNUM8
         iTempCol = iProcCol + ArrDir8X(iDir): iTempRow = iProcRow + ArrDir8Y(iDir)
         If mpFlowDir.IsValidCellValue(iTempCol, iTempRow, dVal) Then
            If mpValley.Cell(iTempCol, iTempRow) <> miVlyTag Then
               iFlowDir = dVal
               iDirTemp = GetESRIDir_ArrayIndex(iFlowDir)
               If iDirTemp > 0 Then
                  If iTempCol + ArrDir8X(iDirTemp) = iProcCol And iTempRow + ArrDir8Y(iDirTemp) = iProcRow Then
                  ' flow dir is (iTempCol, iTempRow) -> (iProcCol, iProcRow)
                     If mpTag.Cell(iTempCol, iTempRow) = 0 Then
                        SearchUpslope_bySFD iTempCol, iTempRow
                     End If
                     
                     If mpSlpLen.Cell(iTempCol, iTempRow) >= 0 And mpDEM.Cell(iTempCol, iTempRow) <> mpDEM.NoData_Value Then
                        If (iDir Mod 2) = 0 Then
                           dTemp = mpSlpLen.Cell(iTempCol, iTempRow) + mdCell
                        Else
                           dTemp = mpSlpLen.Cell(iTempCol, iTempRow) + mdCell * SQRT2
                        End If
                        If mpSlpLen.Cell(iProcCol, iProcRow) < dTemp Then
                           mpSlpLen.Cell(iProcCol, iProcRow) = dTemp
                           mpUpslpCells.Cell(iProcCol, iProcRow) = mpUpslpCells.Cell(iTempCol, iTempRow) + 1
'                           mpUpslpDir.Cell(iProcCol, iProcRow) = ESRIDir(iDir)
                           iUpslpCount = iUpslpCount + 1
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Next
      
      If iUpslpCount = 0 Then
         mpUpslpCells.Cell(iProcCol, iProcRow) = 0
         mpSlpLen.Cell(iProcCol, iProcRow) = 0#
'         mpUpslpDir.Cell(iProcCol, iProcRow) = ESRI_DIR_UNDEF
      End If
   End If
   
   mpTag.Cell(iProcCol, iProcRow) = 1
   SearchUpslope_bySFD = True
End Function

Private Sub cdmQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub

Private Function GetSaveFileName() As String
   comdlg.DialogTitle = "Save GRID"
   comdlg.FileName = ""
   GetSaveFileName = GetFileName(comdlg, False, , ".asc")
End Function

Public Sub DTAFunc(iFuncIndex As Integer)
   Dim i As Integer
   With SSTabFunc
      For i = 0 To .Tabs - 1
         .TabEnabled(i) = False
      Next
      
      If iFuncIndex >= 0 And iFuncIndex < .Tabs Then
         .Tab = iFuncIndex
         .TabEnabled(iFuncIndex) = True
      Else
         For i = 0 To .Tabs - 1
            .TabEnabled(i) = True
         Next
      End If
   End With
End Sub

' verify the input parameters, then call function
Private Sub cmdRun_Click()
   Dim i As Integer
   Dim sInfo As String
      
   If m_bRunning Then Exit Sub
   m_bRunning = True
   Me.MousePointer = 11
   sInfo = "Start: " & Time()
   Call DTAFunc(SSTabFunc.Tab)
   
   On Error GoTo ErrH
   
   Select Case SSTabFunc.Tab
   Case 0
      msDEM = txtSrcGRID0(0).Text
      msFlowDir = txtSrcGRID0(1).Text
      msVly = txtSrcGRID0(2).Text
      If msDEM = "" Or msFlowDir = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the INPUT GRID firstly"
      End If
      With txtValleyTag(0)
         If IsNumeric(.Text) Then
            miVlyTag = CInt(.Text)
            If miVlyTag <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameter: VALLEY TAG"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameter: VALLEY TAG"
         End If
      End With
      msSlpLen = txtSaveGRID0(0).Text
      msUpslpCells = txtSaveGRID0(1).Text
      If msUpslpCells = "" And msSlpLen = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the OUTPUT GRID firstly"
      End If
                  
      If Not SlpLen_SearchUpslope_bySFD() Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error when searching upslope"
      End If
   
   End Select
   
   MsgBox sInfo & vbCrLf & "Done: " & Time(), vbInformation, APP_TITLE
ErrH:
   '
   If Err.Number <> 0 Then
      MsgBox sInfo & vbCrLf & "Error: " & Time() & vbCrLf & Err.Description, vbExclamation, APP_TITLE
   End If
   On Error Resume Next
   With SSTabFunc
      For i = 0 To .Tabs - 1
         .TabEnabled(i) = True
      Next
      '.Tab = iFuncIndex
   End With
   Me.MousePointer = 0
'   CleanUpMemory
   m_bRunning = False
   
   SetProgressBarValue 0
End Sub

' assign OUTPUT GRID for step 1
Private Sub cmdSaveGRID0_Click(Index As Integer)
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveGRID0(Index).Text = strFile
End Sub

' assign source GRID
Private Sub cmdSrcGRID0_Click(Index As Integer)
   Dim strFile As String
   Dim pGrid As clsGrid
   Dim strPath As String, strName As String, strSuffix As String, strFilePre As String
   Dim i As Integer
On Error GoTo ErrH
   If m_bRunning Then Exit Sub
   
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   strFile = GetFileName(comdlg, True, , ".asc")
   If strFile = "" Then Exit Sub
   txtSrcGRID0(Index).Text = strFile
      
   Me.MousePointer = 11
   m_bRunning = True
    
   If Index = 0 Then
      i = InStrRev(strFile, "\")
      strPath = Left(strFile, i)
      strName = Right(strFile, Len(strFile) - i)
      i = InStrRev(strName, ".")
      If i = 0 Then
         strFilePre = strName
      Else
         strFilePre = Left(strName, i - 1)
      End If
         
      txtSaveGRID0(0).Text = strPath & "SlpLen.asc"
      txtSaveGRID0(1).Text = strPath & "SlpLenCells.asc"
      
      ' read parameters in SrcGRID file head
      On Error GoTo ErrH
      Set pGrid = New clsGrid
      With pGrid
         .LoadAscGrid strFile
         txtCols(0).Text = .nCols
         txtRows(0).Text = .nRows
         txtXll(0).Text = .xllcorner
         txtYll(0).Text = .yllcorner
         txtCellSize(0).Text = .CellSize
         txtNoData(0).Text = .NoData_Value
      End With
   End If
ErrH:
   Set pGrid = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   
   Me.MousePointer = 0
   m_bRunning = False
End Sub

Private Sub SetProgressBarValue(iValue As Integer)
   If iValue > 100 Or iValue < 0 Then Exit Sub
   With progbar
      .Value = iValue
      .Refresh
   End With
End Sub

Private Sub Form_Load()
   m_bRunning = False
   SetProgressBarValue 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bRunning Then
      Cancel = 1
      Exit Sub
   End If
   
   CleanUpMemory
End Sub

Private Function CleanUpMemory() As Boolean
   If Not (mpTag Is Nothing) Then Set mpTag = Nothing
   Set mpDEM = Nothing:   Set mpFlowDir = Nothing:   Set mpValley = Nothing
   Set mpUpslpCells = Nothing:   Set mpSlpLen = Nothing
   'Set mpUpslpDir = Nothing
End Function



