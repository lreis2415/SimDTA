VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMulFlowLen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple Flow Direction Flow Distance"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cdmQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   6360
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
      TabCaption(0)   =   "Flow Length"
      TabPicture(0)   =   "frmMulFlowLen.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameInput(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameOutput(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "framePara(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   3240
         Width           =   11055
         Begin VB.TextBox txtParaMFDP 
            Height          =   285
            Index           =   0
            Left            =   6540
            TabIndex        =   35
            Text            =   "1"
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtParaMFDP 
            Height          =   285
            Index           =   1
            Left            =   10140
            TabIndex        =   34
            Text            =   "10"
            Top             =   300
            Width           =   675
         End
         Begin VB.ComboBox cboMFDAlg 
            Height          =   300
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   300
            Width           =   3735
         End
         Begin VB.TextBox txtValleyTag 
            Height          =   315
            Index           =   0
            Left            =   1380
            TabIndex        =   25
            Text            =   "1"
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lblParaMFDP 
            Alignment       =   1  'Right Justify
            Caption         =   "flow exponent:"
            Height          =   315
            Index           =   0
            Left            =   5160
            TabIndex        =   37
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblParaMFDP 
            Alignment       =   2  'Center
            Caption         =   "(when tanb=0) <= p <= (when tanb=1)"
            Height          =   315
            Index           =   1
            Left            =   7140
            TabIndex        =   36
            Top             =   360
            Width           =   3075
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Algorithm"
            Height          =   375
            Index           =   14
            Left            =   180
            TabIndex        =   30
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Valley Tag"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   1155
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1035
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   4500
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID0 
            Caption         =   "Flow Dist"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtSaveGRID0 
            Height          =   375
            Index           =   0
            Left            =   1380
            TabIndex        =   22
            Top             =   360
            Width           =   9555
         End
         Begin VB.Label Label2 
            Caption         =   "(in distance unit)"
            Height          =   255
            Index           =   2
            Left            =   1380
            TabIndex        =   33
            Top             =   720
            Width           =   4095
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   11055
         Begin VB.TextBox txtSrcGRID0 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1380
            TabIndex        =   20
            Top             =   1500
            Width           =   9555
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "Valley"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1500
            Width           =   1275
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "DEM"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1275
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
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   780
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   0
               Left            =   10020
               TabIndex        =   10
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   0
               Left            =   8100
               TabIndex        =   9
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   0
               Left            =   6180
               TabIndex        =   8
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   0
               Left            =   4140
               TabIndex        =   7
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   0
               Left            =   2160
               TabIndex        =   6
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   0
               Left            =   600
               TabIndex        =   5
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   5
               Left            =   8880
               TabIndex        =   16
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   4
               Left            =   7320
               TabIndex        =   15
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   3
               Left            =   5280
               TabIndex        =   14
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   2
               Left            =   3240
               TabIndex        =   13
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   12
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.Label Label2 
            Caption         =   "(Default: No Valley Grid Assigned)"
            Height          =   255
            Index           =   1
            Left            =   1380
            TabIndex        =   32
            Top             =   1860
            Width           =   4815
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmMulFlowLen.frx":001C
         Height          =   495
         Index           =   0
         Left            =   300
         TabIndex        =   27
         Top             =   420
         Width           =   10815
      End
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   960
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   1500
      TabIndex        =   28
      Top             =   5820
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblRunning 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      TabIndex        =   31
      Top             =   5820
      Width           =   1455
   End
End
Attribute VB_Name = "frmMulFlowLen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_VALLEY = 1
'Const FUNC_TYPE_MFD_QUINN91 = "MFD (Quinn et al., 1991)"
'Const FUNC_TYPE_MFD_QIN07 = "MFD-md (Qin et al., 2007)"

Dim m_bRunning As Boolean

Dim mpDEM As clsGrid, mpValley As clsGrid
Dim mpFlowLen As clsGrid

Dim miRows As Integer, miCols As Integer, mdCell As Double
Dim mpTag As clsGrid

Dim miVlyTag As Integer
Dim msDEM As String, msVly As String
Dim msFlowLen As String

Private Sub cboMFDAlg_Click()
   Select Case cboMFDAlg.Text
   Case FUNC_TYPE_MFD_QUINN91
      txtParaMFDP(0).Text = "1"
      lblParaMFDP(1).Visible = False
      txtParaMFDP(1).Visible = False
   Case FUNC_TYPE_MFD_QIN07
      txtParaMFDP(0).Text = "1.1"
      lblParaMFDP(1).Visible = True
      txtParaMFDP(1).Visible = True
      txtParaMFDP(1).Text = "10"
   End Select
End Sub

Private Sub cdmQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub

Private Function GetSaveFileName() As String
   comdlg.DialogTitle = "Save GRID"
   comdlg.FileName = ""
   GetSaveFileName = GetFileName(comdlg, False, , ".asc")
End Function
'
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
   Dim i As Integer, iRow As Integer, iCol As Integer
   Dim sInfo As String, sMFDAlg As String
   Dim dSFD_p As Double, dMFD_p As Double
      
   If m_bRunning Then Exit Sub
   m_bRunning = True
   lblRunning.Caption = "Status: Running"
   Me.MousePointer = 11
   sInfo = "Start: " & Time()
'   Call DTAFunc(SSTabFunc.Tab)
   
   On Error GoTo ErrH
   
   Select Case SSTabFunc.Tab
   Case 0
      msDEM = txtSrcGRID0(0).Text
      msVly = txtSrcGRID0(2).Text
      If msDEM = "" Then
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
      sMFDAlg = cboMFDAlg.Text
      msFlowLen = txtSaveGRID0(0).Text
      If msFlowLen = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the OUTPUT GRID firstly"
      End If
      
      With txtParaMFDP(0)
         If IsNumeric(.Text) Then
            dMFD_p = CDbl(.Text)
            If dMFD_p < 1 Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters, p for MFD should be GE than 1."
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      Select Case sMFDAlg
      Case FUNC_TYPE_MFD_QUINN91
         
      Case FUNC_TYPE_MFD_QIN07
         With txtParaMFDP(1)
            If IsNumeric(.Text) Then
               dSFD_p = CDbl(.Text)
               If dSFD_p <= dMFD_p Then
                  .SetFocus
                  Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters, Value Range of p for MFD."
               End If
            Else
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         End With
      End Select
      
      Set mpDEM = New clsGrid
      mpDEM.LoadAscGrid msDEM
      With mpDEM
         miRows = .nRows: miCols = .nCols: mdCell = .CellSize
      
         Set mpValley = New clsGrid
         If msVly = "" Then
            mpValley.NewGrid miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True
         Else
            mpValley.LoadAscGrid msVly, True
         End If
         
         If miRows <> mpValley.nRows Or miCols <> mpValley.nCols Then
            Err.Raise Number:=vbObjectError + 513, Description:="INPUT GRID are not with same size"
         End If
         
         Set mpTag = New clsGrid
         If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
         Set mpFlowLen = New clsGrid
         If Not mpFlowLen.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End With
      
      'Initialize
      For iRow = 0 To miRows - 1
         For iCol = 0 To miCols - 1
            If mpDEM.Cell(iCol, iRow) = mpDEM.NoData_Value Then
               mpFlowLen.Cell(iCol, iRow) = mpFlowLen.NoData_Value
               mpTag.Cell(iCol, iRow) = True
            Else
               mpFlowLen.Cell(iCol, iRow) = 0#
               If mpValley.Cell(iCol, iRow) = miVlyTag Then
                  mpTag.Cell(iCol, iRow) = True
               Else
                  mpTag.Cell(iCol, iRow) = False
               End If
            End If
         Next
         DoEvents
      Next
      Set mpValley = Nothing
      SetProgressBarValue 20
                  
      If Not modMFD.FlowLen_by_MFD(mpDEM, mpTag, mpFlowLen, sMFDAlg, dMFD_p, dSFD_p - dMFD_p) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error when recursively calculate Multiple Flow Length"
      End If
      
      ''''''''''''''''''''
      ' output
      SetProgressBarValue 90
      If msFlowLen <> "" Then
         If mpFlowLen.SaveAscGrid(msFlowLen, , 2) Then
            'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msFlowLen
         End If
      End If
   End Select
   
   SetProgressBarValue 100
   MsgBox sInfo & vbCrLf & "Done: " & Time(), vbInformation, APP_TITLE
ErrH:
   '
   lblRunning.Caption = "Status: Idle"
   If Err.Number <> 0 Then
      MsgBox sInfo & vbCrLf & "Error: " & Time() & vbCrLf & Err.Description, vbExclamation, APP_TITLE
   End If
   On Error Resume Next
   CleanUpMemory
   
   Me.MousePointer = 0
   SetProgressBarValue 0
   m_bRunning = False
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
         
      txtSaveGRID0(0).Text = strPath & strFilePre & "_FlowLen.asc"
      
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
   
   With cboMFDAlg
      .Clear
      .AddItem FUNC_TYPE_MFD_QUINN91
      .AddItem FUNC_TYPE_MFD_QIN07
      .ListIndex = 1
   End With
   
   lblRunning.Caption = "Status: Idle"
      
   txtParaMFDP(0).Enabled = C_INNER_VERSION
   txtParaMFDP(1).Enabled = C_INNER_VERSION
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
   Set mpDEM = Nothing:    Set mpValley = Nothing
   Set mpFlowLen = Nothing
End Function





