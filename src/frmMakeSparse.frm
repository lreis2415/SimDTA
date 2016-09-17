VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMakeSparse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make GRID Value Sparse"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin TabDlg.SSTab SSTabPara 
      Height          =   2295
      Left            =   0
      TabIndex        =   21
      Top             =   1860
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parameters"
      TabPicture(0)   =   "frmMakeSparse.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "framePara(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame framePara 
         Height          =   1455
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   720
         Width           =   11055
         Begin VB.TextBox txtValueAfter 
            Height          =   315
            Left            =   2400
            TabIndex        =   31
            Text            =   "0"
            Top             =   1020
            Width           =   735
         End
         Begin VB.TextBox txtSrcValueSE 
            Height          =   315
            Left            =   6000
            TabIndex        =   30
            Text            =   "1"
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox txtSrcValueGE 
            Height          =   315
            Left            =   2400
            TabIndex        =   29
            Text            =   "1"
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox txtSparseRatio 
            Height          =   315
            Left            =   2220
            TabIndex        =   23
            Text            =   "2"
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblChangeCount 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9720
            TabIndex        =   36
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblCountInRange 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9720
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Cell count with value changed: "
            Height          =   255
            Index           =   7
            Left            =   6780
            TabIndex        =   34
            Top             =   1140
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "Cell count in value range: "
            Height          =   255
            Index           =   6
            Left            =   7080
            TabIndex        =   33
            Top             =   720
            Width           =   2595
         End
         Begin VB.Label Label2 
            Caption         =   "Value before made sparse:"
            Height          =   435
            Index           =   5
            Left            =   120
            TabIndex        =   32
            Top             =   660
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Value after made sparse:"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   2355
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "<= Value to be made sparse <="
            Height          =   375
            Index           =   3
            Left            =   3120
            TabIndex        =   27
            Top             =   660
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "1 /"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   26
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label2 
            Caption         =   "Sparse ratio:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Label Label2 
         Caption         =   $"frmMakeSparse.frx":001C
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   300
         Width           =   10755
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   495
      Left            =   2940
      TabIndex        =   20
      Top             =   5580
      Width           =   5235
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output GRID"
      Height          =   1035
      Left            =   60
      TabIndex        =   17
      Top             =   4260
      Width           =   11055
      Begin VB.CommandButton cmdSaveGRID 
         Caption         =   "Save GRID..."
         Height          =   375
         Left            =   60
         TabIndex        =   19
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtSaveGRID 
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   360
         Width           =   9615
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   9000
      TabIndex        =   16
      Top             =   5580
      Width           =   1875
   End
   Begin VB.Frame frameSrc 
      Caption         =   "Source"
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11055
      Begin VB.Frame frameFileHead 
         Caption         =   "File Head"
         Enabled         =   0   'False
         Height          =   675
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   10815
         Begin VB.TextBox txtCols 
            Height          =   315
            Left            =   600
            TabIndex        =   9
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtRows 
            Height          =   315
            Left            =   2040
            TabIndex        =   8
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtXll 
            Height          =   315
            Left            =   3660
            TabIndex        =   7
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtYll 
            Height          =   315
            Left            =   5880
            TabIndex        =   6
            Text            =   "0"
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox txtCellSize 
            Height          =   315
            Left            =   8100
            TabIndex        =   5
            Text            =   "1"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtNoData 
            Height          =   315
            Left            =   10020
            TabIndex        =   4
            Text            =   "-9999"
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "nCols"
            Height          =   315
            Index           =   89
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   555
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
            Caption         =   "XllCorner"
            Height          =   315
            Index           =   87
            Left            =   2880
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "YllCorner"
            Height          =   315
            Index           =   86
            Left            =   5160
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "CellSize"
            Height          =   315
            Index           =   85
            Left            =   7380
            TabIndex        =   11
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "NoData_Value"
            Height          =   315
            Index           =   84
            Left            =   8880
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtSrcGRID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   9615
      End
      Begin VB.CommandButton cmdSrcGRID 
         Caption         =   "Src GRID"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   420
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMakeSparse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_bRunning As Boolean
Dim m_strBasePath As String
Dim m_strFilePre As String
Dim m_pBaseGRID As clsGrid

Private Sub cmdRun_Click()
On Error GoTo ErrH
   Dim strSaveGRID As String
   Dim iLoopCount As Integer, bool8Neighbor As Boolean
   Dim pGrid As clsGrid, pSrcGRID As clsGrid, pTemp As clsGrid
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, dValue As Double, iNum As Integer, s As String
   Dim iSparseRatio As Integer, dValueGE As Double, dValueSE As Double, dNewValue As Double
   Dim lCountInRange As Long, lChangeCount As Long
      
   If m_bRunning Then Exit Sub
   If m_pBaseGRID Is Nothing Then
      MsgBox "Assign the INPUT file name first", vbInformation, APP_TITLE
      Exit Sub
   End If
   strSaveGRID = Trim(txtSaveGRID.Text)
   If strSaveGRID = "" Then
      MsgBox "Assign the OUTPUT file name first", vbInformation, APP_TITLE
      Exit Sub
   End If
   
   m_bRunning = True
   Me.MousePointer = 11
      
   ' get parameters
   If txtSrcGRID.Text = "" Then
      Err.Raise Number:=vbObjectError + 513, Description:="Assign the source GRID firstly"
   End If
   strSaveGRID = txtSaveGRID.Text
   
   With txtSparseRatio
      s = .Text
      If IsNumeric(s) Then
         iSparseRatio = CInt(s)
         If iSparseRatio < 2 Or iSparseRatio <> CDbl(s) Then
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Sparse-ratio"
         End If
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Sparse-ratio"
      End If
   End With
   With txtSrcValueGE
      s = .Text
      If IsNumeric(s) Then
         dValueGE = CDbl(s)
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Value-to-be-made-sparse"
      End If
   End With
   With txtSrcValueSE
      s = .Text
      If IsNumeric(s) Then
         dValueSE = CDbl(s)
         If dValueSE < dValueGE Then
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Value-to-be-made-sparse"
         End If
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Value-to-be-made-sparse"
      End If
   End With
   With txtValueAfter
      s = .Text
      If IsNumeric(s) Then
         dNewValue = CDbl(s)
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Value-after-made-sparse"
      End If
   End With
   
   '
   With m_pBaseGRID
      iCols = .nCols: iRows = .nRows
      dXll = .xllcorner: dYll = .yllcorner
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   Set pGrid = New clsGrid
   pGrid.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData
   
   Set pSrcGRID = m_pBaseGRID
   iNum = 0
   lCountInRange = 0: lChangeCount = 0
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         dValue = pSrcGRID.Cell(iCol, iRow)
         If dValue >= dValueGE And dValue <= dValueSE Then
            iNum = iNum + 1
            lCountInRange = lCountInRange + 1
            If iNum = iSparseRatio Then
               pGrid.Cell(iCol, iRow) = dValue
               iNum = 0
            Else
               pGrid.Cell(iCol, iRow) = dNewValue
               lChangeCount = lChangeCount + 1
            End If
         Else
            pGrid.Cell(iCol, iRow) = dValue
         End If
      Next
   Next
   
   lblCountInRange.Caption = lCountInRange
   lblChangeCount.Caption = lChangeCount
   If pGrid.SaveAscGrid(strSaveGRID, , -1) Then
      MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
   End If
   
ErrH:
   Set pGrid = Nothing
   Set pSrcGRID = Nothing
   Set pTemp = Nothing
   Me.MousePointer = 0
   m_bRunning = False
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
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
   
   If m_bRunning Then Exit Sub
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   strBaseDEM = GetFileName(comdlg, True, , ".asc")
   If strBaseDEM = "" Then Exit Sub
   
   Dim strPath As String, strName As String, strSuffix As String
   Dim i As Integer
    
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
   
   txtSaveGRID.Text = m_strBasePath & m_strFilePre & "_Sparse.asc"
End Sub

Private Sub Form_Load()
   'initialize var
   m_bRunning = False
   Set m_pBaseGRID = Nothing
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bRunning Then
      Cancel = 1
      Exit Sub
   End If
   If Not (m_pBaseGRID Is Nothing) Then Set m_pBaseGRID = Nothing
End Sub

