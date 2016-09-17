VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmChangeGRIDRangeValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Range-value in GRID to New Value"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdChangeRangeValue 
      Caption         =   "Assign Range-value [Min, Max] as New Value"
      Height          =   495
      Left            =   2880
      TabIndex        =   31
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output GRID"
      Height          =   1095
      Left            =   60
      TabIndex        =   19
      Top             =   3240
      Width           =   11055
      Begin VB.CommandButton cmdSaveGRID 
         Caption         =   "Save GRID..."
         Height          =   375
         Left            =   60
         TabIndex        =   21
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txtSaveGRID 
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   9615
      End
   End
   Begin VB.Frame framePara 
      Caption         =   "Parameters"
      Height          =   855
      Left            =   60
      TabIndex        =   17
      Top             =   2340
      Width           =   11055
      Begin VB.TextBox txtNewValue 
         Height          =   315
         Left            =   8940
         TabIndex        =   30
         Top             =   300
         Width           =   1875
      End
      Begin VB.TextBox txtRangeTo 
         Height          =   315
         Left            =   6000
         TabIndex        =   28
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox txtRangeFrom 
         Height          =   315
         Left            =   3780
         TabIndex        =   27
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Min (>=)"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "as New Value"
         Height          =   255
         Index           =   4
         Left            =   7560
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Max (<=)"
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   26
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Assign Value-range:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9000
      TabIndex        =   16
      Top             =   4680
      Width           =   1875
   End
   Begin VB.Frame frameSrc 
      Caption         =   "Source"
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11055
      Begin VB.TextBox txtMax 
         Height          =   315
         Left            =   6000
         TabIndex        =   25
         Top             =   1740
         Width           =   1395
      End
      Begin VB.TextBox txtMin 
         Height          =   315
         Left            =   3780
         TabIndex        =   24
         Top             =   1740
         Width           =   1335
      End
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
      Begin VB.Label Label2 
         Caption         =   "Max="
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   23
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Min =  "
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   22
         Top             =   1800
         Width           =   795
      End
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   420
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmChangeGRIDRangeValue"
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

Private Sub cmdChangeRangeValue_Click()
On Error GoTo ErrH
   Dim strSaveGRID As String
   Dim dNewValue As Double, dRangeFrom As Double, dRangeTo As Double
   Dim pGrid As New clsGrid
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, dValue As Double, s As String
   
   If m_bRunning Then Exit Sub
   m_bRunning = True
   Me.MousePointer = 11
         
   ' get parameters
   If txtSrcGRID.Text = "" Then
      Err.Raise Number:=vbObjectError + 513, Description:="Assign the source GRID firstly"
   End If
   s = txtRangeFrom.Text
   If IsNumeric(s) Then
      dRangeFrom = CDbl(s)
'      If dRangeFrom = -9999 Then dRangeFrom = CDbl(txtNoData.Text)
   Else
      txtRangeFrom.SetFocus
      Err.Raise Number:=vbObjectError + 513, Description:="Error in setting change-range-value"
   End If
   s = txtRangeTo.Text
   If IsNumeric(s) Then
      dRangeTo = CDbl(s)
'      If dRangeTo = -9999 Then dRangeTo = CDbl(txtNoData.Text)
   Else
      txtRangeTo.SetFocus
      Err.Raise Number:=vbObjectError + 513, Description:="Error in setting change-range-value"
   End If
   If dRangeTo < dRangeFrom Then
      txtRangeTo.SetFocus
      Err.Raise Number:=vbObjectError + 513, Description:="Error in setting change-range-value"
   End If
   s = txtNewValue.Text
   If IsNumeric(s) Then
      dNewValue = CDbl(s)
'      If dNewValue = -9999 Then dNewValue = CDbl(txtNoData.Text)
   Else
      txtNewValue.SetFocus
      Err.Raise Number:=vbObjectError + 513, Description:="Error in setting change-range-value"
   End If
   strSaveGRID = txtSaveGRID.Text
   
   '
   With m_pBaseGRID
      iCols = .nCols: iRows = .nRows
      dXll = .xllcorner: dYll = .yllcorner
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   pGrid.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         dValue = m_pBaseGRID.Cell(iCol, iRow)
         If (dValue >= dRangeFrom And dValue <= dRangeTo) Then
            pGrid.Cell(iCol, iRow) = dNewValue
         Else
            pGrid.Cell(iCol, iRow) = dValue
         End If
      Next
   Next
   
   If pGrid.SaveAscGrid(strSaveGRID, , -1) Then
      MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
   End If
   
ErrH:
   Set pGrid = Nothing
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
      txtMin.Text = .Minimum
      txtMax.Text = .Maximum
   End With
   
   txtSaveGRID.Text = m_strBasePath & m_strFilePre & "_new.asc"
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

