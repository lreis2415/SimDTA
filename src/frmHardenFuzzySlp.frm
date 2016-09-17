VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmHardenFuzzySlp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Harden fuzzy slope positions"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   6420
      TabIndex        =   8
      Top             =   8460
      Width           =   1335
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   8460
      Width           =   1395
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output"
      Height          =   3075
      Left            =   60
      TabIndex        =   2
      Top             =   4980
      Width           =   9555
      Begin VB.Frame frameSSI 
         Caption         =   "Slope Position Sequence Index (SPSI)"
         Height          =   1215
         Left            =   60
         TabIndex        =   20
         Top             =   1800
         Width           =   9435
         Begin VB.ComboBox cboSSIModel 
            Height          =   315
            Left            =   1740
            TabIndex        =   23
            Top             =   300
            Width           =   7635
         End
         Begin VB.TextBox txtSaveSSI 
            Height          =   375
            Left            =   1740
            TabIndex        =   22
            Top             =   720
            Width           =   7635
         End
         Begin VB.CommandButton cmdSaveSSI 
            Caption         =   "&Slp. Pos. Seq. Index..."
            Height          =   375
            Left            =   60
            TabIndex        =   21
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Model for SPSI:"
            Height          =   315
            Left            =   420
            TabIndex        =   24
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.TextBox txtSave2ndMaxSim 
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   1380
         Width           =   7635
      End
      Begin VB.CommandButton cmdSave2ndMaxSim 
         Caption         =   "2&nd Max Similarity..."
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   1695
      End
      Begin VB.TextBox txtSave2ndHardSlpPos 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   1020
         Width           =   7635
      End
      Begin VB.CommandButton cmdSave2ndHardSlpPos 
         Caption         =   "&2nd Hard Slope Pos..."
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveGRID 
         Caption         =   "&Hard Slope Pos..."
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtSaveGRID 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   300
         Width           =   7875
      End
      Begin VB.CommandButton cmdSaveMaxSim 
         Caption         =   "&Max. Similarity..."
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtSaveMaxSim 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   660
         Width           =   7875
      End
   End
   Begin VB.Frame frameSrc 
      Caption         =   "Source GRIDs of fuzzy slope positions"
      Height          =   4515
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   9555
      Begin VB.ComboBox cboSlpPosType 
         Height          =   300
         Left            =   5100
         TabIndex        =   11
         Top             =   840
         Width           =   4215
      End
      Begin VB.CommandButton cmdOpenGRID 
         Caption         =   "&Assign Fuzzy Slope GRID..."
         Height          =   435
         Left            =   180
         TabIndex        =   10
         Top             =   780
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   4
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   9375
         Begin VB.OptionButton optSlpPosSeq 
            Caption         =   "2nd Level (11 types of slope position)"
            Height          =   315
            Index           =   1
            Left            =   4500
            TabIndex        =   15
            Top             =   120
            Width           =   4815
         End
         Begin VB.OptionButton optSlpPosSeq 
            Caption         =   "1st Level (5 types of slope position)"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Value           =   -1  'True
            Width           =   4335
         End
      End
      Begin VB.Label Label1 
         Caption         =   "type of slope position:"
         Height          =   315
         Index           =   0
         Left            =   2760
         TabIndex        =   12
         Top             =   900
         Width           =   2355
      End
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   8040
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   360
      Top             =   8460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Ref.: 秦承志等, 2007; 秦承志等, in press; Qin et al., in reviewing"
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   60
      Width           =   9135
   End
End
Attribute VB_Name = "frmHardenFuzzySlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Const C_SLPPOS_LVL1 = 0
Const C_SLPPOS_LVL2 = 1
Const C_SLPPOS_LVL1_COUNT = 5
Const C_SLPPOS_LVL2_COUNT = 11

Const C_SSI_MODEL1 = "Model 1: [HardCls] + sgn([2ndHardCls]-[HardCls]) * (1-[MaxSim])/2"
Const C_SSI_MODEL2 = "Model 2: [HardCls] + sgn([2ndHardCls]-[HardCls]) * (1-([MaxSim]-[2ndMaxSim]))/2"

Const C_SLPPOS_TYPE_RDG = "Ridge (RDG)"
Const C_SLPPOS_TYPE_SHD = "Shoulder (SHD)"
Const C_SLPPOS_TYPE_BKS = "Backslope (BKS)"
Const C_SLPPOS_TYPE_FTS = "Footslope (FTS)"
Const C_SLPPOS_TYPE_VLY = "Valley (VLY)"
Const C_SLPPOS_TYPE_DSHD = "Divergent/Convex Shoulder (DSHD)"
Const C_SLPPOS_TYPE_PSHD = "Planar Shoulder (PSHD)"
Const C_SLPPOS_TYPE_CSHD = "Convergent/Concave Shoulder (CSHD)"
Const C_SLPPOS_TYPE_DBKS = "Divergent/Convex Backslope (DBKS)"
Const C_SLPPOS_TYPE_PBKS = "Planar Backslope (PBKS)"
Const C_SLPPOS_TYPE_CBKS = "Convergent/Concave Backslope (CBKS)"
Const C_SLPPOS_TYPE_DFTS = "Divergent/Convex Footslope (DFTS)"
Const C_SLPPOS_TYPE_PFTS = "Planar Footslope (PFTS)"
Const C_SLPPOS_TYPE_CFTS = "Convergent/Concave Footslope (CFTS)"

Dim m_bRunning As Boolean

Dim miRows As Integer, miCols As Integer, mdCell As Double

Dim miSlpPosLvl As Integer
Dim miSlpPosCount As Integer
Dim marriSlpPosType() As Integer
Dim marrsSlpPosCap() As String

Private Function FillList_SlpPosType(iLevel As Integer) As Boolean
   Dim iRow As Integer
   
   FillList_SlpPosType = False
   With cboSlpPosType
      .Clear
      Select Case iLevel
      Case C_SLPPOS_LVL1   '0
         miSlpPosLvl = C_SLPPOS_LVL1
         miSlpPosCount = C_SLPPOS_LVL1_COUNT
         .AddItem C_SLPPOS_TYPE_RDG
         .AddItem C_SLPPOS_TYPE_SHD
         .AddItem C_SLPPOS_TYPE_BKS
         .AddItem C_SLPPOS_TYPE_FTS
         .AddItem C_SLPPOS_TYPE_VLY
         ReDim marrsSlpPosCap(1 To miSlpPosCount)
         marrsSlpPosCap(1) = "RDG"
         marrsSlpPosCap(2) = "SHD"
         marrsSlpPosCap(3) = "BKS"
         marrsSlpPosCap(4) = "FTS"
         marrsSlpPosCap(5) = "VLY"
      Case C_SLPPOS_LVL2   '1
         miSlpPosLvl = C_SLPPOS_LVL2
         miSlpPosCount = C_SLPPOS_LVL2_COUNT
         .AddItem C_SLPPOS_TYPE_RDG
         .AddItem C_SLPPOS_TYPE_DSHD
         .AddItem C_SLPPOS_TYPE_PSHD
         .AddItem C_SLPPOS_TYPE_CSHD
         .AddItem C_SLPPOS_TYPE_DBKS
         .AddItem C_SLPPOS_TYPE_PBKS
         .AddItem C_SLPPOS_TYPE_CBKS
         .AddItem C_SLPPOS_TYPE_DFTS
         .AddItem C_SLPPOS_TYPE_PFTS
         .AddItem C_SLPPOS_TYPE_CFTS
         .AddItem C_SLPPOS_TYPE_VLY
         ReDim marrsSlpPosCap(1 To miSlpPosCount)
         marrsSlpPosCap(1) = "RDG"
         marrsSlpPosCap(2) = "DSHD"
         marrsSlpPosCap(3) = "PSHD"
         marrsSlpPosCap(4) = "CSHD"
         marrsSlpPosCap(5) = "DBKS"
         marrsSlpPosCap(6) = "PBKS"
         marrsSlpPosCap(7) = "CBKS"
         marrsSlpPosCap(8) = "DFTS"
         marrsSlpPosCap(9) = "PFTS"
         marrsSlpPosCap(10) = "CFTS"
         marrsSlpPosCap(11) = "VLY"
      End Select
      .ListIndex = -1
   End With
   
   ' table headings
   With MSFlexGrid1
      .Rows = 1
      .Rows = miSlpPosCount + 1
      For iRow = 1 To miSlpPosCount
         .TextMatrix(iRow, 0) = iRow
         .TextMatrix(iRow, 1) = marrsSlpPosCap(iRow)
         .TextMatrix(iRow, 2) = marriSlpPosType(iRow)
         '.TextMatrix(iRow, 3) = ""
      Next
   End With
   FillList_SlpPosType = True
End Function


Private Function FillList_SSIModel(iLevel As Integer) As Boolean
   Dim iRow As Integer
   
   FillList_SSIModel = False
   With cboSSIModel
      .Clear
      Select Case iLevel
      Case C_SLPPOS_LVL1   '0
         frameSSI.Enabled = True
         .Clear
         .AddItem C_SSI_MODEL1
         .AddItem C_SSI_MODEL2
         .ListIndex = 0
         
      Case C_SLPPOS_LVL2   '1
         txtSaveSSI.Text = ""
         frameSSI.Enabled = False
         .ListIndex = -1
      End Select
   End With
   
   FillList_SSIModel = True
End Function

Private Sub SetProgressBarValue(iValue As Integer)
   If iValue > 100 Or iValue < 0 Then Exit Sub
   With progbar
      .Value = iValue
      .Refresh
   End With
End Sub


Private Sub cmdOpenGRID_Click()
   Dim strBaseDEM As String
   
   If m_bRunning Then Exit Sub
   If cboSlpPosType.ListIndex < 0 Then
      MsgBox "Assign the type of slope position firstly", vbInformation, APP_TITLE
      Exit Sub
   End If
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   strBaseDEM = GetFileName(comdlg, True, , ".asc")
   If strBaseDEM = "" Then Exit Sub
   
   MSFlexGrid1.TextMatrix(cboSlpPosType.ListIndex + 1, 3) = strBaseDEM
End Sub

Private Sub cmdQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub

Private Sub cmdRun_Click()
   On Error GoTo ErrH
   Dim strSrcGRID As String, strSaveGRID As String, strSaveMaxSim As String
   Dim strSaveSSI As String, strSave2ndSlpPos As String, strSave2ndMaxSim As String
   Dim pSrcGRID As clsGrid, pGrid As clsGrid, pGridMaxSim As clsGrid
   Dim pGrid2ndSlpPos As clsGrid, pGrid2ndMaxSim As clsGrid, clsGRIDSSI As clsGrid
   Dim boolPara As Boolean, iLyr As Integer
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer
   Dim iSlpPosTag As Integer
   Dim iSSIModel As Integer
   'Dim dTemp As Double, dValue As Double
   
   If m_bRunning Then Exit Sub
   
   ' verify parameters
   boolPara = True
   If ((miSlpPosLvl = C_SLPPOS_LVL1) And (miSlpPosCount <> C_SLPPOS_LVL1_COUNT)) _
         Or ((miSlpPosLvl = C_SLPPOS_LVL2) And (miSlpPosCount <> C_SLPPOS_LVL2_COUNT)) Then
      boolPara = False
   Else
      With MSFlexGrid1
         If .Rows <> miSlpPosCount + 1 Then
            boolPara = False
         Else
            For iLyr = 1 To miSlpPosCount
               If Trim(.TextMatrix(iLyr, 3)) = "" Then
                  boolPara = False
                  Exit For
               End If
            Next
         End If
      End With
   End If
   If Not boolPara Then
      MsgBox "Wrong in parameter-setting for 'Source GRIDs of fuzzy slope positions'!", vbInformation, APP_TITLE
      Exit Sub
   End If
         
   ' get parameters
   strSaveGRID = Trim(txtSaveGRID.Text)
'   If strSaveGRID = "" Then
'      MsgBox "Assign the Hardened Slope Position file name first", vbInformation, APP_TITLE
'      Exit Sub
'   End If
   strSaveMaxSim = Trim(txtSaveMaxSim.Text)
   strSave2ndSlpPos = Trim(txtSave2ndHardSlpPos.Text)
   strSave2ndMaxSim = Trim(txtSave2ndMaxSim.Text)
         
   If miSlpPosLvl = C_SLPPOS_LVL1 Then
      iSSIModel = cboSSIModel.ListIndex
      strSaveSSI = Trim(txtSaveSSI.Text)
      If strSaveSSI <> "" And iSSIModel = -1 Then
         MsgBox "Assign the Model for computing Slope Position Sequence Index (SPSI) first", vbInformation, APP_TITLE
         Exit Sub
      End If
      If strSaveSSI = "" Then
         iSSIModel = -1
      End If
   Else
      iSSIModel = -1
      strSaveSSI = ""
   End If
      
   m_bRunning = True
   Me.MousePointer = 11
      
   '
   Set pSrcGRID = New clsGrid
   For iLyr = 1 To miSlpPosCount
      strSrcGRID = Trim(MSFlexGrid1.TextMatrix(iLyr, 3))
      iSlpPosTag = MSFlexGrid1.TextMatrix(iLyr, 2)
      
      'Set pSrcGRID = New clsGrid
      If Not pSrcGRID.LoadAscGrid(strSrcGRID) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in openning Source GRID: " & strSrcGRID
      End If
      DoEvents
      
      If iLyr = 1 Then
         With pSrcGRID
            iCols = .nCols: iRows = .nRows
            dXll = .xllcorner: dYll = .yllcorner
            dCellSize = .CellSize
            dNoData = .NoData_Value
         End With
         Set pGrid = New clsGrid
         pGrid.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData, , True
         Set pGridMaxSim = New clsGrid
         pGridMaxSim.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData, , False
         Set pGrid2ndSlpPos = New clsGrid
         pGrid2ndSlpPos.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData, , True
         Set pGrid2ndMaxSim = New clsGrid
         pGrid2ndMaxSim.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData, , False
         
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               If pSrcGRID.Cell(iCol, iRow) = dNoData Then
                  pGrid.Cell(iCol, iRow) = pGrid.NoData_Value
                  pGridMaxSim.Cell(iCol, iRow) = pGridMaxSim.NoData_Value
                  
                  pGrid2ndSlpPos.Cell(iCol, iRow) = pGrid2ndSlpPos.NoData_Value
                  pGrid2ndMaxSim.Cell(iCol, iRow) = pGrid2ndMaxSim.NoData_Value
               Else
                  pGrid.Cell(iCol, iRow) = iSlpPosTag
                  pGridMaxSim.Cell(iCol, iRow) = pSrcGRID.Cell(iCol, iRow)
               End If
            Next
            DoEvents
         Next
      Else
         With pSrcGRID
            If iCols <> .nCols Or iRows <> .nRows Then
               Err.Raise Number:=vbObjectError + 513, Description:="Error in matching GRID size with other GRIDs: " & strSrcGRID
            End If
         End With
         
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               If pSrcGRID.Cell(iCol, iRow) <> pSrcGRID.NoData_Value Then
                  If pSrcGRID.Cell(iCol, iRow) > pGridMaxSim.Cell(iCol, iRow) Then
                     pGrid2ndSlpPos.Cell(iCol, iRow) = pGrid.Cell(iCol, iRow)
                     pGrid2ndMaxSim.Cell(iCol, iRow) = pGridMaxSim.Cell(iCol, iRow)
                     
                     pGrid.Cell(iCol, iRow) = iSlpPosTag
                     pGridMaxSim.Cell(iCol, iRow) = pSrcGRID.Cell(iCol, iRow)
                     
                  ElseIf pSrcGRID.Cell(iCol, iRow) = pGridMaxSim.Cell(iCol, iRow) Then
                     pGrid.Cell(iCol, iRow) = pGrid.Cell(iCol, iRow) + iSlpPosTag
                     
                  Else
                     If pSrcGRID.Cell(iCol, iRow) > pGrid2ndMaxSim.Cell(iCol, iRow) Then
                        pGrid2ndSlpPos.Cell(iCol, iRow) = iSlpPosTag
                        pGrid2ndMaxSim.Cell(iCol, iRow) = pSrcGRID.Cell(iCol, iRow)
                     End If
                     
                  End If
               End If
            Next
            DoEvents
         Next
      End If
      SetProgressBarValue Int(iLyr * 100# / miSlpPosCount)
      DoEvents
   Next
   
   If iSSIModel > -1 And strSaveSSI <> "" Then
      Set clsGRIDSSI = New clsGrid
      clsGRIDSSI.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData, , False
      
      Select Case iSSIModel
      Case 0
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               If pGrid.Cell(iCol, iRow) <> pGrid.NoData_Value And pGrid2ndSlpPos.Cell(iCol, iRow) <> pGrid2ndSlpPos.NoData_Value Then
                     
                  clsGRIDSSI.Cell(iCol, iRow) = (Log(pGrid.Cell(iCol, iRow)) / Log(2) + 1) _
                           + Sgn(pGrid2ndSlpPos.Cell(iCol, iRow) - pGrid.Cell(iCol, iRow)) _
                           * (1 - pGridMaxSim.Cell(iCol, iRow)) / 2
               Else
                  clsGRIDSSI.Cell(iCol, iRow) = clsGRIDSSI.NoData_Value
               End If
            Next
            If (iCol Mod 10) = 0 Then SetProgressBarValue Int(iCol * 100# / iCols)
            DoEvents
         Next
      Case 1
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               If pGrid.Cell(iCol, iRow) <> pGrid.NoData_Value And pGrid2ndSlpPos.Cell(iCol, iRow) <> pGrid2ndSlpPos.NoData_Value Then
                     
                  clsGRIDSSI.Cell(iCol, iRow) = (Log(pGrid.Cell(iCol, iRow)) / Log(2) + 1) _
                           + Sgn(pGrid2ndSlpPos.Cell(iCol, iRow) - pGrid.Cell(iCol, iRow)) _
                           * (1 - (pGridMaxSim.Cell(iCol, iRow) - pGrid2ndMaxSim.Cell(iCol, iRow))) / 2
               Else
                  clsGRIDSSI.Cell(iCol, iRow) = clsGRIDSSI.NoData_Value
               End If
            Next
            If (iCol Mod 10) = 0 Then SetProgressBarValue Int(iCol * 100# / iCols)
            DoEvents
         Next
      Case Else
      
      End Select
      
      If Not clsGRIDSSI.SaveAscGrid(strSaveSSI, , 5) Then
         Err.Raise vbObjectError + 513, , "Failed to save result GRID - Slope Position Sequence Index: " & strSaveSSI
      End If
      SetProgressBarValue 20
      DoEvents
   End If
   
   If strSaveGRID <> "" Then
      If Not pGrid.SaveAscGrid(strSaveGRID, , 0) Then
         Err.Raise vbObjectError + 513, , "Failed to save result GRID - Harden Slope Position: " & strSaveGRID
      End If
   End If
   SetProgressBarValue 40
   DoEvents
   If strSaveMaxSim <> "" Then
      If Not pGridMaxSim.SaveAscGrid(strSaveMaxSim, , 5) Then
         Err.Raise vbObjectError + 513, , "Failed to save result GRID - Maximum Similarity: " & strSaveMaxSim
      End If
   End If
   SetProgressBarValue 60
   DoEvents
   If strSave2ndSlpPos <> "" Then
      If Not pGrid2ndSlpPos.SaveAscGrid(strSave2ndSlpPos, , 0) Then
         Err.Raise vbObjectError + 513, , "Failed to save result GRID - 2nd Harden Slope Position: " & strSave2ndSlpPos
      End If
   End If
   SetProgressBarValue 80
   DoEvents
   If strSave2ndMaxSim <> "" Then
      If Not pGrid2ndMaxSim.SaveAscGrid(strSave2ndMaxSim, , 5) Then
         Err.Raise vbObjectError + 513, , "Failed to save result GRID - 2nd Maximum Similarity: " & strSave2ndMaxSim
      End If
   End If
   SetProgressBarValue 100
   DoEvents
   
   MsgBox "Completed. Save result GRID: " & vbCrLf & "Harden Slope Position: " & strSaveGRID _
         & vbCrLf & "Maximum Similarity: " & strSaveMaxSim _
         & vbCrLf & "2nd Harden Slope Position: " & strSave2ndSlpPos _
         & vbCrLf & "2nd Maximum Similarity: " & strSave2ndMaxSim _
         & vbCrLf & "Slope Position Sequence Index: " & strSaveSSI _
         , vbInformation, APP_TITLE
   
ErrH:
   Set pGrid = Nothing: Set pGridMaxSim = Nothing
   Set pSrcGRID = Nothing
   Set pGrid2ndSlpPos = Nothing: Set pGrid2ndMaxSim = Nothing: Set clsGRIDSSI = Nothing
   
   Me.MousePointer = 0
   m_bRunning = False
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   SetProgressBarValue 0
End Sub

Private Sub cmdSave2ndHardSlpPos_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSave2ndHardSlpPos.Text = strFile
End Sub

Private Sub cmdSave2ndMaxSim_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSave2ndMaxSim.Text = strFile
End Sub

Private Sub cmdSaveGRID_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveGRID.Text = strFile
End Sub

Private Sub cmdSavemaxsim_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveMaxSim.Text = strFile
End Sub

Private Sub cmdSaveSSI_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveSSI.Text = strFile
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   ReDim marriSlpPosType(1 To 11)
   miSlpPosLvl = -9999
   For i = 1 To 11
      marriSlpPosType(i) = 2 ^ (i - 1)
   Next
   
   m_bRunning = False
   SetProgressBarValue 0
   With MSFlexGrid1
      .AllowUserResizing = 1  'flexResizeColumns
      .Cols = 4
      .Rows = 1
      .FixedCols = 3
      '.FixedRows = 1
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Slope Pos."
      .TextMatrix(0, 2) = "Tag"
      .TextMatrix(0, 3) = "Source GRID"
      .ColWidth(3) = 10 * .ColWidth(0)
   End With
   
   FillList_SlpPosType C_SLPPOS_LVL1
   FillList_SSIModel C_SLPPOS_LVL1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bRunning Then
      Cancel = 1
      Exit Sub
   End If
   
   CleanUpMemory
End Sub

Private Function CleanUpMemory() As Boolean
'   If Not (mpTag Is Nothing) Then Set mpTag = Nothing
'   Set mpDEM = Nothing:   Set mpFlowDir = Nothing:   Set mpRidge = Nothing: Set mpValley = Nothing
'   Set mpUpslpCells = Nothing:   Set mpUpslpRlf = Nothing:   Set mpUpslpDir = Nothing
'   Set mpDownslpCells = Nothing:   Set mpDownslpRlf = Nothing
'   Set mpRlfDiffer = Nothing
'   Set mpSlpShape = Nothing
'   Set mpUpRelRlfMax = Nothing:  Set mpUpRelRlfMin = Nothing:   Set mpUpSlpShape = Nothing
'   Set mpDRelRlfMax = Nothing:   Set mpDRelRlfMin = Nothing:   Set mpDSlpShape = Nothing
End Function

Private Function GetSaveFileName() As String
   comdlg.DialogTitle = "Save GRID"
   comdlg.FileName = ""
   GetSaveFileName = GetFileName(comdlg, False, , ".asc")
End Function


Private Sub optSlpPosSeq_Click(Index As Integer)
   
   If miSlpPosLvl = Index Then
      Exit Sub
   Else
      If MsgBox("Sure to change the level of slope position?", vbYesNo + vbQuestion, APP_TITLE) = vbNo Then
         optSlpPosSeq(Index).Value = False
         Exit Sub
      End If
   End If
   
   FillList_SlpPosType Index
   FillList_SSIModel Index
End Sub
