VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmFuzzySlpInfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fuzzy Quantification of Slope Positions"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   9660
      TabIndex        =   25
      Top             =   6000
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTabFunc 
      Height          =   4635
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Input GRID"
      TabPicture(0)   =   "frmFuzzySlpInfer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "comdlg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "framePrototype"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameSrcGRID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGotoSetPara"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Parameters, Output"
      TabPicture(1)   =   "frmFuzzySlpInfer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRun"
      Tab(1).Control(1)=   "frameOutput"
      Tab(1).Control(2)=   "frameSetPara"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run"
         Height          =   495
         Left            =   -65340
         TabIndex        =   24
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output"
         Height          =   1155
         Left            =   -74880
         TabIndex        =   19
         Top             =   3360
         Width           =   9495
         Begin VB.TextBox txtSaveLog 
            Height          =   375
            Left            =   1500
            TabIndex        =   23
            Top             =   660
            Width           =   7815
         End
         Begin VB.CommandButton cmdSaveLog 
            Caption         =   "Save &Log..."
            Height          =   375
            Left            =   180
            TabIndex        =   22
            Top             =   660
            Width           =   1335
         End
         Begin VB.TextBox txtSaveGRID 
            Height          =   375
            Left            =   1500
            TabIndex        =   21
            Top             =   300
            Width           =   7815
         End
         Begin VB.CommandButton cmdSaveGRID 
            Caption         =   "&Output GRID..."
            Height          =   375
            Left            =   180
            TabIndex        =   20
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame frameSetPara 
         Caption         =   "Parameter Setting"
         Height          =   2955
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   10935
         Begin VB.Frame frameDistWeightType 
            Caption         =   "Distance type for Inverse Distance Weighting"
            Height          =   675
            Left            =   3900
            TabIndex        =   17
            Top             =   2160
            Width           =   6795
            Begin VB.OptionButton optDistWeightType 
               Caption         =   "Euclidean distance"
               Height          =   255
               Left            =   300
               TabIndex        =   18
               Top             =   300
               Value           =   -1  'True
               Width           =   2835
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   1935
            Left            =   180
            TabIndex        =   15
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   3413
            _Version        =   393216
            Cols            =   7
         End
         Begin VB.TextBox txtParaR 
            Height          =   375
            Left            =   1140
            TabIndex        =   14
            Text            =   "8"
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "r = "
            Height          =   315
            Index           =   1
            Left            =   540
            TabIndex        =   16
            Top             =   2460
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdGotoSetPara 
         Caption         =   "&Set Parameters"
         Height          =   495
         Left            =   9660
         TabIndex        =   11
         Top             =   3900
         Width           =   1335
      End
      Begin VB.Frame frameSrcGRID 
         Caption         =   "Parameter GRIDs"
         Height          =   2175
         Left            =   60
         TabIndex        =   7
         Top             =   2340
         Width           =   9015
         Begin VB.CommandButton cmdDelGRID 
            Caption         =   "&Delete Current GRID"
            Height          =   495
            Left            =   180
            TabIndex        =   10
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddGRID 
            Caption         =   "&Add GRID..."
            Height          =   495
            Left            =   180
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.ListBox lstSrcGRID 
            Height          =   1425
            ItemData        =   "frmFuzzySlpInfer.frx":0038
            Left            =   1500
            List            =   "frmFuzzySlpInfer.frx":003A
            MultiSelect     =   2  'Extended
            TabIndex        =   8
            Top             =   360
            Width           =   7395
         End
      End
      Begin VB.Frame framePrototype 
         Caption         =   "Prototype GRID"
         Height          =   1935
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Enabled         =   0   'False
            Height          =   675
            Left            =   120
            TabIndex        =   26
            Top             =   1020
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Left            =   10020
               TabIndex        =   32
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Left            =   8100
               TabIndex        =   31
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Left            =   5880
               TabIndex        =   30
               Text            =   "0"
               Top             =   240
               Width           =   1395
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Left            =   3660
               TabIndex        =   29
               Text            =   "0"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Left            =   2040
               TabIndex        =   28
               Top             =   240
               Width           =   795
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Left            =   600
               TabIndex        =   27
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   84
               Left            =   8880
               TabIndex        =   38
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   85
               Left            =   7380
               TabIndex        =   37
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   86
               Left            =   5160
               TabIndex        =   36
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   87
               Left            =   2880
               TabIndex        =   35
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   88
               Left            =   1500
               TabIndex        =   34
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   89
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.TextBox txtPrototypeTag 
            Height          =   375
            Left            =   10380
            TabIndex        =   6
            Text            =   "1"
            Top             =   420
            Width           =   555
         End
         Begin VB.TextBox txtPrototypeGRID 
            Height          =   375
            Left            =   1500
            TabIndex        =   4
            Top             =   420
            Width           =   7455
         End
         Begin VB.CommandButton cmdPrototypeGRID 
            Caption         =   "&Prototype GRID"
            Height          =   375
            Left            =   180
            TabIndex        =   3
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Prototype Tag:"
            Height          =   315
            Index           =   0
            Left            =   9000
            TabIndex        =   5
            Top             =   480
            Width           =   1455
         End
      End
      Begin MSComDlg.CommonDialog comdlg 
         Left            =   10020
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.ListBox lstInfo 
      Height          =   1230
      Left            =   0
      TabIndex        =   0
      Top             =   5460
      Width           =   9435
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   5160
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Ref.: Shi et al., 2005; Qin et al., 2006; «ÿ≥–÷æµ», 2007; Qin et al., in reviewing"
      Height          =   315
      Index           =   2
      Left            =   180
      TabIndex        =   39
      Top             =   180
      Width           =   10875
   End
End
Attribute VB_Name = "frmFuzzySlpInfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FUNCTION_TYPE_BELL = "Bell-shaped"
Const FUNCTION_TYPE_Z = "Z-shaped"
Const FUNCTION_TYPE_S = "S-shaped"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_bRunning As Boolean
Dim m_strBasePath As String
Dim m_strFilePre As String
Dim m_pBaseGRID As clsGrid

Dim PrototypeTag As Integer

Private Sub cmdAddGRID_Click()
   Dim strBaseDEM As String, i As Integer
   
   If m_bRunning Then Exit Sub
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   strBaseDEM = GetFileName(comdlg, True, , ".asc")
   If strBaseDEM = "" Then Exit Sub
   With lstSrcGRID
      .AddItem strBaseDEM
   
      For i = 0 To .ListCount - 2
         If .List(i) = strBaseDEM Then
            .ListIndex = .ListCount - 1
            MsgBox "Source GRIDs are repeatly assigned" & vbCrLf & strBaseDEM, vbInformation, APP_TITLE
         End If
      Next
   End With
End Sub

Private Sub cmdDelGRID_Click()
   Dim i As Integer
   With lstSrcGRID
      For i = .ListCount - 1 To 0 Step -1
         If .Selected(i) Then .RemoveItem (i)
      Next
      .Refresh
   End With
End Sub

Private Sub cmdGotoSetPara_Click()

   On Error GoTo ErrH
    
   If lstSrcGRID.ListCount <= 0 Or txtPrototypeGRID.Text = "" Then
      Err.Raise Number:=vbObjectError + 513, Description:="Input Error!"
   End If
   With txtPrototypeTag
      If IsNumeric(.Text) Then
         PrototypeTag = CInt(.Text)
         If PrototypeTag <> CDbl(.Text) Then
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameter: Prototype Tag"
         End If
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in parameter: Prototype Tag"
      End If
   End With
   
   Dim iCol As Integer, iRow As Integer
   Dim iRowCount As Integer, iColCount As Integer
   Dim strFile As String, strPath As String, strName As String, strPrev As String, strSuffix As String
   Dim i As Integer
      
   iRowCount = lstSrcGRID.ListCount
   ' table headings
   With MSFlexGrid1
      .Rows = 1
      .Rows = iRowCount + 1
      For iRow = 1 To iRowCount
         strFile = lstSrcGRID.List(iRow - 1)
         i = InStrRev(strFile, "\")
         strPath = Left(strFile, i)
         strName = Right(strFile, Len(strFile) - i)
         i = InStrRev(strName, ".")
         If i = 0 Then
            strPrev = strName
         Else
            strPrev = Left(strName, i - 1)
         End If
         
         .TextMatrix(iRow, 0) = iRow
         .TextMatrix(iRow, 1) = strPrev
         ' set default value to parameters w1, r1, k1
         .TextMatrix(iRow, 3) = "6":         .TextMatrix(iRow, 4) = "2":         .TextMatrix(iRow, 5) = "0.5"
         .TextMatrix(iRow, 6) = "6":         .TextMatrix(iRow, 7) = "2":         .TextMatrix(iRow, 8) = "0.5"
      Next
   End With
    
   SSTabFunc.Tab = 1
ErrH:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbExclamation, APP_TITLE
   End If
   On Error Resume Next

End Sub

Private Sub cmdPrototypeGRID_Click()
   Dim strBaseDEM As String
   
On Error GoTo ErrH
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
   
   txtPrototypeGRID.Text = strBaseDEM
   
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
   
   txtSaveGRID.Text = m_strBasePath & m_strFilePre & "_Fuzzy.asc"
   txtSaveLog.Text = m_strBasePath & "log-" & m_strFilePre & "_Fuzzy.log"
ErrH:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbExclamation, APP_TITLE
      If Not (m_pBaseGRID Is Nothing) Then Set m_pBaseGRID = Nothing
   End If
   On Error Resume Next
   
End Sub

Private Sub cmdQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub

Private Sub cmdRun_Click()
   Dim pGridT As clsGrid, strPrototypeFile As String
   Dim pGridAttr() As clsGrid, vDataDEM As Variant
   Dim strFile As String, strLogFile As String
   Dim pGridS As clsGrid, strSFile As String
   ' GRID info variables
   Dim dCellSize As Double
   Dim dNoData As Double, iCols As Integer, iRows As Integer, dXll As Double, dYll As Double
   'parameters
   Dim lPrototypeNum As Long
   Dim dParaR As Double
   Dim dParaW1() As Double, dParaW2() As Double, dParaR1() As Double, dParaR2() As Double, dParaK1() As Double, dParaK2() As Double
   ' algorithm
   Dim iVLyrNum As Integer, iVLyr As Integer, iLyr As Integer
   Dim str As String
   Dim iCol As Integer, iRow As Integer, iColProc As Integer, iRowProc As Integer, n As Integer
   Dim dSijtv As Double, dSijt As Double, dDist_ijt As Double
   Dim vSt As Variant, vTemp As Variant, vDist As Variant
   ' var for Inverse distance weighted by Surface Distance
   Dim boolSurfaceDist As Boolean
   Dim iColDist As Integer, iRowDist As Integer, iSurfaceStepNum As Integer, iSurfaceStep As Integer
   Dim dSurfaceDist As Double, dGridXStep As Double, dGridYStep As Double
   Dim dLastElev As Double, dCurElev As Double
      
   If MsgBox("Start a long running?", vbQuestion + vbOKCancel + vbDefaultButton2, "Fuzzy Quantification of Slope Positions") = vbCancel Then Exit Sub
   cmdRun.Enabled = False
   Me.MousePointer = 11
   m_bRunning = True
   
On Error GoTo ErrH
   ' get parameters
   iVLyrNum = MSFlexGrid1.Rows - 1
   boolSurfaceDist = Not optDistWeightType.Value
   strPrototypeFile = Trim(txtPrototypeGRID.Text)
   strSFile = Trim(txtSaveGRID.Text)
   strLogFile = Trim(txtSaveLog.Text)
   If iVLyrNum < 1 Or strSFile = "" Or strPrototypeFile = "" Then
      Err.Raise Number:=vbObjectError + 513, Description:="Assign GRID wrong!"
   End If
   With txtPrototypeTag
      If IsNumeric(.Text) Then
         PrototypeTag = CInt(.Text)
         If PrototypeTag <> CDbl(.Text) Then Err.Raise Number:=vbObjectError + 513, Description:="Error in parameter: Prototype Tag"
      Else
         Err.Raise Number:=vbObjectError + 513, Description:="Error in parameter: Prototype Tag"
      End If
   End With
   With txtParaR
      If IsNumeric(.Text) Then
         dParaR = CDbl(.Text)
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in parameter: Prototype Tag"
      End If
   End With
   
   ' get Prototype (naming T)
   Set pGridT = New clsGrid
   With pGridT
      .LoadAscGrid strPrototypeFile, True
      dNoData = .NoData_Value
      dCellSize = .CellSize
      iCols = .nCols:      iRows = .nRows
      dXll = .xllcorner: dYll = .yllcorner
   End With
   
   With lstInfo
      .Clear
      .AddItem Date & " " & Time()
      .AddItem "Prototype Layer" & ": " & strPrototypeFile
      .AddItem "  Rows*Cols: " & iRows & " * " & iCols
      .AddItem "  Cell size: " & dCellSize
      .AddItem "  Xllcorner: " & dXll
      .AddItem "  Yllcorner: " & dYll
      .AddItem "  NoData: " & dNoData
      .AddItem "  Prototype position tag = " & PrototypeTag
   End With
   DoEvents
   
   ReDim dParaW1(1 To iVLyrNum)
   ReDim dParaW2(1 To iVLyrNum)
   ReDim dParaR1(1 To iVLyrNum)
   ReDim dParaR2(1 To iVLyrNum)
   ReDim dParaK1(1 To iVLyrNum)
   ReDim dParaK2(1 To iVLyrNum)
   ReDim pGridAttr(1 To iVLyrNum)
   For iVLyr = 1 To iVLyrNum
      ' get parameters of current attribute GRID layer
      With MSFlexGrid1
         str = .TextMatrix(iVLyr, 1)
         strFile = lstSrcGRID.List(iVLyr - 1)
         lstInfo.AddItem "Terrain Attribute Layer No. " & iVLyr & ": " & str & "(" & strFile & ")"
         dParaW1(iVLyr) = CDbl(.TextMatrix(iVLyr, 3)): dParaR1(iVLyr) = CDbl(.TextMatrix(iVLyr, 4)): dParaK1(iVLyr) = CDbl(.TextMatrix(iVLyr, 5))
         dParaW2(iVLyr) = CDbl(.TextMatrix(iVLyr, 6)): dParaR2(iVLyr) = CDbl(.TextMatrix(iVLyr, 7)): dParaK2(iVLyr) = CDbl(.TextMatrix(iVLyr, 8))
         
         If dParaW1(iVLyr) = 0 Or dParaW2(iVLyr) = 0 Or dParaK1(iVLyr) <= 0 Or dParaK2(iVLyr) <= 0 Then
            Err.Raise Number:=vbObjectError + 513, Description:="Parameters Error!"
         Else
            lstInfo.AddItem "  Similarity Function Type: " & .TextMatrix(iVLyr, 2)
            lstInfo.AddItem "  w1=" & dParaW1(iVLyr) & "; r1=" & dParaR1(iVLyr) & "; k1=" & dParaK1(iVLyr)
            lstInfo.AddItem "  w2=" & dParaW2(iVLyr) & "; r2=" & dParaR2(iVLyr) & "; k2=" & dParaK2(iVLyr)
            lstInfo.ListIndex = lstInfo.ListCount - 1
         End If
         
         Set pGridAttr(iVLyr) = New clsGrid
         pGridAttr(iVLyr).LoadAscGrid strFile
         DoEvents
      End With
   Next
      
   lstInfo.AddItem "Decay factor: r = " & dParaR
   If boolSurfaceDist Then
'      str = cboDEMLyr.Text
'      vDataDEM = pBlockDEM.SafeArray(iBandDEM)
'      lstInfo.AddItem "Inverse distance weighted by Surface distance. DEM: " & str
   Else
      lstInfo.AddItem "Inverse distance weighted by Euclidean distance."
      ReDim vDist(0 To iCols - 1, 0 To iRows - 1)
      For iCol = 0 To iCols - 1
         For iRow = 0 To iRows - 1
            vDist(iCol, iRow) = (Sqr(iCol ^ 2 + iRow ^ 2) * dCellSize)
            If vDist(iCol, iRow) <> 0# Then vDist(iCol, iRow) = vDist(iCol, iRow) ^ (-dParaR)
         Next
      Next
   End If
   lstInfo.AddItem "    Computing similarity ..."
   lstInfo.ListIndex = lstInfo.ListCount - 1
   DoEvents
   
   Set pGridS = New clsGrid
   pGridS.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData
   
   ReDim vSt(0 To iCols - 1, 0 To iRows - 1)
   ReDim vTemp(0 To iCols - 1, 0 To iRows - 1)
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
          If (pGridT.Cell(iCol, iRow) = dNoData) Then
            pGridS.Cell(iCol, iRow) = dNoData
         Else
            pGridS.Cell(iCol, iRow) = 0#
         End If
         vTemp(iCol, iRow) = 0#
      Next
   Next
      
On Error GoTo NextAttrLyr
   lPrototypeNum = 0
   For iColProc = 0 To iCols - 1
      For iRowProc = 0 To iRows - 1
         If pGridT.Cell(iColProc, iRowProc) = PrototypeTag Then
            lPrototypeNum = lPrototypeNum + 1
            'dSijt = MAX_SINGLE
            For iCol = 0 To iCols - 1
               For iRow = 0 To iRows - 1
                  If (pGridT.Cell(iCol, iRow) = dNoData) Then
                     vSt(iCol, iRow) = dNoData
                  Else
                     vSt(iCol, iRow) = MAX_SINGLE
                  End If
               Next
            Next
            
            For iVLyr = 1 To iVLyrNum
               For iCol = 0 To iCols - 1
                  For iRow = 0 To iRows - 1
                     If (pGridAttr(iVLyr).Cell(iCol, iRow) = pGridAttr(iVLyr).NoData_Value) Then
                        dSijtv = dNoData
                     Else
                        If pGridAttr(iVLyr).Cell(iCol, iRow) = pGridAttr(iVLyr).Cell(iColProc, iRowProc) Then
                           dSijtv = 1#
                        ElseIf pGridAttr(iVLyr).Cell(iCol, iRow) < pGridAttr(iVLyr).Cell(iColProc, iRowProc) Then
                           If dParaK1(iVLyr) = 1# Then
                              dSijtv = 1#
                           Else
                              dSijtv = Exp(((Abs(pGridAttr(iVLyr).Cell(iCol, iRow) _
                                          - pGridAttr(iVLyr).Cell(iColProc, iRowProc)) _
                                          / dParaW1(iVLyr)) ^ dParaR1(iVLyr)) * Log(dParaK1(iVLyr)))
                           End If
                        Else
                           If dParaK2(iVLyr) = 1# Then
                              dSijtv = 1#
                           Else
                              dSijtv = Exp(((Abs(pGridAttr(iVLyr).Cell(iCol, iRow) _
                                          - pGridAttr(iVLyr).Cell(iColProc, iRowProc)) _
                                          / dParaW2(iVLyr)) ^ dParaR2(iVLyr)) * Log(dParaK2(iVLyr)))
                           End If
                        End If
                        If vSt(iCol, iRow) > dSijtv Then vSt(iCol, iRow) = dSijtv
                     End If
                  Next
               Next
NextAttrLyr:
               If Err.Number <> 0 Then
                  lstInfo.AddItem "!!Error: " & Err.Description
               End If
            Next
            
            If boolSurfaceDist Then
'               For iCol = 0 To iCols - 1
'                     For iRow = 0 To iRows - 1
'                        If (pGridS.Cell(iCol, iRow) <> dNoData And vSt(iCol, iRow) <> dNoData) Then
'                           If iColProc <> iCol Or iRowProc <> iRow Then
'                              dDist_ijt = Sqr((iColProc - iCol) ^ 2 + (iRowProc - iRow) ^ 2)
'                              iSurfaceStepNum = Int(dDist_ijt)
'                              dGridXStep = (iCol - iColProc) / dDist_ijt: dGridYStep = (iRow - iRowProc) / dDist_ijt
'
'                              dSurfaceDist = 0#
'                              dLastElev = vDataDEM(iColProc, iRowProc)
'                              For iSurfaceStep = 1 To iSurfaceStepNum
'                                 iColDist = Int(iColProc + iSurfaceStep * dGridXStep)
'                                 iRowDist = Int(iRowProc + iSurfaceStep * dGridYStep)
'                                 If vDataDEM(iColDist, iRowDist) = dNoData Then
'                                    dCurElev = dLastElev
'                                 Else
'                                    dCurElev = vDataDEM(iColDist, iRowDist)
'                                 End If
'                                 dSurfaceDist = dSurfaceDist + Sqr(dCellSize ^ 2 + (dCurElev - dLastElev) ^ 2)
'                                 dLastElev = dCurElev
'                              Next
'                              If iSurfaceStepNum < dDist_ijt Then
'                                 dSurfaceDist = dSurfaceDist + Sqr((dCellSize * (dDist_ijt - iSurfaceStepNum)) ^ 2 + (vDataDEM(iCol, iRow) - dLastElev) ^ 2)
'                              End If
'                              dDist_ijt = dSurfaceDist ^ (-dParaR)
'                              pGridS.Cell(iCol, iRow) = pGridS.Cell(iCol, iRow) + dDist_ijt * vSt(iCol, iRow)
'                              vTemp(iCol, iRow) = vTemp(iCol, iRow) + dDist_ijt
'                           End If
'                        End If
'                  Next
'               Next
            Else  ' Euclidean distance
               For iCol = 0 To iCols - 1
                  For iRow = 0 To iRows - 1
                     If (pGridS.Cell(iCol, iRow) <> dNoData And vSt(iCol, iRow) <> dNoData) Then
                        If iColProc <> iCol Or iRowProc <> iRow Then
                           dDist_ijt = vDist(Abs(iColProc - iCol), Abs(iRowProc - iRow)) 'dDist_ijt=(Sqr((iColProc - iCol) ^ 2 + (iRowProc - iRow) ^ 2) * dCellSize) ^ (-dParaR)
                           pGridS.Cell(iCol, iRow) = pGridS.Cell(iCol, iRow) + dDist_ijt * vSt(iCol, iRow)
                           vTemp(iCol, iRow) = vTemp(iCol, iRow) + dDist_ijt
                        End If
                     End If
                  Next
               Next
            End If
         End If
         DoEvents
      Next
      
      If iColProc Mod 10 = 0 Then
         SetProgressBarValue Int((iColProc + 1) * 100# / iCols)
         DoEvents
      End If
   Next
   
   vSt = Empty
   vDist = Empty
   
On Error GoTo ErrH
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pGridT.Cell(iCol, iRow) = PrototypeTag Then
            pGridS.Cell(iCol, iRow) = 1#
         Else
            If (pGridS.Cell(iCol, iRow) <> dNoData And vTemp(iCol, iRow) <> 0#) Then
               pGridS.Cell(iCol, iRow) = pGridS.Cell(iCol, iRow) / vTemp(iCol, iRow)
            Else
               pGridS.Cell(iCol, iRow) = pGridS.NoData_Value
            End If
         End If
      Next
   Next
   SetProgressBarValue 100
   lstInfo.AddItem "Finished computation: " & Date & " " & Time()
   lstInfo.AddItem "Done! " & Date & " " & Time()
   lstInfo.AddItem "Total " & lPrototypeNum & " prototype positions were applied into computing similarity."
   lstInfo.ListIndex = lstInfo.ListCount - 1
   DoEvents
   
   '–¥»ÎGRID
   If pGridS.SaveAscGrid(strSFile, , 4) Then
'      MsgBox "Completed. Save result GRID: " & vbCrLf & strSFile, vbInformation, APP_TITLE
      lstInfo.AddItem "Completed. Save similarity GRID: " & strSFile
      lstInfo.ListIndex = lstInfo.ListCount - 1
      DoEvents
   Else
      Err.Raise vbObjectError + 513, , "Failed to save similarity GRID: " & strSFile
   End If
   
   'output log file
   If WriteLogFile(strLogFile) Then
'      MsgBox "Completed. Save Log file: " & vbCrLf & strLogFile, vbInformation, APP_TITLE
      lstInfo.AddItem "Save Log file: " & strLogFile
      lstInfo.ListIndex = lstInfo.ListCount - 1
      DoEvents
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ErrH:
   cmdRun.Enabled = True
   Me.MousePointer = 0
   m_bRunning = False
   SetProgressBarValue 0
   
   'Release memory
   Set pGridS = Nothing
   vTemp = Empty
'   If boolSurfaceDist Then
''      vDataDEM = Empty
'   Else
      vDist = Empty
'   End If
   Set pGridT = Nothing
   For iVLyr = 1 To iVLyrNum
      Set pGridAttr(iVLyr) = Nothing
   Next
   
   If Err.Number <> 0 Then
      lstInfo.AddItem "!!!Error: " & Err.Description
      MsgBox Err.Description, vbExclamation, APP_TITLE
   Else
      MsgBox "Done!", vbInformation, APP_TITLE
   End If
End Sub

Private Function WriteLogFile(outFile As String) As Boolean
On Error GoTo ErrH
   Dim lLine As Long
   Dim fs As FileSystemObject
   'Dim a As FileStream
   Dim ts As TextStream
         
   WriteLogFile = False
   If outFile = "" Then Exit Function
   Set fs = New FileSystemObject 'CreateObject("Scripting.FileSystemObject")
   'Set a = fs.CreateTextFile(outFile, True)
   Set ts = fs.OpenTextFile(outFile, ForWriting, True, TristateUseDefault)
   
   For lLine = 0 To lstInfo.ListCount - 1
      ts.WriteLine (lstInfo.List(lLine))
   Next
   ts.Close
   WriteLogFile = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
   'Set a = Nothing
   Set fs = Nothing
End Function

Private Sub cmdSaveGRID_Click()
   Dim strFile As String
   If m_bRunning Then Exit Sub
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveGRID.Text = strFile
End Sub

Private Sub cmdSavelog_Click()
   Dim strFile As String
   If m_bRunning Then Exit Sub
   strFile = GetSaveFileName("Save Log", ".log")
   If strFile <> "" Then txtSaveLog.Text = strFile
End Sub

Private Sub Form_Load()
   'initialize var
   m_bRunning = False
   SetProgressBarValue 0
   lstSrcGRID.Clear
   lstInfo.Clear
   With MSFlexGrid1
      .AllowUserResizing = 1  'flexResizeColumns
      '.Clear
      .Cols = 9
      .Rows = 1
      .FixedCols = 2
      '.FixedRows = 1
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Attr."
      .TextMatrix(0, 2) = "Func. Shape"
      .TextMatrix(0, 3) = "w1"
      .TextMatrix(0, 4) = "r1"
      .TextMatrix(0, 5) = "k1"
      .TextMatrix(0, 6) = "w2"
      .TextMatrix(0, 7) = "r2"
      .TextMatrix(0, 8) = "k2"
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

Private Function GetSaveFileName(Optional sDialogTitle As String = "Save GRID", Optional sSuffix As String = ".asc") As String
   comdlg.DialogTitle = sDialogTitle
   comdlg.FileName = ""
   GetSaveFileName = GetFileName(comdlg, False, , sSuffix)
End Function

Private Sub MSFlexGrid1_Click()
   Dim iRowCount As Integer, iColCount As Integer
   Dim iCol As Integer, iRow As Integer
   Dim s As String
    
   With MSFlexGrid1
      iRowCount = .Rows - 1
      iColCount = .Cols - 1
      iRow = .Row
      iCol = .Col
      If iRow = 0 Or iCol <= 1 Then
         Exit Sub
      ElseIf iCol = 2 Then
         s = InputBox(.TextMatrix(0, iCol) & vbCrLf & "1: " & FUNCTION_TYPE_Z & vbCrLf & "2: " & FUNCTION_TYPE_BELL _
                                 & vbCrLf & "3: " & FUNCTION_TYPE_S, "Parameter for Similarity Function", "1")
         Select Case s
         Case "1"
            .TextMatrix(iRow, iCol) = FUNCTION_TYPE_Z
            ' para k1=1
            .TextMatrix(iRow, 5) = "1"
            .TextMatrix(iRow, 3) = "1":   .TextMatrix(iRow, 4) = "0" ' not necessary
         Case "2"
            .TextMatrix(iRow, iCol) = FUNCTION_TYPE_BELL
         Case "3"
            .TextMatrix(iRow, iCol) = FUNCTION_TYPE_S
            ' para k2=1
            .TextMatrix(iRow, 8) = "1"
            .TextMatrix(iRow, 6) = "1":   .TextMatrix(iRow, 7) = "0" ' not necessary
         Case Else
            .TextMatrix(iRow, iCol) = ""
         End Select
      Else
         s = InputBox(.TextMatrix(0, iCol), "Parameter for Similarity Function", .TextMatrix(iRow, iCol))
         If IsNumeric(s) Then
            .TextMatrix(iRow, iCol) = s
         ElseIf s <> "" Then
            MsgBox "Wrong parameter!", vbExclamation, APP_TITLE
         End If
      End If
   End With
End Sub
