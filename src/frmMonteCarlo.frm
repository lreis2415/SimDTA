VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMonteCarlo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monte Carlo modeling"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   255
      Left            =   0
      TabIndex        =   74
      Top             =   8640
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSaveLog 
      Caption         =   "SaveLog"
      Height          =   375
      Left            =   1320
      TabIndex        =   27
      Top             =   6120
      Width           =   7515
   End
   Begin VB.CheckBox chkSaveLog 
      Caption         =   "Save log."
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   6120
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   3480
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstLog 
      Height          =   1860
      Left            =   0
      TabIndex        =   3
      Top             =   6600
      Width           =   8895
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   435
      Left            =   4920
      TabIndex        =   2
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   435
      Left            =   1200
      TabIndex        =   1
      Top             =   9000
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTabMonteCarlo 
      Height          =   6075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10716
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Para for Creating Error Surface"
      TabPicture(0)   =   "frmMonteCarlo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frameBaseDEM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameErr"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Output Error Surface"
      TabPicture(1)   =   "frmMonteCarlo.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frameErrOutput"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frameErrOutput 
         Caption         =   "Output Error Surface"
         Height          =   5655
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   8655
         Begin VB.CheckBox chkOutFile 
            Caption         =   "TWI(SCA-md, max downslope)"
            Height          =   435
            Index           =   12
            Left            =   120
            TabIndex        =   73
            Top             =   5160
            Width           =   1635
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   12
            Left            =   1800
            TabIndex        =   72
            Top             =   5160
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "TWI(SCA-md, slope)"
            Height          =   435
            Index           =   11
            Left            =   120
            TabIndex        =   71
            Top             =   4755
            Width           =   1695
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   11
            Left            =   1800
            TabIndex        =   70
            Top             =   4800
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "TWI(Quinn)"
            Height          =   435
            Index           =   10
            Left            =   120
            TabIndex        =   69
            Top             =   4440
            Width           =   1635
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   10
            Left            =   1800
            TabIndex        =   68
            Top             =   4440
            Width           =   6735
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   9
            Left            =   1800
            TabIndex        =   66
            Top             =   4080
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "SCA(MFD-md)"
            Height          =   435
            Index           =   8
            Left            =   120
            TabIndex        =   65
            Top             =   3720
            Width           =   1635
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   8
            Left            =   1800
            TabIndex        =   64
            Top             =   3720
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "SCA(MFD-Quinn)"
            Height          =   435
            Index           =   7
            Left            =   120
            TabIndex        =   63
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   7
            Left            =   1800
            TabIndex        =   62
            Top             =   3360
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "Max. Downslope"
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   61
            Top             =   3000
            Width           =   1635
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   6
            Left            =   1800
            TabIndex        =   60
            Top             =   3000
            Width           =   6735
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   5
            Left            =   1800
            TabIndex        =   58
            Top             =   2640
            Width           =   6735
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   4
            Left            =   1800
            TabIndex        =   56
            Top             =   2280
            Width           =   6735
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   3
            Left            =   1800
            TabIndex        =   54
            Top             =   1680
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "Slope (tanb)"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   53
            Top             =   1260
            Width           =   1635
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   52
            Top             =   1260
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "Error Surface"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   49
            Top             =   480
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "DEM + Error"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   48
            Top             =   900
            Width           =   1635
         End
         Begin VB.TextBox txtOutFile 
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   47
            Top             =   900
            Width           =   6735
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "TWI(SCA, slope)"
            Height          =   435
            Index           =   9
            Left            =   120
            TabIndex        =   67
            Top             =   4080
            Width           =   1755
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "Slope (dep-fill)"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   59
            Top             =   2700
            Width           =   1755
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "Fill Dep. (Planchon and Darboux, 2001)"
            Height          =   555
            Index           =   4
            Left            =   120
            TabIndex        =   57
            Top             =   2160
            Width           =   1755
         End
         Begin VB.CheckBox chkOutFile 
            Caption         =   "Curvatures(Prof., plan.,horiz.)"
            Height          =   435
            Index           =   3
            Left            =   120
            TabIndex        =   55
            Top             =   1620
            Width           =   1755
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   8520
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Label Label1 
            Caption         =   "Assign [Path\PrefixName](Auto-assign [_Number.asc])"
            Height          =   255
            Index           =   7
            Left            =   1980
            TabIndex        =   51
            Top             =   240
            Width           =   6195
         End
      End
      Begin VB.Frame frameErr 
         Caption         =   "Monte Carlo"
         Height          =   3615
         Left            =   -74940
         TabIndex        =   20
         Top             =   2340
         Width           =   8775
         Begin VB.CheckBox chkCalcMoranI 
            Caption         =   "Calc Moran's I by Weight Matrix for Spatial Auto-correlation"
            Height          =   255
            Left            =   2280
            TabIndex        =   43
            Top             =   2880
            Value           =   1  'Checked
            Width           =   6435
         End
         Begin VB.ComboBox cboWeightType 
            Height          =   300
            Left            =   2520
            TabIndex        =   42
            Text            =   "Combo1"
            Top             =   3120
            Width           =   2415
         End
         Begin TabDlg.SSTab SSTabAutoCorr 
            Height          =   2655
            Left            =   2160
            TabIndex        =   30
            Top             =   120
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4683
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Moran's I"
            TabPicture(0)   =   "frmMonteCarlo.frx":0038
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label1(11)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "frameMoranI"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Hunter,Goodchild(1997) Rho"
            TabPicture(1)   =   "frmMonteCarlo.frx":0054
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label1(9)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label1(10)"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "frameRho"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Shi et al(2007) Shuffle Rate"
            TabPicture(2)   =   "frmMonteCarlo.frx":0070
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   "Topo-Correlation"
            TabPicture(3)   =   "frmMonteCarlo.frx":008C
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            Begin VB.Frame frameRho 
               Caption         =   "Rho ([0, 0.25])"
               Height          =   1815
               Left            =   -74880
               TabIndex        =   36
               Top             =   720
               Width           =   3015
               Begin VB.CommandButton cmdDelRho 
                  Caption         =   "Remove"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   40
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.CommandButton cmdAddRho 
                  Caption         =   "Add"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   39
                  Top             =   840
                  Width           =   975
               End
               Begin VB.TextBox txtRho 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   38
                  Text            =   "Text1"
                  Top             =   360
                  Width           =   975
               End
               Begin VB.ListBox lstRho 
                  Height          =   1140
                  Left            =   1200
                  MultiSelect     =   1  'Simple
                  Sorted          =   -1  'True
                  TabIndex        =   37
                  Top             =   360
                  Width           =   1695
               End
            End
            Begin VB.Frame frameMoranI 
               Caption         =   "Moran's I ([0,1])"
               Height          =   1815
               Left            =   120
               TabIndex        =   31
               Top             =   720
               Width           =   2955
               Begin VB.CommandButton cmdDelMoranI 
                  Caption         =   "Remove"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   35
                  Top             =   1320
                  Width           =   855
               End
               Begin VB.CommandButton cmdAddMoranI 
                  Caption         =   "Add"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   34
                  Top             =   840
                  Width           =   855
               End
               Begin VB.ListBox lstMoranI 
                  Height          =   1140
                  Left            =   1080
                  MultiSelect     =   1  'Simple
                  Sorted          =   -1  'True
                  TabIndex        =   33
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.TextBox txtMoranI 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   32
                  Text            =   "Text1"
                  Top             =   360
                  Width           =   855
               End
            End
            Begin VB.Label Label1 
               Caption         =   "Note: Moran's I list must be in order!"
               Height          =   555
               Index           =   11
               Left            =   3120
               TabIndex        =   45
               Top             =   1080
               Width           =   2955
            End
            Begin VB.Label Label1 
               Caption         =   "Note: Rho list must be in order!"
               Height          =   375
               Index           =   10
               Left            =   -71760
               TabIndex        =   44
               Top             =   2100
               Width           =   2955
            End
            Begin VB.Label Label1 
               Height          =   1275
               Index           =   9
               Left            =   -71760
               TabIndex        =   41
               Top             =   840
               Width           =   2355
            End
         End
         Begin VB.ComboBox cboErrType 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton optErrLyrCount 
            Caption         =   "Realization by Rule"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   2460
            Width           =   2055
         End
         Begin VB.TextBox txtErrLyrCount 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Text            =   "30"
            Top             =   2040
            Width           =   1875
         End
         Begin VB.OptionButton optErrLyrCount 
            Caption         =   "Realization Count"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   1620
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.TextBox txtRMSE 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Text            =   "1"
            Top             =   1320
            Width           =   1875
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   2160
            X2              =   2160
            Y1              =   2880
            Y2              =   3480
         End
         Begin VB.Label Label1 
            Caption         =   "Error Type"
            Height          =   315
            Index           =   8
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "RMSE/Range (""Random"")"
            Height          =   315
            Index           =   6
            Left            =   180
            TabIndex        =   22
            Top             =   960
            Width           =   2115
         End
      End
      Begin VB.Frame frameBaseDEM 
         Caption         =   "Base DEM and File Head"
         Height          =   1875
         Left            =   -74940
         TabIndex        =   4
         Top             =   420
         Width           =   8655
         Begin VB.CommandButton cmdBaseDEM 
            Caption         =   "BaseDEM"
            Height          =   375
            Left            =   1260
            TabIndex        =   7
            Top             =   240
            Width           =   7335
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   1155
            Left            =   60
            TabIndex        =   6
            Top             =   660
            Width           =   8535
            Begin VB.TextBox txtFileHead 
               Height          =   315
               Index           =   5
               Left            =   6780
               TabIndex        =   19
               Text            =   "-9999"
               Top             =   600
               Width           =   1155
            End
            Begin VB.TextBox txtFileHead 
               Height          =   315
               Index           =   4
               Left            =   6780
               TabIndex        =   17
               Text            =   "1"
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox txtFileHead 
               Height          =   315
               Index           =   3
               Left            =   3180
               TabIndex        =   15
               Text            =   "0"
               Top             =   600
               Width           =   2115
            End
            Begin VB.TextBox txtFileHead 
               Height          =   315
               Index           =   2
               Left            =   3180
               TabIndex        =   13
               Text            =   "0"
               Top             =   240
               Width           =   2115
            End
            Begin VB.TextBox txtFileHead 
               Height          =   315
               Index           =   1
               Left            =   780
               TabIndex        =   11
               Top             =   600
               Width           =   1155
            End
            Begin VB.TextBox txtFileHead 
               Height          =   315
               Index           =   0
               Left            =   780
               TabIndex        =   9
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   5
               Left            =   5520
               TabIndex        =   18
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   4
               Left            =   5520
               TabIndex        =   16
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   3
               Left            =   2160
               TabIndex        =   14
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   2
               Left            =   2160
               TabIndex        =   12
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   600
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.CheckBox chkBaseDEM 
            Caption         =   "Base DEM"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmMonteCarlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ErrType_Random = "Random"
Const ErrType_RandomNormal = "Normal-Random"
Const ErrType_AutoCorr_MoranI = "AutoCorr-Moran's I"
Const ErrType_AutoCorr_Rho = "AutoCorr-Rho (Hunter&Goodchild, 97)"
Const ErrType_AutoCorr_Shuffle = "AutoCorr-Shuffle Rate (Shi et al, 07)"
Const ErrType_TopoCorr = "Terrain-Correlation"

Const Weight_Binary = "Binary Matrix"
Const Weight_Decreasing = "Decreasing Function"

Const OutputSurfaceCount = 13

Dim m_bRunning As Boolean
Dim m_strBasePath As String
Dim m_strFilePre As String
Dim m_pBaseDEM As clsGrid, m_pErr As clsGrid

Dim m_bLoadBaseDEM As Boolean, m_bSaveLog As Boolean
Dim m_bSaveErr(0 To OutputSurfaceCount - 1) As Boolean, m_strErr(0 To OutputSurfaceCount - 1) As String
Dim m_strErrType As String
Dim m_dRMSE As Double
'
' Goodchild and Openshaw (1980): Calc Moran's I
'iWeightType = 0 (Weight_Binary: Rook's-case neighbor: weight=1; else, weight=0); 1 (Weight_Decreasing)
'
Private Function Calc_MoranI(pGrid As clsGrid, iWeightType As Integer, _
                           dMoranI As Double, Optional dMean As Double, Optional dSD As Double) As Boolean
                           
   Dim dValue As Double, dSum As Double, dSumSq As Double
   Dim lCount As Long
   Dim vSX As Variant, iCol As Integer, iRow As Integer
   Dim dSW As Double, dCon As Double, dWeight As Double
On Error GoTo ErrH
   Calc_MoranI = False
   ReDim vSX(0 To pGrid.nCols - 1, 0 To pGrid.nRows - 1) As Double
   'calc MEAN, SD
   lCount = 0
   dSum = 0#:  dSumSq = 0#
   For iCol = 0 To pGrid.nCols - 1
      For iRow = 0 To pGrid.nRows - 1
         If pGrid.IsValidCellValue(iCol, iRow, dValue) Then
            dSum = dSum + dValue
            dSumSq = dSumSq + dValue ^ 2
            lCount = lCount + 1
         End If
      Next
   Next
   dMean = dSum / lCount
   dSD = Sqr(dSumSq / lCount - dMean ^ 2)
   
   ' calc Moran's I
   dWeight = 1#
   dSW = 0#
   If iWeightType = 0 Then ' Rook's-case neighbor
      For iCol = 0 To pGrid.nCols - 1
         For iRow = 0 To pGrid.nRows - 1
            If pGrid.IsValidCellValue(iCol, iRow, dValue) Then
               vSX(iCol, iRow) = 0#
               If pGrid.IsValidCellValue(iCol, iRow + 1, dValue) Then
                  dSW = dSW + dWeight
                  vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
               End If
               If pGrid.IsValidCellValue(iCol + 1, iRow, dValue) Then
                  dSW = dSW + dWeight
                  vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
               End If
               If pGrid.IsValidCellValue(iCol, iRow - 1, dValue) Then
                  dSW = dSW + dWeight
                  vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
               End If
               If pGrid.IsValidCellValue(iCol - 1, iRow, dValue) Then
                  dSW = dSW + dWeight
                  vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
               End If
            Else ' NoData
               vSX(iCol, iRow) = dValue
            End If
         Next
      Next
      dCon = 1 / (dSW * dSD ^ 2)
      dSum = 0#
      For iCol = 0 To pGrid.nCols - 1
         For iRow = 0 To pGrid.nRows - 1
            If pGrid.IsValidCellValue(iCol, iRow, dValue) Then
               dSum = dSum + dValue * vSX(iCol, iRow)
            End If
         Next
      Next
      dMoranI = dSum * dCon
   Else
   
   
   
   End If
   vSX = Empty
   Calc_MoranI = True
   Exit Function
ErrH:
   vSX = Empty
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Function

Private Sub cboErrType_LostFocus()
   Select Case cboErrType.Text
   Case ErrType_AutoCorr_MoranI
      SSTabAutoCorr.Tab = 0
      chkCalcMoranI.Value = 1
      cboWeightType.Enabled = IIf((chkCalcMoranI.Value = 1), True, False)
   Case ErrType_AutoCorr_Rho
      SSTabAutoCorr.Tab = 1
   Case ErrType_AutoCorr_Shuffle
      SSTabAutoCorr.Tab = 2
   Case ErrType_TopoCorr
      SSTabAutoCorr.Tab = 3
   End Select
End Sub

Private Sub cboWeightType_Change()
   If cboWeightType.ListIndex = 1 Then
      MsgBox "NOT implemented yet."
   End If
End Sub

Private Sub chkBaseDEM_Click()
   Dim bChecked As Boolean
   bChecked = IIf((chkBaseDEM.Value = 1), True, False)
   cmdBaseDEM.Enabled = bChecked
   frameFileHead.Enabled = Not bChecked
   'm_bLoadBaseDEM = bChecked
   If bChecked Then Call cmdBaseDEM_Click
End Sub

Private Sub chkCalcMoranI_Click()
   cboWeightType.Enabled = IIf((chkCalcMoranI.Value = 1), True, False)
End Sub

Private Sub chkOutFile_Click(Index As Integer)
   Dim bChecked As Boolean
   bChecked = IIf((chkOutFile(Index).Value = 1), True, False)
   txtOutFile(Index).Enabled = bChecked
   m_bSaveErr(Index) = bChecked
End Sub

Private Sub chkSaveLog_Click()
   Dim bChecked As Boolean
   bChecked = IIf((chkSaveLog.Value = 1), True, False)
   cmdSaveLog.Enabled = bChecked
   If bChecked Then Call cmdSavelog_Click
   'm_bSaveLog = bChecked
End Sub

Private Sub cmdAddMoranI_Click()
   Dim dI As Double
   Dim i As Integer
On Error GoTo ErrH
   If txtMoranI.Text = "" Then Exit Sub
   If IsNumeric(txtMoranI.Text) Then
      dI = CDbl(txtMoranI.Text)
      If dI < 0 Or dI > 1 Then Err.Raise vbObjectError + 513, , "Moran's I: [0,1]."
      With lstMoranI
         For i = 0 To .ListCount - 1
            If dI = CDbl(.List(i)) Then Err.Raise vbObjectError + 513, , dI & " has been in List of Moran's I."
         Next
         .AddItem Format(dI, "0.0000000")
      End With
   Else
      Err.Raise vbObjectError + 513, "frmMonteCarlo.cmdAddMoranI_Click", "Should be NUMERIC."
   End If
   Exit Sub
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdAddRho_Click()
   Dim dI As Double
   Dim i As Integer
On Error GoTo ErrH
   If txtRho.Text = "" Then Exit Sub
   If IsNumeric(txtRho.Text) Then
      dI = CDbl(txtRho.Text)
      If dI < 0 Or dI > 0.25 Then Err.Raise vbObjectError + 513, , "Rho: [0,0.25]."
      With lstRho
         For i = 0 To .ListCount - 1
            If dI = CDbl(.List(i)) Then Err.Raise vbObjectError + 513, , dI & " has been in List of Rho."
         Next
         .AddItem Format(dI, "0.0000000")
      End With
   Else
      Err.Raise vbObjectError + 513, "frmMonteCarlo.cmdAddMoranI_Click", "Should be NUMERIC."
   End If
   Exit Sub
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdBaseDEM_Click()
   Dim strBaseDEM As String
   
   If m_bRunning Then Exit Sub
   comdlg.DialogTitle = "Open Base DEM"
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
   
   cmdBaseDEM.Caption = strBaseDEM
   txtOutFile(0).Text = m_strBasePath & m_strFilePre & "\e"     'error surface
   txtOutFile(1).Text = m_strBasePath & m_strFilePre & "\eDEM"  ' error+DEM
   txtOutFile(2).Text = m_strBasePath & m_strFilePre & "\eSlp"  ' slope
   txtOutFile(3).Text = m_strBasePath & m_strFilePre & "\eCurv"  'curvatures
   txtOutFile(4).Text = m_strBasePath & m_strFilePre & "\eFil"  'Fill depression with Planchoon & Darbox
   txtOutFile(5).Text = m_strBasePath & m_strFilePre & "\eFilSlp"  'slope after depressions filled
   txtOutFile(6).Text = m_strBasePath & m_strFilePre & "\eFilDslp"  ' max downslope
   txtOutFile(7).Text = m_strBasePath & m_strFilePre & "\eM0"  ' SCA (MFD-Quinn)
   txtOutFile(8).Text = m_strBasePath & m_strFilePre & "\eMd"  ' SCA (MFD-md)
   txtOutFile(9).Text = m_strBasePath & m_strFilePre & "\eM0TWI0"  'TWI (lg(a/tanb)
   txtOutFile(10).Text = m_strBasePath & m_strFilePre & "\eM0TWIQ" 'TWI in Quinn
   txtOutFile(11).Text = m_strBasePath & m_strFilePre & "\eMdTWI0" 'TWI (MFD-md, slope)
   txtOutFile(12).Text = m_strBasePath & m_strFilePre & "\eMdTWI1" 'TWI (MFD-md, max downslope
   cmdSaveLog.Caption = m_strBasePath & "log_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "AMPMhhmm") & ".log"
   
   ' load BaseDEM, read parameters in file head
   If Not (m_pBaseDEM Is Nothing) Then Set m_pBaseDEM = Nothing
   Set m_pBaseDEM = New clsGrid
   With m_pBaseDEM
      .LoadAscGrid strBaseDEM
      txtFileHead(0).Text = .nCols
      txtFileHead(1).Text = .nRows
      txtFileHead(2).Text = .xllcorner
      txtFileHead(3).Text = .yllcorner
      txtFileHead(4).Text = .CellSize
      txtFileHead(5).Text = .NoData_Value
   End With
   
End Sub

Private Sub cmdDelMoranI_Click()
   Dim i As Integer
On Error GoTo ErrH
   With lstMoranI
      If .SelCount <= 0 Then Exit Sub
      i = 0
      Do While i <= .ListCount - 1
         If .Selected(i) Then
            .RemoveItem (i)
         Else
            i = i + 1
         End If
      Loop
   End With
   Exit Sub
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdDelRho_Click()
   Dim i As Integer
On Error GoTo ErrH
   With lstRho
      If .SelCount <= 0 Then Exit Sub
      i = 0
      Do While i <= .ListCount - 1
         If .Selected(i) Then
            .RemoveItem (i)
         Else
            i = i + 1
         End If
      Loop
   End With
   Exit Sub
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub

Private Sub cmdRun_Click()
   Dim iMonteCarloCount As Integer, iMonteCarlo As Integer
   Dim strLog As String, bOutput As Boolean
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim dAutoCorr As Double, arr_dAutoCorr() As Double, iAutoCorrCount As Integer, iAutoCorr As Integer
   Dim i As Integer, iCol As Integer, iRow As Integer
   Dim dTemp As Double, dTemp1 As Double, dTemp2 As Double, strTemp As String, iTemp As Integer
   Dim pGrid As clsGrid
   'for calc Moran's I
   Dim bCalcMoranI As Boolean, iWeightType As Integer   'iWeightType = 0 (Weight_Binary); 1 (Weight_Decreasing)
   Dim dMoranI As Double, dMean As Double, dSD As Double
   Dim dValue As Double, dSum As Double, dSumSq As Double
   Dim lCount As Long
   Dim vSX As Variant
   Dim dSW As Double, dCon As Double, dWeight As Double
   Dim NTRY As Long, NTRYM As Long, NSW As Long, XDiff As Double, DELX As Double, XSUM2 As Double, XNDIFF As Double
   Dim iMCol As Integer, iMRow As Integer, iNCol As Integer, iNRow As Integer
   
On Error GoTo ErrH
   If m_bRunning Then Exit Sub
   m_bRunning = True
   Me.MousePointer = 11
   
   'get parameters and verify them
   m_strErrType = cboErrType.Text
   Select Case m_strErrType
   Case ErrType_AutoCorr_MoranI
      If lstMoranI.ListCount <= 0 Then Err.Raise vbObjectError + 513, , "Assign Moran's I (sequence) firstly."
      If chkCalcMoranI.Value <> 1 Then Err.Raise vbObjectError + 513, , "Must calc Moran's I for " & m_strErrType
      iAutoCorrCount = lstMoranI.ListCount
      ReDim arr_dAutoCorr(0 To iAutoCorrCount - 1)
      For iAutoCorr = 0 To iAutoCorrCount - 1
         arr_dAutoCorr(iAutoCorr) = CDbl(lstMoranI.List(iAutoCorr))
      Next
   Case ErrType_AutoCorr_Rho
      If lstRho.ListCount <= 0 Then Err.Raise vbObjectError + 513, , "Assign Rho (sequence) firstly."
      iAutoCorrCount = lstRho.ListCount
      ReDim arr_dAutoCorr(0 To iAutoCorrCount - 1)
      For iAutoCorr = 0 To iAutoCorrCount - 1
         arr_dAutoCorr(iAutoCorr) = CDbl(lstRho.List(iAutoCorr))
      Next
   Case ErrType_AutoCorr_Shuffle
      Err.Raise vbObjectError + 513, , "Not implemented yet"
      
   Case ErrType_TopoCorr
      Err.Raise vbObjectError + 513, , "Not implemented yet"
      
   Case ErrType_Random
   
   Case ErrType_RandomNormal
   
   Case Else
      Err.Raise vbObjectError + 513, , "Assign Error Type firstly."
   End Select
   
   m_dRMSE = CDbl(txtRMSE.Text)
   If m_dRMSE <= 0# Then Err.Raise vbObjectError + 513, "frmMonteCarlo.cmdRun_Click", "RMSE should be GREATER than zeor."
   iMonteCarloCount = IIf(optErrLyrCount(0).Value, CInt(txtErrLyrCount.Text), 0)
   bCalcMoranI = IIf((chkCalcMoranI.Value = 1), True, False)
   iWeightType = cboWeightType.ListIndex
   
   If Not (IsNumeric(txtFileHead(0).Text) And IsNumeric(txtFileHead(1).Text) And IsNumeric(txtFileHead(2).Text) _
         And IsNumeric(txtFileHead(3).Text) And IsNumeric(txtFileHead(4).Text) And IsNumeric(txtFileHead(5).Text)) Then
      Err.Raise vbObjectError + 513, , "Assign valid parameters firstly."
   End If
   iCols = CInt(txtFileHead(0).Text):      iRows = CInt(txtFileHead(1).Text)
   dXll = CDbl(txtFileHead(2).Text):       dYll = CDbl(txtFileHead(3).Text)
   dCellSize = CDbl(txtFileHead(4).Text):  dNoData = CDbl(txtFileHead(5).Text)
   If iCols < 1 Or iRows < 1 Or dCellSize <= 0 Then
      Err.Raise vbObjectError + 513, , "Assign valid parameters firstly."
   End If
   
   m_bLoadBaseDEM = IIf((chkBaseDEM.Value = 1), True, False)
   m_bSaveLog = IIf((chkSaveLog.Value = 1), True, False)
   If m_bSaveLog Then strLog = cmdSaveLog.Caption
   bOutput = False
   For i = 0 To OutputSurfaceCount - 1
      m_bSaveErr(i) = IIf((chkOutFile(i).Value = 1), True, False)
      m_strErr(i) = txtOutFile(i).Text
      bOutput = bOutput Or m_bSaveErr(i)
   Next
   If Not bOutput Then
      If MsgBox("Run without ANY output?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         m_bRunning = False
         Exit Sub
      End If
   End If
      
   If Not (m_pErr Is Nothing) Then Set m_pErr = Nothing
   Set m_pErr = New clsGrid
   '
   With lstLog
      .Clear
      .AddItem "=== " & Date & " " & Time
      .AddItem "* BaseDEM: " & IIf(m_bLoadBaseDEM, m_pBaseDEM.sAscGridFileName, "[none]")
      .AddItem "  nCols=" & iCols & "; nRows=" & iRows
      .AddItem "  XllCorner=" & dXll & "; YllCorner=" & dYll
      .AddItem "  CellSize=" & dCellSize & "; NoData_Value=" & dNoData
      .AddItem "* Parameters for Monte Carlo"
      .AddItem "  RMSE=" & m_dRMSE
      .AddItem "  Error-Creating Type: " & m_strErrType
      If bCalcMoranI Then
         .AddItem "Type of Weight Matrix for Calc. Moran's I: " & cboWeightType.Text
      End If
      Select Case m_strErrType
      Case ErrType_AutoCorr_MoranI
         strTemp = "  Moran's I: "
         For iAutoCorr = 0 To iAutoCorrCount - 1
            strTemp = strTemp & Format(arr_dAutoCorr(iAutoCorr), "0.0#") & ", "
         Next
         .AddItem strTemp
      Case ErrType_AutoCorr_Rho
         strTemp = "  Rho (Hunter&Goodchild, 1997): "
         For iAutoCorr = 0 To iAutoCorrCount - 1
            strTemp = strTemp & Format(arr_dAutoCorr(iAutoCorr), "0.0#") & ", "
         Next
         .AddItem strTemp
      Case ErrType_AutoCorr_Shuffle
         
      Case ErrType_TopoCorr
         
      End Select
      .AddItem "  Realization Count=" & IIf(optErrLyrCount(0).Value, iMonteCarloCount, "[by rule]")
      .AddItem "=== Creating Error by Monte Carlo ==="
      .ListIndex = .ListCount - 1
   End With
   SetProgressBarValue 0
   
   Randomize   ' Initialize random-number generator.
   m_pErr.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData, 0#
   Set pGrid = New clsGrid
   pGrid.NewGrid iCols, iRows, dXll, dYll, dCellSize, dNoData, 0#
   If m_strErrType = ErrType_AutoCorr_MoranI Then
      ReDim vSX(0 To pGrid.nCols - 1, 0 To pGrid.nRows - 1) As Double
   End If
   For iMonteCarlo = 1 To iMonteCarloCount
      Select Case m_strErrType
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Case ErrType_Random, ErrType_RandomNormal
         If m_strErrType = ErrType_Random Then
            ' complete random error surface: [-m_dRMSE, m_dRMSE]
            For iCol = 0 To iCols - 1
               For iRow = 0 To iRows - 1
                  m_pErr.vData(iCol, iRow) = (Rnd() - 0.5) * 2 * m_dRMSE
               Next
            Next
         ElseIf m_strErrType = ErrType_RandomNormal Then
            ' random error surface with Normal distribution: N (0, m_dRMSE)
            For iCol = 0 To iCols - 1
               For iRow = 0 To iRows - 1
                  dTemp1 = Rnd(): dTemp2 = Rnd()
                  dTemp = Sqr(-2 * Log(dTemp1)) * Cos(2 * PI * dTemp2)
                  'm_pErr.vData(iCol, iRow) = dTemp * m_dRMSE   ' cannot set data of variant array inside class from outside
                  m_pErr.Cell(iCol, iRow) = dTemp * m_dRMSE
               Next
            Next
         End If
         If bCalcMoranI Then Calc_MoranI m_pErr, iWeightType, dMoranI, dMean, dSD
         
         With lstLog
            .AddItem "* No." & iMonteCarlo & " Realization"
            If bCalcMoranI Then .AddItem "  Moran's I = " & dMoranI & "; Mean = " & dMean & "; SD = " & dSD
            .ListIndex = .ListCount - 1
         End With
         OutputErrGRID iMonteCarlo
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Case ErrType_AutoCorr_MoranI ' Goodchild & Openshaw (1980)
      
         ' random error surface with Normal distribution: N (0, m_dRMSE)
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               dTemp1 = Rnd(): dTemp2 = Rnd()
               dTemp = Sqr(-2 * Log(dTemp1)) * Cos(2 * PI * dTemp2)
               pGrid.Cell(iCol, iRow) = dTemp * m_dRMSE
            Next
         Next
         ' make it auto-correlation sequence - pGrid
         'calc MEAN, SD
         lCount = 0
         dSum = 0#:  dSumSq = 0#
         For iCol = 0 To pGrid.nCols - 1
            For iRow = 0 To pGrid.nRows - 1
               If pGrid.IsValidCellValue(iCol, iRow, dValue) Then
                  dSum = dSum + dValue
                  dSumSq = dSumSq + dValue ^ 2
                  lCount = lCount + 1
               End If
            Next
         Next
         dMean = dSum / lCount
         dSD = Sqr(dSumSq / lCount - dMean ^ 2)
         
         ' calc SUMS and INITIAL AUTOCORRELATION (Moran's I)
         dWeight = 1#
         dSW = 0#
         If iWeightType = 0 Then ' Rook's-case neighbor
            For iCol = 0 To pGrid.nCols - 1
               For iRow = 0 To pGrid.nRows - 1
                  If pGrid.IsValidCellValue(iCol, iRow, dValue) Then
                     vSX(iCol, iRow) = 0#
                     If pGrid.IsValidCellValue(iCol, iRow + 1, dValue) Then
                        dSW = dSW + dWeight
                        vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
                     End If
                     If pGrid.IsValidCellValue(iCol + 1, iRow, dValue) Then
                        dSW = dSW + dWeight
                        vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
                     End If
                     If pGrid.IsValidCellValue(iCol, iRow - 1, dValue) Then
                        dSW = dSW + dWeight
                        vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
                     End If
                     If pGrid.IsValidCellValue(iCol - 1, iRow, dValue) Then
                        dSW = dSW + dWeight
                        vSX(iCol, iRow) = vSX(iCol, iRow) + dValue * dWeight
                     End If
                  Else ' NoData
                     vSX(iCol, iRow) = dValue
                  End If
               Next
            Next
            dCon = 1 / (dSW * dSD ^ 2)
            dSum = 0#
            For iCol = 0 To pGrid.nCols - 1
               For iRow = 0 To pGrid.nRows - 1
                  If pGrid.IsValidCellValue(iCol, iRow, dValue) Then
                     dSum = dSum + dValue * vSX(iCol, iRow)
                  End If
               Next
            Next
            dMoranI = dSum * dCon
         Else  ' smooth weight matrix
                  
         End If
         ' SWAP alg. to simulation error surfaces with a list of Moran's I (in UPWARDS order)
         NTRYM = CLng(iCols) * iRows  ' Maximum number of unsuccessful tries at a swap
         NSW = 0     ' Actual number of SWAP
         For iAutoCorr = 0 To iAutoCorrCount - 1
            dAutoCorr = arr_dAutoCorr(iAutoCorr)
            With lstLog
               .AddItem "* No." & iMonteCarlo & " Realization"
               .AddItem "  Target Moran's I = " & dAutoCorr
            End With
            If dAutoCorr <> 0# Then
               'Begin SWAP alg
               NTRY = -1   ' Actual number of unsuccessful tries at a swap
               XDiff = Abs(dMoranI - dAutoCorr)
               Do
                  Do
                     Do
                        iMCol = Int(iCols * Rnd()): iMRow = Int(iRows * Rnd()) ' random int [0, ncols-1]
                        iNCol = Int(iCols * Rnd()): iRow = Int(iRows * Rnd())
                        NTRY = NTRY + 1
                     Loop While (iMCol = iNCol And iMRow = iNRow) Or Not pGrid.IsValidCellValue(iNCol, iNRow) Or Not pGrid.IsValidCellValue(iMCol, iMRow)
                     If iWeightType = 0 Then
                        DELX = 2 * (pGrid.Cell(iNCol, iNRow) - pGrid.Cell(iMCol, iMRow)) _
                              * (vSX(iMCol, iMRow) - vSX(iNCol, iNRow) - (pGrid.Cell(iNCol, iNRow) _
                                 - pGrid.Cell(iMCol, iMRow)) * IIf(Abs(iMCol - iNCol) + Abs(iMRow - iNRow) = 1, 1, 0))
                     Else
                        
                     End If
                     XSUM2 = dSum + DELX
                     XNDIFF = Abs(XSUM2 * dCon - dAutoCorr)
                  Loop While (NTRY <= NTRYM And XNDIFF > XDiff)
                  
                  If XNDIFF <= XDiff Then
                     ' Do swap two cell
                     NSW = NSW + 1
                     XDiff = XNDIFF
                     dSum = XSUM2
                     ' update the SUMs and make the swap
                     '  DO 21 K=1,N
                     '21 SX(K)=SX(K)+(W(K,IM)-W(K,IN))*(X(IN)-X(IM))
                     If iWeightType = 0 Then ' Rook's-case neighbor
                        ' check IM's Rook's-case neighbor which is NOT the IN's Rook's-case neighbor
                        iCol = iMCol: iRow = iMRow + 1
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iNCol) + Abs(iRow - iNRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iNCol, iNRow) - pGrid.Cell(iMCol, iMRow))
                           End If
                        End If
                        iCol = iMCol + 1: iRow = iMRow
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iNCol) + Abs(iRow - iNRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iNCol, iNRow) - pGrid.Cell(iMCol, iMRow))
                           End If
                        End If
                        iCol = iMCol: iRow = iMRow - 1
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iNCol) + Abs(iRow - iNRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iNCol, iNRow) - pGrid.Cell(iMCol, iMRow))
                           End If
                        End If
                        iCol = iMCol - 1: iRow = iMRow
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iNCol) + Abs(iRow - iNRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iNCol, iNRow) - pGrid.Cell(iMCol, iMRow))
                           End If
                        End If
                        ' check IN's Rook's-case neighbor which is NOT the IM's Rook's-case neighbor
                        iCol = iNCol: iRow = iNRow + 1
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iMCol) + Abs(iRow - iMRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iMCol, iMRow) - pGrid.Cell(iNCol, iNRow))
                           End If
                        End If
                        iCol = iNCol + 1: iRow = iNRow
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iMCol) + Abs(iRow - iMRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iMCol, iMRow) - pGrid.Cell(iNCol, iNRow))
                           End If
                        End If
                        iCol = iNCol: iRow = iNRow - 1
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iMCol) + Abs(iRow - iMRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iMCol, iMRow) - pGrid.Cell(iNCol, iNRow))
                           End If
                        End If
                        iCol = iNCol - 1: iRow = iNRow
                        If pGrid.IsValidCell(iCol, iRow) Then
                           If Abs(iCol - iMCol) + Abs(iRow - iMRow) > 1 Then
                              vSX(iCol, iRow) = vSX(iCol, iRow) + (pGrid.Cell(iMCol, iMRow) - pGrid.Cell(iNCol, iNRow))
                           End If
                        End If
                     Else  ' smooth weight matrix
                     
                     End If
                     dTemp = pGrid.Cell(iMCol, iMRow): pGrid.Cell(iMCol, iMRow) = pGrid.Cell(iNCol, iNRow): pGrid.Cell(iNCol, iNRow) = dTemp
                     dMoranI = dSum * dCon
                     'Debug.Print "Moran's I=" & dMoranI & "; Unsuccessfule try number at a SWAP = " & NTRY
                  Else  'NTRY > NTRYM
                     lstLog.AddItem "  ! Unsuccessfule try number at a SWAP exceed the maximum: " & NTRY
                     Exit Do
                  End If
                  NTRY = -1
                  DoEvents
               Loop While XDiff > 0.001 '0.0001
            End If
                        
            With lstLog
               '.AddItem "* No." & iMonteCarlo & " Realization"
               '.AddItem "  Target Moran's I = " & dAutoCorr
               If bCalcMoranI Then
                  .AddItem "  Actual Moran's I = " & dMoranI & "; SWAP Num=" & NSW & "; Mean = " & dMean & "; SD = " & dSD
               End If
               .ListIndex = .ListCount - 1
            End With
            
            For iCol = 0 To iCols - 1
               For iRow = 0 To iRows - 1
                  m_pErr.Cell(iCol, iRow) = pGrid.Cell(iCol, iRow)
               Next
            Next
            
            OutputErrGRID iMonteCarlo, CStr(iAutoCorr)
         Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Case ErrType_AutoCorr_Rho  ' Hunter & Goodchild (1997)
         ' random error surface with Normal distribution: N (0, m_dRMSE)
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               dTemp1 = Rnd(): dTemp2 = Rnd()
               dTemp = Sqr(-2 * Log(dTemp1)) * Cos(2 * PI * dTemp2)
               pGrid.Cell(iCol, iRow) = dTemp * m_dRMSE
            Next
         Next
         ' make it auto-correlation
         For iAutoCorr = 0 To iAutoCorrCount - 1
            dAutoCorr = arr_dAutoCorr(iAutoCorr)
            If dAutoCorr = 0# Then
               For iCol = 0 To iCols - 1
                  For iRow = 0 To iRows - 1
                     m_pErr.Cell(iCol, iRow) = pGrid.vData(iCol, iRow)
                  Next
               Next
            Else
               ' e=rho * W * e + N(0,1)
               For iCol = 0 To iCols - 1
                  For iRow = 0 To iRows - 1
                     dTemp = 0#
                     ' Rook's-case neighbor
                     If pGrid.IsValidCellValue(iCol, iRow + 1, dTemp1) Then dTemp = dTemp + dTemp1 * dAutoCorr
                     If pGrid.IsValidCellValue(iCol + 1, iRow, dTemp1) Then dTemp = dTemp + dTemp1 * dAutoCorr
                     If pGrid.IsValidCellValue(iCol, iRow - 1, dTemp1) Then dTemp = dTemp + dTemp1 * dAutoCorr
                     If pGrid.IsValidCellValue(iCol - 1, iRow, dTemp1) Then dTemp = dTemp + dTemp1 * dAutoCorr
                     dTemp1 = Rnd(): dTemp2 = Rnd()
                     dTemp = dTemp + Sqr(-2 * Log(dTemp1)) * Cos(2 * PI * dTemp2)
                     m_pErr.Cell(iCol, iRow) = dTemp
                  Next
               Next
            End If
            If bCalcMoranI Then Calc_MoranI m_pErr, iWeightType, dMoranI, dMean, dSD
                        
            With lstLog
               .AddItem "* No." & iMonteCarlo & " Realization"
               .AddItem "  rho=" & dAutoCorr
               If bCalcMoranI Then .AddItem "  Moran's I = " & dMoranI & "; Mean = " & dMean & "; SD = " & dSD
               .ListIndex = .ListCount - 1
               '.Refresh
            End With
            OutputErrGRID iMonteCarlo
         Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Case ErrType_AutoCorr_Shuffle
      
      
      
         
      Case ErrType_TopoCorr
      
      
      
            
      End Select
      SetProgressBarValue Int(iMonteCarlo * 100# / iMonteCarloCount)
   Next
      
   lstLog.AddItem "=== Finished at " & Date & " " & Time
   lstLog.ListIndex = lstLog.ListCount - 1
   If m_bSaveLog Then
      Dim fs As New FileSystemObject
      Dim ts As TextStream
      
      If fs.FileExists(strLog) Then
         If MsgBox(strLog & vbCrLf & "exists. Overwrite it?", vbQuestion + vbYesNo + vbDefaultButton2, "Overwrite file?") = vbNo Then
            strLog = GetFileName(comdlg, False, , ".log")
         End If
      End If
      Set ts = fs.OpenTextFile(strLog, ForWriting, True, TristateUseDefault)
      For i = 0 To lstLog.ListCount - 1
         ts.WriteLine lstLog.List(i)
      Next
      ts.Close
      Set fs = Nothing
   End If
   
   MsgBox "Done! " & IIf(m_bSaveLog, vbCrLf & "Log: " & strLog, ""), vbInformation
ErrH:
   Me.MousePointer = 0
   ' release memory
   vSX = Empty
   Set pGrid = Nothing
   m_bRunning = False
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdSavelog_Click()
   Dim str As String
   comdlg.DialogTitle = "Save Log file as"
   comdlg.FileName = ""
   str = GetFileName(comdlg, True, , ".log")
   If str = "" Then Exit Sub
   cmdSaveLog.Caption = str
End Sub

Private Sub Form_Load()
   Dim bChecked As Boolean
   Dim i As Integer
   
   'initalize CommonDialog
   
   'initialize Interface
   SSTabMonteCarlo.Tab = 0
   
   bChecked = IIf((chkBaseDEM.Value = 1), True, False)
   cmdBaseDEM.Enabled = bChecked
   frameFileHead.Enabled = Not bChecked
   
   txtErrLyrCount.Enabled = optErrLyrCount(0).Value
   With cboErrType
      .Clear
      .AddItem ErrType_Random
      .AddItem ErrType_RandomNormal
      .AddItem ErrType_AutoCorr_MoranI
      .AddItem ErrType_AutoCorr_Rho
      .AddItem ErrType_AutoCorr_Shuffle
      .AddItem ErrType_TopoCorr
      .ListIndex = -1
   End With
   txtMoranI.Text = 0.5
   lstMoranI.Clear
   'lstMoranI.Sorted = True   'read-only property
   txtRho.Text = 0.2
   lstRho.Clear
   'lstRho.Sorted = True
   With cboWeightType
      .Clear
      .AddItem Weight_Binary
      .AddItem Weight_Decreasing
      .ListIndex = 0
   End With
   chkCalcMoranI.Value = 1
   cboWeightType.Enabled = IIf((chkCalcMoranI.Value = 1), True, False)
   
   For i = 0 To chkOutFile.Count - 1
      bChecked = IIf((chkOutFile(i).Value = 1), True, False)
      txtOutFile(i).Enabled = bChecked
   Next
   
   bChecked = IIf((chkSaveLog.Value = 1), True, False)
   cmdSaveLog.Enabled = bChecked
   
   lstLog.Clear
   SetProgressBarValue 0
   
   'initialize var
   m_bRunning = False
   Set m_pBaseDEM = Nothing
   Set m_pErr = Nothing
   m_bLoadBaseDEM = IIf((chkBaseDEM.Value = 1), True, False)
   m_bSaveLog = IIf((chkSaveLog.Value = 1), True, False)
   For i = 0 To OutputSurfaceCount - 1
      m_bSaveErr(i) = IIf((chkOutFile(i).Value = 1), True, False)
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not (m_pBaseDEM Is Nothing) Then Set m_pBaseDEM = Nothing
   If Not (m_pErr Is Nothing) Then Set m_pErr = Nothing
End Sub

Private Sub optErrLyrCount_Click(Index As Integer)
   txtErrLyrCount.Enabled = optErrLyrCount(0).Value
End Sub

''''''''''''''''''''''
' error: m_pErr; DEM: m_p
Private Sub OutputErrGRID(iMonteCarlo As Integer, Optional strAutoCorrNo As String = "")
   Dim strFile As String
   Dim i As Integer, iCol As Integer, iRow As Integer, dTemp As Double
   Dim iCols As Integer, iRows As Integer, dNoData As Double
   Dim pDEM As New clsGrid
   Dim pSlope As clsGrid, pFilDep As clsGrid, pFillSlope As clsGrid, pMaxDslp As clsGrid
   Dim pAccum As clsGrid, pSCA0 As clsGrid, pSCA1 As clsGrid, pTWI As clsGrid
   Dim boolCalc As Boolean, boolFilDep As Boolean, boolMFD0 As Boolean, boolMFD_md As Boolean
   
   With m_pBaseDEM
      iCols = .nCols
      iRows = .nRows
      dNoData = .NoData_Value
   End With
   
   boolCalc = False
   For i = 1 To OutputSurfaceCount - 1
      boolCalc = boolCalc Or m_bSaveErr(i)
   Next
   If boolCalc Then
      pDEM.NewGrid iCols, iRows, m_pBaseDEM.xllcorner, m_pBaseDEM.yllcorner, m_pBaseDEM.CellSize, dNoData
      For iCol = 0 To iCols - 1
         For iRow = 0 To iRows - 1
            If m_pBaseDEM.IsValidCellValue(iCol, iRow, dTemp) Then
               pDEM.Cell(iCol, iRow) = m_pErr.Cell(iCol, iRow) + dTemp
            Else
               pDEM.Cell(iCol, iRow) = pDEM.NoData_Value
            End If
         Next
      Next
   End If
   
   boolFilDep = False
   For i = 4 To OutputSurfaceCount - 1
      boolFilDep = boolFilDep Or m_bSaveErr(i)
   Next
   If boolFilDep Then
      Set pFilDep = New clsGrid
      pFilDep.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
      boolFilDep = modMFD.FillDep_RemoveExcessWater_Planchon01(pDEM, pFilDep, 0.01)
   End If
   
   boolMFD0 = False
   If m_bSaveErr(7) Or m_bSaveErr(9) Or m_bSaveErr(10) Then boolMFD0 = True
   If boolFilDep Then
      Set pAccum = New clsGrid
      pAccum.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
      boolMFD0 = modMFD.FlowAccumulation_MFD_Quinn(pFilDep, pAccum)
      If boolMFD0 Then
         Set pSCA0 = New clsGrid
         pSCA0.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
         boolMFD0 = modMFD.SpecificCatchmentArea(pAccum, pSCA0)
      End If
      Set pAccum = Nothing
   End If
      
   boolMFD_md = False
   If m_bSaveErr(8) Or m_bSaveErr(11) Or m_bSaveErr(12) Then boolMFD_md = True
   If boolFilDep Then
      Set pAccum = New clsGrid
      pAccum.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
      boolMFD_md = modMFD.FlowAccumulation_MFD_md(pFilDep, pAccum)
      If boolMFD_md Then
         Set pSCA1 = New clsGrid
         pSCA1.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
         boolMFD_md = modMFD.SpecificCatchmentArea(pAccum, pSCA1)
      End If
      Set pAccum = Nothing
   End If
   
   For i = 0 To OutputSurfaceCount - 1
      Select Case m_strErrType
      Case ErrType_Random
         strFile = m_strErr(i) & "_R0_N" & Format(iMonteCarlo, "000") & ".asc"
      Case ErrType_RandomNormal
         strFile = m_strErr(i) & "_RN_N" & Format(iMonteCarlo, "000") & ".asc"
      Case ErrType_AutoCorr_MoranI
         strFile = m_strErr(i) & "_MI" & strAutoCorrNo & "_N" & Format(iMonteCarlo, "000") & ".asc"
      Case ErrType_AutoCorr_Rho
         strFile = m_strErr(i) & "_Rho" & strAutoCorrNo & "_N" & Format(iMonteCarlo, "000") & ".asc"
      Case ErrType_AutoCorr_Shuffle
         strFile = m_strErr(i) & "_Sh" & strAutoCorrNo & "_N" & Format(iMonteCarlo, "000") & ".asc"
      Case ErrType_TopoCorr
         strFile = m_strErr(i) & "_TC" & strAutoCorrNo & "_N" & Format(iMonteCarlo, "000") & ".asc"
      End Select
         
      Select Case i
      Case 0
         If m_bSaveErr(i) Then
            m_pErr.SaveAscGrid strFile, , 5
            lstLog.AddItem "Error: " & strFile
         End If
         If Not boolCalc Then Exit For
      Case 1
         If m_bSaveErr(i) Then
            pDEM.SaveAscGrid strFile, , 3
            lstLog.AddItem "DEM+Err: " & strFile
         End If
      Case 2
         If m_bSaveErr(i) Then
            Set pSlope = New clsGrid
            pSlope.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
            If modDTA1.Slope_ArcInfo(pDEM, pSlope) Then
               pSlope.SaveAscGrid strFile, , 4
               lstLog.AddItem "Slope: " & strFile
            Else
               lstLog.AddItem "Slope: Unsuccessful"
            End If
            Set pSlope = Nothing
         End If
      Case 3
         If m_bSaveErr(i) Then
         
         
            'lstLog.AddItem "Curvatures: " & strFile
         End If
      Case 4
         If m_bSaveErr(i) Then
            If boolFilDep Then
               pFilDep.SaveAscGrid strFile, , 3
               lstLog.AddItem "Fill Dep.(Planchon & Darbox): " & strFile
            Else
               lstLog.AddItem "Fill Dep.(Planchon & Darbox): Unsuccessful"
            End If
         End If
         If Not boolFilDep Then Exit For
      Case 5
         If m_bSaveErr(i) Then
            Set pSlope = New clsGrid
            pSlope.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
            If modDTA1.Slope_ArcInfo(pFilDep, pSlope) Then
               pSlope.SaveAscGrid strFile, , 4
               lstLog.AddItem "Slope(Dep.fill): " & strFile
            Else
               lstLog.AddItem "Slope(Dep.fill): Unsuccessful"
            End If
            Set pSlope = Nothing
         End If
      Case 6
         If m_bSaveErr(i) Then
            Set pMaxDslp = New clsGrid
            pMaxDslp.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
            If modDTA1.MaximumDownslope(pFilDep, pMaxDslp) Then
               pMaxDslp.SaveAscGrid strFile, , 4
               lstLog.AddItem "Max Downslope(Dep.fill): " & strFile
            Else
               lstLog.AddItem "Max Downslope(Dep.fill): Unsuccessful"
            End If
            Set pMaxDslp = Nothing
         End If
      Case 7
         If m_bSaveErr(i) Then
            If boolMFD0 Then
               pSCA0.SaveAscGrid strFile
               lstLog.AddItem "SCA (MFD-Quinn): " & strFile
            Else
               lstLog.AddItem "SCA (MFD-Quinn): Unsuccessful"
            End If
         End If
      Case 8
         If m_bSaveErr(i) Then
            If boolMFD_md Then
               pSCA1.SaveAscGrid strFile
               lstLog.AddItem "SCA (MFD-md): " & strFile
            Else
               lstLog.AddItem "SCA (MFD-md): Unsuccessful"
            End If
         End If
      Case 9
         If m_bSaveErr(i) Then
            If boolMFD0 Then
               Set pSlope = New clsGrid
               pSlope.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
               If modDTA1.Slope_ArcInfo(pDEM, pSlope) Then
                  Set pTWI = New clsGrid
                  pTWI.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
                  If modMFD.TWI_OriginForm(pSCA0, pSlope, pTWI) Then
                     pTWI.SaveAscGrid strFile
                     lstLog.AddItem "TWI0 (slope before fill dep.): " & strFile
                  Else
                     lstLog.AddItem "TWI0: Unsuccessful"
                  End If
                  Set pTWI = Nothing
               Else
                  lstLog.AddItem "TWI0: Slope: Unsuccessful"
               End If
               Set pSlope = Nothing
            Else
               lstLog.AddItem "TWI0: SCA (MFD-Quinn): Unsuccessful"
            End If
         End If
      Case 10
         If m_bSaveErr(i) Then
            If boolMFD0 Then
               Set pTWI = New clsGrid
               pTWI.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
               If modMFD.TWI_in_MFD_Quinn(pSCA0, pDEM, pTWI) Then
                  pTWI.SaveAscGrid strFile
                  lstLog.AddItem "TWI (Quinn, DEM before fill dep.): " & strFile
               Else
                  lstLog.AddItem "TWI (Quinn, DEM before fill dep.): Unsuccessful"
               End If
               Set pTWI = Nothing
            Else
               lstLog.AddItem "TWI (Quinn, DEM before fill dep.): SCA (MFD-Quinn): Unsuccessful"
            End If
         End If
      Case 11
         If m_bSaveErr(i) Then
            If boolMFD_md Then
               Set pSlope = New clsGrid
               pSlope.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
               If modDTA1.Slope_ArcInfo(pDEM, pSlope) Then
                  Set pTWI = New clsGrid
                  pTWI.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
                  If modMFD.TWI_OriginForm(pSCA1, pSlope, pTWI) Then
                     pTWI.SaveAscGrid strFile
                     lstLog.AddItem "TWI (MFD-md, slope before fill dep.): " & strFile
                  Else
                     lstLog.AddItem "TWI (MFD-md, slope before fill dep.): Unsuccessful"
                  End If
                  Set pTWI = Nothing
               Else
                  lstLog.AddItem "TWI (MFD-md, slope before fill dep.): Slope: Unsuccessful"
               End If
               Set pSlope = Nothing
            Else
               lstLog.AddItem "TWI (MFD-md, slope before fill dep.): SCA (MFD-md): Unsuccessful"
            End If
         End If
      Case 12
         If m_bSaveErr(i) Then
            If boolMFD_md Then
               Set pMaxDslp = New clsGrid
               pMaxDslp.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
               If modDTA1.MaximumDownslope(pFilDep, pMaxDslp) Then
                  Set pTWI = New clsGrid
                  pTWI.NewGrid iCols, iRows, pDEM.xllcorner, pDEM.yllcorner, pDEM.CellSize, dNoData
                  If modMFD.TWI_OriginForm(pSCA1, pMaxDslp, pTWI) Then
                     pTWI.SaveAscGrid strFile
                     lstLog.AddItem "TWI (MFD-md, Max Downslope): " & strFile
                  Else
                     lstLog.AddItem "TWI (MFD-md, Max Downslope): Unsuccessful"
                  End If
                  Set pTWI = Nothing
               Else
                  lstLog.AddItem "TWI (MFD-md, Max Downslope): Max Downslope: Unsuccessful"
               End If
               Set pMaxDslp = Nothing
            Else
               lstLog.AddItem "TWI (MFD-md, Max Downslope): SCA (MFD-md): Unsuccessful"
            End If
         End If
      End Select
      lstLog.ListIndex = lstLog.ListCount - 1
      DoEvents
   Next
   
   Set pDEM = Nothing
   Set pFilDep = Nothing
   Set pSCA0 = Nothing
   Set pSCA1 = Nothing
End Sub

Private Sub SetProgressBarValue(iValue As Integer)
   If iValue > 100 Or iValue < 0 Then Exit Sub
   With progbar
      .Value = iValue
      .Refresh
   End With
End Sub

