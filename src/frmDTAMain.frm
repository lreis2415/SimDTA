VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDTAMain 
   BorderStyle     =   0  'None
   Caption         =   "SimDTA"
   ClientHeight    =   315
   ClientLeft      =   2820
   ClientTop       =   1800
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   14805
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9702
            MinWidth        =   9702
            Text            =   "SimDTA"
            TextSave        =   "SimDTA"
            Object.ToolTipText     =   "Map position"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5821
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2009-1-11"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "21:55"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTabInfo 
      Height          =   8055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   14208
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Info"
      TabPicture(0)   =   "frmDTAMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Table"
      TabPicture(1)   =   "frmDTAMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "MSFlexGridInfo"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Figure"
      TabPicture(2)   =   "frmDTAMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Height          =   7635
         Left            =   -74940
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridInfo 
         Height          =   7635
         Left            =   -71700
         TabIndex        =   3
         Top             =   360
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   13467
         _Version        =   393216
      End
      Begin VB.ListBox lstInfo 
         Height          =   7080
         Left            =   60
         MultiSelect     =   1  'Simple
         TabIndex        =   2
         Top             =   360
         Width           =   13455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_Convert 
         Caption         =   "File Convert..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_Bar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnu_Bar4 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuTool 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTool_Stat 
         Caption         =   "Grid Statistics"
      End
      Begin VB.Menu mnuTool_Hist 
         Caption         =   "Histogram"
      End
      Begin VB.Menu mnuTool_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTool_ChangeValue 
         Caption         =   "Resign Grid Value Range"
      End
      Begin VB.Menu mnuTool_ErosNoData 
         Caption         =   "Erose Cells within Value Range"
      End
      Begin VB.Menu mnuTool_NearestInterpolate 
         Caption         =   "Nearest Interpolation"
      End
      Begin VB.Menu mnuTool_MakeSparse 
         Caption         =   "Make Grid Value Sparse"
      End
      Begin VB.Menu mnuTool_Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTool_AddField 
         Caption         =   "Add Point Field with GRID"
      End
   End
   Begin VB.Menu mnu_Bar1 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuArti 
      Caption         =   "&Artificial Modeling"
      Begin VB.Menu mnuArti_DEM 
         Caption         =   "Artificial Surface...(Zhou and Liu,2002;Pan et al,2004)"
      End
      Begin VB.Menu mnuArti_Err 
         Caption         =   "Error Surface..."
      End
   End
   Begin VB.Menu mnu_Bar2 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuDEM 
      Caption         =   "&DEM Preprocessing"
      Begin VB.Menu mnuDEM_Planchon01 
         Caption         =   "Remove Pit and Flat (Planchon and Darboux,2001)"
      End
   End
   Begin VB.Menu mnu_Bar3 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuDTA1 
      Caption         =   "&Local Topo. Attr."
      Begin VB.Menu mnuDTA1_Slope 
         Caption         =   "Slope/Max downslope/Local Downslope"
      End
      Begin VB.Menu mnuDTA1_Aspect 
         Caption         =   "Aspect"
      End
      Begin VB.Menu mnuDTA1_Curvature 
         Caption         =   "Curvatures"
      End
      Begin VB.Menu mnuDTA1_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTA1_SurfArea 
         Caption         =   "Surface Area, Surface-Area Ratio"
      End
      Begin VB.Menu mnuDTA1_Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTA1_Relief 
         Caption         =   "Relief"
      End
      Begin VB.Menu mnuDTA1_TRI 
         Caption         =   "Terrain Ruggedness Index (TRI)"
      End
      Begin VB.Menu mnuDTA1_ElevPercent 
         Caption         =   "Elevation Percentil Index"
      End
      Begin VB.Menu mnuDTA1_ElevReliefRatio 
         Caption         =   "Elevation-Relief Ratio"
      End
      Begin VB.Menu mnuDTA1_Cs 
         Caption         =   "Surface Curvature Index (Cs)"
      End
      Begin VB.Menu mnuDTA1_LandPos 
         Caption         =   "Landscape Position Index"
      End
      Begin VB.Menu mnuDTA1_TPI 
         Caption         =   "Topographic Position Index (TPI)"
      End
      Begin VB.Menu mnuDTA1_Openness 
         Caption         =   "Openness Angles"
      End
      Begin VB.Menu mnuDTA1_Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTA1_TOPHAT 
         Caption         =   "Hill-Hillslope-Valley Index (TOPHAT)"
      End
      Begin VB.Menu mnuDTA1_Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTA1_Drainage 
         Caption         =   "Extract Drainage Networks"
      End
      Begin VB.Menu mnuDTA1_Ridge 
         Caption         =   "Extract Ridge"
      End
   End
   Begin VB.Menu mnu_Bar6 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuDTA2 
      Caption         =   "&Regional Topo. Attr."
      Begin VB.Menu mnuDTA2_A 
         Caption         =   "Flow Accumulation/Specific Catchment Area (SCA)"
      End
      Begin VB.Menu mnuDTA2_UPNESS 
         Caption         =   "UPNESS Index"
      End
      Begin VB.Menu mnuDTA2_DownslpIndex 
         Caption         =   "Downslope Index"
      End
      Begin VB.Menu mnuDTA2_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTA2_Wet 
         Caption         =   "Topographic Wetness Index"
      End
      Begin VB.Menu mnuDTA2_SPI 
         Caption         =   "Stream Power Index"
      End
      Begin VB.Menu mnuDTA2_TCI 
         Caption         =   "Terrain Characterization Index"
      End
      Begin VB.Menu mnuDTA2_Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTA2_SlpLen 
         Caption         =   "Slope Length"
      End
      Begin VB.Menu mnuDTA2_MulFlowLen 
         Caption         =   "Flow Length based on MFD"
      End
      Begin VB.Menu mnuDTA2_Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDTA2_RPI 
         Caption         =   "Relative Position Index"
      End
      Begin VB.Menu mnuDTA2_RelatRlfIndex 
         Caption         =   "Relative Relief Index"
      End
   End
   Begin VB.Menu mnu_Bar5 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuSlope 
      Caption         =   "&Slope Descrip."
      Begin VB.Menu mnuSlope_SlpPos 
         Caption         =   "Fuzzy Slope Position"
         Begin VB.Menu mnuSlope_SlpPos_FuzzyInfer 
            Caption         =   "Fuzzy Quantification"
         End
         Begin VB.Menu mnuSlope_SlpPos_Fuzzy2Hard 
            Caption         =   "Harden Slope Position"
         End
         Begin VB.Menu mnuSlope_SlpPos_SlpSeqIndex 
            Caption         =   "Slope Sequence Index (SSI)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSlope_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSlope_Shape 
         Caption         =   "Slope Shape"
      End
   End
   Begin VB.Menu mnu_Bar7 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_Contents 
         Caption         =   "Contents..."
      End
      Begin VB.Menu mnuHelp_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmDTAMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuArti_DEM_Click()
   MsgBox "NOT implemented yet", vbInformation, APP_TITLE
End Sub

Private Sub mnuArti_Err_Click()
   frmMonteCarlo.Show vbModal 'default: vbModeless -> can go back main form without quiting form
End Sub

Private Sub mnuDEM_Planchon01_Click()
   With frmDTAFunc
      .DTAFunc FUNC_FILLDEP
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Aspect_Click()
   With frmDTAFunc
      .DTAFunc FUNC_ASPECT
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Cs_Click()
   With frmDTAFunc
      .DTAFunc FUNC_SurfaceCurvature
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Curvature_Click()
   With frmDTAFunc
      .DTAFunc FUNC_CURVATURE
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_DownslpIndex_Click()
'   MsgBox "NOT implemented yet", vbInformation, APP_TITLE
'   Exit Sub
   
   With frmDTAFunc
      .DTAFunc FUNC_DownslopeIndex
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Drainage_Click()
   With frmDTAFunc
      .DTAFunc FUNC_DRAINAGE_Peucker
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_ElevPercent_Click()
   With frmDTAFunc
      .DTAFunc FUNC_ElevPercentile
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_ElevReliefRatio_Click()
   With frmDTAFunc
      .DTAFunc FUNC_ElevReliefRatio
      .Show vbModal
   End With
End Sub

'Private Sub mnuDTA1_HypsomIntegral_Click()
'   With frmDTAFunc
'      .DTAFunc FUNC_HypsomIntegral
'      .Show vbModal
'   End With
'End Sub

Private Sub mnuDTA1_LandPos_Click()
   With frmDTAFunc
      .DTAFunc FUNC_LandPosI
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Openness_Click()
   With frmDTAFunc
      .DTAFunc FUNC_Openness
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Relief_Click()
   With frmDTAFunc
      .DTAFunc FUNC_Relief
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Ridge_Click()
   With frmDTAFunc
      .DTAFunc FUNC_RIDGE_Peucker
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_SurfArea_Click()
   With frmDTAFunc
      .DTAFunc FUNC_SurfaceArea
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_TPI_Click()
   With frmDTAFunc
      .DTAFunc FUNC_TopoPosIndex
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_Slope_Click()
   With frmDTAFunc
      .DTAFunc FUNC_SLOPE
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_TOPHAT_Click()
   With frmDTAFunc
      .DTAFunc FUNC_TOPHAT
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA1_TRI_Click()
   With frmDTAFunc
      .DTAFunc FUNC_TopoRugI
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_RelatRlfIndex_Click()
   With frmDTAFunc
      .DTAFunc FUNC_RelaRlfI
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_UPNESS_Click()
   With frmDTAFunc
      .DTAFunc FUNC_UPNESS
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_A_Click()
   With frmDTAFunc
      .DTAFunc FUNC_MFD
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_SlpLen_Click()
   frmSlopeLength.Show vbModal
End Sub

Private Sub mnuDTA2_MulFlowLen_Click()
   frmMulFlowLen.Show vbModal
End Sub

Private Sub mnuDTA2_RPI_Click()
   With frmDTAFunc
      .DTAFunc FUNC_RelaPosI
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_SPI_Click()
   With frmDTAFunc
      .DTAFunc FUNC_StreamPowerI
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_TCI_Click()
   With frmDTAFunc
      .DTAFunc FUNC_TerrainCharI
      .Show vbModal
   End With
End Sub

Private Sub mnuDTA2_Wet_Click()
   With frmDTAFunc
      .DTAFunc FUNC_TWI
      .Show vbModal
   End With
End Sub

Private Sub mnuFile_Convert_Click()
   MsgBox "NOT implemented yet", vbInformation, APP_TITLE
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Call InitializeGlobalVar
   
   Me.WindowState = vbNormal   ' vbMaximized
   SSTabInfo.Tab = 0
   lstInfo.Clear
    
   ' initial global variant
   Set g_Statusbar = StatusBar1
   g_Statusbar.Panels(1).Text = APP_TITLE & " - " & C_VersionInfo
   
   'mnuTool_Hist.Visible = C_INNER_VERSION
   mnu_Bar1.Visible = C_INNER_VERSION
   mnuArti.Visible = C_INNER_VERSION
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call ReleaseMemory
End Sub

'
Private Sub mnuHelp_About_Click()
    frmAbout.Show vbModal
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case MsgBox("Quit program?", vbOKCancel + vbDefaultButton1 + vbQuestion, APP_TITLE)
        Case vbOK
            Cancel = False
        Case vbCancel
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub mnuHelp_Contents_Click()
   MsgBox "Please refer to the document [SimDTA-manual]", vbInformation, APP_TITLE
End Sub

Private Sub mnuSlope_Shape_Click()
   'modSlopeShape.SlopeDescrib
   frmSlopeShape.Show vbModal
End Sub

Private Sub mnuSlope_SlpPos_Fuzzy2Hard_Click()
   frmHardenFuzzySlp.Show vbModal
End Sub

Private Sub mnuSlope_SlpPos_FuzzyInfer_Click()
   frmFuzzySlpInfer.Show vbModal
End Sub

Private Sub mnuSlope_SlpPos_SlpSeqIndex_Click()
   MsgBox "NOT finished yet, sorry", vbInformation, APP_TITLE
End Sub

Private Sub mnuTool_AddField_Click()
   frmGRIDValue2PtAttr.Show vbModal
End Sub

Private Sub mnuTool_ChangeValue_Click()
   frmChangeGRIDRangeValue.Show vbModal
End Sub

Private Sub mnuTool_ErosNoData_Click()
   frmEroseByNoData.Show vbModal
End Sub

Private Sub mnuTool_Hist_Click()
   MsgBox "NOT implemented yet", vbInformation, APP_TITLE
End Sub

Private Sub mnuTool_MakeSparse_Click()
   frmMakeSparse.Show vbModal
End Sub

Private Sub mnuTool_NearestInterpolate_Click()
   frmNearestInterpolate.Show vbModal
End Sub

Private Sub mnuTool_Stat_Click()
   frmGRIDStatis.Show vbModal
End Sub
