VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmGRIDStatis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GRID Statistics"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   60
      TabIndex        =   83
      Top             =   6960
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   960
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   7380
      Width           =   2055
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Do Statistics"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   7380
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTabStatis 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Single GRID"
      TabPicture(0)   =   "frmGRIDStatis.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frameSrc(0)"
      Tab(0).Control(1)=   "framePara(0)"
      Tab(0).Control(2)=   "frameOutput(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "GRIDs in Folder"
      TabPicture(1)   =   "frmGRIDStatis.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frameSrc(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameOutput(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "framePara(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   2460
         Width           =   11055
         Begin VB.ComboBox cboTransform 
            Height          =   300
            Index           =   1
            Left            =   7620
            TabIndex        =   80
            Top             =   360
            Width           =   3255
         End
         Begin VB.CheckBox chkComputeQuartile 
            Caption         =   "Median, Quartile"
            Height          =   495
            Left            =   2400
            TabIndex        =   77
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Transform before statistics:"
            Height          =   255
            Index           =   2
            Left            =   4620
            TabIndex        =   79
            Top             =   390
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "Statistics includes:"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   78
            Top             =   390
            Width           =   2295
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output"
         Height          =   3495
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   3360
         Width           =   11055
         Begin VB.TextBox txtSave3rdQuartile 
            Height          =   375
            Left            =   1440
            TabIndex        =   76
            Top             =   1920
            Width           =   9495
         End
         Begin VB.TextBox txtSaveMedian 
            Height          =   375
            Left            =   1440
            TabIndex        =   75
            Top             =   2280
            Width           =   9495
         End
         Begin VB.TextBox txtSave1stQuartile 
            Height          =   375
            Left            =   1440
            TabIndex        =   74
            Top             =   2640
            Width           =   9495
         End
         Begin VB.TextBox txtSaveMin 
            Height          =   375
            Left            =   1440
            TabIndex        =   73
            Top             =   3000
            Width           =   9495
         End
         Begin VB.CommandButton cmdSaveMedian 
            Caption         =   "Median"
            Height          =   375
            Left            =   120
            TabIndex        =   72
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdSave1stQuartile 
            Caption         =   "1st Quartile"
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CommandButton cmdSave3rdQuartile 
            Caption         =   "3rd Quartile"
            Height          =   375
            Left            =   120
            TabIndex        =   70
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveMin 
            Caption         =   "Min"
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveMax 
            Caption         =   "Max"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveSTDEV 
            Caption         =   "STDEV"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveSD 
            Caption         =   "SD"
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   660
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveMean 
            Caption         =   "Mean"
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox txtSaveSD 
            Height          =   375
            Left            =   1440
            TabIndex        =   58
            Top             =   660
            Width           =   9495
         End
         Begin VB.TextBox txtSaveMax 
            Height          =   375
            Left            =   1440
            TabIndex        =   57
            Top             =   1560
            Width           =   9495
         End
         Begin VB.TextBox txtSaveSTDEV 
            Height          =   375
            Left            =   1440
            TabIndex        =   56
            Top             =   1020
            Width           =   9495
         End
         Begin VB.TextBox txtSaveMean 
            Height          =   375
            Left            =   1440
            TabIndex        =   55
            Top             =   300
            Width           =   9495
         End
      End
      Begin VB.Frame frameSrc 
         Caption         =   "Source"
         Height          =   2055
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRIDCount 
            Caption         =   "Count of Src GRIDs"
            Height          =   375
            Left            =   8220
            TabIndex        =   63
            Top             =   780
            Width           =   1815
         End
         Begin VB.TextBox txtFileNameLike 
            Height          =   390
            Left            =   2040
            TabIndex        =   62
            Text            =   "*"
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkIncludeSubFolder 
            Caption         =   "Include GRIDs under all sub-folders"
            Height          =   255
            Left            =   3960
            TabIndex        =   61
            Top             =   840
            Width           =   4095
         End
         Begin VB.TextBox txtSrcGRIDCount 
            Height          =   375
            Left            =   10080
            TabIndex        =   59
            Text            =   "0"
            Top             =   780
            Width           =   855
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Enabled         =   0   'False
            Height          =   675
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   1260
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   1
               Left            =   600
               TabIndex        =   47
               Top             =   240
               Width           =   795
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   1
               Left            =   1920
               TabIndex        =   46
               Top             =   240
               Width           =   795
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   1
               Left            =   3600
               TabIndex        =   45
               Text            =   "0"
               Top             =   240
               Width           =   1395
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   1
               Left            =   5880
               TabIndex        =   44
               Text            =   "0"
               Top             =   240
               Width           =   1395
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   1
               Left            =   8100
               TabIndex        =   43
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   1
               Left            =   10020
               TabIndex        =   42
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   5
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   4
               Left            =   1440
               TabIndex        =   52
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   3
               Left            =   2760
               TabIndex        =   51
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   2
               Left            =   5040
               TabIndex        =   50
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   1
               Left            =   7320
               TabIndex        =   49
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   0
               Left            =   8880
               TabIndex        =   48
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   40
            Top             =   240
            Width           =   9495
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "Src Folder"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "GRID Name like        (""*"" means all file)"
            Height          =   555
            Index           =   6
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output"
         Height          =   2295
         Index           =   0
         Left            =   -74880
         TabIndex        =   5
         Top             =   4140
         Width           =   11055
         Begin VB.CheckBox chkKurt 
            Caption         =   "峰度 Kurt = "
            Height          =   255
            Left            =   5400
            TabIndex        =   31
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkSkew 
            Caption         =   "偏度 Skew ="
            Height          =   255
            Left            =   5400
            TabIndex        =   30
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkRMSE 
            Caption         =   "RMSE ="
            Height          =   255
            Left            =   5400
            TabIndex        =   28
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkSD 
            Caption         =   "均方差 SD ="
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkMean 
            Caption         =   "Mean ="
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkSTDEV 
            Caption         =   "样本标准偏差 STDEV ="
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Cell count for statistics"
            Height          =   255
            Index           =   4
            Left            =   540
            TabIndex        =   85
            Top             =   360
            Width           =   1875
         End
         Begin VB.Label lblCount 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   84
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblSTDEV 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   35
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblKurt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7080
            TabIndex        =   37
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblSkew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7080
            TabIndex        =   36
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblRMSE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7080
            TabIndex        =   34
            Top             =   720
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblSD 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   33
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblMean 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   32
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "for RMSE, Src GRIS as Error GRID, assumption: A zero mean error, i.e. no systematic bias"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   25
            Top             =   1800
            Visible         =   0   'False
            Width           =   8295
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1695
         Index           =   0
         Left            =   -74880
         TabIndex        =   4
         Top             =   2280
         Width           =   11055
         Begin VB.ComboBox cboTransform 
            Height          =   300
            Index           =   0
            Left            =   3120
            TabIndex        =   81
            Top             =   1020
            Width           =   3495
         End
         Begin VB.CheckBox chkIncludeNearNoData 
            Caption         =   "NoData-neighboring Cells"
            Height          =   255
            Left            =   4200
            TabIndex        =   23
            Top             =   480
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkIncludeEdge 
            Caption         =   "Edge Cells"
            Height          =   255
            Left            =   2520
            TabIndex        =   22
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkIncludeNoDataCell 
            Caption         =   "Cells with NoData"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Transform before statistics:"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   82
            Top             =   1020
            Width           =   2835
         End
         Begin VB.Label Label2 
            Caption         =   "Statistics includes:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame frameSrc 
         Caption         =   "Source"
         Height          =   1815
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "Src GRID"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Enabled         =   0   'False
            Height          =   675
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   0
               Left            =   10020
               TabIndex        =   12
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   0
               Left            =   8100
               TabIndex        =   11
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   0
               Left            =   5880
               TabIndex        =   10
               Text            =   "0"
               Top             =   240
               Width           =   1395
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   0
               Left            =   3600
               TabIndex        =   9
               Text            =   "0"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   0
               Left            =   1920
               TabIndex        =   8
               Top             =   240
               Width           =   795
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   0
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
               Left            =   7320
               TabIndex        =   17
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   86
               Left            =   5040
               TabIndex        =   16
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   87
               Left            =   2760
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   88
               Left            =   1440
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
      End
   End
End
Attribute VB_Name = "frmGRIDStatis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Const TRANS_NO = "No"
Private Const TRANS_LN = "Natural Logarithm: ln()"
Private Const TRANS_LOG10 = "Common Logarithm: lg()"

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_bRunning As Boolean
Dim m_strBasePath As String
Dim m_strFilePre As String
Dim m_pBaseGRID As clsGrid


Private Function GetFileCount(strRootFolder As String, strFolderLike As String, strFileLike As String, Optional colFilesFound As Collection = Nothing, Optional boolIncludeSubFolder As Boolean = False) As Long
    
   '*********************************************************
   '* Author: Dragon <sebastian.strand@pp.inet.fi>          *
   '*         http://personal.inet.fi/cool/dragon/vb/       *
   '*                                                       *
   '* Last updated: August 14, 1998                         *
   '*                                                       *
   '* This recursive routine searches for a specified       *
   '* file/files starting from a specified rootfolder       *
   '* You can specify folder and file info with pattern     *
   '* matching (*, ?, # and so on). For more info on        *
   '* pattern matching please refer to the VB documentation *
   '* for the 'Like' function                               *
   '*                                                       *
   '* This function has the following arguments:            *
   '*                                                       *
   '*   strRootFolder  =  the folder from which the search  *
   '*                     starts. The search will only find *
   '*                     files in this directory or it's   *
   '*                     subdirectories                    *
   '*                                                       *
   '*   strFolderLike = folder information for the files        *
   '*               searched. Specify * to allow files in   *
   '*               any folder. Pattern matching allowed.   *
   '*                                                       *
   '*   strFileLike = the filename to search for. Pattern       *
   '*             matching allowed.                         *
   '*                                                       *
   '*   colFilesFound = the files found will be placed in   *
   '*                   this collection                     *
   '*                                                       *
   '* Example usage:                                        *
   '* Dim colFiles as New Collection 'Note 'New' keyword!!  *
   '* Call FindFiles("C:\Windows\System","*","doc[123].txt")*
   '*                                                       *
   '* Then colFiles will be filled with all the text files  *
   '* named doc1.txt or doc2.txt or doc3.txt in the Windows\*
   '* System dir and all it's subdirs.                      *
   '*                                                       *
   '*********************************************************

    Dim lngSearchHandle As Long
    Dim udtFindData As WIN32_FIND_DATA
    Dim strTemp As String, lngRet As Long
    Dim lFileCount As Long
        
    'Check that folder name ends with "\"
    If Right$(strRootFolder, 1) <> "\" Then strRootFolder = strRootFolder & "\"
    
    'Find first file/folder in current folder
    lngSearchHandle = FindFirstFile(strRootFolder & "*", udtFindData)
    
    'Check that we received a valid handle
    If lngSearchHandle = INVALID_HANDLE_VALUE Then Exit Function
    
    lngRet = 1
    lFileCount = 0
    Do While lngRet <> 0
        
        'Trim nulls from filename
        strTemp = TrimNulls(udtFindData.cFileName)
        
        If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            'It's a dir - make sure it isn't . or .. dirs
            If strTemp <> "." And strTemp <> ".." Then
                'It's a normal dir: let's dive straight
                'into it...
               If boolIncludeSubFolder Then
                  If (colFilesFound Is Nothing) Then
                     lFileCount = lFileCount + GetFileCount(strRootFolder & strTemp, strFolderLike, strFileLike, , boolIncludeSubFolder)
                  Else
                     lFileCount = lFileCount + GetFileCount(strRootFolder & strTemp, strFolderLike, strFileLike, colFilesFound, boolIncludeSubFolder)
                  End If
               End If
            End If
        Else
            'It's a file. First check if the current folder matches
            'the folder path in strFolderLike
            If (strRootFolder Like strFolderLike) Then
                'Folder matches, what about file?
                If (strTemp Like strFileLike) Then
                    'Found one!
                    If Not (colFilesFound Is Nothing) Then colFilesFound.Add strRootFolder & strTemp
                    lFileCount = lFileCount + 1
                End If
            End If
        End If
        
        'Get next file/folder
        lngRet = FindNextFile(lngSearchHandle, udtFindData)
        
    Loop
    
    'Close find handle
    Call FindClose(lngSearchHandle)
    GetFileCount = lFileCount
End Function


Private Sub chkComputeQuartile_Click()
   Dim boolChk As Boolean
   boolChk = IIf(chkComputeQuartile.Value = vbChecked, True, False)
   cmdSaveMedian.Enabled = boolChk
   txtSaveMedian.Enabled = boolChk
   txtSaveMedian.Text = ""
   cmdSave1stQuartile.Enabled = boolChk
   txtSave1stQuartile.Text = ""
   txtSave1stQuartile.Enabled = boolChk
   cmdSave3rdQuartile.Enabled = boolChk
   txtSave3rdQuartile.Text = ""
   txtSave3rdQuartile.Enabled = boolChk
End Sub

'Private Sub chkMean_Click(Index As Integer)
'   If Index = 1 Then
'      If chkMean(Index).Value = 1 Then
'         txtSaveMean.Enabled = True
'      Else
'         txtSaveMean.Enabled = False
'      End If
'   End If
'End Sub
'
'Private Sub chkRMSE_Click(Index As Integer)
'   If Index = 1 Then
'      If chkRMSE(Index).Value = 1 Then
'         txtSaveRMSE.Enabled = True
'      Else
'         txtSaveRMSE.Enabled = False
'      End If
'   End If
'End Sub
'
'Private Sub chkSD_Click(Index As Integer)
'   If Index = 1 Then
'      If chkSD(Index).Value = 1 Then
'         txtSaveSD.Enabled = True
'      Else
'         txtSaveSD.Enabled = False
'      End If
'   End If
'End Sub
'
'Private Sub chkSTDEV_Click(Index As Integer)
'   If Index = 1 Then
'      If chkSTDEV(Index).Value = 1 Then
'         txtSaveSTDEV.Enabled = True
'      Else
'         txtSaveSTDEV.Enabled = False
'      End If
'   End If
'End Sub

Private Sub cmdQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub

Private Sub cmdSave1stQuartile_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSave1stQuartile.Text = strFile
End Sub

Private Sub cmdSave3rdQuartile_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSave3rdQuartile.Text = strFile
End Sub

Private Sub cmdSaveMax_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveMax.Text = strFile
End Sub

Private Sub cmdSaveMean_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveMean.Text = strFile
End Sub

Private Sub cmdSaveMedian_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveMedian.Text = strFile
End Sub

Private Sub cmdSaveMin_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveMin.Text = strFile
End Sub

Private Sub cmdSaveSD_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveSD.Text = strFile
End Sub

Private Sub cmdSaveSTDEV_Click()
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveSTDEV.Text = strFile
End Sub

Private Sub cmdSrcGRID_Click(Index As Integer)
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
   
   txtSrcGRID(Index).Text = IIf(Index = 0, strBaseDEM, m_strBasePath)
   
   ' load BaseDEM, read parameters in file head
   If Not (m_pBaseGRID Is Nothing) Then Set m_pBaseGRID = Nothing
   Set m_pBaseGRID = New clsGrid
   With m_pBaseGRID
      .LoadAscGrid strBaseDEM
      txtCols(Index).Text = .nCols
      txtRows(Index).Text = .nRows
      txtXll(Index).Text = .xllcorner
      txtYll(Index).Text = .yllcorner
      txtCellSize(Index).Text = .CellSize
      txtNoData(Index).Text = .NoData_Value
   End With
   If Index = 1 Then
      txtSaveMean.Text = m_strBasePath & "Mean.asc"
      txtSaveSD.Text = m_strBasePath & "SD.asc"
'      txtSaveRMSE.Text = m_strBasePath & "RMSE.asc"
'      txtSaveSTDEV.Text = m_strBasePath & "STDEV.asc"
   Else
      Me.Refresh
      Call cmdStatistics_Click
   End If
End Sub

Private Sub cmdSrcGRIDCount_Click()
   Dim boolIncludeSubFolder As Boolean
   boolIncludeSubFolder = IIf(chkIncludeSubFolder.Value = 1, True, False)
   txtSrcGRIDCount.Text = str(GetFileCount(txtSrcGRID(1).Text, "*", txtFileNameLike.Text, , boolIncludeSubFolder))
End Sub

Private Sub cmdStatistics_Click()
   Dim iStatisType As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dMean As Double, dSD As Double, dRMSE As Double, dSTDEV As Double, dSkew As Double, dKurt As Double
   Dim dSum As Double, dSum2 As Double
   Dim iCol As Integer, iRow As Integer, n As Double
   Dim boolIncludeNoDataCell As Boolean, boolIncludeEdge As Boolean, boolIncludeNearNoData As Boolean
   Dim boolMean As Boolean, boolSD As Boolean, boolRMSE As Boolean, boolSTDEV As Boolean, boolSkew As Boolean, boolKurt As Boolean
   Dim boolMax As Boolean, boolMin As Boolean, boolMedian As Boolean, bool1stQuartile As Boolean, bool3rdQuartile As Boolean
   Dim iColFrom As Integer, iColTo As Integer, iRowFrom As Integer, iRowTo As Integer
   'Variants for Folder statistics
   Dim strRootFolder As String, strFolderLike As String, strFileLike As String, boolIncludeSubFolder As Boolean
   Dim strMeanFile As String, strSDFile As String, strRMSEFile As String, strSTDEVFile As String
   Dim strMaxFile As String, strMinFile As String, strMedianFile As String, str1stQuartileFile As String, str3rdQuartileFile As String
   Dim pGrid As clsGrid, pGridMean As clsGrid, pGridSD As clsGrid, pGridRMSE As clsGrid, pGridSTDEV As clsGrid
   Dim pGridMax As clsGrid, pGridMin As clsGrid, pGridMedian As clsGrid, pGrid1stQuartile As clsGrid, pGrid3rdQuartile As clsGrid
   Dim colFindFiles As Collection, lFileNo As Long, lFileCount As Long
   Dim vMax As Variant, vMin As Variant, vSum As Variant, vSum2 As Variant, vCount As Variant, dValue As Double
   Dim str As String
   
On Error GoTo ErrH
   If m_bRunning Then Exit Sub
   m_bRunning = True
   str = "Done!   " & Time()
   Me.MousePointer = 11
   
   SetProgressBarValue 0
   iStatisType = SSTabStatis.Tab
   If txtSrcGRID(iStatisType).Text = "" Then
      Err.Raise Number:=vbObjectError + 513, Description:="Assign the source GRID firstly"
   End If
   If iStatisType = 0 Then
      boolMean = chkMean.Value
      boolSD = chkSD.Value
      boolRMSE = chkRMSE.Value
      boolSTDEV = chkSTDEV.Value
   Else
      boolMean = IIf(Trim(txtSaveMean.Text) = "", False, True)
      boolSD = IIf(Trim(txtSaveSD.Text) = "", False, True)
      boolRMSE = False
      boolSTDEV = IIf(Trim(txtSaveSTDEV.Text) = "", False, True)
      boolMax = IIf(Trim(txtSaveMax.Text) = "", False, True)
      boolMin = IIf(Trim(txtSaveMin.Text) = "", False, True)
      boolMedian = IIf(Trim(txtSaveMedian.Text) = "", False, True)
      bool1stQuartile = IIf(Trim(txtSave1stQuartile.Text) = "", False, True)
      bool3rdQuartile = IIf(Trim(txtSave3rdQuartile.Text) = "", False, True)
   End If
   
   Select Case iStatisType
   Case 0   ' single grid statistics
      With m_pBaseGRID
         iCols = .nCols: iRows = .nRows
         dCellSize = .CellSize
         dNoData = .NoData_Value
      End With
      
      boolIncludeNoDataCell = chkIncludeNoDataCell.Value
      boolIncludeEdge = chkIncludeEdge.Value
      boolIncludeNearNoData = chkIncludeNearNoData.Value
      boolSkew = chkSkew.Value:  boolKurt = chkKurt.Value
      
      If boolIncludeEdge Then
         iColFrom = 0: iColTo = iCols - 1: iRowFrom = 0: iRowTo = iRows - 1
      Else
         iColFrom = 1: iColTo = iCols - 2: iRowFrom = 1: iRowTo = iRows - 2
      End If
      
      dSum = 0#: dSum2 = 0#: n = 0
      For iCol = iColFrom To iColTo
         For iRow = iRowFrom To iRowTo
            If m_pBaseGRID.Cell(iCol, iRow) = dNoData Then
               If boolIncludeNoDataCell Then n = n + 1
            Else
               If boolIncludeNearNoData Or Not m_pBaseGRID.NearNoDataCell(iCol, iRow) Then
                  n = n + 1
                  dValue = m_pBaseGRID.Cell(iCol, iRow)
                  dSum = dSum + dValue
                  dSum2 = dSum2 + dValue ^ 2
               End If
            End If
         Next
         SetProgressBarValue Int(iCol * 70# / iColTo)
      Next
      dMean = dSum / n
      If boolMean Then
         lblMean.Caption = dMean
      Else
         lblMean.Caption = ""
      End If
      
      dSD = Sqr(dSum2 / n - dMean ^ 2)
      If boolSD Then
         lblSD.Caption = dSD
      Else
         lblSD.Caption = ""
      End If
            
      dRMSE = Sqr(dSum2 / n)
      If boolRMSE Then
         lblRMSE.Caption = dRMSE
      Else
         lblRMSE.Caption = ""
      End If
      dSTDEV = dSum2 / (n - 1) - (dSum ^ 2 / n) / (n - 1)
      If dSTDEV < 0# Then
         dSTDEV = 0#
      Else
         dSTDEV = Sqr(dSTDEV)
      End If
      If boolSTDEV Then
         lblSTDEV.Caption = dSTDEV
      Else
         lblSTDEV.Caption = ""
      End If
      
      If boolSkew Then
         dSum2 = 0#
         For iCol = iColFrom To iColTo
            For iRow = iRowFrom To iRowTo
               If m_pBaseGRID.Cell(iCol, iRow) <> dNoData Then
                  If boolIncludeNearNoData Or Not m_pBaseGRID.NearNoDataCell(iCol, iRow) Then
                     dSum2 = dSum2 + ((m_pBaseGRID.Cell(iCol, iRow) - dMean) / dSD) ^ 3
                  End If
               End If
            Next
         Next
         dSkew = dSum2 / n
         lblSkew.Caption = dSkew
      Else
         lblSkew.Caption = ""
      End If
      SetProgressBarValue 85
      
      If boolKurt Then
         dSum2 = 0#
         For iCol = iColFrom To iColTo
            For iRow = iRowFrom To iRowTo
               If m_pBaseGRID.Cell(iCol, iRow) <> dNoData Then
                  If boolIncludeNearNoData Or Not m_pBaseGRID.NearNoDataCell(iCol, iRow) Then
                     dSum2 = dSum2 + ((m_pBaseGRID.Cell(iCol, iRow) - dMean) / dSD) ^ 4
                  End If
               End If
            Next
         Next
         dKurt = dSum2 / n
         lblKurt.Caption = dKurt
      Else
         lblKurt.Caption = ""
      End If
      
      lblCount.Caption = n
      SetProgressBarValue 100
      
   Case 1   ' GRIDs in a file folder
      iCols = CInt(txtCols(1).Text): iRows = CInt(txtRows(1).Text):  dNoData = CDbl(txtNoData(1).Text)
      boolIncludeSubFolder = IIf(chkIncludeSubFolder.Value = 1, True, False)
      strRootFolder = txtSrcGRID(1).Text
      strFolderLike = "*"
      strFileLike = txtFileNameLike.Text
      Set colFindFiles = New Collection
      GetFileCount strRootFolder, strFolderLike, strFileLike, colFindFiles, boolIncludeSubFolder
      lFileCount = colFindFiles.Count
      If lFileCount = 0 Then
         Err.Raise vbObjectError + 513, "frmGRIDStatis.cmdStatistics_Click", "No files like " & strFileLike & " in folder " & strRootFolder
      End If
      If boolMean Or boolSD Then
         strMeanFile = txtSaveMean.Text
         Set pGridMean = New clsGrid
         pGridMean.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
      End If
      If boolSD Then
         strSDFile = txtSaveSD.Text
         Set pGridSD = New clsGrid
         pGridSD.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
      End If
'      If boolRMSE Then
'         strRMSEFile = txtSaveRMSE.Text
'         Set pGridRMSE = New clsGrid
'         pGridRMSE.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
'      End If
      If boolSTDEV Then
         strSTDEVFile = txtSaveSTDEV.Text
         Set pGridSTDEV = New clsGrid
         pGridSTDEV.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
      End If
      If boolMin Then
         strMinFile = txtSaveMin.Text
         Set pGridMin = New clsGrid
         pGridMin.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
      End If
      If boolMax Then
         strMaxFile = txtSaveMax.Text
         Set pGridMax = New clsGrid
         pGridMax.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
      End If
      
      Set pGrid = New clsGrid
      ReDim vSum(0 To iCols - 1, 0 To iRows - 1)
      ReDim vSum2(0 To iCols - 1, 0 To iRows - 1)
      ReDim vCount(0 To iCols - 1, 0 To iRows - 1)
      ReDim vMax(0 To iCols - 1, 0 To iRows - 1)
      ReDim vMin(0 To iCols - 1, 0 To iRows - 1)
      For iCol = 0 To iCols - 1
         For iRow = 0 To iRows - 1
            vSum(iCol, iRow) = CDbl(0):  vSum2(iCol, iRow) = CDbl(0): vCount(iCol, iRow) = CLng(0)
            vMax(iCol, iRow) = MIN_SINGLE:   vMin(iCol, iRow) = MAX_SINGLE
         Next
      Next
      
      iColFrom = 0: iColTo = iCols - 1: iRowFrom = 0: iRowTo = iRows - 1
      For lFileNo = 1 To lFileCount
         pGrid.LoadAscGrid colFindFiles.Item(lFileNo)
         For iCol = iColFrom To iColTo
            For iRow = iRowFrom To iRowTo
               dValue = pGrid.Cell(iCol, iRow)
               If dValue <> pGrid.NoData_Value Then
                  vCount(iCol, iRow) = vCount(iCol, iRow) + 1
                  vSum(iCol, iRow) = vSum(iCol, iRow) + dValue
                  vSum2(iCol, iRow) = vSum2(iCol, iRow) + dValue ^ 2
                  If dValue > vMax(iCol, iRow) Then vMax(iCol, iRow) = dValue
                  If dValue < vMin(iCol, iRow) Then vMin(iCol, iRow) = dValue
               End If
            Next
         Next
         DoEvents
         SetProgressBarValue Int(lFileNo * 30# / lFileCount)
      Next
      
      'mean
      If boolMean Or boolSD Then
         For iCol = iColFrom To iColTo
            For iRow = iRowFrom To iRowTo
               If vCount(iCol, iRow) = 0 Then
                  pGridMean.Cell(iCol, iRow) = pGridMean.NoData_Value
               Else
                  pGridMean.Cell(iCol, iRow) = vSum(iCol, iRow) / vCount(iCol, iRow)
               End If
            Next
         Next
         If boolMean Then pGridMean.SaveAscGrid strMeanFile, , 5
      End If
      SetProgressBarValue 35
      DoEvents
'      If boolRMSE Then
'         For iCol = iColFrom To iColTo
'            For iRow = iRowFrom To iRowTo
'               If vCount(iCol, iRow) = 0 Then
'                  pGridRMSE.Cell(iCol, iRow) = pGridRMSE.NoData_Value
'               Else
'                  pGridRMSE.Cell(iCol, iRow) = Sqr(vSum2(iCol, iRow) / vCount(iCol, iRow))
'               End If
'            Next
'         Next
'         pGridRMSE.SaveAscGrid strRMSEFile, , 5
'         DoEvents
'      End If
      
      If boolSTDEV Then
         For iCol = iColFrom To iColTo
            For iRow = iRowFrom To iRowTo
               If vCount(iCol, iRow) <= 1 Then
                  pGridSTDEV.Cell(iCol, iRow) = pGridSTDEV.NoData_Value
               Else
                  dValue = vSum2(iCol, iRow) / (vCount(iCol, iRow) - 1) - (vSum(iCol, iRow) ^ 2 / vCount(iCol, iRow)) / (vCount(iCol, iRow) - 1)
                  If dValue < 0# Then dValue = 0#
                  pGridSTDEV.Cell(iCol, iRow) = Sqr(dValue)
               End If
            Next
         Next
         pGridSTDEV.SaveAscGrid strSTDEVFile, , 5
         DoEvents
      End If
      SetProgressBarValue 40
      
      If boolSD Then
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               If vCount(iCol, iRow) = 0 Then
                  pGridSD.Cell(iCol, iRow) = pGridSD.NoData_Value
               Else
                  pGridSD.Cell(iCol, iRow) = Sqr(vSum2(iCol, iRow) / vCount(iCol, iRow) - pGridMean.Cell(iCol, iRow) ^ 2)
               End If
            Next
         Next
         pGridSD.SaveAscGrid strSDFile, , 5
      End If
      SetProgressBarValue 45
            
      If boolMax Then
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               If vCount(iCol, iRow) = 0 Then
                  pGridMax.Cell(iCol, iRow) = pGridMax.NoData_Value
               Else
                  pGridMax.Cell(iCol, iRow) = vMax(iCol, iRow)
               End If
            Next
         Next
         pGridMax.SaveAscGrid strMaxFile
      End If
      SetProgressBarValue 48
            
      If boolMin Then
         For iCol = 0 To iCols - 1
            For iRow = 0 To iRows - 1
               If vCount(iCol, iRow) = 0 Then
                  pGridMin.Cell(iCol, iRow) = pGridMin.NoData_Value
               Else
                  pGridMin.Cell(iCol, iRow) = vMin(iCol, iRow)
               End If
            Next
         Next
         pGridMin.SaveAscGrid strMinFile
      End If
      SetProgressBarValue 50
      
      If boolMean Then str = str & vbCrLf & "Mean: " & strMeanFile
      If boolSD Then str = str & vbCrLf & "SD: " & strSDFile
      If boolSTDEV Then str = str & vbCrLf & "StDev: " & strSTDEVFile
      If boolRMSE Then str = str & vbCrLf & "RMSE: " & strRMSEFile
      If boolMax Then str = str & vbCrLf & "Max: " & strMaxFile
      If boolMin Then str = str & vbCrLf & "Min: " & strMinFile
      
      Set pGrid = Nothing
      Set pGridMean = Nothing:   Set pGridSD = Nothing:   Set pGridRMSE = Nothing:   Set pGridSTDEV = Nothing
      Set pGridMax = Nothing:   Set pGridMin = Nothing
      vSum = Empty: vSum2 = Empty: vMin = Empty:  vMax = Empty
      DoEvents
      
      If boolMedian Or bool1stQuartile Or bool3rdQuartile Then
         strMedianFile = txtSaveMedian.Text
         Set pGridMedian = New clsGrid
         pGridMedian.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
         
         str1stQuartileFile = txtSave1stQuartile.Text
         Set pGrid1stQuartile = New clsGrid
         pGrid1stQuartile.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
         
         str3rdQuartileFile = txtSave3rdQuartile.Text
         Set pGrid3rdQuartile = New clsGrid
         pGrid3rdQuartile.NewGrid iCols, iRows, CDbl(txtXll(1).Text), CDbl(txtYll(1).Text), CDbl(txtCellSize(1).Text), dNoData
                  
         Dim arrGRID() As clsGrid
         Dim arrValue As Variant
         Dim iValueCount As Integer, iValuePos As Integer
         
         ReDim arrGRID(1 To lFileCount)
         For lFileNo = 1 To lFileCount
            Set arrGRID(lFileNo) = New clsGrid
            arrGRID(lFileNo).LoadAscGrid colFindFiles.Item(lFileNo)
            SetProgressBarValue 50 + Int(lFileNo * 20# / lFileCount)
            DoEvents
         Next
         
         ReDim arrValue(1 To lFileCount)
         
         For iCol = iColFrom To iColTo
            For iRow = iRowFrom To iRowTo
               If vCount(iCol, iRow) = 0 Then
                  pGridMedian.Cell(iCol, iRow) = pGridMedian.NoData_Value
                  pGrid1stQuartile.Cell(iCol, iRow) = pGrid1stQuartile.NoData_Value
                  pGrid3rdQuartile.Cell(iCol, iRow) = pGrid3rdQuartile.NoData_Value
               Else
                  'ReDim arrValue(1 To vCount(iCol, iRow))
                  iValueCount = 0
                  For lFileNo = 1 To lFileCount
                     If arrGRID(lFileNo).Cell(iCol, iRow) <> arrGRID(lFileNo).NoData_Value Then
                        iValueCount = iValueCount + 1
                        arrValue(iValueCount) = arrGRID(lFileNo).Cell(iCol, iRow)
                     End If
                  Next
                  'sort arrValue(1 to iValueCount)
                  If QuickSort_NonRecursive(arrValue, 1, iValueCount) Then
                      iValuePos = Int(iValueCount / 2)
                     If iValuePos = iValueCount / 2 Then
                        pGridMedian.Cell(iCol, iRow) = (arrValue(iValuePos) + arrValue(iValuePos + 1)) / 2
                     Else
                        pGridMedian.Cell(iCol, iRow) = arrValue(iValuePos + 1)
                     End If
                     iValuePos = Int(iValueCount / 4)
                     If iValueCount / 4 < Round(iValueCount / 4, 0) Then
                        pGrid1stQuartile.Cell(iCol, iRow) = arrValue(iValuePos + 1)
                     Else
                        pGrid1stQuartile.Cell(iCol, iRow) = (arrValue(iValuePos) + arrValue(iValuePos + 1)) / 2
                     End If
                     iValuePos = Int(iValueCount * 0.75)
                     If iValueCount * 0.75 < Round(iValueCount * 0.75, 0) Then
                        pGrid3rdQuartile.Cell(iCol, iRow) = arrValue(iValuePos + 1)
                     Else
                        pGrid3rdQuartile.Cell(iCol, iRow) = (arrValue(iValuePos) + arrValue(iValuePos + 1)) / 2
                     End If
                  Else
                     Debug.Print ""
                  End If
                  DoEvents
               End If
            Next
            SetProgressBarValue 70 + Int(iCol * 30# / iColTo)
         Next
         
         For lFileNo = 1 To lFileCount
            Set arrGRID(lFileNo) = Nothing
         Next
         
         If boolMedian Then
            pGridMedian.SaveAscGrid strMedianFile
            str = str & vbCrLf & "Median: " & strMedianFile
         End If
         If bool1stQuartile Then
            pGrid1stQuartile.SaveAscGrid str1stQuartileFile
            str = str & vbCrLf & "1st Quartile: " & str1stQuartileFile
         End If
         If bool3rdQuartile Then
            pGrid3rdQuartile.SaveAscGrid str3rdQuartileFile
            str = str & vbCrLf & "3rd Quartile: " & str3rdQuartileFile
         End If
      End If
      SetProgressBarValue 100
   End Select
   str = str & vbCrLf & "End. " & Time()
   MsgBox str, vbInformation, APP_TITLE
   SetProgressBarValue 0
ErrH:
   Me.MousePointer = 0
   vSum = Empty: vSum2 = Empty: vCount = Empty: vMin = Empty:  vMax = Empty
   arrValue = Empty
   Set colFindFiles = Nothing
   Set pGrid = Nothing
   Set pGridMean = Nothing:   Set pGridSD = Nothing:   Set pGridRMSE = Nothing:   Set pGridSTDEV = Nothing
   Set pGridMax = Nothing:   Set pGridMin = Nothing:   Set pGridMedian = Nothing
   Set pGrid1stQuartile = Nothing:    Set pGrid3rdQuartile = Nothing
   m_bRunning = False
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub Form_Load()
   Dim i As Integer
   'initialize var
   m_bRunning = False
   Set m_pBaseGRID = Nothing
   
   Call chkComputeQuartile_Click
   For i = 0 To 1
      With cboTransform(i)
         .Clear
         .AddItem TRANS_NO
         .AddItem TRANS_LN
         '.AddItem = TRANS_LOG10
         .ListIndex = 0
         .Enabled = False
      End With
   Next
   progbar.Value = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bRunning Then
      Cancel = 1
      Exit Sub
   End If
   If Not (m_pBaseGRID Is Nothing) Then Set m_pBaseGRID = Nothing
End Sub

Private Function GetSaveFileName() As String
   comdlg.DialogTitle = "Save GRID"
   comdlg.FileName = ""
   GetSaveFileName = GetFileName(comdlg, False, , ".asc")
End Function

Private Sub SetProgressBarValue(iValue As Integer)
   If iValue > 100 Or iValue < 0 Then Exit Sub
   With progbar
      .Value = iValue
      .Refresh
   End With
End Sub
