VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSlopeShape 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Slope Shape Description"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   2580
      TabIndex        =   3
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cdmQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6900
      TabIndex        =   2
      Top             =   7920
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTabFunc 
      Height          =   7395
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   13044
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Step 1: Search upslope"
      TabPicture(0)   =   "frmSlopeShape.frx":0000
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
      TabCaption(1)   =   "Step 2: Search downslope"
      TabPicture(1)   =   "frmSlopeShape.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblInfo(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameInput(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frameOutput(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "framePara(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Step 3: Calculate slope shape"
      TabPicture(2)   =   "frmSlopeShape.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblInfo(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frameInput(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "frameOutput(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "framePara(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   795
         Index           =   2
         Left            =   -74880
         TabIndex        =   96
         Top             =   3180
         Width           =   11055
         Begin VB.Label Label2 
            Caption         =   $"frmSlopeShape.frx":0054
            Height          =   435
            Index           =   2
            Left            =   240
            TabIndex        =   97
            Top             =   300
            Width           =   10635
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   3315
         Index           =   2
         Left            =   -74880
         TabIndex        =   81
         Top             =   4020
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID2 
            Caption         =   "Downslope Shape"
            Height          =   375
            Index           =   6
            Left            =   60
            TabIndex        =   95
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox txtSaveGRID2 
            Height          =   375
            Index           =   6
            Left            =   2220
            TabIndex        =   94
            Top             =   2760
            Width           =   8715
         End
         Begin VB.CommandButton cmdSaveGRID2 
            Caption         =   "Min Downslope Rel. Relief"
            Height          =   375
            Index           =   5
            Left            =   60
            TabIndex        =   93
            Top             =   2400
            Width           =   2175
         End
         Begin VB.TextBox txtSaveGRID2 
            Height          =   375
            Index           =   5
            Left            =   2220
            TabIndex        =   92
            Top             =   2400
            Width           =   8715
         End
         Begin VB.CommandButton cmdSaveGRID2 
            Caption         =   "Max Downslope Rel. Relief"
            Height          =   375
            Index           =   4
            Left            =   60
            TabIndex        =   91
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txtSaveGRID2 
            Height          =   375
            Index           =   4
            Left            =   2220
            TabIndex        =   90
            Top             =   2040
            Width           =   8715
         End
         Begin VB.CommandButton cmdSaveGRID2 
            Caption         =   "Upslope Shape"
            Height          =   375
            Index           =   3
            Left            =   60
            TabIndex        =   89
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtSaveGRID2 
            Height          =   375
            Index           =   3
            Left            =   2220
            TabIndex        =   88
            Top             =   1560
            Width           =   8715
         End
         Begin VB.CommandButton cmdSaveGRID2 
            Caption         =   "Min Upslope Relative Relief"
            Height          =   375
            Index           =   2
            Left            =   60
            TabIndex        =   87
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox txtSaveGRID2 
            Height          =   375
            Index           =   2
            Left            =   2220
            TabIndex        =   86
            Top             =   1200
            Width           =   8715
         End
         Begin VB.CommandButton cmdSaveGRID2 
            Caption         =   "Max Upslope Relative Relief"
            Height          =   375
            Index           =   1
            Left            =   60
            TabIndex        =   85
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtSaveGRID2 
            Height          =   375
            Index           =   1
            Left            =   2220
            TabIndex        =   84
            Top             =   840
            Width           =   8715
         End
         Begin VB.CommandButton cmdSaveGRID2 
            Caption         =   "Slope Shape"
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   83
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtSaveGRID2 
            Height          =   375
            Index           =   0
            Left            =   2220
            TabIndex        =   82
            Top             =   360
            Width           =   8715
         End
         Begin VB.Line Line2 
            X1              =   60
            X2              =   10920
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   10920
            Y1              =   780
            Y2              =   780
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2355
         Index           =   2
         Left            =   -74880
         TabIndex        =   69
         Top             =   720
         Width           =   11055
         Begin VB.TextBox txtSrcGRID2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1620
            TabIndex        =   79
            Top             =   1080
            Width           =   9315
         End
         Begin VB.TextBox txtSrcGRID2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1620
            TabIndex        =   78
            Top             =   720
            Width           =   9315
         End
         Begin VB.CommandButton cmdSrcGRID2 
            Caption         =   "Upslope Direction"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   77
            Top             =   1080
            Width           =   1515
         End
         Begin VB.CommandButton cmdSrcGRID2 
            Caption         =   "Upslope Cells"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   1515
         End
         Begin VB.CommandButton cmdSrcGRID2 
            Caption         =   "Relief Difference"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox txtSrcGRID2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1620
            TabIndex        =   74
            Top             =   360
            Width           =   9315
         End
         Begin VB.TextBox txtSrcGRID2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1620
            TabIndex        =   73
            Top             =   1440
            Width           =   9315
         End
         Begin VB.CommandButton cmdSrcGRID2 
            Caption         =   "Downslope Cells"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   72
            Top             =   1440
            Width           =   1515
         End
         Begin VB.TextBox txtSrcGRID2 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1620
            TabIndex        =   71
            Top             =   1800
            Width           =   9315
         End
         Begin VB.CommandButton cmdSrcGRID2 
            Caption         =   "ArcInfo FlowDir"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   70
            Top             =   1800
            Width           =   1515
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   1
         Left            =   -74880
         TabIndex        =   62
         Top             =   4080
         Width           =   11055
         Begin VB.TextBox txtValleyTag 
            Height          =   315
            Index           =   1
            Left            =   1380
            TabIndex        =   63
            Text            =   "1"
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Valley Tag"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1875
         Index           =   1
         Left            =   -74880
         TabIndex        =   55
         Top             =   5220
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   0
            Left            =   1500
            TabIndex        =   61
            Top             =   360
            Width           =   9435
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Downslope Cells"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   1395
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   1
            Left            =   1500
            TabIndex        =   59
            Top             =   780
            Width           =   9435
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Downslope Relief"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   780
            Width           =   1395
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   2
            Left            =   1500
            TabIndex        =   57
            Top             =   1200
            Width           =   9435
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Relief Difference"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   1395
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   3075
         Index           =   1
         Left            =   -74880
         TabIndex        =   34
         Top             =   840
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID1 
            Caption         =   "Upslope Relief"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   68
            Top             =   2580
            Width           =   1275
         End
         Begin VB.TextBox txtSrcGRID1 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1380
            TabIndex        =   67
            Top             =   2580
            Width           =   9555
         End
         Begin VB.CommandButton cmdSrcGRID1 
            Caption         =   "Upslope Cells"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   66
            Top             =   2220
            Width           =   1275
         End
         Begin VB.TextBox txtSrcGRID1 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1380
            TabIndex        =   65
            Top             =   2220
            Width           =   9555
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   780
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   1
               Left            =   600
               TabIndex        =   47
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   1
               Left            =   2160
               TabIndex        =   46
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   1
               Left            =   4140
               TabIndex        =   45
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   1
               Left            =   6180
               TabIndex        =   44
               Text            =   "0"
               Top             =   240
               Width           =   1035
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
               Index           =   11
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   10
               Left            =   1620
               TabIndex        =   52
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   9
               Left            =   3240
               TabIndex        =   51
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   8
               Left            =   5280
               TabIndex        =   50
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   7
               Left            =   7320
               TabIndex        =   49
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   6
               Left            =   8880
               TabIndex        =   48
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID1 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1380
            TabIndex        =   40
            Top             =   360
            Width           =   9555
         End
         Begin VB.CommandButton cmdSrcGRID1 
            Caption         =   "DEM"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1275
         End
         Begin VB.CommandButton cmdSrcGRID1 
            Caption         =   "ArcInfo FlowDir"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   1500
            Width           =   1275
         End
         Begin VB.CommandButton cmdSrcGRID1 
            Caption         =   "Valley"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   1860
            Width           =   1275
         End
         Begin VB.TextBox txtSrcGRID1 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1380
            TabIndex        =   36
            Top             =   1500
            Width           =   9555
         End
         Begin VB.TextBox txtSrcGRID1 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1380
            TabIndex        =   35
            Top             =   1860
            Width           =   9555
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   3600
         Width           =   11055
         Begin VB.TextBox txtValleyTag 
            Height          =   315
            Index           =   0
            Left            =   1380
            TabIndex        =   29
            Text            =   "1"
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Valley Tag"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1875
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   4920
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID0 
            Caption         =   "Upslope Direction"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   1395
         End
         Begin VB.TextBox txtSaveGRID0 
            Height          =   375
            Index           =   2
            Left            =   1500
            TabIndex        =   32
            Top             =   1200
            Width           =   9435
         End
         Begin VB.CommandButton cmdSaveGRID0 
            Caption         =   "Upslope Relief"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   780
            Width           =   1395
         End
         Begin VB.TextBox txtSaveGRID0 
            Height          =   375
            Index           =   1
            Left            =   1500
            TabIndex        =   30
            Top             =   780
            Width           =   9435
         End
         Begin VB.CommandButton cmdSaveGRID0 
            Caption         =   "Upslope Cells"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1395
         End
         Begin VB.TextBox txtSaveGRID0 
            Height          =   375
            Index           =   0
            Left            =   1500
            TabIndex        =   22
            Top             =   360
            Width           =   9435
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2355
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   11055
         Begin VB.TextBox txtSrcGRID0 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1380
            TabIndex        =   27
            Top             =   1860
            Width           =   9555
         End
         Begin VB.TextBox txtSrcGRID0 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1380
            TabIndex        =   26
            Top             =   1500
            Width           =   9555
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "Valley"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1860
            Width           =   1275
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "ArcInfo FlowDir"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   1500
            Width           =   1275
         End
         Begin VB.CommandButton cmdSrcGRID0 
            Caption         =   "DEM"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtSrcGRID0 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1380
            TabIndex        =   18
            Top             =   360
            Width           =   9555
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   780
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   0
               Left            =   10020
               TabIndex        =   11
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   0
               Left            =   8100
               TabIndex        =   10
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   0
               Left            =   6180
               TabIndex        =   9
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   0
               Left            =   4140
               TabIndex        =   8
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   0
               Left            =   2160
               TabIndex        =   7
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   0
               Left            =   600
               TabIndex        =   6
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   5
               Left            =   8880
               TabIndex        =   17
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   4
               Left            =   7320
               TabIndex        =   16
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   3
               Left            =   5280
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   2
               Left            =   3240
               TabIndex        =   14
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   13
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   "Calculate slope shape"
         Height          =   375
         Index           =   2
         Left            =   -74820
         TabIndex        =   80
         Top             =   420
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Search downslope"
         Height          =   375
         Index           =   1
         Left            =   -74820
         TabIndex        =   54
         Top             =   420
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Search upslope"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   420
         Width           =   10935
      End
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   1020
      Top             =   7980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   60
      TabIndex        =   98
      Top             =   7440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSlopeShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_VALLEY = 1

Const C_SLPSHAPE_CONCAVE = 1
Const C_SLPSHAPE_STRAIGHT = 3
Const C_SLPSHAPE_CONVEX = 4
Const C_SLPSHAPE_UPCONCAVE_DOWNCONVEX = 2
Const C_SLPSHAPE_UPCONVEX_DOWNCONCAVE = 5

Dim m_bRunning As Boolean

Dim mpDEM As clsGrid, mpFlowDir As clsGrid, mpRidge As clsGrid, mpValley As clsGrid
Dim mpUpslpCells As clsGrid, mpUpslpRlf As clsGrid, mpUpslpDir As clsGrid
Dim mpDownslpCells As clsGrid, mpDownslpRlf As clsGrid
Dim mpRlfDiffer As clsGrid
Dim mpSlpShape As clsGrid
Dim mpUpRelRlfMax As clsGrid, mpUpRelRlfMin As clsGrid, mpUpSlpShape As clsGrid
Dim mpDRelRlfMax As clsGrid, mpDRelRlfMin As clsGrid, mpDSlpShape As clsGrid

Dim miRows As Integer, miCols As Integer, mdCell As Double
Dim mpTag As clsGrid

Dim miVlyTag As Integer
Dim msDEM As String, msFlowDir As String, msVly As String
Dim msUpslpCells As String, msUpslpRlf As String, msUpslpDir As String
Dim msDwnslpCells As String, msDwnslpRlf As String, msRlfDiff As String
Dim msSlpShape As String
Dim msUpslpRelRlfMax As String, msUpslpRelRlfMin As String, msUpslpShape As String
Dim msDwnslpRelRlfMax As String, msDwnslpRelRlfMin As String, msDwnslpShape As String

'Step 1:
Private Function SlopeDescrib_SearchUpslope() As Boolean
   Dim iRow As Integer, iCol As Integer
On Error GoTo ErrH
   SlopeDescrib_SearchUpslope = False
   
   Set mpDEM = New clsGrid
   mpDEM.LoadAscGrid msDEM
   Set mpFlowDir = New clsGrid
   mpFlowDir.LoadAscGrid msFlowDir, True
   Set mpValley = New clsGrid
   mpValley.LoadAscGrid msVly, True
   
   With mpDEM
      miRows = .nRows: miCols = .nCols: mdCell = .CellSize
      
      If miRows <> mpFlowDir.nRows Or miCols <> mpFlowDir.nCols _
            Or miRows <> mpValley.nRows Or miCols <> mpValley.nCols Then
         Err.Raise Number:=vbObjectError + 513, Description:="INPUT GRID are not with same size"
      End If
      
      Set mpUpslpCells = New clsGrid
      If Not mpUpslpCells.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpUpslpRlf = New clsGrid
      If Not mpUpslpRlf.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpUpslpDir = New clsGrid
      If Not mpUpslpDir.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpTag = New clsGrid
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, 0, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   
   '''''''''''''''''''''''''
   ' search upslope
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = 0 Then SearchUpslope iCol, iRow
         DoEvents
      Next
      SetProgressBarValue Int((iRow + 1) * 100# / miRows)
      DoEvents
   Next
   
   ''''''''''''''''''''
   ' output para
   If msUpslpCells <> "" Then
      If mpUpslpCells.SaveAscGrid(msUpslpCells, , 0) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpCells
      End If
   End If
   If msUpslpRlf <> "" Then
      If mpUpslpRlf.SaveAscGrid(msUpslpRlf, , 2) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpRlf
      End If
   End If
   If msUpslpDir <> "" Then
      If mpUpslpDir.SaveAscGrid(msUpslpDir, , 0) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpDir
      End If
   End If
   
   SlopeDescrib_SearchUpslope = True
ErrH:
   If Not SlopeDescrib_SearchUpslope Then
      If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   End If
   On Error Resume Next
   Set mpTag = Nothing
   Set mpDEM = Nothing:   Set mpFlowDir = Nothing: Set mpValley = Nothing
   Set mpUpslpCells = Nothing:   Set mpUpslpRlf = Nothing:   Set mpUpslpDir = Nothing
End Function

'
Private Function SearchUpslope(iProcCol As Integer, iProcRow As Integer) As Boolean
   Dim dElev As Double, dRelief As Double
   Dim iTempRow As Integer, iTempCol As Integer
   Dim iDir As Integer, iDirTemp As Integer, iFlowDir As Integer
   Dim dVal As Double
   Dim iUpslpCount As Integer
   
   SearchUpslope = False
   
   dElev = mpDEM.Cell(iProcCol, iProcRow)
   If dElev = mpDEM.NoData_Value Then
      mpUpslpCells.Cell(iProcCol, iProcRow) = mpUpslpCells.NoData_Value
      mpUpslpRlf.Cell(iProcCol, iProcRow) = mpUpslpRlf.NoData_Value
      mpUpslpDir.Cell(iProcCol, iProcRow) = mpUpslpDir.NoData_Value
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
                        SearchUpslope iTempCol, iTempRow
                     End If
                     
                     If mpUpslpRlf.Cell(iTempCol, iTempRow) >= 0 And mpDEM.Cell(iTempCol, iTempRow) <> mpDEM.NoData_Value Then
                        dRelief = mpUpslpRlf.Cell(iTempCol, iTempRow) + mpDEM.Cell(iTempCol, iTempRow) - dElev
                        If mpUpslpRlf.Cell(iProcCol, iProcRow) < dRelief Then
                           mpUpslpRlf.Cell(iProcCol, iProcRow) = dRelief
                           mpUpslpCells.Cell(iProcCol, iProcRow) = mpUpslpCells.Cell(iTempCol, iTempRow) + 1
                           mpUpslpDir.Cell(iProcCol, iProcRow) = ESRIDir(iDir)
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
         mpUpslpRlf.Cell(iProcCol, iProcRow) = 0#
         mpUpslpDir.Cell(iProcCol, iProcRow) = ESRI_DIR_UNDEF
      End If
   End If
   
   mpTag.Cell(iProcCol, iProcRow) = 1
   SearchUpslope = True
End Function

' Step 2:
Private Function SlopeDescrib_SearchDownslope() As Boolean
   Dim iRow As Integer, iCol As Integer
On Error GoTo ErrH
   SlopeDescrib_SearchDownslope = False
   
   Set mpDEM = New clsGrid
   mpDEM.LoadAscGrid msDEM
   Set mpFlowDir = New clsGrid
   mpFlowDir.LoadAscGrid msFlowDir, True
   Set mpValley = New clsGrid
   mpValley.LoadAscGrid msVly, True
   Set mpUpslpCells = New clsGrid
   mpUpslpCells.LoadAscGrid msUpslpCells, True
   Set mpUpslpRlf = New clsGrid
   mpUpslpRlf.LoadAscGrid msUpslpRlf
   
   With mpDEM
      miRows = .nRows: miCols = .nCols: mdCell = .CellSize
      
      If miRows <> mpFlowDir.nRows Or miCols <> mpFlowDir.nCols _
            Or miRows <> mpValley.nRows Or miCols <> mpValley.nCols _
            Or miRows <> mpUpslpCells.nRows Or miCols <> mpUpslpCells.nCols _
            Or miRows <> mpUpslpRlf.nRows Or miCols <> mpUpslpRlf.nCols Then
         Err.Raise Number:=vbObjectError + 513, Description:="INPUT GRID are not with same size"
      End If
      
      Set mpDownslpCells = New clsGrid
      If Not mpDownslpCells.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpDownslpRlf = New clsGrid
      If Not mpDownslpRlf.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpRlfDiffer = New clsGrid
      If Not mpRlfDiffer.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpTag = New clsGrid
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, 0, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   '''''''''''''''''''''''''
   ' search downslope
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = 0 Then SearchDownslope iCol, iRow
         DoEvents
      Next
      SetProgressBarValue Int((iRow + 1) * 100# / miRows)
      DoEvents
   Next
      
   ' calc Relief Difference
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpDownslpCells.Cell(iCol, iRow) = mpDownslpCells.NoData_Value _
               Or mpDownslpRlf.Cell(iCol, iRow) = mpDownslpRlf.NoData_Value _
               Or mpUpslpCells.Cell(iCol, iRow) = mpUpslpCells.NoData_Value _
               Or mpUpslpRlf.Cell(iCol, iRow) = mpUpslpRlf.NoData_Value Then
            mpRlfDiffer.Cell(iCol, iRow) = mpRlfDiffer.NoData_Value
         Else
            mpRlfDiffer.Cell(iCol, iRow) = mpDownslpRlf.Cell(iCol, iRow) _
                  - (mpUpslpRlf.Cell(iCol, iRow) + mpDownslpRlf.Cell(iCol, iRow)) * mpDownslpCells.Cell(iCol, iRow) _
                  / (mpDownslpCells.Cell(iCol, iRow) + 1 + mpUpslpCells.Cell(iCol, iRow))
         End If
      Next
      SetProgressBarValue Int((iRow + 1) * 100# / miRows)
      DoEvents
   Next
   
   ''''''''''''''''''''
   ' output para
   If msDwnslpCells <> "" Then
      If mpDownslpCells.SaveAscGrid(msDwnslpCells, , 0) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msDwnslpCells
      End If
   End If
   If msDwnslpRlf <> "" Then
      If mpDownslpRlf.SaveAscGrid(msDwnslpRlf, , 2) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msDwnslpRlf
      End If
   End If
   If msRlfDiff <> "" Then
      If mpRlfDiffer.SaveAscGrid(msRlfDiff, , 2) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msRlfDiff
      End If
   End If
   
   SlopeDescrib_SearchDownslope = True
ErrH:
   If Not SlopeDescrib_SearchDownslope Then
      If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   End If
   On Error Resume Next
   Set mpTag = Nothing
   Set mpDEM = Nothing:   Set mpFlowDir = Nothing: Set mpValley = Nothing
   Set mpUpslpCells = Nothing:   Set mpUpslpRlf = Nothing
   Set mpDownslpCells = Nothing:   Set mpDownslpRlf = Nothing
   Set mpRlfDiffer = Nothing
      
End Function
'
Private Function SearchDownslope(iProcCol As Integer, iProcRow As Integer) As Boolean
   Dim dElev As Double, dRelief As Double
   Dim iTempRow As Integer, iTempCol As Integer
   Dim iDir As Integer, iDirTemp As Integer, iFlowDir As Integer
   Dim dVal As Double
   
   SearchDownslope = False
   
   dElev = mpDEM.Cell(iProcCol, iProcRow)
   If dElev = mpDEM.NoData_Value Then
      mpDownslpCells.Cell(iProcCol, iProcRow) = mpDownslpCells.NoData_Value
      mpDownslpRlf.Cell(iProcCol, iProcRow) = mpDownslpRlf.NoData_Value
   Else
      If mpFlowDir.IsValidCellValue(iProcCol, iProcRow, dVal) And mpValley.Cell(iProcCol, iProcRow) <> miVlyTag Then
         iFlowDir = dVal
         iDir = GetESRIDir_ArrayIndex(iFlowDir)
         If iDir > 0 Then
            iTempCol = iProcCol + ArrDir8X(iDir): iTempRow = iProcRow + ArrDir8Y(iDir)
            
            If mpTag.IsValidCellValue(iTempCol, iTempRow, dVal) Then
               If dVal = 0 Then
                  SearchDownslope iTempCol, iTempRow
               End If
               'If mpDownslpRlf.Cell(iTempCol, iTempRow) >= 0 And mpDEM.Cell(iTempCol, iTempRow) <> mpDEM.NoData_Value Then
               If mpDownslpRlf.Cell(iTempCol, iTempRow) <> mpDownslpRlf.NoData_Value And mpDEM.Cell(iTempCol, iTempRow) <> mpDEM.NoData_Value Then
                  dRelief = mpDownslpRlf.Cell(iTempCol, iTempRow) + dElev - mpDEM.Cell(iTempCol, iTempRow)
                  mpDownslpRlf.Cell(iProcCol, iProcRow) = dRelief
                  mpDownslpCells.Cell(iProcCol, iProcRow) = mpDownslpCells.Cell(iTempCol, iTempRow) + 1
               End If
            Else  ' flow-out point of all study area
               mpDownslpCells.Cell(iProcCol, iProcRow) = 0
               mpDownslpRlf.Cell(iProcCol, iProcRow) = 0#
            End If
         Else
            mpDownslpCells.Cell(iProcCol, iProcRow) = 0
            mpDownslpRlf.Cell(iProcCol, iProcRow) = 0#
         End If
      Else
         mpDownslpCells.Cell(iProcCol, iProcRow) = 0
         mpDownslpRlf.Cell(iProcCol, iProcRow) = 0#
      End If
   End If
      
   mpTag.Cell(iProcCol, iProcRow) = 1
   SearchDownslope = True
End Function

' Step 3:
Private Function SlopeDescrib_CalcSlopeShape() As Boolean

On Error GoTo ErrH
   SlopeDescrib_CalcSlopeShape = False
   
   Set mpRlfDiffer = New clsGrid
   mpRlfDiffer.LoadAscGrid msRlfDiff
   Set mpUpslpCells = New clsGrid
   mpUpslpCells.LoadAscGrid msUpslpCells, True
   Set mpUpslpDir = New clsGrid
   mpUpslpDir.LoadAscGrid msUpslpDir, True
   Set mpDownslpCells = New clsGrid
   mpDownslpCells.LoadAscGrid msDwnslpCells, True
   Set mpFlowDir = New clsGrid
   mpFlowDir.LoadAscGrid msFlowDir, True

   'create para Grid
   With mpRlfDiffer
      miRows = .nRows: miCols = .nCols: mdCell = .CellSize
      
      If miRows <> mpDownslpCells.nRows Or miCols <> mpDownslpCells.nCols _
            Or miRows <> mpFlowDir.nRows Or miCols <> mpFlowDir.nCols _
            Or miRows <> mpUpslpCells.nRows Or miCols <> mpUpslpCells.nCols _
            Or miRows <> mpUpslpDir.nRows Or miCols <> mpUpslpDir.nCols Then
         Err.Raise Number:=vbObjectError + 513, Description:="INPUT GRID are not with same size"
      End If
      
      If msSlpShape <> "" Then
         Set mpSlpShape = New clsGrid
         If Not mpSlpShape.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End If
      ' upslope part
      If msUpslpRelRlfMax <> "" Then
         Set mpUpRelRlfMax = New clsGrid
         If Not mpUpRelRlfMax.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End If
      If msUpslpRelRlfMin <> "" Then
         Set mpUpRelRlfMin = New clsGrid
         If Not mpUpRelRlfMin.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End If
      If msUpslpShape <> "" Then
         Set mpUpSlpShape = New clsGrid
         If Not mpUpSlpShape.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End If
      'downslope part
      If msDwnslpRelRlfMax <> "" Then
         Set mpDRelRlfMax = New clsGrid
         If Not mpDRelRlfMax.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End If
      If msDwnslpRelRlfMin <> "" Then
         Set mpDRelRlfMin = New clsGrid
         If Not mpDRelRlfMin.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End If
      If msDwnslpShape <> "" Then
         Set mpDSlpShape = New clsGrid
         If Not mpDSlpShape.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
         End If
      End If
   End With
   
   ' calc slopeShape
   Call SlopeShape((msSlpShape <> ""), (msUpslpShape <> ""), (msUpslpRelRlfMax <> ""), (msUpslpRelRlfMin <> ""), _
                  (msDwnslpShape <> ""), (msDwnslpRelRlfMax <> ""), (msDwnslpRelRlfMin <> ""))
   
   ' output slope-shape result
   If msSlpShape <> "" Then
      If mpSlpShape.SaveAscGrid(msSlpShape, , 0) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msSlpShape
      End If
   End If
   
   If msUpslpRelRlfMax <> "" Then
      If mpUpRelRlfMax.SaveAscGrid(msUpslpRelRlfMax, , 2) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpRelRlfMax
      End If
   End If
   
   If msUpslpRelRlfMin <> "" Then
      If mpUpRelRlfMin.SaveAscGrid(msUpslpRelRlfMin, , 2) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpRelRlfMin
      End If
   End If
   
   If msUpslpShape <> "" Then
      If mpUpSlpShape.SaveAscGrid(msUpslpShape, , 0) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msUpslpShape
      End If
   End If
   
   If msDwnslpRelRlfMax <> "" Then
      If mpDRelRlfMax.SaveAscGrid(msDwnslpRelRlfMax, , 2) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msDwnslpRelRlfMax
      End If
   End If
   
   If msDwnslpRelRlfMin <> "" Then
      If mpDRelRlfMin.SaveAscGrid(msDwnslpRelRlfMin, , 2) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msDwnslpRelRlfMin
      End If
   End If
   
   If msDwnslpShape <> "" Then
      If mpDSlpShape.SaveAscGrid(msDwnslpShape, , 0) Then
         'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
      Else
         Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & msDwnslpShape
      End If
   End If
   
   SlopeDescrib_CalcSlopeShape = True
ErrH:
   If Not SlopeDescrib_CalcSlopeShape Then
      If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   End If
   On Error Resume Next
   Set mpRlfDiffer = Nothing:   Set mpUpslpCells = Nothing: Set mpUpslpDir = Nothing
   Set mpDownslpCells = Nothing:   Set mpFlowDir = Nothing
   Set mpSlpShape = Nothing
   Set mpUpRelRlfMax = Nothing:  Set mpUpRelRlfMin = Nothing:   Set mpUpSlpShape = Nothing
   Set mpDRelRlfMax = Nothing:   Set mpDRelRlfMin = Nothing:   Set mpDSlpShape = Nothing
   
End Function
'
' search along both upslope and downslope flow-path
' find both position and value with max and min of RlfDiffer
'
Public Function SlopeShape(Optional blnCalcSlpShape As Boolean = True, _
                           Optional blnCalcUpslpShape As Boolean = False, _
                           Optional blnCalcUpRelRlfMax As Boolean = False, _
                           Optional blnCalcUpRelRlfMin As Boolean = False, _
                           Optional blnCalcDownslpShape As Boolean = False, _
                           Optional blnCalcDownRelRlfMax As Boolean = False, _
                           Optional blnCalcDownRelRlfMin As Boolean = False) As Boolean
   Dim iRow As Integer, iCol As Integer, iProcRow As Integer, iProcCol As Integer
   'Dim iTopRow As Integer, iTopCol As Integer, iBottomRow As Integer, iBottomCol As Integer
   'Dim dTop As Double, dBottom As Double, dHeight As Double, dRelief As Double
   'Dim iTempRow As Integer, iTempCol As Integer
'   Dim iDir As Integer
   'Dim lCells As Long
   Dim blnProc As Boolean
   
   Dim dElev As Double, dRelief As Double
   Dim iTempRow As Integer, iTempCol As Integer
   Dim iDir As Integer, iDirTemp As Integer, iFlowDir As Integer
   Dim dVal As Double, dTemp As Double
   Dim iUpslpCount As Integer, iDslpCount As Integer
   
   Dim dUpRelRlfMax As Double, iUpRelRlfMaxPos As Integer
   Dim dUpRelRlfMin As Double, iUpRelRlfMinPos As Integer
   Dim dDRelRlfMax As Double, iDRelRlfMaxPos As Integer
   Dim dDRelRlfMin As Double, iDRelRlfMinPos As Integer
   
   Dim iHighRow As Integer, iHighCol As Integer, dHigh As Double, iHighPos As Integer
   Dim iLowRow As Integer, iLowCol As Integer, dLow As Double, iLowPos As Integer
   
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         dVal = mpRlfDiffer.Cell(iCol, iRow)
         If dVal = mpRlfDiffer.NoData_Value Then
            If blnCalcSlpShape Then mpSlpShape.Cell(iCol, iRow) = mpSlpShape.NoData_Value
            
            If blnCalcUpslpShape Then mpUpSlpShape.Cell(iCol, iRow) = mpUpSlpShape.NoData_Value
            If blnCalcUpRelRlfMax Then mpUpRelRlfMax.Cell(iCol, iRow) = mpUpRelRlfMax.NoData_Value
            If blnCalcUpRelRlfMin Then mpUpRelRlfMin.Cell(iCol, iRow) = mpUpRelRlfMin.NoData_Value
                        
            If blnCalcDownslpShape Then mpDSlpShape.Cell(iCol, iRow) = mpDSlpShape.NoData_Value
            If blnCalcDownRelRlfMax Then mpDRelRlfMax.Cell(iCol, iRow) = mpDRelRlfMax.NoData_Value
            If blnCalcDownRelRlfMin Then mpDRelRlfMin.Cell(iCol, iRow) = mpDRelRlfMin.NoData_Value
         Else
            ' search along upslope direction of flow path
            iHighRow = iRow: iHighCol = iCol
            dHigh = dVal: iHighPos = 0
            dUpRelRlfMax = dVal: iUpRelRlfMaxPos = 0
            dUpRelRlfMin = dVal: iUpRelRlfMinPos = 0
            iProcRow = iRow: iProcCol = iCol
            iUpslpCount = mpUpslpCells.Cell(iCol, iRow)
            
            Do While iUpslpCount > 0
               iFlowDir = mpUpslpDir.Cell(iProcCol, iProcRow)
               iDirTemp = GetESRIDir_ArrayIndex(iFlowDir)
               If iDirTemp <= 0 Then Err.Raise vbObjectError + 513, , "Failed in SlopeShape function"
               iProcCol = iProcCol + ArrDir8X(iDirTemp): iProcRow = iProcRow + ArrDir8Y(iDirTemp)
               iUpslpCount = iUpslpCount - 1
               
               dTemp = mpRlfDiffer.Cell(iProcCol, iProcRow)
               If dTemp <> mpRlfDiffer.NoData_Value Then
                  If dUpRelRlfMax <= dTemp Then
                     dUpRelRlfMax = dTemp
                     iUpRelRlfMaxPos = mpUpslpCells.Cell(iCol, iRow) - iUpslpCount
                  End If
                  If dUpRelRlfMin > dTemp Then
                     dUpRelRlfMin = dTemp
                     iUpRelRlfMinPos = mpUpslpCells.Cell(iCol, iRow) - iUpslpCount
                  End If
                  If dHigh <= dTemp Then
                     dHigh = dTemp
                     iHighRow = iProcRow: iHighCol = iProcCol
                     iHighPos = mpUpslpCells.Cell(iCol, iRow) - iUpslpCount
                  End If
                  If dLow > dTemp Then
                     dLow = dTemp
                     iLowRow = iProcRow: iLowCol = iProcCol
                     iLowPos = mpUpslpCells.Cell(iCol, iRow) - iUpslpCount
                  End If
               End If
            Loop
            If blnCalcUpRelRlfMax Then mpUpRelRlfMax.Cell(iCol, iRow) = dUpRelRlfMax
            If blnCalcUpRelRlfMin Then mpUpRelRlfMin.Cell(iCol, iRow) = dUpRelRlfMin
            
            If blnCalcUpslpShape Then
               'judge slope shape
               With mpUpSlpShape
                  If (dUpRelRlfMax <= 0 And dUpRelRlfMin < 0) Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_CONCAVE
                  ElseIf dUpRelRlfMax = 0 And dUpRelRlfMin = 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_STRAIGHT
                  ElseIf dUpRelRlfMax > 0 And dUpRelRlfMin >= 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_CONVEX
                  ElseIf dUpRelRlfMax > 0 And dUpRelRlfMin < 0 Then
                     If iUpRelRlfMaxPos < iUpRelRlfMinPos Then
                        .Cell(iCol, iRow) = C_SLPSHAPE_UPCONCAVE_DOWNCONVEX
                     Else
                        .Cell(iCol, iRow) = C_SLPSHAPE_UPCONVEX_DOWNCONCAVE
                     End If
                  End If
               End With
            End If
            
            ' search along downslope direction of flow path
            ' pos(in cells) is negative
            iLowRow = iRow: iLowCol = iCol
            dLow = dVal: iLowPos = 0
            dDRelRlfMax = dVal: iDRelRlfMaxPos = 0
            dDRelRlfMin = dVal: iDRelRlfMinPos = 0
            iProcRow = iRow: iProcCol = iCol
            iDslpCount = mpDownslpCells.Cell(iCol, iRow)
            
            Do While iDslpCount > 0
               iFlowDir = mpFlowDir.Cell(iProcCol, iProcRow)
               iDirTemp = GetESRIDir_ArrayIndex(iFlowDir)
               If iDirTemp <= 0 Then Err.Raise vbObjectError + 513, , "Failed in SlopeShape function"
               iProcCol = iProcCol + ArrDir8X(iDirTemp): iProcRow = iProcRow + ArrDir8Y(iDirTemp)
               iDslpCount = iDslpCount - 1
               
               dTemp = mpRlfDiffer.Cell(iProcCol, iProcRow)
               If dTemp <> mpRlfDiffer.NoData_Value Then
                  If dDRelRlfMax < dTemp Then
                     dDRelRlfMax = dTemp
                     iDRelRlfMaxPos = -(mpDownslpCells.Cell(iCol, iRow) - iDslpCount)
                  End If
                  If dDRelRlfMin >= dTemp Then
                     dDRelRlfMin = dTemp
                     iDRelRlfMinPos = -(mpDownslpCells.Cell(iCol, iRow) - iDslpCount)
                  End If
                  If dHigh < dTemp Then
                     dHigh = dTemp
                     iHighRow = iProcRow: iHighCol = iProcCol
                     iHighPos = -(mpDownslpCells.Cell(iCol, iRow) - iDslpCount)
                  End If
                  If dLow >= dTemp Then
                     dLow = dTemp
                     iLowRow = iProcRow: iLowCol = iProcCol
                     iLowPos = -(mpDownslpCells.Cell(iCol, iRow) - iDslpCount)
                  End If
               End If
            Loop
                        
            If blnCalcDownRelRlfMax Then mpDRelRlfMax.Cell(iCol, iRow) = dDRelRlfMax
            If blnCalcDownRelRlfMin Then mpDRelRlfMin.Cell(iCol, iRow) = dDRelRlfMin
                 
            If blnCalcDownslpShape Then
               'judge slope shape
               With mpDSlpShape
                  If dDRelRlfMax <= 0 And dDRelRlfMin < 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_CONCAVE
                  ElseIf dDRelRlfMax = 0 And dDRelRlfMin = 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_STRAIGHT
                  ElseIf dDRelRlfMax > 0 And dDRelRlfMin >= 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_CONVEX
                  ElseIf dDRelRlfMax > 0 And dDRelRlfMin < 0 Then
                     If iDRelRlfMaxPos < iDRelRlfMinPos Then
                        .Cell(iCol, iRow) = C_SLPSHAPE_UPCONCAVE_DOWNCONVEX
                     Else
                        .Cell(iCol, iRow) = C_SLPSHAPE_UPCONVEX_DOWNCONCAVE
                     End If
                  End If
               End With
            End If
            
            If blnCalcSlpShape Then
               With mpSlpShape
                  If dHigh = 0 And dLow < 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_CONCAVE
                  ElseIf dHigh = 0 And dLow = 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_STRAIGHT
                  ElseIf dHigh > 0 And dLow = 0 Then
                     .Cell(iCol, iRow) = C_SLPSHAPE_CONVEX
                  ElseIf dHigh > 0 And dLow < 0 Then
                     If iHighPos < iLowPos Then
                        .Cell(iCol, iRow) = C_SLPSHAPE_UPCONCAVE_DOWNCONVEX
                     Else
                        .Cell(iCol, iRow) = C_SLPSHAPE_UPCONVEX_DOWNCONCAVE
                     End If
                  End If
               End With
            End If
            
         End If
         DoEvents
      Next
      SetProgressBarValue Int((iRow + 1) * 100# / miRows)
      DoEvents
   Next
   
End Function

Private Sub cdmQuit_Click()
   If m_bRunning Then Exit Sub
   Unload Me
End Sub


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
      If msDEM = "" Or msFlowDir = "" Or msVly = "" Then
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
      msUpslpCells = txtSaveGRID0(0).Text
      msUpslpRlf = txtSaveGRID0(1).Text
      msUpslpDir = txtSaveGRID0(2).Text
      If msUpslpCells = "" And msUpslpRlf = "" And msUpslpDir = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the OUTPUT GRID firstly"
      End If
      
      txtSrcGRID1(0).Text = msDEM
      txtSrcGRID1(1).Text = msFlowDir
      txtSrcGRID1(2).Text = msVly
      txtSrcGRID1(3).Text = msUpslpCells
      txtSrcGRID1(4).Text = msUpslpRlf
      txtSrcGRID2(1).Text = msUpslpCells
      txtSrcGRID2(2).Text = msUpslpDir
      txtSrcGRID2(4).Text = msFlowDir
            
      If Not SlopeDescrib_SearchUpslope() Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error when searching upslope"
      End If
   Case 1
      msDEM = txtSrcGRID1(0).Text
      msFlowDir = txtSrcGRID1(1).Text
      msVly = txtSrcGRID1(2).Text
      msUpslpCells = txtSrcGRID1(3).Text
      msUpslpRlf = txtSrcGRID1(4).Text
      If msDEM = "" Or msFlowDir = "" Or msVly = "" Or msUpslpCells = "" Or msUpslpRlf = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the INPUT GRID firstly"
      End If
      With txtValleyTag(1)
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
      msDwnslpCells = txtSaveGRID1(0).Text
      msDwnslpRlf = txtSaveGRID1(1).Text
      msRlfDiff = txtSaveGRID1(2).Text
      If msDwnslpCells = "" And msDwnslpRlf = "" And msRlfDiff = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the OUTPUT GRID firstly"
      End If
      
      txtSrcGRID2(0).Text = msRlfDiff
      txtSrcGRID2(1).Text = msUpslpCells
      txtSrcGRID2(2).Text = msUpslpDir
      txtSrcGRID2(3).Text = msDwnslpCells
      txtSrcGRID2(4).Text = msFlowDir
      
      If Not SlopeDescrib_SearchDownslope() Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error when searching downslope"
      End If
   Case 2
      msRlfDiff = txtSrcGRID2(0).Text
      msUpslpCells = txtSrcGRID2(1).Text
      msUpslpDir = txtSrcGRID2(2).Text
      msDwnslpCells = txtSrcGRID2(3).Text
      msFlowDir = txtSrcGRID2(4).Text
      If msRlfDiff = "" Or msUpslpCells = "" Or msUpslpDir = "" Or msDwnslpCells = "" Or msFlowDir = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the INPUT GRID firstly"
      End If
      
      msSlpShape = txtSaveGRID2(0).Text
      msUpslpRelRlfMax = txtSaveGRID2(1).Text
      msUpslpRelRlfMin = txtSaveGRID2(2).Text
      msUpslpShape = txtSaveGRID2(3).Text
      msDwnslpRelRlfMax = txtSaveGRID2(4).Text
      msDwnslpRelRlfMin = txtSaveGRID2(5).Text
      msDwnslpShape = txtSaveGRID2(6).Text
      
      If msSlpShape = "" And msUpslpRelRlfMax = "" And msUpslpRelRlfMin = "" And msUpslpShape = "" _
            And msDwnslpRelRlfMax = "" And msDwnslpRelRlfMin = "" And msDwnslpShape = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign the OUTPUT GRID firstly"
      End If
      If Not ((msUpslpRelRlfMax = "" And msUpslpRelRlfMin = "" And msUpslpShape = "") _
            Or (msUpslpRelRlfMax <> "" And msUpslpRelRlfMin <> "" And msUpslpShape <> "")) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Three OUTPUT Upslope GRID should be with same status"
      End If
      If Not ((msDwnslpRelRlfMax = "" And msDwnslpRelRlfMin = "" And msDwnslpShape = "") _
            Or (msDwnslpRelRlfMax <> "" And msDwnslpRelRlfMin <> "" And msDwnslpShape <> "")) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Three OUTPUT Downslope GRID should be with same status"
      End If
      
      If Not SlopeDescrib_CalcSlopeShape() Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error when calculating slope shape"
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
' assign OUTPUT GRID for step 2
Private Sub cmdSaveGRID1_Click(Index As Integer)
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveGRID1(Index).Text = strFile
End Sub
' assign OUTPUT GRID for step 3
Private Sub cmdSaveGRID2_Click(Index As Integer)
   Dim strFile As String
   strFile = GetSaveFileName()
   If strFile <> "" Then txtSaveGRID2(Index).Text = strFile
End Sub

' assign source GRID for step 1
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
         
      txtSaveGRID0(0).Text = strPath & "UpslpCells.asc"
      txtSaveGRID0(1).Text = strPath & "UpslpRlf.asc"
      txtSaveGRID0(2).Text = strPath & "UpslpDir.asc"
      
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
' assign source GRID for step 2
Private Sub cmdSrcGRID1_Click(Index As Integer)
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
   txtSrcGRID1(Index).Text = strFile
      
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
         
      txtSaveGRID1(0).Text = strPath & "DwnslpCells.asc"
      txtSaveGRID1(1).Text = strPath & "DwnslpRlf.asc"
      txtSaveGRID1(2).Text = strPath & "RlfDiffer.asc"
      
      ' read parameters in SrcGRID file head
      On Error GoTo ErrH
      Set pGrid = New clsGrid
      With pGrid
         .LoadAscGrid strFile
         txtCols(1).Text = .nCols
         txtRows(1).Text = .nRows
         txtXll(1).Text = .xllcorner
         txtYll(1).Text = .yllcorner
         txtCellSize(1).Text = .CellSize
         txtNoData(1).Text = .NoData_Value
      End With
   End If
ErrH:
   Set pGrid = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   
   Me.MousePointer = 0
   m_bRunning = False
End Sub
' assign source GRID for step 3
Private Sub cmdSrcGRID2_Click(Index As Integer)
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
   txtSrcGRID2(Index).Text = strFile
   
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
         
      txtSaveGRID2(0).Text = strPath & "SlpShape.asc"
      txtSaveGRID2(1).Text = strPath & "UpRelRlfMax.asc"
      txtSaveGRID2(2).Text = strPath & "UpRelRlfMin.asc"
      txtSaveGRID2(3).Text = strPath & "UpslpShape.asc"
      txtSaveGRID2(4).Text = strPath & "DRelRlfMax.asc"
      txtSaveGRID2(5).Text = strPath & "DRelRlfMin.asc"
      txtSaveGRID2(6).Text = strPath & "DwnSlpShape.asc"
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
   Set mpDEM = Nothing:   Set mpFlowDir = Nothing:   Set mpRidge = Nothing: Set mpValley = Nothing
   Set mpUpslpCells = Nothing:   Set mpUpslpRlf = Nothing:   Set mpUpslpDir = Nothing
   Set mpDownslpCells = Nothing:   Set mpDownslpRlf = Nothing
   Set mpRlfDiffer = Nothing
   Set mpSlpShape = Nothing
   Set mpUpRelRlfMax = Nothing:  Set mpUpRelRlfMin = Nothing:   Set mpUpSlpShape = Nothing
   Set mpDRelRlfMax = Nothing:   Set mpDRelRlfMin = Nothing:   Set mpDSlpShape = Nothing
End Function

Private Function GetSaveFileName() As String
   comdlg.DialogTitle = "Save GRID"
   comdlg.FileName = ""
   GetSaveFileName = GetFileName(comdlg, False, , ".asc")
End Function

