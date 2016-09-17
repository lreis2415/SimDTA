VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDTAFunc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DTA Functions"
   ClientHeight    =   8115
   ClientLeft      =   3810
   ClientTop       =   3690
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   780
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cdmQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   7440
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTabFunc 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   24
      Tab             =   9
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Fill dep."
      TabPicture(0)   =   "frmDTAFunc.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(1)=   "frameInput(0)"
      Tab(0).Control(2)=   "framePara(0)"
      Tab(0).Control(3)=   "frameOutput(0)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Slope"
      TabPicture(1)   =   "frmDTAFunc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblInfo(1)"
      Tab(1).Control(1)=   "frameInput(1)"
      Tab(1).Control(2)=   "framePara(1)"
      Tab(1).Control(3)=   "frameOutput(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Aspect"
      TabPicture(2)   =   "frmDTAFunc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblInfo(2)"
      Tab(2).Control(1)=   "frameOutput(2)"
      Tab(2).Control(2)=   "framePara(2)"
      Tab(2).Control(3)=   "frameInput(2)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Curvatures"
      TabPicture(3)   =   "frmDTAFunc.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblInfo(3)"
      Tab(3).Control(1)=   "frameOutput(3)"
      Tab(3).Control(2)=   "framePara(3)"
      Tab(3).Control(3)=   "frameInput(3)"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Surface Curvature Index"
      TabPicture(4)   =   "frmDTAFunc.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frameInput(4)"
      Tab(4).Control(1)=   "optCs_WinShape(4)"
      Tab(4).Control(2)=   "frameOutput(4)"
      Tab(4).Control(3)=   "lblInfo(4)"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Elevation Percentil Index"
      TabPicture(5)   =   "frmDTAFunc.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "frameInput(5)"
      Tab(5).Control(1)=   "framePara(5)"
      Tab(5).Control(2)=   "frameOutput(5)"
      Tab(5).Control(3)=   "lblInfo(5)"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Hill/Hillslope/Valley Index"
      TabPicture(6)   =   "frmDTAFunc.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "frameInput(6)"
      Tab(6).Control(1)=   "framePara(6)"
      Tab(6).Control(2)=   "frameOutput(6)"
      Tab(6).Control(3)=   "lblInfo(6)"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Terrain Ruggedness Index"
      TabPicture(7)   =   "frmDTAFunc.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblInfo(7)"
      Tab(7).Control(1)=   "frameOutput(7)"
      Tab(7).Control(2)=   "framePara(7)"
      Tab(7).Control(3)=   "frameInput(7)"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "Landscape Position Index"
      TabPicture(8)   =   "frmDTAFunc.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "frameInput(8)"
      Tab(8).Control(1)=   "framePara(8)"
      Tab(8).Control(2)=   "frameOutput(8)"
      Tab(8).Control(3)=   "lblInfo(8)"
      Tab(8).ControlCount=   4
      TabCaption(9)   =   "UPNESS Index"
      TabPicture(9)   =   "frmDTAFunc.frx":00FC
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "lblInfo(9)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "frameOutput(9)"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "framePara(9)"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "frameInput(9)"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).ControlCount=   4
      TabCaption(10)  =   "Relative Position Index"
      TabPicture(10)  =   "frmDTAFunc.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lblInfo(10)"
      Tab(10).Control(1)=   "frameOutput(10)"
      Tab(10).Control(2)=   "framePara(10)"
      Tab(10).Control(3)=   "frameInput(10)"
      Tab(10).ControlCount=   4
      TabCaption(11)  =   "Downslope Index"
      TabPicture(11)  =   "frmDTAFunc.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "frameInput(11)"
      Tab(11).Control(1)=   "framePara(11)"
      Tab(11).Control(2)=   "frameOutput(11)"
      Tab(11).Control(3)=   "lblInfo(11)"
      Tab(11).ControlCount=   4
      TabCaption(12)  =   "Extract Ridge"
      TabPicture(12)  =   "frmDTAFunc.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "frameOutput(12)"
      Tab(12).Control(1)=   "framePara(12)"
      Tab(12).Control(2)=   "frameInput(12)"
      Tab(12).Control(3)=   "lblInfo(12)"
      Tab(12).ControlCount=   4
      TabCaption(13)  =   "Extract Drainage Networks"
      TabPicture(13)  =   "frmDTAFunc.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "lblInfo(13)"
      Tab(13).Control(1)=   "frameInput(13)"
      Tab(13).Control(2)=   "framePara(13)"
      Tab(13).Control(3)=   "frameOutput(13)"
      Tab(13).ControlCount=   4
      TabCaption(14)  =   "MFD"
      TabPicture(14)  =   "frmDTAFunc.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "frameOutput(14)"
      Tab(14).Control(1)=   "framePara(14)"
      Tab(14).Control(2)=   "frameInput(14)"
      Tab(14).Control(3)=   "lblInfo(14)"
      Tab(14).ControlCount=   4
      TabCaption(15)  =   "TWI"
      TabPicture(15)  =   "frmDTAFunc.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "frameOutput(15)"
      Tab(15).Control(1)=   "framePara(15)"
      Tab(15).Control(2)=   "frameInput(15)"
      Tab(15).Control(3)=   "lblInfo(15)"
      Tab(15).ControlCount=   4
      TabCaption(16)  =   "Terrain Char. Index"
      TabPicture(16)  =   "frmDTAFunc.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "lblInfo(16)"
      Tab(16).Control(1)=   "frameInput(16)"
      Tab(16).Control(2)=   "framePara(16)"
      Tab(16).Control(3)=   "frameOutput(16)"
      Tab(16).ControlCount=   4
      TabCaption(17)  =   "Stream Power Index"
      TabPicture(17)  =   "frmDTAFunc.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "lblInfo(17)"
      Tab(17).Control(1)=   "frameOutput(17)"
      Tab(17).Control(2)=   "frameInput(17)"
      Tab(17).Control(3)=   "framePara(17)"
      Tab(17).ControlCount=   4
      TabCaption(18)  =   "Elevation-Relief Ratio"
      TabPicture(18)  =   "frmDTAFunc.frx":01F8
      Tab(18).ControlEnabled=   0   'False
      Tab(18).Control(0)=   "lblInfo(18)"
      Tab(18).Control(1)=   "frameInput(18)"
      Tab(18).Control(2)=   "framePara(4)"
      Tab(18).Control(3)=   "frameOutput(18)"
      Tab(18).ControlCount=   4
      TabCaption(19)  =   "Relief"
      TabPicture(19)  =   "frmDTAFunc.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "lblInfo(19)"
      Tab(19).Control(1)=   "frameInput(19)"
      Tab(19).Control(2)=   "optCs_WinShape(0)"
      Tab(19).Control(3)=   "frameOutput(19)"
      Tab(19).ControlCount=   4
      TabCaption(20)  =   "Topographic Position Index"
      TabPicture(20)  =   "frmDTAFunc.frx":0230
      Tab(20).ControlEnabled=   0   'False
      Tab(20).Control(0)=   "frameOutput(20)"
      Tab(20).Control(1)=   "optCs_WinShape(1)"
      Tab(20).Control(2)=   "frameInput(20)"
      Tab(20).Control(3)=   "lblInfo(20)"
      Tab(20).ControlCount=   4
      TabCaption(21)  =   "Surface Area"
      TabPicture(21)  =   "frmDTAFunc.frx":024C
      Tab(21).ControlEnabled=   0   'False
      Tab(21).Control(0)=   "frameOutput(21)"
      Tab(21).Control(1)=   "framePara(21)"
      Tab(21).Control(2)=   "frameInput(21)"
      Tab(21).Control(3)=   "lblInfo(21)"
      Tab(21).ControlCount=   4
      TabCaption(22)  =   "Openness Angle"
      TabPicture(22)  =   "frmDTAFunc.frx":0268
      Tab(22).ControlEnabled=   0   'False
      Tab(22).Control(0)=   "lblInfo(22)"
      Tab(22).Control(1)=   "frameInput(22)"
      Tab(22).Control(2)=   "framePara(18)"
      Tab(22).Control(3)=   "frameOutput(22)"
      Tab(22).ControlCount=   4
      TabCaption(23)  =   "Relative Relief Index"
      TabPicture(23)  =   "frmDTAFunc.frx":0284
      Tab(23).ControlEnabled=   0   'False
      Tab(23).Control(0)=   "lblInfo(23)"
      Tab(23).Control(0).Enabled=   0   'False
      Tab(23).Control(1)=   "frameInput(23)"
      Tab(23).Control(1).Enabled=   0   'False
      Tab(23).Control(2)=   "framePara(19)"
      Tab(23).Control(2).Enabled=   0   'False
      Tab(23).Control(3)=   "frameOutput(23)"
      Tab(23).Control(3).Enabled=   0   'False
      Tab(23).ControlCount=   4
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1815
         Index           =   23
         Left            =   -74880
         TabIndex        =   619
         Top             =   5280
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   23
            Left            =   1320
            TabIndex        =   625
            Top             =   300
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "RRI"
            Height          =   375
            Index           =   23
            Left            =   120
            TabIndex        =   624
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveRlf2RdgGRID 
            Caption         =   "Relief to Ridge"
            Height          =   375
            Left            =   120
            TabIndex        =   623
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtSaveRlf2RdgGRID 
            Height          =   375
            Left            =   1320
            TabIndex        =   622
            Top             =   720
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveRlf2VlyGRID 
            Caption         =   "Relief to Valley"
            Height          =   375
            Left            =   120
            TabIndex        =   621
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtSaveRlf2VlyGRID 
            Height          =   375
            Left            =   1320
            TabIndex        =   620
            Top             =   1080
            Width           =   9615
         End
         Begin VB.Label Label2 
            Caption         =   "(in elevation unit)"
            Height          =   315
            Index           =   36
            Left            =   1380
            TabIndex        =   626
            Top             =   1440
            Width           =   2355
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   795
         Index           =   19
         Left            =   -74880
         TabIndex        =   614
         Top             =   4500
         Width           =   11055
         Begin VB.ComboBox cboRRIAlg 
            Height          =   315
            Left            =   6300
            Style           =   2  'Dropdown List
            TabIndex        =   644
            Top             =   240
            Width           =   4515
         End
         Begin VB.TextBox txtRRI_RidgeTag 
            Height          =   375
            Left            =   2280
            TabIndex        =   616
            Text            =   "1"
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox txtRRI_ValleyTag 
            Height          =   375
            Left            =   4320
            TabIndex        =   615
            Text            =   "1"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Algorithm"
            Height          =   315
            Index           =   41
            Left            =   5220
            TabIndex        =   643
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Ridge Tag"
            Height          =   315
            Index           =   34
            Left            =   1080
            TabIndex        =   618
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Valley Tag"
            Height          =   315
            Index           =   33
            Left            =   3240
            TabIndex        =   617
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2235
         Index           =   23
         Left            =   -74880
         TabIndex        =   592
         Top             =   2280
         Width           =   11055
         Begin VB.CommandButton cmdRRI_OpenRidge 
            Caption         =   "Ridge"
            Height          =   375
            Left            =   120
            TabIndex        =   630
            Top             =   660
            Width           =   1215
         End
         Begin VB.TextBox txtRRI_OpenRidge 
            Height          =   375
            Left            =   1320
            TabIndex        =   629
            Top             =   660
            Width           =   9615
         End
         Begin VB.TextBox txtRRI_OpenValley 
            Height          =   375
            Left            =   1320
            TabIndex        =   628
            Top             =   1020
            Width           =   9615
         End
         Begin VB.CommandButton cmdRRI_OpenValley 
            Caption         =   "Valley"
            Height          =   375
            Left            =   120
            TabIndex        =   627
            Top             =   1020
            Width           =   1215
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   23
            Left            =   120
            TabIndex        =   607
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   23
            Left            =   1320
            TabIndex        =   606
            Top             =   300
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   23
            Left            =   120
            TabIndex        =   593
            Top             =   1440
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   23
               Left            =   10020
               TabIndex        =   599
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   23
               Left            =   8100
               TabIndex        =   598
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   23
               Left            =   6180
               TabIndex        =   597
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   23
               Left            =   4140
               TabIndex        =   596
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   23
               Left            =   2160
               TabIndex        =   595
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   23
               Left            =   600
               TabIndex        =   594
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   143
               Left            =   8880
               TabIndex        =   605
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   142
               Left            =   7320
               TabIndex        =   604
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   141
               Left            =   5280
               TabIndex        =   603
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   140
               Left            =   3240
               TabIndex        =   602
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   139
               Left            =   1620
               TabIndex        =   601
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   138
               Left            =   120
               TabIndex        =   600
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1695
         Index           =   22
         Left            =   -74880
         TabIndex        =   585
         Top             =   5340
         Width           =   11055
         Begin VB.TextBox txtSaveNegOpen 
            Height          =   375
            Left            =   1380
            TabIndex        =   590
            Top             =   900
            Width           =   9555
         End
         Begin VB.CommandButton cmdSaveNegOpen 
            Caption         =   "Negative Openness"
            Height          =   495
            Left            =   120
            TabIndex        =   589
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   22
            Left            =   1380
            TabIndex        =   587
            Top             =   420
            Width           =   9555
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Positive Openness"
            Height          =   495
            Index           =   22
            Left            =   120
            TabIndex        =   586
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "(in degree)"
            Height          =   315
            Index           =   32
            Left            =   1620
            TabIndex        =   591
            Top             =   1320
            Width           =   3855
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   18
         Left            =   -74880
         TabIndex        =   582
         Top             =   4380
         Width           =   11055
         Begin VB.TextBox txtOpenness_CirRCells 
            Height          =   375
            Left            =   6300
            TabIndex        =   583
            Text            =   "3"
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Radius (in cells) for searching Zenith and Nadir Angles"
            Height          =   435
            Index           =   31
            Left            =   720
            TabIndex        =   584
            Top             =   360
            Width           =   5535
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   22
         Left            =   -74880
         TabIndex        =   566
         Top             =   2700
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   22
            Left            =   120
            TabIndex        =   569
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   22
               Left            =   600
               TabIndex        =   575
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   22
               Left            =   2160
               TabIndex        =   574
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   22
               Left            =   4140
               TabIndex        =   573
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   22
               Left            =   6180
               TabIndex        =   572
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   22
               Left            =   8100
               TabIndex        =   571
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   22
               Left            =   10020
               TabIndex        =   570
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   137
               Left            =   120
               TabIndex        =   581
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   136
               Left            =   1620
               TabIndex        =   580
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   135
               Left            =   3240
               TabIndex        =   579
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   134
               Left            =   5280
               TabIndex        =   578
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   133
               Left            =   7320
               TabIndex        =   577
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   132
               Left            =   8880
               TabIndex        =   576
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   22
            Left            =   1200
            TabIndex        =   568
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   22
            Left            =   120
            TabIndex        =   567
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1455
         Index           =   21
         Left            =   -74880
         TabIndex        =   560
         Top             =   5460
         Width           =   11055
         Begin VB.CommandButton cmdSaveSARGRID 
            Caption         =   "Surface-Area Ratio"
            Height          =   435
            Left            =   120
            TabIndex        =   565
            Top             =   720
            Width           =   1395
         End
         Begin VB.TextBox txtSaveSARGRID 
            Height          =   375
            Left            =   1500
            TabIndex        =   564
            Top             =   780
            Width           =   9315
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   21
            Left            =   1500
            TabIndex        =   562
            Top             =   360
            Width           =   9315
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Surface Area"
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   561
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   975
         Index           =   21
         Left            =   -74880
         TabIndex        =   558
         Top             =   4440
         Visible         =   0   'False
         Width           =   11055
         Begin VB.Label Label2 
            Caption         =   "Upness above Delta. Elev."
            Height          =   375
            Index           =   30
            Left            =   240
            TabIndex        =   559
            Top             =   420
            Width           =   3015
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   21
         Left            =   -74880
         TabIndex        =   542
         Top             =   2760
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   21
            Left            =   120
            TabIndex        =   545
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   21
               Left            =   600
               TabIndex        =   551
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   21
               Left            =   2160
               TabIndex        =   550
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   21
               Left            =   4140
               TabIndex        =   549
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   21
               Left            =   6180
               TabIndex        =   548
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   21
               Left            =   8100
               TabIndex        =   547
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   21
               Left            =   10020
               TabIndex        =   546
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   131
               Left            =   120
               TabIndex        =   557
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   130
               Left            =   1620
               TabIndex        =   556
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   129
               Left            =   3240
               TabIndex        =   555
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   128
               Left            =   5280
               TabIndex        =   554
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   127
               Left            =   7320
               TabIndex        =   553
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   126
               Left            =   8880
               TabIndex        =   552
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   21
            Left            =   1200
            TabIndex        =   544
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   543
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   975
         Index           =   20
         Left            =   -74880
         TabIndex        =   528
         Top             =   6060
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   20
            Left            =   1200
            TabIndex        =   530
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "TPI"
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   529
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame optCs_WinShape 
         Caption         =   "Parameters"
         Height          =   1455
         Index           =   1
         Left            =   -74880
         TabIndex        =   522
         Top             =   4560
         Width           =   11055
         Begin VB.TextBox txtTPI_Inner_HalfWinCells 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7620
            TabIndex        =   534
            Text            =   "-1"
            Top             =   960
            Width           =   675
         End
         Begin VB.OptionButton optTPIWinShape 
            Caption         =   "Annulus (or Ring shape, the inner radius is needed.)"
            Height          =   255
            Index           =   2
            Left            =   5400
            TabIndex        =   532
            Top             =   300
            Width           =   5535
         End
         Begin VB.TextBox txtTPI_HalfWinCells 
            Height          =   315
            Left            =   7620
            TabIndex        =   525
            Text            =   "3"
            Top             =   600
            Width           =   675
         End
         Begin VB.OptionButton optTPIWinShape 
            Caption         =   "Square"
            Height          =   255
            Index           =   0
            Left            =   3060
            TabIndex        =   524
            Top             =   300
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optTPIWinShape 
            Caption         =   "Circle"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   523
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "(must be smaller than the above)"
            Height          =   435
            Index           =   26
            Left            =   8340
            TabIndex        =   535
            Top             =   960
            Width           =   2595
         End
         Begin VB.Label Label2 
            Caption         =   "Inner Radius (in cells. -1 means NO inner. 0 means that the center cell is excluded.)"
            Height          =   375
            Index           =   25
            Left            =   240
            TabIndex        =   533
            Top             =   960
            Width           =   7335
         End
         Begin VB.Label Label2 
            Caption         =   "HALF (or Radius) size of window for neighbor-searching (in cells) (e.g. 1 for 3x3 window or diameter=3)"
            Height          =   435
            Index           =   24
            Left            =   240
            TabIndex        =   527
            Top             =   600
            Width           =   7395
         End
         Begin VB.Label Label2 
            Caption         =   "Shape of neighboring window"
            Height          =   375
            Index           =   23
            Left            =   240
            TabIndex        =   526
            Top             =   300
            Width           =   2895
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   20
         Left            =   -74880
         TabIndex        =   506
         Top             =   2880
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   20
            Left            =   120
            TabIndex        =   509
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   20
               Left            =   600
               TabIndex        =   515
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   20
               Left            =   2160
               TabIndex        =   514
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   20
               Left            =   4140
               TabIndex        =   513
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   20
               Left            =   6180
               TabIndex        =   512
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   20
               Left            =   8100
               TabIndex        =   511
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   20
               Left            =   10020
               TabIndex        =   510
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   125
               Left            =   120
               TabIndex        =   521
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   124
               Left            =   1620
               TabIndex        =   520
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   123
               Left            =   3240
               TabIndex        =   519
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   122
               Left            =   5280
               TabIndex        =   518
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   121
               Left            =   7320
               TabIndex        =   517
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   120
               Left            =   8880
               TabIndex        =   516
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   20
            Left            =   1200
            TabIndex        =   508
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   507
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1035
         Index           =   19
         Left            =   -74880
         TabIndex        =   502
         Top             =   6000
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   19
            Left            =   1200
            TabIndex        =   504
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Relief"
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   503
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame optCs_WinShape 
         Caption         =   "Parameters"
         Height          =   1335
         Index           =   0
         Left            =   -74880
         TabIndex        =   496
         Top             =   4500
         Width           =   11055
         Begin VB.TextBox txtRelief_HalfWinCells 
            Height          =   375
            Left            =   8160
            TabIndex        =   499
            Text            =   "3"
            Top             =   780
            Width           =   735
         End
         Begin VB.OptionButton optReliefWinShape 
            Caption         =   "Square"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   498
            Top             =   420
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optReliefWinShape 
            Caption         =   "Circle"
            Height          =   255
            Index           =   1
            Left            =   4680
            TabIndex        =   497
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "HALF(Radius) size of window for neighbor-searching (in cells) (e.g. 1 for 3x3 window or diameter=3)"
            Height          =   375
            Index           =   22
            Left            =   240
            TabIndex        =   501
            Top             =   840
            Width           =   7875
         End
         Begin VB.Label Label2 
            Caption         =   "Shape of neighboring window"
            Height          =   375
            Index           =   21
            Left            =   240
            TabIndex        =   500
            Top             =   420
            Width           =   3135
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   19
         Left            =   -74880
         TabIndex        =   480
         Top             =   2700
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   19
            Left            =   120
            TabIndex        =   483
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   19
               Left            =   600
               TabIndex        =   489
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   19
               Left            =   2160
               TabIndex        =   488
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   19
               Left            =   4140
               TabIndex        =   487
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   19
               Left            =   6180
               TabIndex        =   486
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   19
               Left            =   8100
               TabIndex        =   485
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   19
               Left            =   10020
               TabIndex        =   484
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   119
               Left            =   120
               TabIndex        =   495
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   118
               Left            =   1620
               TabIndex        =   494
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   117
               Left            =   3240
               TabIndex        =   493
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   116
               Left            =   5280
               TabIndex        =   492
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   115
               Left            =   7320
               TabIndex        =   491
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   114
               Left            =   8880
               TabIndex        =   490
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   19
            Left            =   1200
            TabIndex        =   482
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   481
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   18
         Left            =   -74880
         TabIndex        =   476
         Top             =   5820
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   18
            Left            =   1260
            TabIndex        =   478
            Top             =   360
            Width           =   9555
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Elev-Relief Ratio"
            Height          =   495
            Index           =   18
            Left            =   120
            TabIndex        =   477
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "-1: Flat area"
            Height          =   255
            Index           =   29
            Left            =   1320
            TabIndex        =   541
            Top             =   780
            Width           =   5715
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   4
         Left            =   -74880
         TabIndex        =   473
         Top             =   4740
         Width           =   11055
         Begin VB.TextBox txtElevReliefR_CirRCells 
            Height          =   375
            Left            =   6300
            TabIndex        =   474
            Text            =   "3"
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Radius of circle window for neighbor-searching (in cells) (e.g. 1 when diameter=3)"
            Height          =   375
            Index           =   20
            Left            =   240
            TabIndex        =   475
            Top             =   360
            Width           =   5775
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   18
         Left            =   -74880
         TabIndex        =   457
         Top             =   2940
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   18
            Left            =   120
            TabIndex        =   460
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   18
               Left            =   600
               TabIndex        =   466
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   18
               Left            =   2160
               TabIndex        =   465
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   18
               Left            =   4140
               TabIndex        =   464
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   18
               Left            =   6180
               TabIndex        =   463
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   18
               Left            =   8100
               TabIndex        =   462
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   18
               Left            =   10020
               TabIndex        =   461
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   113
               Left            =   120
               TabIndex        =   472
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   112
               Left            =   1620
               TabIndex        =   471
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   111
               Left            =   3240
               TabIndex        =   470
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   110
               Left            =   5280
               TabIndex        =   469
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   109
               Left            =   7320
               TabIndex        =   468
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   108
               Left            =   8880
               TabIndex        =   467
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   18
            Left            =   1260
            TabIndex        =   459
            Top             =   360
            Width           =   9555
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   458
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   17
         Left            =   -74880
         TabIndex        =   454
         Top             =   4860
         Width           =   11055
         Begin VB.TextBox Text16 
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   455
            Text            =   "0.01"
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label2 
            Caption         =   "Delta. Elev."
            Height          =   375
            Index           =   19
            Left            =   240
            TabIndex        =   456
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2055
         Index           =   17
         Left            =   -74880
         TabIndex        =   435
         Top             =   2760
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "SCA"
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   452
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   17
            Left            =   1200
            TabIndex        =   451
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   17
            Left            =   120
            TabIndex        =   438
            Top             =   1200
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   17
               Left            =   10020
               TabIndex        =   444
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   17
               Left            =   8100
               TabIndex        =   443
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   17
               Left            =   6180
               TabIndex        =   442
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   17
               Left            =   4140
               TabIndex        =   441
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   17
               Left            =   2160
               TabIndex        =   440
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   17
               Left            =   600
               TabIndex        =   439
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   107
               Left            =   8880
               TabIndex        =   450
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   106
               Left            =   7320
               TabIndex        =   449
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   105
               Left            =   5280
               TabIndex        =   448
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   104
               Left            =   3240
               TabIndex        =   447
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   103
               Left            =   1620
               TabIndex        =   446
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   102
               Left            =   120
               TabIndex        =   445
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.TextBox txtSPI_OpenSlope 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   437
            Top             =   720
            Width           =   9615
         End
         Begin VB.CommandButton cmdSPI_OpenSlope 
            Caption         =   "Tan(Slope)"
            Height          =   375
            Left            =   120
            TabIndex        =   436
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   17
         Left            =   -74880
         TabIndex        =   432
         Top             =   5820
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "SPI"
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   434
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   17
            Left            =   1200
            TabIndex        =   433
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   16
         Left            =   -74880
         TabIndex        =   400
         Top             =   5940
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "TCI"
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   402
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   16
            Left            =   1200
            TabIndex        =   401
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   16
         Left            =   -74880
         TabIndex        =   397
         Top             =   4920
         Width           =   11055
         Begin VB.TextBox Text16 
            Height          =   375
            Index           =   0
            Left            =   1560
            TabIndex        =   398
            Text            =   "0.01"
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label2 
            Caption         =   "Delta. Elev."
            Height          =   375
            Index           =   16
            Left            =   240
            TabIndex        =   399
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2115
         Index           =   16
         Left            =   -74880
         TabIndex        =   381
         Top             =   2700
         Width           =   11055
         Begin VB.CommandButton cmdTCI_OpenCs 
            Caption         =   "Surface Curv. Index"
            Height          =   495
            Left            =   120
            TabIndex        =   421
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtTCI_OpenCs 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   420
            Top             =   780
            Width           =   9495
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "SCA"
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   396
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   16
            Left            =   1320
            TabIndex        =   395
            Top             =   360
            Width           =   9495
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   16
            Left            =   120
            TabIndex        =   382
            Top             =   1320
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   16
               Left            =   10020
               TabIndex        =   388
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   16
               Left            =   8100
               TabIndex        =   387
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   16
               Left            =   6180
               TabIndex        =   386
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   16
               Left            =   4140
               TabIndex        =   385
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   16
               Left            =   2160
               TabIndex        =   384
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   16
               Left            =   600
               TabIndex        =   383
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   101
               Left            =   8880
               TabIndex        =   394
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   100
               Left            =   7320
               TabIndex        =   393
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   99
               Left            =   5280
               TabIndex        =   392
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   98
               Left            =   3240
               TabIndex        =   391
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   97
               Left            =   1680
               TabIndex        =   390
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   96
               Left            =   120
               TabIndex        =   389
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   15
         Left            =   -74880
         TabIndex        =   377
         Top             =   5940
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   15
            Left            =   1200
            TabIndex        =   379
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "TWI"
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   378
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1035
         Index           =   15
         Left            =   -74880
         TabIndex        =   374
         Top             =   4800
         Width           =   11055
         Begin VB.TextBox Text15 
            Height          =   375
            Left            =   1560
            TabIndex        =   375
            Text            =   "0.01"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Delta. Elev."
            Height          =   375
            Index           =   15
            Left            =   240
            TabIndex        =   376
            Top             =   420
            Width           =   1335
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2055
         Index           =   15
         Left            =   -74880
         TabIndex        =   358
         Top             =   2580
         Width           =   11055
         Begin VB.CommandButton cmdTWI_OpenSlope 
            Caption         =   "Tan(Slope)"
            Height          =   375
            Left            =   120
            TabIndex        =   419
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTWI_OpenSlope 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   418
            Top             =   720
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   15
            Left            =   120
            TabIndex        =   361
            Top             =   1200
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   15
               Left            =   600
               TabIndex        =   367
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   15
               Left            =   2160
               TabIndex        =   366
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   15
               Left            =   4140
               TabIndex        =   365
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   15
               Left            =   6180
               TabIndex        =   364
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   15
               Left            =   8100
               TabIndex        =   363
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   15
               Left            =   10020
               TabIndex        =   362
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   95
               Left            =   120
               TabIndex        =   373
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   94
               Left            =   1620
               TabIndex        =   372
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   93
               Left            =   3240
               TabIndex        =   371
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   92
               Left            =   5280
               TabIndex        =   370
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   91
               Left            =   7320
               TabIndex        =   369
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   90
               Left            =   8880
               TabIndex        =   368
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   15
            Left            =   1200
            TabIndex        =   360
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "SCA"
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   359
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1755
         Index           =   14
         Left            =   -74880
         TabIndex        =   354
         Top             =   5340
         Width           =   11055
         Begin VB.TextBox txtSaveSCAGRID 
            Height          =   375
            Left            =   1560
            TabIndex        =   429
            Top             =   1020
            Width           =   9255
         End
         Begin VB.CommandButton cmdSaveSCAGRID 
            Caption         =   "SCA"
            Height          =   375
            Left            =   60
            TabIndex        =   428
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   14
            Left            =   1560
            TabIndex        =   356
            Top             =   360
            Width           =   9255
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Accumulation"
            Height          =   375
            Index           =   14
            Left            =   60
            TabIndex        =   355
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label2 
            Caption         =   "(unit: count of cells)"
            Height          =   315
            Index           =   39
            Left            =   1680
            TabIndex        =   640
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label Label2 
            Caption         =   "(unit: m^2/m)"
            Height          =   315
            Index           =   38
            Left            =   1680
            TabIndex        =   639
            Top             =   1380
            Width           =   1695
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1155
         Index           =   14
         Left            =   -74880
         TabIndex        =   352
         Top             =   4080
         Width           =   11055
         Begin VB.ComboBox cboMFD_EffectContourLen 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   641
            Top             =   720
            Width           =   4035
         End
         Begin VB.TextBox txtParaMFDP 
            Height          =   285
            Index           =   1
            Left            =   10140
            TabIndex        =   634
            Text            =   "10"
            Top             =   360
            Width           =   675
         End
         Begin VB.TextBox txtParaMFDP 
            Height          =   285
            Index           =   0
            Left            =   6540
            TabIndex        =   632
            Text            =   "1"
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboMFDAlg 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   422
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Effective contour length"
            Height          =   375
            Index           =   40
            Left            =   120
            TabIndex        =   642
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblParaMFDP 
            Alignment       =   2  'Center
            Caption         =   "(when tanb=0) <= p <= (when tanb=1)"
            Height          =   315
            Index           =   1
            Left            =   7140
            TabIndex        =   633
            Top             =   420
            Width           =   3075
         End
         Begin VB.Label lblParaMFDP 
            Alignment       =   1  'Right Justify
            Caption         =   "flow exponent:"
            Height          =   315
            Index           =   0
            Left            =   5160
            TabIndex        =   631
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Flow Direction Algorithm"
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   353
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   14
         Left            =   -74880
         TabIndex        =   336
         Top             =   2400
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   14
            Left            =   120
            TabIndex        =   339
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   14
               Left            =   600
               TabIndex        =   345
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   14
               Left            =   2160
               TabIndex        =   344
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   14
               Left            =   4140
               TabIndex        =   343
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   14
               Left            =   6180
               TabIndex        =   342
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   14
               Left            =   8100
               TabIndex        =   341
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   14
               Left            =   10020
               TabIndex        =   340
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   89
               Left            =   120
               TabIndex        =   351
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   88
               Left            =   1620
               TabIndex        =   350
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   87
               Left            =   3240
               TabIndex        =   349
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   86
               Left            =   5280
               TabIndex        =   348
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   85
               Left            =   7320
               TabIndex        =   347
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   84
               Left            =   8880
               TabIndex        =   346
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   14
            Left            =   1560
            TabIndex        =   338
            Top             =   360
            Width           =   9255
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   337
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   13
         Left            =   -74880
         TabIndex        =   332
         Top             =   5640
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   13
            Left            =   1200
            TabIndex        =   334
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Valley"
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   333
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1035
         Index           =   13
         Left            =   -74880
         TabIndex        =   329
         Top             =   4500
         Width           =   11055
         Begin VB.TextBox txtValley_UppestElev 
            Height          =   375
            Left            =   4260
            TabIndex        =   330
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Uppest elevation for valley"
            Height          =   375
            Index           =   13
            Left            =   1080
            TabIndex        =   331
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   13
         Left            =   -74880
         TabIndex        =   313
         Top             =   2700
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   13
            Left            =   120
            TabIndex        =   316
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   13
               Left            =   600
               TabIndex        =   322
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   13
               Left            =   2160
               TabIndex        =   321
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   13
               Left            =   4140
               TabIndex        =   320
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   13
               Left            =   6180
               TabIndex        =   319
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   13
               Left            =   8100
               TabIndex        =   318
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   13
               Left            =   10020
               TabIndex        =   317
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   83
               Left            =   120
               TabIndex        =   328
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   82
               Left            =   1620
               TabIndex        =   327
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   81
               Left            =   3240
               TabIndex        =   326
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   80
               Left            =   5280
               TabIndex        =   325
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   79
               Left            =   7320
               TabIndex        =   324
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   78
               Left            =   8880
               TabIndex        =   323
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   13
            Left            =   1200
            TabIndex        =   315
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   314
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   12
         Left            =   -74880
         TabIndex        =   309
         Top             =   5640
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   12
            Left            =   1200
            TabIndex        =   311
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Ridge"
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   310
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1035
         Index           =   12
         Left            =   -74880
         TabIndex        =   306
         Top             =   4500
         Width           =   11055
         Begin VB.TextBox txtRidge_LowestElev 
            Height          =   375
            Left            =   4260
            TabIndex        =   307
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Lowest elevation for ridge"
            Height          =   375
            Index           =   12
            Left            =   1080
            TabIndex        =   308
            Top             =   360
            Width           =   3075
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   12
         Left            =   -74880
         TabIndex        =   290
         Top             =   2700
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   12
            Left            =   120
            TabIndex        =   293
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   12
               Left            =   600
               TabIndex        =   299
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   12
               Left            =   2160
               TabIndex        =   298
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   12
               Left            =   4140
               TabIndex        =   297
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   12
               Left            =   6180
               TabIndex        =   296
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   12
               Left            =   8100
               TabIndex        =   295
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   12
               Left            =   10020
               TabIndex        =   294
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   77
               Left            =   120
               TabIndex        =   305
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   76
               Left            =   1560
               TabIndex        =   304
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   75
               Left            =   3240
               TabIndex        =   303
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   74
               Left            =   5280
               TabIndex        =   302
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   73
               Left            =   7320
               TabIndex        =   301
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   72
               Left            =   8880
               TabIndex        =   300
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   12
            Left            =   1200
            TabIndex        =   292
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   291
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2055
         Index           =   11
         Left            =   -74880
         TabIndex        =   273
         Top             =   2580
         Width           =   11055
         Begin VB.CommandButton cmdDslpI_OpenD8 
            Caption         =   "ArcInfo FlowDir"
            Height          =   375
            Left            =   120
            TabIndex        =   417
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtDslpI_OpenD8 
            Height          =   375
            Left            =   1320
            TabIndex        =   416
            Top             =   720
            Width           =   9495
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   288
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   11
            Left            =   1320
            TabIndex        =   287
            Top             =   360
            Width           =   9495
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   11
            Left            =   120
            TabIndex        =   274
            Top             =   1200
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   11
               Left            =   10020
               TabIndex        =   280
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   11
               Left            =   8100
               TabIndex        =   279
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   11
               Left            =   6180
               TabIndex        =   278
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   11
               Left            =   4140
               TabIndex        =   277
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   11
               Left            =   2160
               TabIndex        =   276
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   11
               Left            =   600
               TabIndex        =   275
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   71
               Left            =   8880
               TabIndex        =   286
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   70
               Left            =   7320
               TabIndex        =   285
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   69
               Left            =   5280
               TabIndex        =   284
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   68
               Left            =   3240
               TabIndex        =   283
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   67
               Left            =   1620
               TabIndex        =   282
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   66
               Left            =   120
               TabIndex        =   281
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1035
         Index           =   11
         Left            =   -74880
         TabIndex        =   270
         Top             =   4800
         Width           =   11055
         Begin VB.TextBox txtDslpI_DeltaElev 
            Height          =   375
            Left            =   4260
            TabIndex        =   271
            Text            =   "2"
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Delta downslope elevation (m)"
            Height          =   435
            Index           =   11
            Left            =   840
            TabIndex        =   272
            Top             =   360
            Width           =   3195
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   11
         Left            =   -74880
         TabIndex        =   267
         Top             =   5940
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "DownslopeI"
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   269
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   11
            Left            =   1200
            TabIndex        =   268
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   2175
         Index           =   10
         Left            =   -74880
         TabIndex        =   250
         Top             =   2400
         Width           =   11055
         Begin VB.TextBox txtRPI_OpenRidge 
            Height          =   375
            Left            =   1320
            TabIndex        =   636
            Top             =   660
            Width           =   9615
         End
         Begin VB.CommandButton cmdRPI_OpenRidge 
            Caption         =   "Ridge"
            Height          =   375
            Left            =   120
            TabIndex        =   635
            Top             =   660
            Width           =   1215
         End
         Begin VB.CommandButton cmdRPI_OpenValley 
            Caption         =   "Valley"
            Height          =   375
            Left            =   120
            TabIndex        =   413
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtRPI_OpenValley 
            Height          =   375
            Left            =   1320
            TabIndex        =   412
            Top             =   1020
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   265
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   10
            Left            =   1320
            TabIndex        =   264
            Top             =   300
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   10
            Left            =   120
            TabIndex        =   251
            Top             =   1440
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   10
               Left            =   10020
               TabIndex        =   257
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   10
               Left            =   8100
               TabIndex        =   256
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   10
               Left            =   6180
               TabIndex        =   255
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   10
               Left            =   4140
               TabIndex        =   254
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   10
               Left            =   2160
               TabIndex        =   253
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   10
               Left            =   600
               TabIndex        =   252
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   65
               Left            =   8880
               TabIndex        =   263
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   64
               Left            =   7320
               TabIndex        =   262
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   63
               Left            =   5280
               TabIndex        =   261
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   62
               Left            =   3240
               TabIndex        =   260
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   61
               Left            =   1620
               TabIndex        =   259
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   60
               Left            =   120
               TabIndex        =   258
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   735
         Index           =   10
         Left            =   -74880
         TabIndex        =   247
         Top             =   4560
         Width           =   11055
         Begin VB.ComboBox cboRPIAlg 
            Height          =   315
            Left            =   6300
            Style           =   2  'Dropdown List
            TabIndex        =   637
            Top             =   240
            Width           =   4515
         End
         Begin VB.TextBox txtRPI_ValleyTag 
            Height          =   375
            Left            =   4320
            TabIndex        =   414
            Text            =   "1"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtRPI_RidgeTag 
            Height          =   375
            Left            =   2280
            TabIndex        =   248
            Text            =   "1"
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Algorithm"
            Height          =   375
            Index           =   37
            Left            =   5340
            TabIndex        =   638
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Valley Tag"
            Height          =   315
            Index           =   7
            Left            =   3240
            TabIndex        =   415
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Ridge Tag"
            Height          =   315
            Index           =   10
            Left            =   1080
            TabIndex        =   249
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1815
         Index           =   10
         Left            =   -74880
         TabIndex        =   244
         Top             =   5340
         Width           =   11055
         Begin VB.TextBox txtSaveDist2VlyGRID 
            Height          =   375
            Left            =   1320
            TabIndex        =   611
            Top             =   1080
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveDist2VlyGRID 
            Caption         =   "Dist to Valley"
            Height          =   375
            Left            =   120
            TabIndex        =   610
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtSaveDist2RdgGRID 
            Height          =   375
            Left            =   1320
            TabIndex        =   609
            Top             =   720
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveDist2RdgGRID 
            Caption         =   "Dist to Ridge"
            Height          =   375
            Left            =   120
            TabIndex        =   608
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "RPI"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   246
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   10
            Left            =   1320
            TabIndex        =   245
            Top             =   300
            Width           =   9615
         End
         Begin VB.Label Label2 
            Caption         =   "(in distance unit)"
            Height          =   315
            Index           =   35
            Left            =   1380
            TabIndex        =   612
            Top             =   1440
            Width           =   1695
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   9
         Left            =   120
         TabIndex        =   227
         Top             =   2700
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   242
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   9
            Left            =   1200
            TabIndex        =   241
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   9
            Left            =   120
            TabIndex        =   228
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   9
               Left            =   10020
               TabIndex        =   234
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   9
               Left            =   8100
               TabIndex        =   233
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   9
               Left            =   6180
               TabIndex        =   232
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   9
               Left            =   4140
               TabIndex        =   231
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   9
               Left            =   2160
               TabIndex        =   230
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   9
               Left            =   600
               TabIndex        =   229
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   59
               Left            =   8880
               TabIndex        =   240
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   58
               Left            =   7320
               TabIndex        =   239
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   57
               Left            =   5280
               TabIndex        =   238
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   56
               Left            =   3240
               TabIndex        =   237
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   55
               Left            =   1620
               TabIndex        =   236
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   54
               Left            =   120
               TabIndex        =   235
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   975
         Index           =   9
         Left            =   120
         TabIndex        =   225
         Top             =   4500
         Width           =   11055
         Begin VB.Label Label2 
            Caption         =   "Upness above Delta. Elev."
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   226
            Top             =   420
            Width           =   2295
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   9
         Left            =   120
         TabIndex        =   222
         Top             =   5580
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "UPNESS"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   224
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   9
            Left            =   1200
            TabIndex        =   223
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   8
         Left            =   -74880
         TabIndex        =   205
         Top             =   2700
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   220
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   8
            Left            =   1200
            TabIndex        =   219
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   8
            Left            =   120
            TabIndex        =   206
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   8
               Left            =   10020
               TabIndex        =   212
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   8
               Left            =   8100
               TabIndex        =   211
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   8
               Left            =   6180
               TabIndex        =   210
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   8
               Left            =   4140
               TabIndex        =   209
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   8
               Left            =   2160
               TabIndex        =   208
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   8
               Left            =   600
               TabIndex        =   207
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   53
               Left            =   8880
               TabIndex        =   218
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   52
               Left            =   7320
               TabIndex        =   217
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   51
               Left            =   5280
               TabIndex        =   216
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   50
               Left            =   3240
               TabIndex        =   215
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   49
               Left            =   1620
               TabIndex        =   214
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   48
               Left            =   120
               TabIndex        =   213
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1035
         Index           =   8
         Left            =   -74880
         TabIndex        =   202
         Top             =   4440
         Width           =   11055
         Begin VB.TextBox txtLPos_CirRCells 
            Height          =   375
            Left            =   6300
            TabIndex        =   203
            Text            =   "3"
            Top             =   420
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Radius of circle window for neighbor-searching (in cells) (e.g. 1 when diameter=3)"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   204
            Top             =   420
            Width           =   6255
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   8
         Left            =   -74880
         TabIndex        =   199
         Top             =   5580
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "LPos"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   201
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   8
            Left            =   1200
            TabIndex        =   200
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   7
         Left            =   -74880
         TabIndex        =   182
         Top             =   2760
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   197
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   1200
            TabIndex        =   196
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   7
            Left            =   120
            TabIndex        =   183
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   7
               Left            =   10020
               TabIndex        =   189
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   7
               Left            =   8100
               TabIndex        =   188
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   7
               Left            =   6180
               TabIndex        =   187
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   7
               Left            =   4140
               TabIndex        =   186
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   7
               Left            =   2160
               TabIndex        =   185
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   7
               Left            =   600
               TabIndex        =   184
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   47
               Left            =   8880
               TabIndex        =   195
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   46
               Left            =   7320
               TabIndex        =   194
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   45
               Left            =   5280
               TabIndex        =   193
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   44
               Left            =   3240
               TabIndex        =   192
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   43
               Left            =   1620
               TabIndex        =   191
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   42
               Left            =   120
               TabIndex        =   190
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1395
         Index           =   7
         Left            =   -74880
         TabIndex        =   181
         Top             =   4440
         Width           =   11055
         Begin VB.OptionButton optTRIWinShape 
            Caption         =   "Circle"
            Height          =   255
            Index           =   1
            Left            =   4500
            TabIndex        =   538
            Top             =   360
            Width           =   1155
         End
         Begin VB.OptionButton optTRIWinShape 
            Caption         =   "Square"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   537
            Top             =   360
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.TextBox txtTRI_HalfWinCells 
            Height          =   375
            Left            =   8220
            TabIndex        =   536
            Text            =   "1"
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label2 
            Caption         =   "Shape of neighboring window"
            Height          =   375
            Index           =   28
            Left            =   240
            TabIndex        =   540
            Top             =   360
            Width           =   2955
         End
         Begin VB.Label Label2 
            Caption         =   "HALF(Radius) size of window for neighbor-searching (in cells) (e.g. 1 for 3x3 window or diameter=3)"
            Height          =   375
            Index           =   27
            Left            =   240
            TabIndex        =   539
            Top             =   780
            Width           =   7935
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   7
         Left            =   -74880
         TabIndex        =   178
         Top             =   5880
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "TRI"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   180
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   7
            Left            =   1200
            TabIndex        =   179
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   6
         Left            =   -74880
         TabIndex        =   161
         Top             =   2460
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   176
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   1200
            TabIndex        =   175
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   6
            Left            =   120
            TabIndex        =   162
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   6
               Left            =   10020
               TabIndex        =   168
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   6
               Left            =   8100
               TabIndex        =   167
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   6
               Left            =   6180
               TabIndex        =   166
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   6
               Left            =   4140
               TabIndex        =   165
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   6
               Left            =   2160
               TabIndex        =   164
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   6
               Left            =   600
               TabIndex        =   163
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   41
               Left            =   8880
               TabIndex        =   174
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   40
               Left            =   7320
               TabIndex        =   173
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   39
               Left            =   5280
               TabIndex        =   172
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   38
               Left            =   3240
               TabIndex        =   171
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   37
               Left            =   1620
               TabIndex        =   170
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   36
               Left            =   120
               TabIndex        =   169
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   975
         Index           =   6
         Left            =   -74880
         TabIndex        =   158
         Top             =   4200
         Width           =   11055
         Begin VB.TextBox txtTOPHAT_HalfWinCells 
            Height          =   375
            Left            =   10080
            TabIndex        =   404
            Text            =   "7"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtTOPHAT_thresh 
            Height          =   375
            Left            =   2280
            TabIndex        =   159
            Text            =   "0.05"
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "HALF length of rectangle window for neighbor-searching (in cells) (e.g. 1 for 3x3 window)"
            Height          =   375
            Index           =   17
            Left            =   3420
            TabIndex        =   405
            Top             =   360
            Width           =   6615
         End
         Begin VB.Label Label2 
            Caption         =   "Elev. threshold"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   160
            Top             =   360
            Width           =   1755
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1635
         Index           =   6
         Left            =   -74880
         TabIndex        =   155
         Top             =   5280
         Width           =   11055
         Begin VB.TextBox txtTOPHAT_SaveValleyI 
            Height          =   375
            Left            =   1680
            TabIndex        =   409
            Top             =   1080
            Width           =   9135
         End
         Begin VB.CommandButton cmdTOPHAT_ValleyI 
            Caption         =   "Valley Index"
            Height          =   375
            Left            =   120
            TabIndex        =   408
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton cmdTOPHAT_HillslpI 
            Caption         =   "Hillslope Index"
            Height          =   375
            Left            =   120
            TabIndex        =   407
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtTOPHAT_SaveHillslpI 
            Height          =   375
            Left            =   1680
            TabIndex        =   406
            Top             =   720
            Width           =   9135
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Hill Index"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   157
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   6
            Left            =   1680
            TabIndex        =   156
            Top             =   360
            Width           =   9135
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   5
         Left            =   -74880
         TabIndex        =   138
         Top             =   2820
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   153
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   1200
            TabIndex        =   152
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   5
            Left            =   120
            TabIndex        =   139
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   5
               Left            =   10020
               TabIndex        =   145
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   5
               Left            =   8100
               TabIndex        =   144
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   5
               Left            =   6180
               TabIndex        =   143
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   5
               Left            =   4140
               TabIndex        =   142
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   5
               Left            =   2160
               TabIndex        =   141
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   5
               Left            =   600
               TabIndex        =   140
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   35
               Left            =   8880
               TabIndex        =   151
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   34
               Left            =   7320
               TabIndex        =   150
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   33
               Left            =   5280
               TabIndex        =   149
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   32
               Left            =   3240
               TabIndex        =   148
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   31
               Left            =   1620
               TabIndex        =   147
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   30
               Left            =   120
               TabIndex        =   146
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   915
         Index           =   5
         Left            =   -74880
         TabIndex        =   135
         Top             =   4560
         Width           =   11055
         Begin VB.TextBox txtElevPctl_CirRCells 
            Height          =   375
            Left            =   6300
            TabIndex        =   136
            Text            =   "3"
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Radius of circle window for neighbor-searching (in cells) (e.g. 1 when diameter=3)"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   137
            Top             =   360
            Width           =   6075
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   5
         Left            =   -74880
         TabIndex        =   132
         Top             =   5580
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "ElevPerc"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   134
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   5
            Left            =   1200
            TabIndex        =   133
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   4
         Left            =   -74880
         TabIndex        =   115
         Top             =   2880
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   130
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1200
            TabIndex        =   129
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   4
            Left            =   120
            TabIndex        =   116
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   4
               Left            =   10020
               TabIndex        =   122
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   4
               Left            =   8100
               TabIndex        =   121
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   4
               Left            =   6180
               TabIndex        =   120
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   4
               Left            =   4140
               TabIndex        =   119
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   4
               Left            =   2160
               TabIndex        =   118
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   4
               Left            =   600
               TabIndex        =   117
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   29
               Left            =   8880
               TabIndex        =   128
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   28
               Left            =   7320
               TabIndex        =   127
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   27
               Left            =   5280
               TabIndex        =   126
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   26
               Left            =   3240
               TabIndex        =   125
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   25
               Left            =   1620
               TabIndex        =   124
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   24
               Left            =   120
               TabIndex        =   123
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame optCs_WinShape 
         Caption         =   "Parameters"
         Height          =   1335
         Index           =   4
         Left            =   -74880
         TabIndex        =   112
         Top             =   4620
         Width           =   11055
         Begin VB.OptionButton optCsWinShape 
            Caption         =   "Circle"
            Height          =   255
            Index           =   1
            Left            =   4500
            TabIndex        =   426
            Top             =   420
            Width           =   1155
         End
         Begin VB.OptionButton optCsWinShape 
            Caption         =   "Square"
            Height          =   255
            Index           =   0
            Left            =   3300
            TabIndex        =   425
            Top             =   420
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.TextBox txtCs_HalfWinCells 
            Height          =   375
            Left            =   8160
            TabIndex        =   113
            Text            =   "3"
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Shape of neighboring window"
            Height          =   375
            Index           =   18
            Left            =   240
            TabIndex        =   427
            Top             =   420
            Width           =   3075
         End
         Begin VB.Label Label2 
            Caption         =   "HALF (or Radius) size of window for neighbor-searching (in cells) (e.g. 1 for 3x3 window or diameter=3)"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   114
            Top             =   780
            Width           =   7815
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   975
         Index           =   4
         Left            =   -74880
         TabIndex        =   109
         Top             =   6060
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Cs"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   111
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   4
            Left            =   1200
            TabIndex        =   110
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1575
         Index           =   3
         Left            =   -74880
         TabIndex        =   86
         Top             =   2100
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   100
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   3
            Left            =   120
            TabIndex        =   87
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   3
               Left            =   10020
               TabIndex        =   93
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   3
               Left            =   8100
               TabIndex        =   92
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   3
               Left            =   6180
               TabIndex        =   91
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   3
               Left            =   4140
               TabIndex        =   90
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   3
               Left            =   2160
               TabIndex        =   89
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   3
               Left            =   600
               TabIndex        =   88
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   23
               Left            =   8880
               TabIndex        =   99
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   22
               Left            =   7320
               TabIndex        =   98
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   21
               Left            =   5280
               TabIndex        =   97
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   20
               Left            =   3240
               TabIndex        =   96
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   19
               Left            =   1620
               TabIndex        =   95
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   18
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   615
         Index           =   3
         Left            =   -74880
         TabIndex        =   83
         Top             =   3660
         Width           =   11055
         Begin VB.ComboBox cboCurvatureAlg 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label2 
            Caption         =   "Curvature algorithm:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   2895
         Index           =   3
         Left            =   -74880
         TabIndex        =   76
         Top             =   4260
         Width           =   11055
         Begin VB.TextBox txtSaveMaxCurv 
            Height          =   375
            Left            =   1200
            TabIndex        =   411
            Top             =   2400
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveMaxCurv 
            Caption         =   "Max Curv."
            Height          =   375
            Left            =   120
            TabIndex        =   410
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdSaveMinCurv 
            Caption         =   "Min Curv."
            Height          =   375
            Left            =   120
            TabIndex        =   108
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txtSaveMinCurv 
            Height          =   375
            Left            =   1200
            TabIndex        =   107
            Top             =   2040
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveUnspher 
            Caption         =   "Unsphericity"
            Height          =   375
            Left            =   120
            TabIndex        =   106
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtSaveUnspher 
            Height          =   375
            Left            =   1200
            TabIndex        =   105
            Top             =   1680
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveMeanCurv 
            Caption         =   "Mean Curv."
            Height          =   375
            Left            =   120
            TabIndex        =   104
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtSaveMeanCurv 
            Height          =   375
            Left            =   1200
            TabIndex        =   103
            Top             =   1320
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Profile"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   3
            Left            =   1200
            TabIndex        =   81
            Top             =   240
            Width           =   9615
         End
         Begin VB.TextBox txtSavePlanCurv 
            Height          =   375
            Left            =   1200
            TabIndex        =   80
            Top             =   600
            Width           =   9615
         End
         Begin VB.CommandButton cmdSavePlanCurv 
            Caption         =   "Plan"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdSaveHorizCurv 
            Caption         =   "Horizontal"
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtSaveHorizCurv 
            Height          =   375
            Left            =   1200
            TabIndex        =   77
            Top             =   960
            Width           =   9615
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   2
         Left            =   -74880
         TabIndex        =   55
         Top             =   2460
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   2
               Left            =   600
               TabIndex        =   64
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   2
               Left            =   2160
               TabIndex        =   63
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   2
               Left            =   4140
               TabIndex        =   62
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   2
               Left            =   6180
               TabIndex        =   61
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   2
               Left            =   8100
               TabIndex        =   60
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   2
               Left            =   10020
               TabIndex        =   59
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   17
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   16
               Left            =   1620
               TabIndex        =   69
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   15
               Left            =   3240
               TabIndex        =   68
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   14
               Left            =   5280
               TabIndex        =   67
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   13
               Left            =   7320
               TabIndex        =   66
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   12
               Left            =   8880
               TabIndex        =   65
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   57
            Top             =   360
            Width           =   9495
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   855
         Index           =   2
         Left            =   -74880
         TabIndex        =   52
         Top             =   4140
         Width           =   11055
         Begin VB.ComboBox cboAspect 
            Height          =   315
            Left            =   1320
            TabIndex        =   53
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label Label2 
            Caption         =   "Aspect in"
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   54
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1995
         Index           =   2
         Left            =   -74880
         TabIndex        =   49
         Top             =   5100
         Width           =   11055
         Begin VB.TextBox txtSaveArcInfoAspect 
            Height          =   375
            Left            =   1320
            TabIndex        =   431
            Top             =   720
            Width           =   9495
         End
         Begin VB.CommandButton cmdSaveArcInfoAspect 
            Caption         =   "ArcInfo Aspect"
            Height          =   375
            Left            =   60
            TabIndex        =   430
            Top             =   720
            Width           =   1275
         End
         Begin VB.TextBox txtSaveCosAspect 
            Height          =   375
            Left            =   1320
            TabIndex        =   75
            Top             =   1440
            Width           =   9495
         End
         Begin VB.CommandButton cmdSaveCosAspect 
            Caption         =   "Cos(Aspect)"
            Height          =   375
            Left            =   60
            TabIndex        =   74
            Top             =   1440
            Width           =   1275
         End
         Begin VB.CommandButton cmdSaveSinAspect 
            Caption         =   "Sin(Aspect)"
            Height          =   375
            Left            =   60
            TabIndex        =   73
            Top             =   1080
            Width           =   1275
         End
         Begin VB.TextBox txtSaveSinAspect 
            Height          =   375
            Left            =   1320
            TabIndex        =   72
            Top             =   1080
            Width           =   9495
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   51
            Top             =   360
            Width           =   9495
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Aspect (degree)"
            Height          =   375
            Index           =   2
            Left            =   60
            TabIndex        =   50
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   1
         Left            =   -74880
         TabIndex        =   45
         Top             =   5700
         Width           =   11055
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "Output"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   46
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1335
         Index           =   1
         Left            =   -74880
         TabIndex        =   43
         Top             =   4200
         Width           =   11055
         Begin VB.OptionButton optSlopeFormat 
            Caption         =   "in degree"
            Height          =   315
            Index           =   1
            Left            =   3240
            TabIndex        =   424
            Top             =   840
            Width           =   1395
         End
         Begin VB.OptionButton optSlopeFormat 
            Caption         =   "in Tan(.)"
            Height          =   315
            Index           =   0
            Left            =   1260
            TabIndex        =   423
            Top             =   840
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.ComboBox cboSlopeType 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label Label2 
            Caption         =   "Slope type:"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame frameOutput 
         Caption         =   "Output GRID"
         Height          =   1095
         Index           =   0
         Left            =   -74880
         TabIndex        =   40
         Top             =   5580
         Width           =   11055
         Begin VB.TextBox txtSaveGRID1 
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   42
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSaveGRID1 
            Caption         =   "New DEM"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame framePara 
         Caption         =   "Parameters"
         Height          =   1035
         Index           =   0
         Left            =   -74880
         TabIndex        =   37
         Top             =   4320
         Width           =   11055
         Begin VB.TextBox txtFilDep_DeltaElev 
            Height          =   375
            Left            =   1560
            TabIndex        =   39
            Text            =   "0.01"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Delta. Elev."
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   1
         Left            =   -74880
         TabIndex        =   20
         Top             =   2460
         Width           =   11055
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   34
            Top             =   360
            Width           =   9615
         End
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   1
               Left            =   10020
               TabIndex        =   27
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtCellSize 
               Height          =   315
               Index           =   1
               Left            =   8100
               TabIndex        =   26
               Text            =   "1"
               Top             =   240
               Width           =   675
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   1
               Left            =   6180
               TabIndex        =   25
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   1
               Left            =   4140
               TabIndex        =   24
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   1
               Left            =   2160
               TabIndex        =   23
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   1
               Left            =   600
               TabIndex        =   22
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   11
               Left            =   8880
               TabIndex        =   33
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CellSize"
               Height          =   315
               Index           =   10
               Left            =   7320
               TabIndex        =   32
               Top             =   240
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   9
               Left            =   5280
               TabIndex        =   31
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   8
               Left            =   3240
               TabIndex        =   30
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   7
               Left            =   1620
               TabIndex        =   29
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   6
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   555
            End
         End
      End
      Begin VB.Frame frameInput 
         Caption         =   "Input GRID"
         Height          =   1635
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   2460
         Width           =   11055
         Begin VB.Frame frameFileHead 
            Caption         =   "File Head"
            Height          =   675
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   10815
            Begin VB.TextBox txtCols 
               Height          =   315
               Index           =   0
               Left            =   600
               TabIndex        =   13
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtRows 
               Height          =   315
               Index           =   0
               Left            =   2160
               TabIndex        =   12
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtXll 
               Height          =   315
               Index           =   0
               Left            =   4140
               TabIndex        =   11
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin VB.TextBox txtYll 
               Height          =   315
               Index           =   0
               Left            =   6180
               TabIndex        =   10
               Text            =   "0"
               Top             =   240
               Width           =   1035
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
            Begin VB.TextBox txtNoData 
               Height          =   315
               Index           =   0
               Left            =   10020
               TabIndex        =   8
               Text            =   "-9999"
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "nCols"
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "nRows"
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   18
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "XllCorner"
               Height          =   315
               Index           =   2
               Left            =   3240
               TabIndex        =   17
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "YllCorner"
               Height          =   315
               Index           =   3
               Left            =   5280
               TabIndex        =   16
               Top             =   240
               Width           =   975
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
               Caption         =   "NoData_Value"
               Height          =   315
               Index           =   5
               Left            =   8880
               TabIndex        =   14
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtSrcGRID 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton cmdSrcGRID 
            Caption         =   "DEM"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":02A0
         Height          =   435
         Index           =   23
         Left            =   -74820
         TabIndex        =   613
         Top             =   1980
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0335
         Height          =   615
         Index           =   22
         Left            =   -74760
         TabIndex        =   588
         Top             =   2100
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Surface Area (Jenness, 2004);   Surface-area ratio = Surface area / grid area"
         Height          =   495
         Index           =   21
         Left            =   -74760
         TabIndex        =   563
         Top             =   2040
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0424
         Height          =   975
         Index           =   20
         Left            =   -74760
         TabIndex        =   531
         Top             =   1920
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Relief (or roughness): Max(elev)-Min(elev)"
         Height          =   555
         Index           =   19
         Left            =   -74640
         TabIndex        =   505
         Top             =   2100
         Width           =   10815
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0603
         Height          =   855
         Index           =   18
         Left            =   -74760
         TabIndex        =   479
         Top             =   2160
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":06F5
         Height          =   555
         Index           =   17
         Left            =   -74760
         TabIndex        =   453
         Top             =   2040
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0791
         Height          =   675
         Index           =   16
         Left            =   -74760
         TabIndex        =   403
         Top             =   1980
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0878
         Height          =   435
         Index           =   15
         Left            =   -74760
         TabIndex        =   380
         Top             =   2040
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0917
         Height          =   495
         Index           =   14
         Left            =   -74760
         TabIndex        =   357
         Top             =   1920
         Width           =   10875
      End
      Begin VB.Label lblInfo 
         Caption         =   "Extract Drainage Networks (Peucker and Douglas, 1995): Mark as ""1"""
         Height          =   375
         Index           =   13
         Left            =   -74760
         TabIndex        =   335
         Top             =   2160
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Extract Ridge (Peucker and Douglas, 1995): Mark as ""1"""
         Height          =   375
         Index           =   12
         Left            =   -74760
         TabIndex        =   312
         Top             =   2160
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Downslope Index (Hjerdt,2004)"
         Height          =   375
         Index           =   11
         Left            =   -74760
         TabIndex        =   289
         Top             =   2100
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":09DE
         Height          =   615
         Index           =   10
         Left            =   -74760
         TabIndex        =   266
         Top             =   1920
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0A8F
         Height          =   615
         Index           =   9
         Left            =   240
         TabIndex        =   243
         Top             =   1980
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0B7C
         Height          =   675
         Index           =   8
         Left            =   -74760
         TabIndex        =   221
         Top             =   1980
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0C85
         Height          =   735
         Index           =   7
         Left            =   -74760
         TabIndex        =   198
         Top             =   1980
         Width           =   10875
      End
      Begin VB.Label lblInfo 
         Caption         =   "Hill-Hillslope-Valley Index (TOPHAT)(Schmidt and Hewitt, 2004)"
         Height          =   375
         Index           =   6
         Left            =   -74760
         TabIndex        =   177
         Top             =   2100
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0DA5
         Height          =   675
         Index           =   5
         Left            =   -74760
         TabIndex        =   154
         Top             =   2100
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmDTAFunc.frx":0E59
         Height          =   915
         Index           =   4
         Left            =   -74760
         TabIndex        =   131
         Top             =   1980
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Profile, horizontal, plan... curvatures by Shary et al. (2002), also ref. Young & Evans (1978); Pennock et al. (1987)"
         Height          =   375
         Index           =   3
         Left            =   -74760
         TabIndex        =   102
         Top             =   1860
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Aspect: Based on algorithm in ArcInfo. Aspect is expressed in positive degrees from 0 to 360, measured clockwise from the north."
         Height          =   435
         Index           =   2
         Left            =   -74760
         TabIndex        =   71
         Top             =   1980
         Width           =   10875
      End
      Begin VB.Label lblInfo 
         Caption         =   "1) Slope: ArcInfo; 2) Max. downslope; 3) local downslope=Sum[tan(downslope_i)*Li]/Sum(Li)  (Quinn et al., 1991)"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   36
         Top             =   2100
         Width           =   10935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Removel depressions and flat areas (Planchon and Darbox, 2001)"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   6
         Top             =   2100
         Width           =   10935
      End
   End
End
Attribute VB_Name = "frmDTAFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const FUNC_TYPE_SLOPE_DEGREE = "Slope in ArcInfo (in degree)"
'Const FUNC_TYPE_SLOPE = "Slope in ArcInfo"
'Const FUNC_TYPE_MAXDOWNSLOPE = "Max Downslope"
'Const FUNC_TYPE_MFD_QUINN91 = "MFD (Quinn et al., 1991)"
'Const FUNC_TYPE_MFD_QIN07 = "MFD-md (Qin et al., 2007)"

Dim m_bRunning As Boolean
Dim m_strBasePath As String
Dim m_strFilePre As String
Dim m_pBaseDEM As clsGrid

Public Sub DTAFunc(iFuncIndex As Integer)
   Dim i As Integer
   With SSTabFunc
      For i = 0 To .Tabs - 1
         .TabEnabled(i) = False
      Next
      .Tab = iFuncIndex
      .TabEnabled(iFuncIndex) = True
   End With
End Sub


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

Private Sub cmdDslpI_OpenD8_Click()
'   Dim str As String
'   comdlg.DialogTitle = "Open Src GRID"
'   comdlg.FileName = ""
'   str = GetFileName(comdlg, True, , ".asc")
'   If str <> "" Then txtDslpI_OpenD8.Text = str
   txtDslpI_OpenD8.Text = GetSaveFileName(txtDslpI_OpenD8.Text)
End Sub

Private Sub cmdRPI_OpenRidge_Click()
   txtRPI_OpenRidge.Text = GetSaveFileName(txtRPI_OpenRidge.Text)
End Sub

Private Sub cmdRPI_OpenValley_Click()
   txtRPI_OpenValley.Text = GetSaveFileName(txtRPI_OpenValley.Text)
End Sub

Private Sub cmdRRI_OpenRidge_Click()
   txtRRI_OpenRidge.Text = GetSaveFileName(txtRRI_OpenRidge.Text)
End Sub

Private Sub cmdRRI_OpenValley_Click()
   txtRRI_OpenValley.Text = GetSaveFileName(txtRRI_OpenValley.Text)
End Sub

Private Sub cmdRun_Click()
On Error GoTo ErrH
   Dim iFuncIndex As Integer
   Dim dPara1 As Double, iPara1 As Integer, iPara2 As Integer
   Dim strSaveGRID As String
   Dim sSrcTanb As String, pSrcGRIDTanb As clsGrid
   Dim pGrid As clsGrid
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim boolResult As Boolean
   
   If m_bRunning Then Exit Sub
   m_bRunning = True
   frameInput(iFuncIndex).Enabled = Not m_bRunning
   framePara(iFuncIndex).Enabled = Not m_bRunning
   frameOutput(iFuncIndex).Enabled = Not m_bRunning
   
   Me.MousePointer = 11
      
   iFuncIndex = SSTabFunc.Tab
   If txtSrcGRID(iFuncIndex).Text = "" Then
      Err.Raise Number:=vbObjectError + 513, Description:="Assign the source GRID firstly"
   End If
   
   iCols = CInt(txtCols(iFuncIndex).Text):      iRows = CInt(txtRows(iFuncIndex).Text)
   dXll = CDbl(txtXll(iFuncIndex).Text):       dYll = CDbl(txtYll(iFuncIndex).Text)
   dCellSize = CDbl(txtCellSize(iFuncIndex).Text):  dNoData = CDbl(txtNoData(iFuncIndex).Text)
   
   strSaveGRID = Trim(txtSaveGRID1(iFuncIndex).Text)
   
   Select Case iFuncIndex
   Case FUNC_FILLDEP
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtFilDep_DeltaElev
         If IsNumeric(.Text) Then
            dPara1 = CDbl(.Text)
            If dPara1 < 0 Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      If modMFD.FillDep_RemoveExcessWater_Planchon01(m_pBaseDEM, pGrid, dPara1) Then
         If pGrid.SaveAscGrid(strSaveGRID, , 4) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_SLOPE
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      Select Case cboSlopeType
      Case FUNC_TYPE_SLOPE
         If optSlopeFormat(0).Value Then
            boolResult = modDTA1.Slope_ArcInfo(m_pBaseDEM, pGrid, False)
         Else
            boolResult = modDTA1.Slope_ArcInfo(m_pBaseDEM, pGrid, True)
         End If
      Case FUNC_TYPE_MAXDOWNSLOPE
         If optSlopeFormat(0).Value Then
            boolResult = modDTA1.MaximumDownslope(m_pBaseDEM, pGrid, False)
         Else
            boolResult = modDTA1.MaximumDownslope(m_pBaseDEM, pGrid, True)
         End If
      Case FUNC_TYPE_LOCALDOWNSLOPE
         If optSlopeFormat(0).Value Then
            boolResult = modDTA1.LocalDownslope(m_pBaseDEM, pGrid, False)
         Else
            boolResult = modDTA1.LocalDownslope(m_pBaseDEM, pGrid, True)
         End If
      Case Else
         boolResult = False
      End Select
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 4) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_ASPECT
      Dim sSaveESRIAspect As String, sSaveSinAspect As String, sSaveCosAspect As String
      Dim pSaveESRIAspect As clsGrid, pSaveSinAspect As clsGrid, pSaveCosAspect As clsGrid
            
      sSaveESRIAspect = Trim(txtSaveArcInfoAspect.Text)
      sSaveSinAspect = Trim(txtSaveSinAspect.Text)
      sSaveCosAspect = Trim(txtSaveCosAspect.Text)
      
      If strSaveGRID = "" And sSaveESRIAspect = "" And sSaveSinAspect = "" And sSaveCosAspect = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      End If
      If strSaveGRID = "" Then
         Set pGrid = Nothing
      Else
         Set pGrid = New clsGrid
         If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
         End If
      End If
      If sSaveESRIAspect = "" Then
         Set pSaveESRIAspect = Nothing
      Else
         Set pSaveESRIAspect = New clsGrid
         If Not pSaveESRIAspect.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveESRIAspect
         End If
      End If
      If sSaveSinAspect = "" Then
         Set pSaveSinAspect = Nothing
      Else
         Set pSaveSinAspect = New clsGrid
         If Not pSaveSinAspect.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveSinAspect
         End If
      End If
      If sSaveCosAspect = "" Then
         Set pSaveCosAspect = Nothing
      Else
         Set pSaveCosAspect = New clsGrid
         If Not pSaveCosAspect.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveCosAspect
         End If
      End If
      
      boolResult = modDTA1.Aspect(m_pBaseDEM, pGrid, pSaveESRIAspect, pSaveSinAspect, pSaveCosAspect)
      If boolResult Then
         If strSaveGRID <> "" Then
            If Not pGrid.SaveAscGrid(strSaveGRID, , 2) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If sSaveESRIAspect <> "" Then
            If Not pSaveESRIAspect.SaveAscGrid(sSaveESRIAspect, , 0) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveESRIAspect
            End If
         End If
         If sSaveSinAspect <> "" Then
            If Not pSaveSinAspect.SaveAscGrid(sSaveSinAspect, , 4) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveSinAspect
            End If
         End If
         If sSaveCosAspect <> "" Then
            If Not pSaveCosAspect.SaveAscGrid(sSaveCosAspect, , 4) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveCosAspect
            End If
         End If
                  
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & sSaveESRIAspect & vbCrLf _
               & sSaveSinAspect & vbCrLf & sSaveCosAspect
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
            
   Case FUNC_CURVATURE
      Dim sPlanc As String, sHorizc As String, sMeanc As String, sUnspher As String, sMinc As String, sMaxc As String
      Dim pPlanc As clsGrid, pHorizc As clsGrid, pMeanc As clsGrid, pUnspher As clsGrid, pMinc As clsGrid, pMaxc As clsGrid
            
      sPlanc = txtSavePlanCurv.Text: sHorizc = txtSaveHorizCurv.Text
      sMeanc = txtSaveMeanCurv.Text: sUnspher = txtSaveUnspher.Text
      sMinc = txtSaveMinCurv.Text: sMaxc = txtSaveMaxCurv.Text
      
      If strSaveGRID = "" And sPlanc = "" And sHorizc = "" And sMeanc = "" And sUnspher = "" And sMinc = "" And sMaxc = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      End If
      If strSaveGRID = "" Then
         Set pGrid = Nothing
      Else
         Set pGrid = New clsGrid
         If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
         End If
      End If
      If sPlanc = "" Then
         Set pPlanc = Nothing
      Else
         Set pPlanc = New clsGrid
         If Not pPlanc.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sPlanc
         End If
      End If
      If sHorizc = "" Then
         Set pHorizc = Nothing
      Else
         Set pHorizc = New clsGrid
         If Not pHorizc.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sHorizc
         End If
      End If
      If sMeanc = "" Then
         Set pMeanc = Nothing
      Else
         Set pMeanc = New clsGrid
         If Not pMeanc.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sMeanc
         End If
      End If
      If sUnspher = "" Then
         Set pUnspher = Nothing
      Else
         Set pUnspher = New clsGrid
         If Not pUnspher.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sUnspher
         End If
      End If
      If sMinc = "" Then
         Set pMinc = Nothing
      Else
         Set pMinc = New clsGrid
         If Not pMinc.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sMinc
         End If
      End If
      If sMaxc = "" Then
         Set pMaxc = Nothing
      Else
         Set pMaxc = New clsGrid
         If Not pMaxc.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sMaxc
         End If
      End If
      
      boolResult = modDTA1.Curvatures_Shary(m_pBaseDEM, pGrid, pPlanc, pHorizc, pMeanc, pUnspher, pMinc, pMaxc)
      If boolResult Then
         If strSaveGRID <> "" Then
            If Not pGrid.SaveAscGrid(strSaveGRID, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If sPlanc <> "" Then
            If Not pPlanc.SaveAscGrid(sPlanc, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sPlanc
            End If
         End If
         If sHorizc <> "" Then
            If Not pHorizc.SaveAscGrid(sHorizc, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sHorizc
            End If
         End If
         If sMeanc <> "" Then
            If Not pMeanc.SaveAscGrid(sMeanc, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sMeanc
            End If
         End If
         If sUnspher <> "" Then
            If Not pUnspher.SaveAscGrid(sUnspher, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sUnspher
            End If
         End If
         If sMinc <> "" Then
            If Not pMinc.SaveAscGrid(sMinc, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sMinc
            End If
         End If
         If sMaxc <> "" Then
            If Not pMaxc.SaveAscGrid(sMaxc, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sMaxc
            End If
         End If
         
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & sPlanc & vbCrLf _
               & sHorizc & vbCrLf & sMeanc & vbCrLf & sUnspher & vbCrLf & sMinc & vbCrLf & sMaxc
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_SurfaceCurvature
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtCs_HalfWinCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.SurfaceCurvatureIndex(m_pBaseDEM, optCsWinShape(1).Value, iPara1, pGrid)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 5) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_TopoPosIndex
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtTPI_HalfWinCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      With txtTPI_Inner_HalfWinCells
         If IsNumeric(.Text) Then
            iPara2 = CInt(.Text)
            If iPara2 <> CDbl(.Text) Or iPara2 >= iPara1 Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.TopoPosIndex(m_pBaseDEM, pGrid, optTPIWinShape(1).Value Or optTPIWinShape(2).Value, iPara1, iPara2)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 3) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_Relief
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtRelief_HalfWinCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.Relief(m_pBaseDEM, optReliefWinShape(1).Value, iPara1, pGrid)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_ElevPercentile
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtElevPctl_CirRCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.ElevPercentile(m_pBaseDEM, pGrid, iPara1)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 4) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_ElevReliefRatio
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtElevReliefR_CirRCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.ElevReliefRatio(m_pBaseDEM, pGrid, iPara1)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_TOPHAT
      Dim sSaveHillslpI As String, sSaveVlyI As String
      Dim pSaveHillslpI As clsGrid, pSaveVlyI As clsGrid
      Dim iTOPHAT_HalfWin As Integer
      
      sSaveHillslpI = txtTOPHAT_SaveHillslpI.Text
      sSaveVlyI = txtTOPHAT_SaveValleyI.Text
      If strSaveGRID = "" And sSaveHillslpI = "" And sSaveVlyI = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      End If
      With txtTOPHAT_thresh
         If IsNumeric(.Text) Then
            dPara1 = CDbl(.Text)
            If dPara1 < 0 Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      With txtTOPHAT_HalfWinCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      Set pSaveHillslpI = New clsGrid
      If Not pSaveHillslpI.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveHillslpI
      End If
      Set pSaveVlyI = New clsGrid
      If Not pSaveVlyI.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveVlyI
      End If
      
      boolResult = modDTA1.Terrain_TOPHAT(m_pBaseDEM, iPara1, dPara1, pGrid, pSaveHillslpI, pSaveVlyI)
      If boolResult Then
         If strSaveGRID <> "" Then
            If Not pGrid.SaveAscGrid(strSaveGRID, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If sSaveHillslpI <> "" Then
            If Not pSaveHillslpI.SaveAscGrid(sSaveHillslpI, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveHillslpI
            End If
         End If
         If sSaveVlyI <> "" Then
            If Not pSaveVlyI.SaveAscGrid(sSaveVlyI, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveVlyI
            End If
         End If
         
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & sSaveHillslpI & vbCrLf & sSaveVlyI
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
   
   Case FUNC_TopoRugI
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtTRI_HalfWinCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.TopoRuggednessIndex(m_pBaseDEM, optTRIWinShape(1).Value, iPara1, pGrid)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 4) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_LandPosI
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtLPos_CirRCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.LandscapePosition(m_pBaseDEM, pGrid, iPara1)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 4) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
   
   Case FUNC_UPNESS
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      If modDTA1.UPNESSIndex(m_pBaseDEM, pGrid) Then
         If pGrid.SaveAscGrid(strSaveGRID, , 0) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
   
   Case FUNC_RelaPosI
      Dim sSrcRdgGRID As String, pSrcRdgGRID As clsGrid
      Dim sSrcVlyGRID As String, pSrcVlyGRID As clsGrid
      Dim iRdgTag As Integer, iVlyTag As Integer
      Dim sSaveDist2Rdg As String, sSaveDist2Vly As String
      Dim pSaveDist2Rdg As clsGrid, pSaveDist2Vly As clsGrid
      
      sSrcVlyGRID = txtRPI_OpenValley.Text
      sSrcRdgGRID = txtRPI_OpenRidge.Text
      sSaveDist2Rdg = Trim(txtSaveDist2RdgGRID.Text): sSaveDist2Vly = Trim(txtSaveDist2VlyGRID.Text)
      If cboRPIAlg.Text <> FUNC_TYPE_RPI_routing Then
         If sSrcVlyGRID = "" Or sSrcRdgGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign source GRID firstly"
      End If
      If strSaveGRID = "" And sSaveDist2Rdg = "" And sSaveDist2Vly = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      End If
      With txtRPI_RidgeTag
         If IsNumeric(.Text) Then
            iRdgTag = CInt(.Text)
            If iRdgTag <= 0 Or iRdgTag <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      With txtRPI_ValleyTag
         If IsNumeric(.Text) Then
            iVlyTag = CInt(.Text)
            If iVlyTag <= 0 Or iVlyTag <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
            
      Set pSrcRdgGRID = New clsGrid
      If cboRPIAlg.Text = FUNC_TYPE_RPI_routing And sSrcRdgGRID = "" Then
         If Not pSrcRdgGRID.NewGrid(iCols, iRows, dXll, dYll, dCellSize, -9999, -9999, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID "
         End If
      Else
         If Not pSrcRdgGRID.LoadAscGrid(sSrcRdgGRID) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Failed to open Ridge GRID: " & sSrcRdgGRID
         End If
      End If
      Set pSrcVlyGRID = New clsGrid
      If cboRPIAlg.Text = FUNC_TYPE_RPI_routing And sSrcVlyGRID = "" Then
         If Not pSrcVlyGRID.NewGrid(iCols, iRows, dXll, dYll, dCellSize, -9999, -9999, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID "
         End If
      Else
         If Not pSrcVlyGRID.LoadAscGrid(sSrcVlyGRID) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Failed to open Valley GRID: " & sSrcVlyGRID
         End If
      End If
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      Set pSaveDist2Rdg = New clsGrid
      If Not pSaveDist2Rdg.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveDist2Rdg
      End If
      Set pSaveDist2Vly = New clsGrid
      If Not pSaveDist2Vly.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveDist2Vly
      End If
      
      Select Case cboRPIAlg.Text
      Case FUNC_TYPE_RPI_Skidmore90
         boolResult = modDTA1.RelativePosition(pSrcRdgGRID, pSrcVlyGRID, pGrid, pSaveDist2Rdg, pSaveDist2Vly, iRdgTag, iVlyTag)
      Case FUNC_TYPE_RPI_relief
         boolResult = modDTA1.RelativePositionIndex_KeepRelief(m_pBaseDEM, pSrcRdgGRID, pSrcVlyGRID, pGrid, pSaveDist2Rdg, pSaveDist2Vly, iRdgTag, iVlyTag)
      Case FUNC_TYPE_RPI_routing
         boolResult = modSlopeShape.RelativePositionIndex_KeepRouting(m_pBaseDEM, pSrcRdgGRID, pSrcVlyGRID, pGrid, pSaveDist2Rdg, pSaveDist2Vly, iRdgTag, iVlyTag)
      Case Else
         boolResult = False
      End Select
      
      If boolResult Then
         If strSaveGRID <> "" Then
            If Not pGrid.SaveAscGrid(strSaveGRID, , 4) Then
               'MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If sSaveDist2Rdg <> "" Then
            If Not pSaveDist2Rdg.SaveAscGrid(sSaveDist2Rdg, , 4) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveDist2Rdg
            End If
         End If
         If sSaveDist2Vly <> "" Then
            If Not pSaveDist2Vly.SaveAscGrid(sSaveDist2Vly, , 4) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveDist2Vly
            End If
         End If
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & sSaveDist2Rdg & vbCrLf & sSaveDist2Vly
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
   
   Case FUNC_DownslopeIndex
      Dim sSrcD8 As String, pSrcGRIDD8 As clsGrid
      
      With txtDslpI_DeltaElev
         If IsNumeric(.Text) Then
            dPara1 = CDbl(.Text)
            If dPara1 <= 0 Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      sSrcD8 = txtDslpI_OpenD8.Text
      If sSrcD8 = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign source GRID firstly"
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
                  
      Set pSrcGRIDD8 = New clsGrid
      If Not pSrcGRIDD8.LoadAscGrid(sSrcD8, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Failed to open ArcInfo FlowDir GRID: " & sSrcD8
      End If
           
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      If modDTA1.DownslopeIndex(m_pBaseDEM, pSrcGRIDD8, dPara1, pGrid) Then
         If pGrid.SaveAscGrid(strSaveGRID, , 5) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_RIDGE_Peucker
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtRidge_LowestElev
         If .Text = "" Then
            dPara1 = MIN_SINGLE
         Else
            If IsNumeric(.Text) Then
               dPara1 = CDbl(.Text)
               If dPara1 < 0 Then
                  .SetFocus
                  Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
               End If
            Else
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.FindRidge_Peucker(m_pBaseDEM, pGrid, dPara1)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 0) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
   
   Case FUNC_DRAINAGE_Peucker
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      With txtValley_UppestElev
         If .Text = "" Then
            dPara1 = MAX_SINGLE
         Else
            If IsNumeric(.Text) Then
               dPara1 = CDbl(.Text)
               If dPara1 < 0 Then
                  .SetFocus
                  Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
               End If
            Else
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      boolResult = modDTA1.FindDrainageNetwork_Peucker(m_pBaseDEM, pGrid, dPara1)
      If boolResult Then
         If pGrid.SaveAscGrid(strSaveGRID, , 0) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_MFD
      Dim strSaveSCAGRID As String
      Dim pSaveSCAGRID As clsGrid
      Dim dMFD_p As Double, dSFD_p As Double
      
      strSaveSCAGRID = txtSaveSCAGRID.Text
      If strSaveGRID = "" And strSaveSCAGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      
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
      Select Case cboMFDAlg.Text
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
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      
      Select Case cboMFDAlg.Text
      Case FUNC_TYPE_MFD_QUINN91
         boolResult = modMFD.FlowAccumulation_MFD_Quinn(m_pBaseDEM, pGrid, dMFD_p)
      Case FUNC_TYPE_MFD_QIN07
         boolResult = modMFD.FlowAccumulation_MFD_md(m_pBaseDEM, pGrid, dMFD_p, dSFD_p - dMFD_p)
      Case Else
         boolResult = False
      End Select
      
      If boolResult Then
         If strSaveGRID <> "" Then
            If pGrid.SaveAscGrid(strSaveGRID) Then
               'MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
            Else
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If strSaveSCAGRID <> "" Then
            Set pSaveSCAGRID = New clsGrid
            If Not pSaveSCAGRID.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
               Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveSCAGRID
            End If
            boolResult = modMFD.SpecificCatchmentArea(pGrid, pSaveSCAGRID, cboMFD_EffectContourLen.Text, m_pBaseDEM)
            If boolResult Then
               If Not pSaveSCAGRID.SaveAscGrid(strSaveSCAGRID, , 4) Then
                  Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveSCAGRID
               End If
            Else
               Err.Raise vbObjectError + 513, , "Failed in DTA function"
            End If
         End If
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & strSaveSCAGRID
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_TWI
      
      sSrcTanb = txtTWI_OpenSlope.Text
      If sSrcTanb = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign source GRID firstly"
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
                  
      Set pSrcGRIDTanb = New clsGrid
      If Not pSrcGRIDTanb.LoadAscGrid(sSrcTanb) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Failed to open Tan(Slope) GRID: " & sSrcTanb
      End If
           
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      If modMFD.TWI_OriginForm(m_pBaseDEM, pSrcGRIDTanb, pGrid) Then
         If pGrid.SaveAscGrid(strSaveGRID) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_StreamPowerI
      
      sSrcTanb = txtSPI_OpenSlope.Text
      If sSrcTanb = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign source GRID firstly"
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
                  
      Set pSrcGRIDTanb = New clsGrid
      If Not pSrcGRIDTanb.LoadAscGrid(sSrcTanb) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Failed to open Tan(Slope) GRID: " & sSrcTanb
      End If
           
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      If modMFD.StreamPowerIndex(m_pBaseDEM, pSrcGRIDTanb, pGrid) Then
         If pGrid.SaveAscGrid(strSaveGRID) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_TerrainCharI
      Dim sSrcCs As String, pSrcGRIDCs As clsGrid
      
      sSrcCs = txtTCI_OpenCs.Text
      If sSrcCs = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign source GRID firstly"
      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
                  
      Set pSrcGRIDCs = New clsGrid
      If Not pSrcGRIDCs.LoadAscGrid(sSrcCs) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Failed to open Cs GRID: " & sSrcCs
      End If
           
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      If modDTA1.TerrainCharI(m_pBaseDEM, pSrcGRIDCs, pGrid) Then
         If pGrid.SaveAscGrid(strSaveGRID) Then
            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
         Else
            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
         End If
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_SurfaceArea
      Dim strSaveSAR As String
      Dim pSaveSAR As clsGrid
      
      strSaveSAR = txtSaveSARGRID.Text
      If strSaveGRID = "" And strSaveSAR = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      If strSaveSAR = "" Then
         Set pSaveSAR = Nothing
      Else
         Set pSaveSAR = New clsGrid
         If Not pSaveSAR.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveSAR
         End If
      End If
      
      If modDTA1.SurfaceArea(m_pBaseDEM, pGrid, pSaveSAR) Then
         If strSaveGRID <> "" Then
            If Not pGrid.SaveAscGrid(strSaveGRID, , 3) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If strSaveSAR <> "" Then
            If Not pSaveSAR.SaveAscGrid(strSaveSAR, , 5) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveSAR
            End If
         End If
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & strSaveSAR
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
      
   Case FUNC_Openness
      Dim strSaveNegOpen As String
      Dim pSaveNegOpen As clsGrid
      
      strSaveNegOpen = txtSaveNegOpen.Text
      If strSaveGRID = "" And strSaveNegOpen = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      
      With txtOpenness_CirRCells
         If IsNumeric(.Text) Then
            iPara1 = CInt(.Text)
            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      Set pSaveNegOpen = New clsGrid
      If Not pSaveNegOpen.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveNegOpen
      End If
      
      If modDTA1.Openness(m_pBaseDEM, pGrid, pSaveNegOpen, iPara1) Then
         If strSaveGRID <> "" Then
            If Not pGrid.SaveAscGrid(strSaveGRID, , 2) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If strSaveNegOpen <> "" Then
            If Not pSaveNegOpen.SaveAscGrid(strSaveNegOpen, , 2) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveNegOpen
            End If
         End If
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & strSaveNegOpen
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
   Case FUNC_RelaRlfI
'      Dim sSrcVlyGRID As String, pSrcVlyGRID As clsGrid
'      Dim sSrcRdgGRID As String, pSrcRdgGRID As clsGrid
'      Dim iRdgTag As Integer, iVlyTag As Integer
      Dim sSaveRlf2Rdg As String, sSaveRlf2Vly As String
      Dim pSaveRlf2Rdg As clsGrid, pSaveRlf2Vly As clsGrid
      
      sSrcVlyGRID = txtRRI_OpenValley.Text
      sSrcRdgGRID = txtRRI_OpenRidge.Text
      sSaveRlf2Rdg = Trim(txtSaveRlf2RdgGRID.Text): sSaveRlf2Vly = Trim(txtSaveRlf2VlyGRID.Text)
      If cboRRIAlg.Text <> FUNC_TYPE_RRI_routing Then
         If sSrcVlyGRID = "" Or sSrcRdgGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign source GRID firstly"
      End If
      If strSaveGRID = "" And sSaveRlf2Rdg = "" And sSaveRlf2Vly = "" Then
         Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
      End If
      With txtRRI_RidgeTag
         If IsNumeric(.Text) Then
            iRdgTag = CInt(.Text)
            If iRdgTag <= 0 Or iRdgTag <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
      With txtRRI_ValleyTag
         If IsNumeric(.Text) Then
            iVlyTag = CInt(.Text)
            If iVlyTag <= 0 Or iVlyTag <> CDbl(.Text) Then
               .SetFocus
               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
            End If
         Else
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
         End If
      End With
            
      Set pSrcRdgGRID = New clsGrid
      If cboRRIAlg.Text = FUNC_TYPE_RRI_routing And sSrcRdgGRID = "" Then
         If Not pSrcRdgGRID.NewGrid(iCols, iRows, dXll, dYll, dCellSize, -9999, -9999, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID "
         End If
      Else
         If Not pSrcRdgGRID.LoadAscGrid(sSrcRdgGRID) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Failed to open Ridge GRID: " & sSrcRdgGRID
         End If
      End If
      If cboRRIAlg.Text = FUNC_TYPE_RRI_routing And sSrcVlyGRID = "" Then
         If Not pSrcVlyGRID.NewGrid(iCols, iRows, dXll, dYll, dCellSize, -9999, -9999, True) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID "
         End If
      Else
         Set pSrcVlyGRID = New clsGrid
         If Not pSrcVlyGRID.LoadAscGrid(sSrcVlyGRID) Then
            Err.Raise Number:=vbObjectError + 513, Description:="Failed to open Valley GRID: " & sSrcVlyGRID
         End If
      End If
          
      Set pGrid = New clsGrid
      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
      End If
      Set pSaveRlf2Rdg = New clsGrid
      If Not pSaveRlf2Rdg.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveRlf2Rdg
      End If
      Set pSaveRlf2Vly = New clsGrid
      If Not pSaveRlf2Vly.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & sSaveRlf2Vly
      End If
      
      Select Case cboRRIAlg.Text
      Case FUNC_TYPE_RRI_relief
         boolResult = modDTA1.RelativeRelief(m_pBaseDEM, pSrcRdgGRID, pSrcVlyGRID, pGrid, pSaveRlf2Rdg, pSaveRlf2Vly, iRdgTag, iVlyTag)
      Case FUNC_TYPE_RRI_routing
         boolResult = modSlopeShape.RelativeReliefIndex_KeepRouting(m_pBaseDEM, pSrcRdgGRID, pSrcVlyGRID, pGrid, pSaveRlf2Rdg, pSaveRlf2Vly, iRdgTag, iVlyTag)
      Case Else
         boolResult = False
      End Select
      
      If boolResult Then
         If strSaveGRID <> "" Then
            If Not pGrid.SaveAscGrid(strSaveGRID, , 4) Then
               'MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
            End If
         End If
         If sSaveRlf2Rdg <> "" Then
            If Not pSaveRlf2Rdg.SaveAscGrid(sSaveRlf2Rdg, , 4) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveRlf2Rdg
            End If
         End If
         If sSaveRlf2Vly <> "" Then
            If Not pSaveRlf2Vly.SaveAscGrid(sSaveRlf2Vly, , 4) Then
               Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & sSaveRlf2Vly
            End If
         End If
         MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID & vbCrLf & sSaveRlf2Rdg & vbCrLf & sSaveRlf2Vly
      Else
         Err.Raise vbObjectError + 513, , "Failed in DTA function"
      End If
         
'   Case FUNC_HypsomIntegral
'      If strSaveGRID = "" Then Err.Raise Number:=vbObjectError + 513, Description:="Assign output GRID firstly"
'      With txtHypsomIntegral_CirRCells
'         If IsNumeric(.Text) Then
'            iPara1 = CInt(.Text)
'            If iPara1 <= 0 Or iPara1 <> CDbl(.Text) Then
'               .SetFocus
'               Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
'            End If
'         Else
'            .SetFocus
'            Err.Raise Number:=vbObjectError + 513, Description:="Error in parameters"
'         End If
'      End With
'
'      Set pGrid = New clsGrid
'      If Not pGrid.NewGrid(iCols, iRows, dXll, dYll, dCellSize, dNoData) Then
'         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID " & strSaveGRID
'      End If
'      boolResult = modDTA1.HypsometricIntegral(m_pBaseDEM, pGrid, iPara1)
'      If boolResult Then
'         If pGrid.SaveAscGrid(strSaveGRID) Then
'            MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID
'         Else
'            Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
'         End If
'      Else
'         Err.Raise vbObjectError + 513, , "Failed in DTA function"
'      End If
   End Select
   
ErrH:
   Set pSrcVlyGRID = Nothing: Set pSrcRdgGRID = Nothing
   Set pSrcGRIDTanb = Nothing
   Set pSrcGRIDCs = Nothing
   
   Set pSaveESRIAspect = Nothing: Set pSaveSinAspect = Nothing: Set pSaveCosAspect = Nothing
   Set pPlanc = Nothing: Set pHorizc = Nothing: Set pMeanc = Nothing
   Set pUnspher = Nothing: Set pMinc = Nothing: Set pMaxc = Nothing
   Set pSaveHillslpI = Nothing: Set pSaveVlyI = Nothing
   Set pSaveSCAGRID = Nothing
   Set pSaveSAR = Nothing
   Set pSaveNegOpen = Nothing
   Set pSaveDist2Rdg = Nothing:  Set pSaveDist2Vly = Nothing
   Set pSaveRlf2Rdg = Nothing:  Set pSaveRlf2Vly = Nothing
   Set pGrid = Nothing
   
   Me.MousePointer = 0
   m_bRunning = False
   frameInput(iFuncIndex).Enabled = Not m_bRunning
   framePara(iFuncIndex).Enabled = Not m_bRunning
   frameOutput(iFuncIndex).Enabled = Not m_bRunning
   
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Sub

Private Function GetSaveFileName(Optional strOldFileName As String = "") As String
   comdlg.DialogTitle = "Save GRID"
   comdlg.FileName = ""
   GetSaveFileName = GetFileName(comdlg, False, , ".asc")
   If GetSaveFileName = "" Then GetSaveFileName = strOldFileName
End Function

Private Sub cmdSaveCosAspect_Click()
   txtSaveCosAspect.Text = GetSaveFileName(txtSaveCosAspect.Text)
End Sub

Private Sub cmdSaveDist2RdgGRID_Click()
   txtSaveDist2RdgGRID.Text = GetSaveFileName(txtSaveDist2RdgGRID.Text)
End Sub

Private Sub cmdSaveDist2VlyGRID_Click()
   txtSaveDist2VlyGRID.Text = GetSaveFileName(txtSaveDist2VlyGRID.Text)
End Sub

Private Sub cmdSaveGRID1_Click(index As Integer)
   txtSaveGRID1(index).Text = GetSaveFileName(txtSaveGRID1(index).Text)
End Sub

Private Sub cmdSaveHorizCurv_Click()
   txtSaveHorizCurv.Text = GetSaveFileName(txtSaveHorizCurv.Text)
End Sub

Private Sub cmdSaveMaxCurv_Click()
   txtSaveMaxCurv.Text = GetSaveFileName(txtSaveMaxCurv.Text)
End Sub

Private Sub cmdSaveMeanCurv_Click()
   txtSaveMeanCurv.Text = GetSaveFileName(txtSaveMeanCurv.Text)
End Sub

Private Sub cmdSaveMinCurv_Click()
   txtSaveMinCurv.Text = GetSaveFileName(txtSaveMinCurv.Text)
End Sub

Private Sub cmdSaveNegOpen_Click()
   txtSaveNegOpen.Text = GetSaveFileName(txtSaveNegOpen.Text)
End Sub

Private Sub cmdSavePlanCurv_Click()
   txtSavePlanCurv.Text = GetSaveFileName(txtSavePlanCurv.Text)
End Sub

Private Sub cmdSaveRlf2RdgGRID_Click()
   txtSaveRlf2RdgGRID.Text = GetSaveFileName(txtSaveRlf2RdgGRID.Text)
End Sub

Private Sub cmdSaveRlf2VlyGRID_Click()
   txtSaveRlf2VlyGRID.Text = GetSaveFileName(txtSaveRlf2VlyGRID.Text)
End Sub

Private Sub cmdSaveSARGRID_Click()
   txtSaveSARGRID.Text = GetSaveFileName(txtSaveSARGRID.Text)
End Sub

Private Sub cmdSaveSCAGRID_Click()
   txtSaveSCAGRID.Text = GetSaveFileName(txtSaveSCAGRID.Text)
End Sub

Private Sub cmdSaveSinAspect_Click()
   txtSaveSinAspect.Text = GetSaveFileName(txtSaveSinAspect.Text)
End Sub

Private Sub cmdSaveUnspher_Click()
   txtSaveUnspher.Text = GetSaveFileName(txtSaveUnspher.Text)
End Sub

Private Sub cmdSPI_OpenSlope_Click()
   Dim str As String
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   str = GetFileName(comdlg, True, , ".asc")
   If str <> "" Then txtSPI_OpenSlope.Text = str
End Sub

Private Sub cmdSrcGRID_Click(index As Integer)
   Dim strBaseDEM As String
On Error GoTo ErrH
   If m_bRunning Then Exit Sub
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   strBaseDEM = GetFileName(comdlg, True, , ".asc")
   If strBaseDEM = "" Then Exit Sub
   
   Dim strPath As String, strName As String, strSuffix As String
   Dim i As Integer
   
   Me.MousePointer = 11
   m_bRunning = True
   
   i = InStrRev(strBaseDEM, "\")
   m_strBasePath = Left(strBaseDEM, i)
   strName = Right(strBaseDEM, Len(strBaseDEM) - i)
   i = InStrRev(strName, ".")
   If i = 0 Then
      m_strFilePre = strName
   Else
      m_strFilePre = Left(strName, i - 1)
   End If
   
   txtSrcGRID(index).Text = strBaseDEM
   Select Case index
   Case FUNC_FILLDEP
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Fil.asc"
   Case FUNC_SLOPE
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Slp.asc"
   Case FUNC_ASPECT
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Aspect.asc"
      txtSaveArcInfoAspect.Text = m_strBasePath & m_strFilePre & "ESRIAsp.asc"
      txtSaveSinAspect.Text = m_strBasePath & m_strFilePre & "SinAsp.asc"
      txtSaveCosAspect.Text = m_strBasePath & m_strFilePre & "CosAsp.asc"
   Case FUNC_CURVATURE
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Profc.asc"
      txtSavePlanCurv.Text = m_strBasePath & m_strFilePre & "Planc.asc"
      txtSaveHorizCurv.Text = m_strBasePath & m_strFilePre & "Horizc.asc"
      txtSaveMeanCurv.Text = m_strBasePath & m_strFilePre & "Meanc.asc"
      txtSaveUnspher.Text = m_strBasePath & m_strFilePre & "Unspher.asc"
      txtSaveMinCurv.Text = m_strBasePath & m_strFilePre & "Minc.asc"
      txtSaveMaxCurv.Text = m_strBasePath & m_strFilePre & "Maxc.asc"
   Case FUNC_SurfaceCurvature
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Cs.asc"
   Case FUNC_ElevPercentile
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "ElevPctl.asc"
   Case FUNC_TOPHAT
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "TOPHAT_HillI.asc"
      txtTOPHAT_SaveHillslpI.Text = m_strBasePath & m_strFilePre & "TOPHAT_HillslpI.asc"
      txtTOPHAT_SaveValleyI.Text = m_strBasePath & m_strFilePre & "TOPHAT_VlyI.asc"
   Case FUNC_TopoRugI
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "TRI.asc"
   Case FUNC_LandPosI
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "LPos.asc"
   Case FUNC_UPNESS
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "UPNESS.asc"
   Case FUNC_RelaPosI
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "RPI.asc"
   Case FUNC_DownslopeIndex
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "DSlpI.asc"
   Case FUNC_RIDGE_Peucker
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Rdg.asc"
   Case FUNC_DRAINAGE_Peucker
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Vly.asc"
   Case FUNC_MFD
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "MFD.asc"
      txtSaveSCAGRID.Text = m_strBasePath & m_strFilePre & "SCA.asc"
   Case FUNC_TWI
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "TWI.asc"
   Case FUNC_TerrainCharI
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "TCI.asc"
   Case FUNC_StreamPowerI
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "SPI.asc"
   Case FUNC_ElevReliefRatio
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "ElevRlfR.asc"
   Case FUNC_Relief
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "Relief.asc"
   Case FUNC_TopoPosIndex
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "TPI.asc"
   Case FUNC_SurfaceArea
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "SurfArea.asc"
      txtSaveSARGRID.Text = m_strBasePath & m_strFilePre & "SAR.asc"
   Case FUNC_Openness
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "PosOpen.asc"
      txtSaveNegOpen.Text = m_strBasePath & m_strFilePre & "NegOpen.asc"
   Case FUNC_RelaRlfI
      txtSaveGRID1(index).Text = m_strBasePath & m_strFilePre & "RRI.asc"
'   Case FUNC_HypsomIntegral
'      txtSaveGRID1(Index).Text = m_strBasePath & m_strFilePre & "HypsoInt.asc"
   End Select
   
   ' load BaseDEM, read parameters in file head
   If Not (m_pBaseDEM Is Nothing) Then Set m_pBaseDEM = Nothing
   Set m_pBaseDEM = New clsGrid
   With m_pBaseDEM
      .LoadAscGrid strBaseDEM
      txtCols(index).Text = .nCols
      txtRows(index).Text = .nRows
      txtXll(index).Text = .xllcorner
      txtYll(index).Text = .yllcorner
      txtCellSize(index).Text = .CellSize
      txtNoData(index).Text = .NoData_Value
   End With
   
ErrH:
   
   Me.MousePointer = 0
   m_bRunning = False
End Sub

Private Sub cmdTCI_OpenCs_Click()
   Dim str As String
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   str = GetFileName(comdlg, True, , ".asc")
   If str <> "" Then txtTCI_OpenCs.Text = str
End Sub

Private Sub cmdTOPHAT_HillslpI_Click()
   txtTOPHAT_SaveHillslpI.Text = GetSaveFileName(txtTOPHAT_SaveHillslpI.Text)
End Sub

Private Sub cmdTOPHAT_ValleyI_Click()
   txtTOPHAT_SaveValleyI.Text = GetSaveFileName(txtTOPHAT_SaveValleyI.Text)
End Sub

Private Sub cmdTWI_OpenSlope_Click()
   Dim str As String
   comdlg.DialogTitle = "Open Src GRID"
   comdlg.FileName = ""
   str = GetFileName(comdlg, True, , ".asc")
   If str <> "" Then txtTWI_OpenSlope.Text = str
End Sub

Private Sub Form_Load()
   Dim i As Integer

   'initialize var
   m_bRunning = False
   Set m_pBaseDEM = Nothing
   
   '
   With cboSlopeType
      .Clear
      .AddItem FUNC_TYPE_SLOPE
      '.AddItem FUNC_TYPE_SLOPE_DEGREE
      .AddItem FUNC_TYPE_MAXDOWNSLOPE
      .AddItem FUNC_TYPE_LOCALDOWNSLOPE
      .ListIndex = 0
   End With
   '
   With cboCurvatureAlg
      .Clear
      .AddItem "Shary et al. (2002)"
      .ListIndex = 0
   End With
   
   With cboMFDAlg
      .Clear
      .AddItem FUNC_TYPE_MFD_QUINN91
      .AddItem FUNC_TYPE_MFD_QIN07
      .ListIndex = 1
   End With
   With cboMFD_EffectContourLen
      .Clear
      .AddItem FUNC_TYPE_EffectContourLen_Cell
      .AddItem FUNC_TYPE_EffectContourLen_UpslopeWeighted
      .ListIndex = 0
   End With
   
   With cboRPIAlg
      .Clear
      .AddItem FUNC_TYPE_RPI_Skidmore90
      .AddItem FUNC_TYPE_RPI_relief
      .AddItem FUNC_TYPE_RPI_routing
      .ListIndex = 2
   End With
   With cboRRIAlg
      .Clear
      .AddItem FUNC_TYPE_RRI_relief
      .AddItem FUNC_TYPE_RRI_routing
      .ListIndex = 1
   End With
   '
   framePara(FUNC_ASPECT).Visible = False
   framePara(FUNC_UPNESS).Visible = False
   framePara(FUNC_TerrainCharI).Visible = False
   framePara(FUNC_TWI).Visible = False
   framePara(FUNC_TerrainCharI).Visible = False
   framePara(FUNC_StreamPowerI).Visible = False
   framePara(FUNC_SurfaceArea).Visible = False
   
   For i = 0 To frameFileHead.Count - 1
      frameFileHead(i).Enabled = False
   Next
   
   txtParaMFDP(0).Enabled = C_INNER_VERSION
   txtParaMFDP(1).Enabled = C_INNER_VERSION
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_bRunning Then
      Cancel = 1
      Exit Sub
   End If
   If Not (m_pBaseDEM Is Nothing) Then Set m_pBaseDEM = Nothing
End Sub

Private Sub optTPIWinShape_Click(index As Integer)
   txtTPI_Inner_HalfWinCells.Enabled = (index = 2)
   If Not txtTPI_Inner_HalfWinCells.Enabled Then
      txtTPI_Inner_HalfWinCells.Text = "-1"
   End If
End Sub
