VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEroseByNoData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Erose cells within assigned value range"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame frameSrc 
      Caption         =   "Source"
      Height          =   1755
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   11055
      Begin VB.CommandButton cmdSrcGRID 
         Caption         =   "&Src GRID"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSrcGRID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   9615
      End
      Begin VB.Frame frameFileHead 
         Caption         =   "File Head"
         Enabled         =   0   'False
         Height          =   675
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   10815
         Begin VB.TextBox txtNoData 
            Height          =   315
            Left            =   10020
            TabIndex        =   16
            Text            =   "-9999"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCellSize 
            Height          =   315
            Left            =   8100
            TabIndex        =   15
            Text            =   "1"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtYll 
            Height          =   315
            Left            =   5880
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox txtXll 
            Height          =   315
            Left            =   3660
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtRows 
            Height          =   315
            Left            =   2040
            TabIndex        =   12
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtCols 
            Height          =   315
            Left            =   600
            TabIndex        =   11
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "NoData_Value"
            Height          =   315
            Index           =   84
            Left            =   8880
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "CellSize"
            Height          =   315
            Index           =   85
            Left            =   7380
            TabIndex        =   21
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "YllCorner"
            Height          =   315
            Index           =   86
            Left            =   5160
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "XllCorner"
            Height          =   315
            Index           =   87
            Left            =   2880
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "nRows"
            Height          =   315
            Index           =   88
            Left            =   1500
            TabIndex        =   18
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "nCols"
            Height          =   315
            Index           =   89
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   4800
      Width           =   1875
   End
   Begin VB.Frame framePara 
      Caption         =   "Parameters"
      Height          =   1575
      Left            =   60
      TabIndex        =   4
      Top             =   1860
      Width           =   11055
      Begin VB.TextBox txtSrcValueGE 
         Height          =   315
         Left            =   3780
         TabIndex        =   29
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtSrcValueSE 
         Height          =   315
         Left            =   6660
         TabIndex        =   28
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtValueAfter 
         Height          =   315
         Left            =   2940
         TabIndex        =   27
         Text            =   "-9999"
         Top             =   1140
         Width           =   735
      End
      Begin VB.OptionButton optNeighbor 
         Caption         =   "by 4-neighboring-cell"
         Height          =   255
         Index           =   1
         Left            =   5460
         TabIndex        =   26
         Top             =   360
         Width           =   2595
      End
      Begin VB.OptionButton optNeighbor 
         Caption         =   "by 8-neighboring-cell"
         Height          =   255
         Index           =   0
         Left            =   2940
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.TextBox txtLoop 
         Height          =   315
         Left            =   10140
         TabIndex        =   5
         Text            =   "1"
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "<= Value range <="
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   34
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "New value after erosion:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   33
         Top             =   1140
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "Cells to be erosed should be within:"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   3555
      End
      Begin VB.Label Label2 
         Caption         =   "Cell count with value changed: "
         Height          =   255
         Index           =   7
         Left            =   5820
         TabIndex        =   31
         Top             =   1140
         Width           =   3135
      End
      Begin VB.Label lblChangeCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9000
         TabIndex        =   30
         Top             =   1140
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Erose cells in GRID:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Count (Erosion Loop)"
         Height          =   255
         Index           =   3
         Left            =   8100
         TabIndex        =   6
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output GRID"
      Height          =   1095
      Left            =   60
      TabIndex        =   1
      Top             =   3480
      Width           =   11055
      Begin VB.TextBox txtSaveGRID 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   9615
      End
      Begin VB.CommandButton cmdSaveGRID 
         Caption         =   "S&ave GRID..."
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdErose 
      Caption         =   "&Run"
      Height          =   495
      Left            =   2940
      TabIndex        =   0
      Top             =   4800
      Width           =   5235
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   420
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEroseByNoData"
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

Private Sub cmdErose_Click()
On Error GoTo ErrH
   Dim strSaveGRID As String
   Dim iLoopCount As Integer, bool8Neighbor As Boolean
   Dim pGrid As clsGrid, pSrcGRID As clsGrid, pTemp As clsGrid
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim dValueGE As Double, dValueSE As Double, dNewValue As Double
   Dim iCol As Integer, iRow As Integer, dValue As Double, iNum As Integer, s As String
   Dim iCol1 As Integer, iRow1 As Integer, dValue1 As Double
   Dim iLoop As Integer, k As Integer
   Dim lChangeCount As Long
   
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
   bool8Neighbor = optNeighbor(0).Value
   s = txtLoop.Text
   If txtSrcGRID.Text = "" Then
      Err.Raise Number:=vbObjectError + 513, Description:="Assign the source GRID firstly"
   End If
   If IsNumeric(s) Then
      iLoopCount = CInt(s)
      If iLoopCount < 1 Or iLoopCount <> CDbl(s) Then
         txtLoop.SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Loop-count"
      End If
   Else
      txtLoop.SetFocus
      Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Loop-count"
   End If
   strSaveGRID = txtSaveGRID.Text
   
   With txtSrcValueGE
      s = .Text
      If IsNumeric(s) Then
         dValueGE = CDbl(s)
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Value-range"
      End If
   End With
   With txtSrcValueSE
      s = .Text
      If IsNumeric(s) Then
         dValueSE = CDbl(s)
         If dValueSE < dValueGE Then
            .SetFocus
            Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Value-range"
         End If
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting Value-range"
      End If
   End With
   With txtValueAfter
      s = .Text
      If IsNumeric(s) Then
         dNewValue = CDbl(s)
      Else
         .SetFocus
         Err.Raise Number:=vbObjectError + 513, Description:="Error in setting New-value-after-erosion"
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
   
   lChangeCount = 0
   If iLoopCount = 1 Then
      Set pSrcGRID = m_pBaseGRID
   Else
      Set pSrcGRID = New clsGrid
      pSrcGRID.LoadAscGrid m_pBaseGRID.sAscGridFileName
   End If
   
   For iLoop = 1 To iLoopCount
      If iLoop > 1 Then
         Set pTemp = pSrcGRID
         Set pSrcGRID = pGrid
         Set pGrid = pTemp
'      Else ' erose the margin
'         For iRow = 0 To iRows - 1
'            dValue = pSrcGRID.Cell(0, iRow)
'            If dValue < dValueGE Or dValue > dValueSE Then
'               pGrid.Cell(0, iRow) = dValue
'            Else
'               pGrid.Cell(0, iRow) = dNewValue
'               lChangeCount = lChangeCount + 1
'            End If
'
'            dValue = pSrcGRID.Cell(iCols - 1, iRow)
'            If dValue < dValueGE Or dValue > dValueSE Then
'               pGrid.Cell(iCols - 1, iRow) = dValue
'            Else
'               pGrid.Cell(iCols - 1, iRow) = dNewValue
'               lChangeCount = lChangeCount + 1
'            End If
'         Next
'
'         For iCol = 1 To iCols - 2
'            dValue = pSrcGRID.Cell(iCol, 0)
'            If dValue < dValueGE Or dValue > dValueSE Then
'               pGrid.Cell(iCol, 0) = dValue
'            Else
'               pGrid.Cell(iCol, 0) = dNewValue
'               lChangeCount = lChangeCount + 1
'            End If
'
'            dValue = pSrcGRID.Cell(iCol, iRows - 1)
'            If dValue < dValueGE Or dValue > dValueSE Then
'               pGrid.Cell(iCol, iRows - 1) = dValue
'            Else
'               pGrid.Cell(iCol, iRows - 1) = dNewValue
'               lChangeCount = lChangeCount + 1
'            End If
'         Next
      End If
      
      If bool8Neighbor Then
         For iRow = 0 To iRows - 1
            For iCol = 0 To iCols - 1
               dValue = pSrcGRID.Cell(iCol, iRow)
               If dValue >= dValueGE And dValue <= dValueSE Then
                  iNum = 0
                  For k = 1 To DIRNUM8
                     iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                     If pSrcGRID.IsValidCellValue(iCol1, iRow1, dValue1) Then
                        If dValue1 >= dValueGE And dValue1 <= dValueSE Then
                           iNum = iNum + 1
                        Else
                           Exit For
                        End If
                     Else
                        Exit For
                     End If
                  Next
                  If iNum = 8 Then
                     pGrid.Cell(iCol, iRow) = dValue
                  Else
                     pGrid.Cell(iCol, iRow) = dNewValue
                     lChangeCount = lChangeCount + 1
                  End If
               Else
                  pGrid.Cell(iCol, iRow) = dValue
               End If
            Next
         Next
      Else  ' 4-neighboring search
         For iRow = 0 To iRows - 1
            For iCol = 0 To iCols - 1
               dValue = pSrcGRID.Cell(iCol, iRow)
               If dValue >= dValueGE And dValue <= dValueSE Then
                  iNum = 0
                  For k = 2 To DIRNUM8 Step 2
                     iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                     If pSrcGRID.IsValidCellValue(iCol1, iRow1, dValue1) Then
                        If dValue1 >= dValueGE And dValue1 <= dValueSE Then
                           iNum = iNum + 1
                        Else
                           Exit For
                        End If
                     Else
                        Exit For
                     End If
                  Next
                  If iNum = 4 Then
                     pGrid.Cell(iCol, iRow) = dValue
                  Else
                     pGrid.Cell(iCol, iRow) = dNewValue
                     lChangeCount = lChangeCount + 1
                  End If
               Else
                  pGrid.Cell(iCol, iRow) = dValue
               End If
            Next
         Next
      End If
   Next  'loop
                        
'            'erose the non-Nodata area
'            If (dValue = dNoData) Then
'               pGrid.Cell(iCol, iRow) = dNoData
'            Else
'               iNum = 0
'               If bool8Neighbor Then
'                  For k = 1 To DIRNUM8
'                     iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
'                     If pSrcGRID.IsValidCellValue(iCol1, iRow1, dValue1) Then
'                        iNum = iNum + 1
'                     Else
'                        Exit For
'                     End If
'                  Next
'                  If iNum = 8 Then
'                     pGrid.Cell(iCol, iRow) = dValue
'                  Else
'                     pGrid.Cell(iCol, iRow) = dNoData
'                  End If
'               Else
'                  For k = 2 To DIRNUM8 Step 2
'                     iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
'                     If pSrcGRID.IsValidCellValue(iCol1, iRow1, dValue1) Then
'                        iNum = iNum + 1
'                     Else
'                        Exit For
'                     End If
'                  Next
'                  If iNum = 4 Then
'                     pGrid.Cell(iCol, iRow) = dValue
'                  Else
'                     pGrid.Cell(iCol, iRow) = dNoData
'                  End If
'               End If
'            End If
   
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
   
   txtSaveGRID.Text = m_strBasePath & m_strFilePre & "_Erose.asc"
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


