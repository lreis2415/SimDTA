VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNearestInterpolate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nearest Interpolation"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Do Nearest Interpolation"
      Height          =   495
      Left            =   2940
      TabIndex        =   20
      Top             =   3480
      Width           =   5235
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output GRID"
      Height          =   1095
      Left            =   60
      TabIndex        =   17
      Top             =   1860
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
      Top             =   3480
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
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   315
      Left            =   60
      TabIndex        =   21
      Top             =   2940
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmNearestInterpolate"
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
   Dim pGrid As clsGrid, pSrcGRID As clsGrid
   Dim iCols As Integer, iRows As Integer, dXll As Double, dYll As Double, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer
   Dim iCol2 As Integer, iRow2 As Integer
   Dim dTemp As Double, dValue As Double, dNearestValue As Double, dDistUnit2 As Double, dDist2 As Double
   Dim iDistUnit As Integer, iDistUnit1 As Integer, iDistUnit2 As Integer
   
   If m_bRunning Then Exit Sub
         
   ' get parameters
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
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pSrcGRID.Cell(iCol, iRow) <> dNoData Then GoTo HasValidValue
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Valid Value in Source GRID"
   
HasValidValue:
    
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pSrcGRID.Cell(iCol, iRow) = dNoData Then
         ' find the nearest valid value
            dDist2 = MAX_SINGLE
            iDistUnit = 0
            Do
               iDistUnit = iDistUnit + 1
               dDistUnit2 = (iDistUnit) ^ 2
               
               iCol2 = iCol - iDistUnit
               If iCol2 >= 0 Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                           dTemp = (iRow2 - iRow) ^ 2 + dDistUnit2
                           If dTemp < dDist2 Then
                              dDist2 = dTemp
                              dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                           End If
                        End If
                     End If
                  Next
               End If
               iCol2 = iCol + iDistUnit
               If iCol2 < iCols Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                           dTemp = (iRow2 - iRow) ^ 2 + dDistUnit2
                           If dTemp < dDist2 Then
                              dDist2 = dTemp
                              dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                           End If
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow - iDistUnit
               If iRow2 >= 0 Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                           dTemp = (iCol2 - iCol) ^ 2 + dDistUnit2
                           If dTemp < dDist2 Then
                              dDist2 = dTemp
                              dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                           End If
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow + iDistUnit
               If iRow2 < iRows Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                           dTemp = (iCol2 - iCol) ^ 2 + dDistUnit2
                           If dTemp < dDist2 Then
                              dDist2 = dTemp
                              dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                           End If
                        End If
                     End If
                  Next
               End If
               
               ' search the other part of circle neighbor region besides the rectangle neighbor region
               If dDist2 < MAX_SINGLE Then
                  iDistUnit1 = iDistUnit + 1
                  iDistUnit2 = Int(Sqr(dDist2)) + 1 ' Int(Sqr(2) * iDistUnit) + 1
                  
                  For iDistUnit = iDistUnit1 To iDistUnit2
                     dDistUnit2 = (iDistUnit) ^ 2
                     
                     iCol2 = iCol - iDistUnit
                     If iCol2 >= 0 Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                                 dTemp = (iRow2 - iRow) ^ 2 + dDistUnit2
                                 If dTemp < dDist2 Then
                                    dDist2 = dTemp
                                    dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iCol2 = iCol + iDistUnit
                     If iCol2 < iCols Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                                 dTemp = (iRow2 - iRow) ^ 2 + dDistUnit2
                                 If dTemp < dDist2 Then
                                    dDist2 = dTemp
                                    dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                                 End If
                              End If
                           End If
                        Next
                     End If
                     
                     iRow2 = iRow - iDistUnit
                     If iRow2 >= 0 Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                                 dTemp = (iCol2 - iCol) ^ 2 + dDistUnit2
                                 If dTemp < dDist2 Then
                                    dDist2 = dTemp
                                    dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow + iDistUnit
                     If iRow2 < iRows Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If pSrcGRID.Cell(iCol2, iRow2) <> dNoData Then
                                 dTemp = (iCol2 - iCol) ^ 2 + dDistUnit2
                                 If dTemp < dDist2 Then
                                    dDist2 = dTemp
                                    dNearestValue = pSrcGRID.Cell(iCol2, iRow2)
                                 End If
                              End If
                           End If
                        Next
                     End If
                  Next
               End If
            Loop Until (dDist2 < MAX_SINGLE)  ' Or (iDistUnit0 >= nRows And iDistUnit0 >= nCols)
            
            pGrid.Cell(iCol, iRow) = dNearestValue
         Else
            pGrid.Cell(iCol, iRow) = pSrcGRID.Cell(iCol, iRow)
         End If
         'SetProgressBarValue Int((iRow + 1) * 100# / iRows)
         DoEvents
      Next
      SetProgressBarValue Int((iCol + 1) * 100# / iCols)
      DoEvents
   Next
   
   If pGrid.SaveAscGrid(strSaveGRID, , -1) Then
      MsgBox "Completed. Save result GRID: " & vbCrLf & strSaveGRID, vbInformation, APP_TITLE
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strSaveGRID
   End If
   
ErrH:
   Set pGrid = Nothing
   Set pSrcGRID = Nothing
   
   Me.MousePointer = 0
   m_bRunning = False
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
   SetProgressBarValue 0
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
   
   txtSaveGRID.Text = m_strBasePath & m_strFilePre & "_FillNone.asc"
ErrH:
   Me.MousePointer = 0
   m_bRunning = False
End Sub

Private Sub Form_Load()
   'initialize var
   m_bRunning = False
   Set m_pBaseGRID = Nothing
   SetProgressBarValue 0
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

