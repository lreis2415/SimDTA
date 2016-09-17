Attribute VB_Name = "modDTA1"
Option Explicit
Option Base 1  ' index of array begins from 1

Dim m_vTag As Variant, UPNESS_ElevThresh As Double  ' for UPNESSIndex


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Hypsometric Integral (Strahler, 1952):
' HI=Sum[Zi - Min(elev)] / {n * [Max(elev)-Min(elev)]} in circle window. Value: [0,1]..
'
Public Function HypsometricIntegral(pDEM As clsGrid, pHI As clsGrid, iCirRCells As Integer, Optional dFlatValue As Double = -1) As Boolean
On Error GoTo ErrH
   Dim iRow As Integer, iCol As Integer, iDeltaCol As Integer, iDeltaRow As Integer, iCol2 As Integer, iRow2 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dMin As Double, dMax As Double, dSum As Double, dMean As Double, iNum As Integer, dElev As Double
      
   HypsometricIntegral = False
   If iCirRCells <= 0 Then Err.Raise Number:=vbObjectError + 513, Description:="iCirRCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pDEM.Cell(iCol, iRow) = dNoData Then
            pHI.Cell(iCol, iRow) = pHI.NoData_Value
         Else
            dMax = MIN_SINGLE
            dMin = MAX_SINGLE
            dSum = 0
            iNum = 0
            
            For iDeltaCol = -iCirRCells To iCirRCells
               iCol2 = iCol + iDeltaCol
               For iDeltaRow = -iCirRCells To iCirRCells
                  iRow2 = iRow + iDeltaRow
                  If pDEM.IsValidCellValue(iCol2, iRow2, dElev) Then
                     If Sqr(iDeltaCol ^ 2 + iDeltaRow ^ 2) <= iCirRCells Then
                        iNum = iNum + 1
                        dSum = dSum + dElev
                        If dMax < dElev Then dMax = dElev
                        If dMin > dElev Then dMin = dElev
                     End If
                  End If
               Next
            Next
            If iNum <= 1 Then
               pHI.Cell(iCol, iRow) = pHI.NoData_Value
            ElseIf dMax = dMin Then
               pHI.Cell(iCol, iRow) = dFlatValue
            Else
               pHI.Cell(iCol, iRow) = (pDEM.Cell(iCol, iRow) - dMin) / (dMax - dMin)
            End If
         End If
      Next
   Next
   HypsometricIntegral = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'''''''''''''''''''''''''''''
' Openness (Yokoyama et al., 2002) for visualizing topography:
'     Positive Openness (in degree) = Sum(Zenith Angle along each of the eight azimuths) / 8;
'     Negative Openness (in degree) = Sum(Nadir Angle along each of the eight azimuths) / 8
'
Public Function Openness(pDEM As clsGrid, pPosOpen As clsGrid, pNegOpen As clsGrid, iCirRCells As Integer) As Boolean
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim k As Integer, m As Integer, n As Integer
   Dim dValue As Double, dValue1 As Double, dValue2 As Double, dTemp As Double, dDist As Integer
   Dim dZenith As Double, dSumZenith As Double, dNadir As Double, dSumNadir As Double
   
On Error GoTo ErrH
   Openness = False
   If iCirRCells < 1 Then Err.Raise Number:=vbObjectError + 513, Description:="iCirRCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         dValue = pDEM.Cell(iCol, iRow)
         If (dValue <> dNoData) Then
            dSumZenith = 0#:  dSumNadir = 0#
            For k = 1 To DIRNUM8
               iCol1 = iCol: iRow1 = iRow
               n = 0
               dZenith = MIN_SINGLE: dNadir = MAX_SINGLE
               For m = 1 To iCirRCells
                  iCol1 = iCol1 + ArrDir8X(k): iRow1 = iRow1 + ArrDir8Y(k)
                  If pDEM.IsValidCell(iCol1, iRow1) Then
                     dValue1 = pDEM.Cell(iCol1, iRow1)
                     If dValue1 <> dNoData Then
                        dTemp = (dValue1 - dValue) / (dCellSize * m)
                        If k Mod 2 = 1 Then dTemp = dTemp / SQRT2
                        If dZenith < dTemp Then
                           dZenith = dTemp
                           n = m
                        End If
                        If dNadir > dTemp Then dNadir = dTemp
                     End If
                  Else
                     Exit For
                  End If
               Next
               
               If n = 0 Then
                  dZenith = 0#:  dNadir = 0#
               End If
               
               dZenith = Atn(dZenith) * COEF_2AngleDegree
               dNadir = Atn(dNadir) * COEF_2AngleDegree
               dSumZenith = dSumZenith + (90 - dZenith)
               dSumNadir = dSumNadir + (90 + dNadir)
            Next
            pPosOpen.Cell(iCol, iRow) = dSumZenith / 8
            pNegOpen.Cell(iCol, iRow) = dSumNadir / 8
         Else
            pPosOpen.Cell(iCol, iRow) = pPosOpen.NoData_Value
            pNegOpen.Cell(iCol, iRow) = pNegOpen.NoData_Value
         End If
      Next
   Next
      
   Openness = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'''''''''''''''''''''''''''''
' Surface Area (Jenness, 2004); Surface-area ratio = Surface area / cell area
'
Public Function SurfaceArea(pDEM As clsGrid, pSurfArea As clsGrid, Optional pSurfAreaRatio As clsGrid = Nothing) As Boolean
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim k As Integer, m As Integer
   Dim dValue As Double, dValue1 As Double, dValue2 As Double, dCellSize2 As Double
   Dim dS As Double, dA As Double, dB As Double, dC As Double
   Dim deltaA As Double, deltaB As Double, deltaC As Double
   Dim dSurfArea As Double
   
On Error GoTo ErrH
   SurfaceArea = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   dCellSize2 = dCellSize ^ 2
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         dValue = pDEM.Cell(iCol, iRow)
         If (dValue <> dNoData) Then
            dSurfArea = 0#
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If Not pDEM.IsValidCellValue(iCol1, iRow1, dValue1) Then
                  dValue1 = dValue
               End If
               If k = DIRNUM8 Then
                  m = 1
               Else
                  m = k + 1
               End If
               iCol2 = iCol + ArrDir8X(m): iRow2 = iRow + ArrDir8Y(m)
               If Not pDEM.IsValidCellValue(iCol2, iRow2, dValue2) Then
                  dValue2 = dValue
               End If
               
               If k Mod 2 = 1 Then
                  dA = Sqr((dValue - dValue1) ^ 2 + 2 * dCellSize2) / 2
                  dB = Sqr((dValue2 - dValue1) ^ 2 + dCellSize2) / 2
                  dC = Sqr((dValue - dValue2) ^ 2 + dCellSize2) / 2
               Else
                  dA = Sqr((dValue - dValue1) ^ 2 + dCellSize2) / 2
                  dB = Sqr((dValue2 - dValue1) ^ 2 + dCellSize2) / 2
                  dC = Sqr((dValue - dValue2) ^ 2 + 2 * dCellSize2) / 2
               End If
               dS = (dA + dB + dC) / 2
               dSurfArea = dSurfArea + Sqr(dS * (dS - dA) * (dS - dB) * (dS - dC))
            Next
            pSurfArea.Cell(iCol, iRow) = dSurfArea
         Else
            pSurfArea.Cell(iCol, iRow) = pSurfArea.NoData_Value
         End If
      Next
   Next
   
   If Not (pSurfAreaRatio Is Nothing) Then
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            dValue = pSurfArea.Cell(iCol, iRow)
            If (dValue = pSurfArea.NoData_Value) Then
               pSurfAreaRatio.Cell(iCol, iRow) = pSurfAreaRatio.NoData_Value
            Else
               pSurfAreaRatio.Cell(iCol, iRow) = pSurfArea.Cell(iCol, iRow) / dCellSize2
            End If
         Next
      Next
   End If
   
   SurfaceArea = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


''////////////////////Primary topographic attributes//////////////////////

'----Aspect based on slope in ArcInfo-------------
' Aspect is expressed in positive degrees from 0 to 360, measured clockwise from the north.
' Cells in the input grid of zero slope (flat) are assigned an aspect of -1.
' If the center cell in the immediate neighborhood is NODATA, the output is NODATA.
' If any neighborhood cells are NODATA, they are assigned the value of the center cell then the aspect is computed
'
'  rise_run = SQRT(SQR(dz/dx)+SQR(dz/dy))
'  degree_slope = ATAN(rise_run) * 57.29578  (degree)
'where the deltas are calculated using a 3x3 roving window.
'a through i represent the z_values in the window:
'     (0,0) a b c
'           d e f
'           g h i
'(dz/dx) = ((a + 2d + g) - (c + 2f + i)) / (8 * x_mesh_spacing)
'(dz/dy) = ((a + 2b + c) - (g + 2h + i)) / (8 * y_mesh_spacing)
' ref: Burrough, P.A., (1986). Principles of Geographical Information Systems for Land Resources Assessment. Oxford University Press, New York, p. 50.
'
' Aspect(degree)=180 - arctan((dz/dy) / (dz/dx)) + 90 * SGN(dz/dx)
' because the y direction in ArcInfo is to south, Aspect in ArcInfo = 360-Aspect(degree)
'
Public Function Aspect(pDEM As clsGrid, _
      Optional pAspectDegree As clsGrid = Nothing, Optional pArcInfoAspect As clsGrid = Nothing, _
      Optional pSinAspect As clsGrid = Nothing, Optional pCosAspect As clsGrid = Nothing) As Boolean
      
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim a As Double, b As Double, c As Double, d As Double, e As Double, f As Double, g As Double, h As Double, i As Double, dzdx As Double, dzdy As Double
   Dim dAspect As Double
'   Dim sCodeLine As String
   
On Error GoTo ErrH
   Aspect = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         e = pDEM.Cell(iCol, iRow)
'         sCodeLine = "row" & iRow & "col" & iCol & "-"
         If (e <> dNoData) Then
'            sCodeLine = "row" & iRow & "col" & iCol & "-e=" & e
            If Not pDEM.IsValidCellValue(iCol - 1, iRow - 1, a) Then a = e
'            sCodeLine = "row" & iRow & "col" & iCol & "-a=" & a
            If Not pDEM.IsValidCellValue(iCol, iRow - 1, b) Then b = e
            If Not pDEM.IsValidCellValue(iCol + 1, iRow - 1, c) Then c = e
            If Not pDEM.IsValidCellValue(iCol - 1, iRow, d) Then d = e
            If Not pDEM.IsValidCellValue(iCol + 1, iRow, f) Then f = e
            If Not pDEM.IsValidCellValue(iCol - 1, iRow + 1, g) Then g = e
            If Not pDEM.IsValidCellValue(iCol, iRow + 1, h) Then h = e
            If Not pDEM.IsValidCellValue(iCol + 1, iRow + 1, i) Then i = e
'            sCodeLine = "row" & iRow & "col" & iCol & "-i=" & i
            dzdx = ((a + 2 * d + g) - (c + 2 * f + i)) / (8 * dCellSize)
            dzdy = ((a + 2 * b + c) - (g + 2 * h + i)) / (8 * dCellSize)
'            sCodeLine = "row" & iRow & "col" & iCol & "-dz/dx=" & dzdx
            
            If dzdx = 0 Then
               If dzdy = 0 Then
                  dAspect = -1
               ElseIf dzdy > 0 Then
                  dAspect = 180
               Else
                  dAspect = 0
               End If
            Else
               ' because the y direction in ArcInfo is to south
               dAspect = 360 - (180 - Atn(dzdy / dzdx) * 57.29578 + 90 * Sgn(dzdx)) '57.29578: 180/PI
            End If
            
'            sCodeLine = "row" & iRow & "col" & iCol & "-dAspect=" & dAspect
            If Not (pAspectDegree Is Nothing) Then
               pAspectDegree.Cell(iCol, iRow) = dAspect
            End If
            
'            sCodeLine = "row" & iRow & "col" & iCol & "-pAspectDegree=" & dAspect
            If Not (pArcInfoAspect Is Nothing) Then
'               sCodeLine = "row" & iRow & "col" & iCol & "-pArcInfoAspect not nothing: " & dAspect
               If dAspect < 0 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_UNDEF
               ElseIf dAspect > 22.5 And dAspect <= 67.5 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_NE
               ElseIf dAspect > 67.5 And dAspect <= 112.5 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_E
               ElseIf dAspect > 112.5 And dAspect <= 157.5 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_SE
               ElseIf dAspect > 157.5 And dAspect <= 202.5 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_S
               ElseIf dAspect > 202.5 And dAspect <= 247.5 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_SW
               ElseIf dAspect > 247.5 And dAspect <= 292.5 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_W
               ElseIf dAspect > 292.5 And dAspect <= 337.5 Then
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_NW
               Else
                  pArcInfoAspect.Cell(iCol, iRow) = ESRI_DIR_N
               End If
'               sCodeLine = "row" & iRow & "col" & iCol & "-pArcInfoAspect=" & pArcInfoAspect.Cell(iCol, iRow)
            End If
                        
            
            If Not (pSinAspect Is Nothing) Then
               If dAspect < 0 Then
                  pSinAspect.Cell(iCol, iRow) = pSinAspect.NoData_Value
               Else
                  pSinAspect.Cell(iCol, iRow) = Sin(dAspect / 57.29578)
               End If
'               sCodeLine = "row" & iRow & "col" & iCol & "-pSinAspect=" & pSinAspect.Cell(iCol, iRow)
            End If
            
            
            If Not (pCosAspect Is Nothing) Then
               If dAspect < 0 Then
                  pCosAspect.Cell(iCol, iRow) = pCosAspect.NoData_Value
               Else
                  pCosAspect.Cell(iCol, iRow) = Cos(dAspect / 57.29578)
               End If
'               sCodeLine = "row" & iRow & "col" & iCol & "-pCosAspect=" & pCosAspect.Cell(iCol, iRow)
            End If
            
         Else
'            sCodeLine = "row" & iRow & "col" & iCol & "-NoDataValue"
            If Not (pAspectDegree Is Nothing) Then
'               sCodeLine = "row" & iRow & "col" & iCol & "-pAspectDegree" & pAspectDegree.NoData_Value
               pAspectDegree.Cell(iCol, iRow) = pAspectDegree.NoData_Value
            End If
            If Not (pArcInfoAspect Is Nothing) Then
'               sCodeLine = "row" & iRow & "col" & iCol & "-pArcInfoAspect" & pArcInfoAspect.NoData_Value
               pArcInfoAspect.Cell(iCol, iRow) = pArcInfoAspect.NoData_Value
            End If
            If Not (pSinAspect Is Nothing) Then
'               sCodeLine = "row" & iRow & "col" & iCol & "-pSinAspect" & pSinAspect.NoData_Value
               pSinAspect.Cell(iCol, iRow) = pSinAspect.NoData_Value
            End If
            If Not (pCosAspect Is Nothing) Then
'               sCodeLine = "row" & iRow & "col" & iCol & "-pCosAspect" & pCosAspect.NoData_Value
               pCosAspect.Cell(iCol, iRow) = pCosAspect.NoData_Value
            End If
            
         End If
      Next
   Next
   Aspect = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
'----Slope in ArcInfo-------------
'If the center cell in the immediate neighborhood (3x3 window) is NODATA, the output is NODATA.
'If any neighborhood cells are NODATA, they are assigned the value of the center cell then the slope is computed.
'  rise_run = SQRT(SQR(dz/dx)+SQR(dz/dy))
'  degree_slope = ATAN(rise_run) * 57.29578  (degree)
'where the deltas are calculated using a 3x3 roving window.
'a through i represent the z_values in the window:
'     (0,0) a b c
'           d e f
'           g h i
'(dz/dx) = ((a + 2d + g) - (c + 2f + i)) / (8 * x_mesh_spacing)
'(dz/dy) = ((a + 2b + c) - (g + 2h + i)) / (8 * y_mesh_spacing)
' ref: Burrough, P.A., (1986). Principles of Geographical Information Systems for Land Resources Assessment. Oxford University Press, New York, p. 50.
'
Public Function Slope_ArcInfo(pDEM As clsGrid, pSlope As clsGrid, Optional bInDegree As Boolean = False) As Boolean
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim a As Double, b As Double, c As Double, d As Double, e As Double, f As Double, g As Double, h As Double, i As Double, dzdx As Double, dzdy As Double
   Dim dSlope As Double
On Error GoTo ErrH
   Slope_ArcInfo = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         e = pDEM.Cell(iCol, iRow)
         If (e <> dNoData) Then
            If Not pDEM.IsValidCellValue(iCol - 1, iRow - 1, a) Then a = e
            If Not pDEM.IsValidCellValue(iCol, iRow - 1, b) Then b = e
            If Not pDEM.IsValidCellValue(iCol + 1, iRow - 1, c) Then c = e
            If Not pDEM.IsValidCellValue(iCol - 1, iRow, d) Then d = e
            If Not pDEM.IsValidCellValue(iCol + 1, iRow, f) Then f = e
            If Not pDEM.IsValidCellValue(iCol - 1, iRow + 1, g) Then g = e
            If Not pDEM.IsValidCellValue(iCol, iRow + 1, h) Then h = e
            If Not pDEM.IsValidCellValue(iCol + 1, iRow + 1, i) Then i = e
            dzdx = ((a + 2 * d + g) - (c + 2 * f + i)) / (8 * dCellSize)
            dzdy = ((a + 2 * b + c) - (g + 2 * h + i)) / (8 * dCellSize)
            If bInDegree Then
               pSlope.Cell(iCol, iRow) = Atn(Sqr(dzdx ^ 2 + dzdy ^ 2)) * 57.29578
            Else
               pSlope.Cell(iCol, iRow) = Sqr(dzdx ^ 2 + dzdy ^ 2)
            End If
         Else
            pSlope.Cell(iCol, iRow) = pSlope.NoData_Value
         End If
      Next
   Next
   Slope_ArcInfo = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

' ---Maximum Downslope
Public Function MaximumDownslope(pDEM As clsGrid, pMaxDownslope As clsGrid, Optional bInDegree As Boolean = False) As Boolean
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dMax As Double, n As Integer, dValue As Double, dValue1 As Double
   Dim dSlope As Double
On Error GoTo ErrH
   MaximumDownslope = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         dValue = pDEM.Cell(iCol, iRow)
         If (dValue <> dNoData) Then
            dMax = 0#
            n = 0
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If pDEM.IsValidCellValue(iCol1, iRow1, dValue1) Then
                  If dValue > dValue1 Then
                     dSlope = (dValue - dValue1) / dCellSize    ' tan(SlopeDegree): (iCol, y) -> k
                     If k Mod 2 = 1 Then dSlope = dSlope / SQRT2
                     If dSlope > dMax Then dMax = dSlope
                     n = n + 1
                  End If
               End If
            Next
            If bInDegree Then
               If n = 0 Then
                  pMaxDownslope.Cell(iCol, iRow) = 0
               Else
                  pMaxDownslope.Cell(iCol, iRow) = Atn(dMax) * 57.29578
               End If
            Else
               pMaxDownslope.Cell(iCol, iRow) = IIf(n = 0, 0#, dMax)
            End If
         Else
            pMaxDownslope.Cell(iCol, iRow) = pMaxDownslope.NoData_Value
         End If
      Next
   Next
   MaximumDownslope = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

' ---Local slope in the downslope direction (Quinn et al., 1991)
' tanb=Sum(tanbi*Li)/Sum(Li)
'
Public Function LocalDownslope(pDEM As clsGrid, pLocalDownslope As clsGrid, Optional bInDegree As Boolean = False) As Boolean
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dSlope As Double, dSum_Tanbi_x_Li As Double, dSum_Li As Double
   Dim n As Integer, dValue As Double, dValue1 As Double
On Error GoTo ErrH
   LocalDownslope = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         dValue = pDEM.Cell(iCol, iRow)
         If (dValue <> dNoData) Then
            dSum_Tanbi_x_Li = 0#:   dSum_Li = 0#
            n = 0
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If pDEM.IsValidCellValue(iCol1, iRow1, dValue1) Then
                  If dValue > dValue1 Then
                     dSlope = (dValue - dValue1) / dCellSize    ' tan(SlopeDegree): (iCol, y) -> k
                     If k Mod 2 = 1 Then dSlope = dSlope / SQRT2
                        
                     If k Mod 2 = 1 Then
                        dSum_Tanbi_x_Li = dSum_Tanbi_x_Li + dSlope * SQRT2 / 4
                        dSum_Li = dSum_Li + SQRT2 / 4
                     Else
                        dSum_Tanbi_x_Li = dSum_Tanbi_x_Li + dSlope * 0.5
                        dSum_Li = dSum_Li + 0.5
                     End If
                     n = n + 1
                  End If
               End If
            Next
            If bInDegree Then
               If n = 0 Then
                  pLocalDownslope.Cell(iCol, iRow) = 0
               Else
                  pLocalDownslope.Cell(iCol, iRow) = Atn(dSum_Tanbi_x_Li / dSum_Li) * 57.29578
               End If
            Else
               pLocalDownslope.Cell(iCol, iRow) = IIf(n = 0, 0#, dSum_Tanbi_x_Li / dSum_Li)
            End If
         Else
            pLocalDownslope.Cell(iCol, iRow) = pLocalDownslope.NoData_Value
         End If
      Next
   Next
   LocalDownslope = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function



' Compute profile and planform curvature by Shary et al. (2002), also ref. Young & Evans (1978); Pennock et al. (1987)
' z=(rx^2)/2+sxy+(ty^2)/2+px+qy+z0
'subgrid:          z1 z2 z3
'                  z4 z5 z6
'                  z7 z8 z9
'
' p=(z3+z6+z9-z1-z4-z7)/(6*w^2)
' q=(z1+z2+z3-z7-z8-z9)/(6*w^2)
' r=(z1+z3+z4+z6+z7+z9-2*(z2+z5+z8))/(3*w^2)
' s=(-z1+z3+z7-z9)/(4*w^2)
' t=(z1+z2+z3+z7+z8+z9-2*(z4+z5+z6))/(3*w^2)
''
' Prof. Curvature (1/m) = -(r*p^2+t*q^2+2*p*q*s)/[(p^2+q^2)*(1+p^2+q^2)^1.5]
' Plan. Curvature (1/m) = -(r*q^2+t*p^2-2*p*q*s)/[(p^2+q^2)^1.5]
' Horizontal Curvature (1/m) (similar with Plan Curvature) = - (q^2 * r - 2*p*q*s+p^2*t)/[(p^2+q^2)*(1+p^2+q^2)^1.5]
' Gradient =  (p^2+q^2)^0.5
' Mean Curvature (1/m) = -((1+q^2)*r-2*p*q*s+(1+p^2)*t)/(2*(1+p^2+q^2)^1.5)
' Unsphericity (1/m) = sqr((r*sqr((1+q^2)/(1+p^2))-t/sqr((1+q^2)/(1+p^2)))^2/(1+p^2+q^2) + (p*q*r*sqr((1+q^2)/(1+p^2))-2*sqr((1+q^2)*(1+p^2))*s+p*q*t/sqr((1+q^2)/(1+p^2)))^2)/(2*(1+p^2+q^2)^1.5)
' Minimal curvature (1/m) =Mean Curvature - Unsphericity
' Maximal curvature (1/m) =Mean Curvature + Unsphericity
''
Public Function Curvatures_Shary(pDEM As clsGrid, _
      Optional pProfCurv As clsGrid = Nothing, Optional pPlanCurv As clsGrid = Nothing, Optional pHorizCurv As clsGrid = Nothing, _
      Optional pMeanCurv As clsGrid = Nothing, Optional pUnspher As clsGrid = Nothing, _
      Optional pMinCurv As clsGrid = Nothing, Optional pMaxCurv As clsGrid = Nothing) As Boolean
      
On Error GoTo ErrH
   Dim dCellSize2 As Double, p As Double, q As Double, r As Double, s As Double, t As Double
   Dim z1 As Double, z2 As Double, z3 As Double, z4 As Double, z5 As Double, z6 As Double, z7 As Double, z8 As Double, z9 As Double
   Dim iRow As Integer, iCol As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dMeanCurve As Double, dUnspher As Double
      
   Curvatures_Shary = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   dCellSize2 = dCellSize ^ 2
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (pDEM.Cell(iCol, iRow) = dNoData) Then
            If Not (pPlanCurv Is Nothing) Then pPlanCurv.Cell(iCol, iRow) = pPlanCurv.NoData_Value
            If Not (pProfCurv Is Nothing) Then pProfCurv.Cell(iCol, iRow) = pProfCurv.NoData_Value
            If Not (pHorizCurv Is Nothing) Then pHorizCurv.Cell(iCol, iRow) = pHorizCurv.NoData_Value
            If Not (pMeanCurv Is Nothing) Then pMeanCurv.Cell(iCol, iRow) = pMeanCurv.NoData_Value
            If Not (pUnspher Is Nothing) Then pUnspher.Cell(iCol, iRow) = pUnspher.NoData_Value
            If Not (pMinCurv Is Nothing) Then pMinCurv.Cell(iCol, iRow) = pMinCurv.NoData_Value
            If Not (pMaxCurv Is Nothing) Then pMaxCurv.Cell(iCol, iRow) = pMaxCurv.NoData_Value
         Else
            z5 = pDEM.Cell(iCol, iRow)
            If pDEM.IsValidCell(iCol - 1, iRow - 1) Then
               z1 = pDEM.Cell(iCol - 1, iRow - 1)
            Else
               z1 = z5
            End If
            If pDEM.IsValidCell(iCol, iRow - 1) Then
               z2 = pDEM.Cell(iCol, iRow - 1)
            Else
               z2 = z5
            End If
            If pDEM.IsValidCell(iCol + 1, iRow - 1) Then
               z3 = pDEM.Cell(iCol + 1, iRow - 1)
            Else
               z3 = z5
            End If
            If pDEM.IsValidCell(iCol - 1, iRow) Then
               z4 = pDEM.Cell(iCol - 1, iRow)
            Else
               z4 = z5
            End If
            If pDEM.IsValidCell(iCol + 1, iRow) Then
               z6 = pDEM.Cell(iCol + 1, iRow)
            Else
               z6 = z5
            End If
            If pDEM.IsValidCell(iCol - 1, iRow + 1) Then
               z7 = pDEM.Cell(iCol - 1, iRow + 1)
            Else
               z7 = z5
            End If
            If pDEM.IsValidCell(iCol, iRow + 1) Then
               z8 = pDEM.Cell(iCol, iRow + 1)
            Else
               z8 = z5
            End If
            If pDEM.IsValidCell(iCol + 1, iRow + 1) Then
               z9 = pDEM.Cell(iCol + 1, iRow + 1)
            Else
               z9 = z5
            End If
            
            If z1 = dNoData Or z2 = dNoData Or z3 = dNoData Or z4 = dNoData _
                  Or z6 = dNoData Or z7 = dNoData Or z8 = dNoData Or z9 = dNoData Then 'IsValidXY_Value(i, j) Then
               If Not (pPlanCurv Is Nothing) Then pPlanCurv.Cell(iCol, iRow) = pPlanCurv.NoData_Value
               If Not (pProfCurv Is Nothing) Then pProfCurv.Cell(iCol, iRow) = pProfCurv.NoData_Value
               If Not (pHorizCurv Is Nothing) Then pHorizCurv.Cell(iCol, iRow) = pHorizCurv.NoData_Value
               If Not (pMeanCurv Is Nothing) Then pMeanCurv.Cell(iCol, iRow) = pMeanCurv.NoData_Value
               If Not (pUnspher Is Nothing) Then pUnspher.Cell(iCol, iRow) = pUnspher.NoData_Value
               If Not (pMinCurv Is Nothing) Then pMinCurv.Cell(iCol, iRow) = pMinCurv.NoData_Value
               If Not (pMaxCurv Is Nothing) Then pMaxCurv.Cell(iCol, iRow) = pMaxCurv.NoData_Value
            Else
               p = (z3 + z6 + z9 - z1 - z4 - z7) / (6 * dCellSize2)
               q = (z1 + z2 + z3 - z7 - z8 - z9) / (6 * dCellSize2)
               r = (z1 + z3 + z4 + z6 + z7 + z9 - 2 * (z2 + z5 + z8)) / (3 * dCellSize2)
               s = (-z1 + z3 + z7 - z9) / (4 * dCellSize2)
               t = (z1 + z2 + z3 + z7 + z8 + z9 - 2 * (z4 + z5 + z6)) / (3 * dCellSize2)
                           
               If p = 0 And q = 0 Then
                  If Not (pPlanCurv Is Nothing) Then pPlanCurv.Cell(iCol, iRow) = 0#
                  If Not (pProfCurv Is Nothing) Then pProfCurv.Cell(iCol, iRow) = 0#
                  If Not (pHorizCurv Is Nothing) Then pHorizCurv.Cell(iCol, iRow) = 0#
               Else
                  If Not (pProfCurv Is Nothing) Then pProfCurv.Cell(iCol, iRow) = -(r * p ^ 2 + t * q ^ 2 + 2 * p * q * s) / ((p ^ 2 + q ^ 2) * (1 + p ^ 2 + q ^ 2) ^ 1.5)
                  If Not (pPlanCurv Is Nothing) Then pPlanCurv.Cell(iCol, iRow) = -(r * q ^ 2 + t * p ^ 2 - 2 * p * q * s) / ((p ^ 2 + q ^ 2) ^ 1.5)
                  If Not (pHorizCurv Is Nothing) Then pHorizCurv.Cell(iCol, iRow) = -(q ^ 2 * r - 2 * p * q * s + p ^ 2 * t) / ((p ^ 2 + q ^ 2) * (1 + p ^ 2 + q ^ 2) ^ 1.5)
               End If
               
               dMeanCurve = -((1 + q ^ 2) * r - 2 * p * q * s + (1 + p ^ 2) * t) / (2 * (1 + p ^ 2 + q ^ 2) ^ 1.5)
               If Not (pMeanCurv Is Nothing) Then pMeanCurv.Cell(iCol, iRow) = dMeanCurve
               dUnspher = Sqr((r * Sqr((1 + q ^ 2) / (1 + p ^ 2)) - t / Sqr((1 + q ^ 2) / (1 + p ^ 2))) ^ 2 / (1 + p ^ 2 + q ^ 2) + (p * q * r * Sqr((1 + q ^ 2) / (1 + p ^ 2)) - 2 * Sqr((1 + q ^ 2) * (1 + p ^ 2)) * s + p * q * t / Sqr((1 + q ^ 2) / (1 + p ^ 2))) ^ 2) / (2 * (1 + p ^ 2 + q ^ 2) ^ 1.5)
               If Not (pUnspher Is Nothing) Then pUnspher.Cell(iCol, iRow) = dUnspher
               If Not (pMinCurv Is Nothing) Then pMinCurv.Cell(iCol, iRow) = dMeanCurve - dUnspher
               If Not (pMaxCurv Is Nothing) Then pMaxCurv.Cell(iCol, iRow) = dMeanCurve + dUnspher
            End If
         End If
      Next
   Next
   
   Curvatures_Shary = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Summerell et al, 2004; 2005
' function: count UPNESS Index, application usually use the  natural log (ln) of UPNESS Index
' UPNESS Index: Accumulation of upslope area at any given point, i.e., by the se of points that are connected by a continuous monotonic uphill path"
' Correct but very very slow
'
Public Function UPNESSIndex(pDEM As clsGrid, pUPNESS As clsGrid) As Boolean
On Error GoTo ErrH
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
      
   UPNESSIndex = False
   UPNESS_ElevThresh = 0#
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   ' initialize
   ReDim m_vTag(0 To iCols - 1, 0 To iRows - 1)
      
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (pDEM.Cell(iCol, iRow) = dNoData) Then
            pUPNESS.Cell(iCol, iRow) = pUPNESS.NoData_Value
         Else
            'taking vDataOut as Mask during algorithm before output GRID
            For iRow1 = 0 To iRows - 1
               For iCol1 = 0 To iCols - 1
                  m_vTag(iCol1, iRow1) = False
               Next
            Next
            pUPNESS.Cell(iCol, iRow) = 0
            
            UPNESS_Process_cell pDEM, pUPNESS, iCol, iRow, iCol, iRow
         End If
      Next
      DoEvents
   Next
   UPNESSIndex = True
ErrH:
   m_vTag = Empty
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

' called only by sub UPNESSIndex
Private Sub UPNESS_Process_cell(pDEM As clsGrid, pUPNESS As clsGrid, tarCol As Integer, tarRow As Integer, curCol As Integer, curRow As Integer)
   Dim iCol As Integer, iRow As Integer, k As Integer
   
   m_vTag(curCol, curRow) = True '1
   pUPNESS.Cell(tarCol, tarRow) = pUPNESS.Cell(tarCol, tarRow) + 1
   For k = 1 To DIRNUM8
      iCol = curCol + ArrDir8X(k): iRow = curRow + ArrDir8Y(k)
      If pDEM.IsValidCell(iCol, iRow) Then
         If (pDEM.Cell(iCol, iRow) = pDEM.NoData_Value) Then
            m_vTag(iCol, iRow) = True '1
         ElseIf pDEM.Cell(iCol, iRow) - pDEM.Cell(curCol, curRow) >= UPNESS_ElevThresh Then
            If Not m_vTag(iCol, iRow) Then
               Call UPNESS_Process_cell(pDEM, pUPNESS, tarCol, tarRow, iCol, iRow)
            End If
         End If
      End If
   Next
End Sub


'
' Relief: Max(elev)-Min(elev)
'
Public Function Relief(pDEM As clsGrid, bWinShapeIsCircle As Boolean, iHalfWinCells As Integer, pRelief As clsGrid) As Boolean
   On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, i As Integer, j As Integer
   Dim dMax As Double, dMin As Double, n As Integer, dElev As Double
   
   Relief = False
   If iHalfWinCells < 1 Then Err.Raise Number:=vbObjectError + 513, Description:="iHalfWinCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   If bWinShapeIsCircle Then
      For iCol = 0 To iCols - 1
         For iRow = 0 To iRows - 1
            If pDEM.Cell(iCol, iRow) = dNoData Then
               pRelief.Cell(iCol, iRow) = pRelief.NoData_Value
            Else
               dMax = MIN_SINGLE: dMin = MAX_SINGLE
               n = 0
               For i = -iHalfWinCells To iHalfWinCells
                  iCol1 = iCol + i
                  For j = -iHalfWinCells To iHalfWinCells
                     iRow1 = iRow + j
                     If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                        If Sqr(i ^ 2 + j ^ 2) <= iHalfWinCells Then
                           n = n + 1
                           If dMax < dElev Then dMax = dElev
                           If dMin > dElev Then dMin = dElev
                        End If
                     End If
                  Next
               Next
                   
               If n = 0 Then
                  pRelief.Cell(iCol, iRow) = pRelief.NoData_Value
               Else
                  pRelief.Cell(iCol, iRow) = dMax - dMin
               End If
            End If
         Next
      Next
   Else  ' square shape of window
      For iCol = 0 To iCols - 1
         For iRow = 0 To iRows - 1
            If pDEM.Cell(iCol, iRow) = dNoData Then
               pRelief.Cell(iCol, iRow) = pRelief.NoData_Value
            Else
               dMax = MIN_SINGLE: dMin = MAX_SINGLE
               n = 0
               For i = -iHalfWinCells To iHalfWinCells
                  iCol1 = iCol + i
                  For j = -iHalfWinCells To iHalfWinCells
                     iRow1 = iRow + j
                     If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                        If dMax < dElev Then dMax = dElev
                        If dMin > dElev Then dMin = dElev
                        n = n + 1
                     End If
                  Next
               Next
               
               If n = 0 Then
                  pRelief.Cell(iCol, iRow) = pRelief.NoData_Value
               Else
                  pRelief.Cell(iCol, iRow) = dMax - dMin
               End If
            End If
         Next
      Next
   End If
    
   Relief = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
' Surface Curvature Index: Cs=(Sum((Zi-Zaverage)/dist))/n
' origin type is for square shape of window
'
Public Function SurfaceCurvatureIndex(pDEM As clsGrid, bWinShapeIsCircle As Boolean, iHalfWinCells As Integer, pCs As clsGrid) As Boolean
   On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, i As Integer, j As Integer
   Dim dSum As Double, dAve As Double, n As Integer, dValue As Double, dElev As Double
   
   SurfaceCurvatureIndex = False
   If iHalfWinCells < 1 Then Err.Raise Number:=vbObjectError + 513, Description:="iHalfWinCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   If bWinShapeIsCircle Then
      For iCol = 0 To iCols - 1
        For iRow = 0 To iRows - 1
            dElev = pDEM.Cell(iCol, iRow)
            If dElev <> dNoData Then
                dSum = 0#
                n = 0
                For i = -iHalfWinCells To iHalfWinCells
                    iCol1 = iCol + i
                    For j = -iHalfWinCells To iHalfWinCells
                        iRow1 = iRow + j
                        If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                           If Sqr(i ^ 2 + j ^ 2) <= iHalfWinCells Then
                              dSum = dSum + dValue
                              n = n + 1
                           End If
                        End If
                    Next
                Next
                dSum = dSum - dElev
                n = n - 1
                If n = 0 Then
                    pCs.Cell(iCol, iRow) = pCs.NoData_Value
                Else
                    dAve = dSum / n
                    dSum = 0
                    For i = -iHalfWinCells To iHalfWinCells
                        iCol1 = iCol + i
                        For j = -iHalfWinCells To iHalfWinCells
                           iRow1 = iRow + j
                           If (iCol <> iCol1 Or iRow <> iRow1) And pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                              If Sqr(i ^ 2 + j ^ 2) <= iHalfWinCells Then
                                 dSum = dSum + (dValue - dAve) / (dCellSize * Sqr((iCol1 - iCol) ^ 2 + (iRow1 - iRow) ^ 2))
                              End If
                           End If
                        Next
                    Next
                    pCs.Cell(iCol, iRow) = dSum / n
                End If
            Else
                pCs.Cell(iCol, iRow) = pCs.NoData_Value
            End If
        Next
      Next
   Else  ' square shape of window
      For iCol = 0 To iCols - 1
        For iRow = 0 To iRows - 1
            dElev = pDEM.Cell(iCol, iRow)
            If dElev <> dNoData Then
                dSum = 0#
                n = 0
                For i = -iHalfWinCells To iHalfWinCells
                    iCol1 = iCol + i
                    For j = -iHalfWinCells To iHalfWinCells
                        iRow1 = iRow + j
                        If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                           dSum = dSum + dValue
                           n = n + 1
                        End If
                    Next
                Next
                dSum = dSum - dElev
                n = n - 1
                If n = 0 Then
                    pCs.Cell(iCol, iRow) = pCs.NoData_Value
                Else
                    dAve = dSum / n
                    dSum = 0
                    For i = -iHalfWinCells To iHalfWinCells
                        iCol1 = iCol + i
                        For j = -iHalfWinCells To iHalfWinCells
                           iRow1 = iRow + j
                           If (iCol <> iCol1 Or iRow <> iRow1) And pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                              dSum = dSum + (dValue - dAve) / (dCellSize * Sqr((iCol1 - iCol) ^ 2 + (iRow1 - iRow) ^ 2))
                           End If
                        Next
                    Next
                    pCs.Cell(iCol, iRow) = dSum / n
                End If
            Else
                pCs.Cell(iCol, iRow) = pCs.NoData_Value
            End If
        Next
      Next
    End If
    
    SurfaceCurvatureIndex = True
    Exit Function
ErrH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function
'
' Topographic Position Index: TPI=Z - focalmean(Zi)   (Weiss, 2001; Jenness, 2005, 2006)
' origin type is for Annulus (Ring) shape of window
'
Public Function TopoPosIndex(pDEM As clsGrid, pTPI As clsGrid, _
         bWinShapeIsCircle As Boolean, iHalfWinCells As Integer, Optional iExcludeHalfWinCells As Integer = -1) As Boolean
         
   On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, i As Integer, j As Integer
   Dim dSum As Double, dAve As Double, n As Integer, dValue As Double, dElev As Double, dTemp As Double
   
   TopoPosIndex = False
   If iHalfWinCells < 1 Then Err.Raise Number:=vbObjectError + 513, Description:="iHalfWinCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   If bWinShapeIsCircle Then
      For iCol = 0 To iCols - 1
        For iRow = 0 To iRows - 1
            dElev = pDEM.Cell(iCol, iRow)
            If dElev <> dNoData Then
                dSum = 0#
                n = 0
                For i = -iHalfWinCells To iHalfWinCells
                    iCol1 = iCol + i
                    For j = -iHalfWinCells To iHalfWinCells
                        iRow1 = iRow + j
                        If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                           dTemp = Sqr(i ^ 2 + j ^ 2)
                           If dTemp <= iHalfWinCells And dTemp > iExcludeHalfWinCells Then
                              dSum = dSum + dValue
                              n = n + 1
                           End If
                        End If
                    Next
                Next
'                dSum = dSum - dElev
'                n = n - 1
                If n = 0 Then
                    pTPI.Cell(iCol, iRow) = pTPI.NoData_Value
                Else
                    dAve = dSum / n
                    pTPI.Cell(iCol, iRow) = dElev - dAve
                End If
            Else
                pTPI.Cell(iCol, iRow) = pTPI.NoData_Value
            End If
        Next
      Next
   Else  ' square shape of window
      For iCol = 0 To iCols - 1
        For iRow = 0 To iRows - 1
            dElev = pDEM.Cell(iCol, iRow)
            If dElev <> dNoData Then
                dSum = 0#
                n = 0
                For i = -iHalfWinCells To iHalfWinCells
                    iCol1 = iCol + i
                    For j = -iHalfWinCells To iHalfWinCells
                        iRow1 = iRow + j
                        If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                           dSum = dSum + dValue
                           n = n + 1
                        End If
                    Next
                Next
'                dSum = dSum - dElev
'                n = n - 1
                If n = 0 Then
                    pTPI.Cell(iCol, iRow) = pTPI.NoData_Value
                Else
                    dAve = dSum / n
                    pTPI.Cell(iCol, iRow) = dElev - dAve
                End If
            Else
                pTPI.Cell(iCol, iRow) = pTPI.NoData_Value
            End If
        Next
      Next
    End If
    
    TopoPosIndex = True
    Exit Function
ErrH:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
' Compute the Topographic Ruggedness Index (Riley et al., 1999)
' TRI=sqrt{Sum[(delta Elevation) ^ 2] / 8}
' Multiscale TRI = Sum[(Zij-Z0)^2]/(n^2 - 1)
'
Public Function TopoRuggednessIndex(pDEM As clsGrid, bWinShapeIsCircle As Boolean, iHalfWinCells As Integer, pTRI As clsGrid) As Boolean
   Dim iRow As Integer, iCol As Integer, iCol1 As Integer, iRow1 As Integer, i As Integer, j As Integer, k As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dSum As Double, dElev As Double, n As Long, dValue As Double
   
   TopoRuggednessIndex = False
   If iHalfWinCells < 1 Then Err.Raise Number:=vbObjectError + 513, Description:="iHalfWinCells should be GREATER than 0"
On Error GoTo ErrH
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If bWinShapeIsCircle Then
      For iCol = 0 To iCols - 1
        For iRow = 0 To iRows - 1
            dElev = pDEM.Cell(iCol, iRow)
            If dElev <> dNoData Then
                dSum = 0#
                n = 0
                For i = -iHalfWinCells To iHalfWinCells
                    iCol1 = iCol + i
                    For j = -iHalfWinCells To iHalfWinCells
                        iRow1 = iRow + j
                        If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                           If Sqr(i ^ 2 + j ^ 2) <= iHalfWinCells Then
                              dSum = dSum + (dElev - dValue) ^ 2
                              n = n + 1
                           End If
                        End If
                    Next
                Next
                n = n - 1
                If n = 0 Then
                    pTRI.Cell(iCol, iRow) = pTRI.NoData_Value
                Else
                    pTRI.Cell(iCol, iRow) = Sqr(dSum / n)
                End If
            Else
                pTRI.Cell(iCol, iRow) = pTRI.NoData_Value
            End If
        Next
      Next
   Else  ' square shape of window
      For iCol = 0 To iCols - 1
        For iRow = 0 To iRows - 1
            dElev = pDEM.Cell(iCol, iRow)
            If dElev <> dNoData Then
                dSum = 0#
                n = 0
                For i = -iHalfWinCells To iHalfWinCells
                    iCol1 = iCol + i
                    For j = -iHalfWinCells To iHalfWinCells
                        iRow1 = iRow + j
                        If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                           dSum = dSum + (dElev - dValue) ^ 2
                           n = n + 1
                        End If
                    Next
                Next
                n = n - 1
                If n = 0 Then
                    pTRI.Cell(iCol, iRow) = pTRI.NoData_Value
                Else
                    pTRI.Cell(iCol, iRow) = Sqr(dSum / n)
                End If
            Else
                pTRI.Cell(iCol, iRow) = pTRI.NoData_Value
            End If
        Next
      Next
   '   For iRow = 0 To iRows - 1
   '      For iCol = 0 To iCols - 1
   '         If (pDEM.Cell(iCol, iRow) <> dNoData) Then
   '            n = 0
   '            dSum = 0#
   '            For k = 1 To DIRNUM8
   '               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
   '               If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
   '                  n = n + 1
   '                  dSum = dSum + (dElev - pDEM.Cell(iCol, iRow)) ^ 2
   '               End If
   '            Next
   '            If n = 0 Then
   '               pTRI.Cell(iCol, iRow) = 0#
   '            Else
   '               pTRI.Cell(iCol, iRow) = Sqr(dSum / n)
   '            End If
   '         Else
   '            pTRI.Cell(iCol, iRow) = pTRI.NoData_Value
   '         End If
   '      Next
   '   Next
   End If
   
   TopoRuggednessIndex = True
   Exit Function
ErrH:
   If Err.Number > 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Wilson & Gallant, 2000
' function: count the  number of points lower than the central point in a circle window.
'           pctl=100/Nc * count(Z<Zc)
'
Public Function ElevPercentile(pDEM As clsGrid, pEPI As clsGrid, iCirRCells As Integer) As Boolean
On Error GoTo ErrH
   Dim iRow As Integer, iCol As Integer, iDeltaCol As Integer, iDeltaRow As Integer, iCol2 As Integer, iRow2 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dCenter As Double, iNum As Integer, iNumLower As Integer, dElev As Double
      
   ElevPercentile = False
   If iCirRCells < 1 Then Err.Raise Number:=vbObjectError + 513, Description:="iCirRCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (pDEM.Cell(iCol, iRow) = dNoData) Then
            pEPI.Cell(iCol, iRow) = pEPI.NoData_Value
         Else
            dCenter = pDEM.Cell(iCol, iRow)
            iNum = 0
            iNumLower = 0
            For iDeltaCol = -iCirRCells To iCirRCells
               iCol2 = iCol + iDeltaCol
               For iDeltaRow = -iCirRCells To iCirRCells
                  iRow2 = iRow + iDeltaRow
                  If pDEM.IsValidCellValue(iCol2, iRow2, dElev) Then
                     If Sqr(iDeltaCol ^ 2 + iDeltaRow ^ 2) <= iCirRCells Then
                        If dElev < dCenter Then iNumLower = iNumLower + 1
                        iNum = iNum + 1
                     End If
                  End If
               Next
            Next
            pEPI.Cell(iCol, iRow) = (iNumLower * 100#) / iNum
         End If
      Next
   Next
   ElevPercentile = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Elevation-Relief Ratio (Mark, 1975; Pike and Wilson, 1971):
' ER=[Mean(elev)-Min(elev)] / [Max(elev)-Min(elev)] in circle window. Value: [0,1].
' This is a measure of the hypsography or elevation mass distribution (Etzelmuller, et al., 2007).
'
Public Function ElevReliefRatio(pDEM As clsGrid, pERR As clsGrid, iCirRCells As Integer, Optional dFlatValue As Double = -1) As Boolean
On Error GoTo ErrH
   Dim iRow As Integer, iCol As Integer, iDeltaCol As Integer, iDeltaRow As Integer, iCol2 As Integer, iRow2 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim dMin As Double, dMax As Double, dSum As Double, dMean As Double, iNum As Integer, dElev As Double
      
   ElevReliefRatio = False
   If iCirRCells <= 0 Then Err.Raise Number:=vbObjectError + 513, Description:="iCirRCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pDEM.Cell(iCol, iRow) = dNoData Then
            pERR.Cell(iCol, iRow) = pERR.NoData_Value
         Else
            dMax = MIN_SINGLE
            dMin = MAX_SINGLE
            dSum = 0
            iNum = 0
            
            For iDeltaCol = -iCirRCells To iCirRCells
               iCol2 = iCol + iDeltaCol
               For iDeltaRow = -iCirRCells To iCirRCells
                  iRow2 = iRow + iDeltaRow
                  If pDEM.IsValidCellValue(iCol2, iRow2, dElev) Then
                     If Sqr(iDeltaCol ^ 2 + iDeltaRow ^ 2) <= iCirRCells Then
                        iNum = iNum + 1
                        dSum = dSum + dElev
                        If dMax < dElev Then dMax = dElev
                        If dMin > dElev Then dMin = dElev
                     End If
                  End If
               Next
            Next
            If iNum <= 1 Then
               pERR.Cell(iCol, iRow) = pERR.NoData_Value
            ElseIf dMax = dMin Then
               pERR.Cell(iCol, iRow) = dFlatValue
            Else
               pERR.Cell(iCol, iRow) = (dSum / iNum - dMin) / (dMax - dMin)
            End If
         End If
      Next
   Next
   ElevReliefRatio = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Landscape Position Index (Fels & Matson, 1996)
' function: LPos= [Sum(Elev_iDeltaCol - Elev_0) / Dist_iDeltaCol] / n
'    The value calculated is the mean of the distance-weighted elevation differences between a given point and all other model points within a specified search radius.
' Greater positive values indicate lower topographic positions (proximal to streams)
' and greater negative values indicate higher landscape positions (ridges, summits) while values approaching zero indicate mid-slope positions.
' Where relief is minimal within the search radius, values will also tend to approach zero.
'    The extent of the search area is an important consideration, since the evaluation of landscape position will be most meaningful when confined to a single landform.
' In principle, the radius of search should be one-half of the fractal dimension of the landscape,
' that is, one half of the ridge-to-stream distance in that landscape.
'
Public Function LandscapePosition(pDEM As clsGrid, pLPos As clsGrid, iCirRCells As Integer) As Boolean
On Error GoTo ErrH
   Dim dCenter As Double, iNum As Integer, dSum As Double, dDist As Double, dElev As Double
   Dim iRow As Integer, iCol As Integer, iDeltaCol As Integer, iDeltaRow As Integer, iCol2 As Integer, iRow2 As Integer
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
      
   LandscapePosition = False
   If iCirRCells < 1 Then Err.Raise Number:=vbObjectError + 513, Description:="iCirRCells should be GREATER than 0"
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         dCenter = pDEM.Cell(iCol, iRow)
         If (dCenter = dNoData) Then
            pLPos.Cell(iCol, iRow) = pLPos.NoData_Value
         Else
            iNum = 0
            dSum = 0#
            For iDeltaCol = -iCirRCells To iCirRCells
               iCol2 = iCol + iDeltaCol
               For iDeltaRow = -iCirRCells To iCirRCells
                  iRow2 = iRow + iDeltaRow
                  If pDEM.IsValidCellValue(iCol2, iRow2, dElev) And (iDeltaCol <> 0 Or iDeltaRow <> 0) Then
                     dDist = Sqr(iDeltaCol ^ 2 + iDeltaRow ^ 2)
                     If dDist <= iCirRCells Then
                        dSum = dSum + (dElev - dCenter) / (dDist * dCellSize)
                        iNum = iNum + 1
                     End If
                  End If
               Next
            Next
            If iNum = 0 Then
               pLPos.Cell(iCol, iRow) = pLPos.NoData_Value
            Else
               pLPos.Cell(iCol, iRow) = dSum / iNum
            End If
         End If
      Next
   Next
   LandscapePosition = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
'(Hjerdt et al., 2004)
'One potentially important feature not considered in the implementation of the ln(a/tanb) index
' is the enhancement or impedance of local drainage by downslope topography.
'在wetness的计算中，以downslope index: tanαd=d/Ld（沿着最陡下坡方向垂直距离降低d米时水平距离L米）
' 代替local slope: tanb。The parameter d controls the deviation of hydraulic gradient from surface slope。
' 这样可更好地表示下坡对地下水导水的影响。
'
'算法优化: 沿流路处理
Public Function DownslopeIndex(pDEM As clsGrid, pDirD8 As clsGrid, dDownslopeElev As Double, pDSlpI As clsGrid) As Boolean
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer
   Dim dLd As Double, temp As Double, dElev As Double
   Dim boolSearch As Boolean
   Dim iCol0 As Integer, iRow0 As Integer, h0 As Double, flowdir0 As Integer
   Dim iCol1 As Integer, iRow1 As Integer, h1 As Double, flowdir1 As Integer
   Dim i As Integer
   Dim UnProcess As Double
      
On Error GoTo ErrH
   DownslopeIndex = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   If iCols <> pDirD8.nCols Or iRows <> pDirD8.nRows Or dCellSize <> pDirD8.CellSize _
         Or pDEM.xllcorner <> pDirD8.xllcorner Or pDEM.yllcorner <> pDirD8.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID DEM and FLOWDIRECTION should be with same position and same size."
   End If
   
   UnProcess = -100#
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (pDEM.Cell(iCol, iRow) = dNoData) _
            Or (pDirD8.Cell(iCol, iRow) = pDirD8.NoData_Value) Or pDirD8.Cell(iCol, iRow) = ESRI_DIR_UNDEF Then
            pDSlpI.Cell(iCol, iRow) = pDSlpI.NoData_Value
         Else
            pDSlpI.Cell(iCol, iRow) = UnProcess
         End If
      Next
   Next
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pDSlpI.Cell(iCol, iRow) = UnProcess Then    ' has not computed the downslope index
            boolSearch = True
            iCol0 = iCol: iRow0 = iRow: h0 = pDEM.Cell(iCol0, iRow0): flowdir0 = Int(pDirD8.Cell(iCol0, iRow0))
            iCol1 = iCol0: iRow1 = iRow0: h1 = h0: flowdir1 = flowdir0
            dLd = 0#
            ' along flow path, to get next downslope cell (iCol1,iRow1)
            For i = 1 To DIRNUM8
               If flowdir1 = ESRIDir(i) Then
                  iCol1 = iCol1 + ArrDir8X(i): iRow1 = iRow1 + ArrDir8Y(i)
                  Exit For
               End If
            Next
                     
            While boolSearch
               If Not pDEM.IsValidCell(iCol1, iRow1) Then
                  boolSearch = False
                  pDSlpI.Cell(iCol0, iRow0) = pDSlpI.NoData_Value
               ElseIf (pDirD8.Cell(iCol1, iRow1) = pDirD8.NoData_Value) Or (pDirD8.Cell(iCol1, iRow1) = ESRI_DIR_UNDEF) Then
                  boolSearch = False
                  pDSlpI.Cell(iCol0, iRow0) = pDSlpI.NoData_Value
               Else
                  If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If h0 - dElev >= dDownslopeElev Then
                        If flowdir1 = ESRI_DIR_N Or flowdir1 = ESRI_DIR_S Or flowdir1 = ESRI_DIR_E Or flowdir1 = ESRI_DIR_W Then
                           temp = dLd + dCellSize * (1 - (h0 - dDownslopeElev - dElev) / (h1 - dElev))
                        Else
                           temp = dLd + dCellSize * SQRT2 * (1 - (h0 - dDownslopeElev - dElev) / (h1 - dElev))
                        End If
                        pDSlpI.Cell(iCol0, iRow0) = dDownslopeElev / temp
                        
                        ' along flow path, set (iCol0,iRow0) as the neighbouring downslope cell
                        ' if the new (iCol0,iRow0) has not been computed Downslope Index, go on the loop for computing
                        For i = 1 To DIRNUM8
                           If flowdir0 = ESRIDir(i) Then
                              iCol0 = iCol0 + ArrDir8X(i): iRow0 = iRow0 + ArrDir8Y(i)
                              If Not pDEM.IsValidCellValue(iCol0, iRow0, h0) Then
                                 boolSearch = False
                                 Exit For
                              ElseIf (pDirD8.Cell(iCol0, iRow0) = pDirD8.NoData_Value) Or (pDirD8.Cell(iCol0, iRow0) = ESRI_DIR_UNDEF) Then
                                 boolSearch = False
                                 Exit For
                              ElseIf pDSlpI.Cell(iCol0, iRow0) <> UnProcess Then
                                 boolSearch = False
                                 Exit For
                              End If
                              'h0 = pDEM.Cell (iCol0, iRow0)
                              If dLd <= 0# Then ' the downslope neighboring cell of (iCol0,iRow0) fall down more than dDownslopeelev (m)
                                 iCol1 = iCol0: iRow1 = iRow0: h1 = h0: flowdir1 = flowdir0
                              ElseIf flowdir0 = ESRI_DIR_N Or flowdir0 = ESRI_DIR_S Or flowdir0 = ESRI_DIR_E Or flowdir0 = ESRI_DIR_W Then
                                 dLd = dLd - dCellSize
                              Else
                                 dLd = dLd - dCellSize * SQRT2
                              End If
                              flowdir0 = pDirD8.Cell(iCol0, iRow0)
                              Exit For
                           End If
                        Next
                     Else
                     ' (iCol1,iRow1) has not been downslope for more than dDownslopeelev(m) from (iCol0,iRow0),
                     ' go on downslope searching
                        If flowdir1 = ESRI_DIR_N Or flowdir1 = ESRI_DIR_S Or flowdir1 = ESRI_DIR_E Or flowdir1 = ESRI_DIR_W Then
                           dLd = dLd + dCellSize
                        Else
                           dLd = dLd + dCellSize * SQRT2
                        End If
                        h1 = pDEM.Cell(iCol1, iRow1)
                        flowdir1 = pDirD8.Cell(iCol1, iRow1)
                        ' along flow path, to get next downslope cell (iCol1,iRow1)
                        For i = 1 To DIRNUM8
                           If flowdir1 = ESRIDir(i) Then
                              iCol1 = iCol1 + ArrDir8X(i): iRow1 = iRow1 + ArrDir8Y(i)
                              Exit For
                           End If
                        Next
                     End If
                  Else
                     boolSearch = False
                     pDSlpI.Cell(iCol0, iRow0) = pDSlpI.NoData_Value
                  End If
               End If
            Wend
         End If
      Next
   Next
   
   DownslopeIndex = True
ErrH:
   If Err.Number > 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


' extract ridge
' by means of Peucker & Douglas, 1995
' For 2-dimensional binary image (initial value=1), move 3*3 window
' and mark the pixel with highest elevation as 0
' After scan completed, the ridges were marked as 1.
'
Public Function FindRidge_Peucker(pDEM As clsGrid, pRidge As clsGrid, Optional dLowestElev As Double = MIN_SINGLE) As Boolean
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer, k As Integer
   Dim dMin As Double, dElev As Double
On Error GoTo ErrH
   FindRidge_Peucker = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   For iCol1 = 0 To iCols - 1
        For iRow1 = 0 To iRows - 1
         If pDEM.Cell(iCol1, iRow1) = dNoData Then
            pRidge.Cell(iCol1, iRow1) = pRidge.NoData_Value    ' 若输出为整数型，此处将溢出错误，包括用pRPropOut.NoDataValue(0)
         Else
            pRidge.Cell(iCol1, iRow1) = 1
         End If
        Next
    Next
    
    For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pDEM.Cell(iCol, iRow) <> dNoData Then
            'find pixel with max value in 3*3 neighbours
            dMin = pDEM.Cell(iCol, iRow)
            iCol2 = iCol: iRow2 = iRow
            For k = 1 To DIRNUM8
                  iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                  If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If pRidge.Cell(iCol1, iRow1) = 1 Then ' and pdem(icol1, irow1) <> dgridnodata Then
                        If dMin > dElev Then
                           dMin = dElev
                           iCol2 = iCol1: iRow2 = iRow1
                        End If
                     End If
                  End If
            Next
            pRidge.Cell(iCol2, iRow2) = 0
         End If
      Next
    Next
    
    If dLowestElev > MIN_SINGLE Then
      For iCol = 0 To iCols - 1
         For iRow = 0 To iRows - 1
            If pRidge.Cell(iCol, iRow) = 1 Then
               If pDEM.Cell(iCol, iRow) < dLowestElev Then pRidge.Cell(iCol, iRow) = 0
            End If
         Next
      Next
   End If
   FindRidge_Peucker = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

' extract drainage network
' by means of Peucker & Douglas, 1995
' For 2-dimensional binary image (initial value=1), move 3*3 window
' and mark the pixel with highest elevation as 0
' After scan completed, the stream channels were marked as 1.
'
Public Function FindDrainageNetwork_Peucker(pDEM As clsGrid, pChannel As clsGrid, Optional dUppestElev As Double = MAX_SINGLE) As Boolean
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer, k As Integer
   Dim dMax As Double, dElev As Double
On Error GoTo ErrH
   FindDrainageNetwork_Peucker = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   For iCol1 = 0 To iCols - 1
        For iRow1 = 0 To iRows - 1
         If pDEM.Cell(iCol1, iRow1) = dNoData Then
            pChannel.Cell(iCol1, iRow1) = pChannel.NoData_Value
         Else
            pChannel.Cell(iCol1, iRow1) = 1
         End If
        Next
    Next
    
    For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pDEM.Cell(iCol, iRow) <> dNoData Then
            'find pixel with max value in 3*3 neighbours
            dMax = pDEM.Cell(iCol, iRow)
            iCol2 = iCol: iRow2 = iRow
            For k = 1 To DIRNUM8
                  iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                  If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If pChannel.Cell(iCol1, iRow1) = 1 Then ' and pDEM(iCol1, iRow1) <> dgridnodata Then
                        If dMax < dElev Then
                           dMax = dElev
                           iCol2 = iCol1: iRow2 = iRow1
                        End If
                     End If
                  End If
            Next
            pChannel.Cell(iCol2, iRow2) = 0
         End If
      Next
   Next
           
   If dUppestElev < MAX_SINGLE Then
      For iCol = 0 To iCols - 1
         For iRow = 0 To iRows - 1
            If pChannel.Cell(iCol, iRow) = 1 Then
               If pDEM.Cell(iCol, iRow) > dUppestElev Then pChannel.Cell(iCol, iRow) = 0
            End If
         Next
      Next
   End If
   
   FindDrainageNetwork_Peucker = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'
' Skidmore A K. Terrain position as mapped from gridded digital elevation data. IJGIS, 1990, 4(1): 33-49.
' Relative position:
'        Pij = (Euclidean distance to the nearest valley)
'              / (Euclidean distance to the nearest valley + Euclidean distance to the nearest ridge)
'
' Pij<0.1 → valley; 0.1≤Pij<0.4 → lower mid-slope; 0.4≤Pij<0.6 → mid-slope; 0.6≤Pij<0.8 → an upper mid-slope; Pij≥0.8 → ridge.
'
Public Function RelativePosition(pRidge As clsGrid, pChannel As clsGrid, pRPI As clsGrid, _
      pDist2Ridge As clsGrid, pDist2Valley As clsGrid, _
      Optional iRidgeTag As Integer = 1, Optional iChannelTag As Integer = 1) As Boolean
On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dNoData As Double
   'Dim vDist2Ridge As Variant, vDist2Valley As Variant   'vTemp As Variant
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer
   Dim dCellSize As Double, dDist As Double, iDistUnit As Integer, dTemp As Double, iDistUnit1 As Integer, iDistUnit2 As Integer, iDistUnit0 As Integer
         
   RelativePosition = False
   With pRidge
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   If iCols <> pChannel.nCols Or iRows <> pChannel.nRows Or dCellSize <> pChannel.CellSize Then
      '   Or pRidge.xllcorner <> pChannel.xllcorner Or pRidge.yllcorner <> pChannel.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID Ridge and Channel should be with same position and same size."
   End If
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pChannel.Cell(iCol, iRow) = iChannelTag Then GoTo HasValley
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Valley-cell in given Valley GRID"
HasValley:
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then GoTo HasRidge
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Ridge-cell in given Ridge GRID"
HasRidge:
   DoEvents
   
   ' initialize   ' pSizeIn==pSize
   'ReDim vDist2Ridge(0 To iCols - 1, 0 To iRows - 1)
   'ReDim vDist2Valley(0 To iCols - 1, 0 To iRows - 1)
      
   ' compute the distance from the every pixel to the nearest valley pixel
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pChannel.Cell(iCol, iRow) = iChannelTag Then
            pDist2Valley.Cell(iCol, iRow) = 0# 'vDist2Valley(iCol, iRow) = 0#
         Else
            dDist = MAX_SINGLE
            iDistUnit0 = 0
            Do
               iDistUnit0 = iDistUnit0 + 1
               iDistUnit = iDistUnit0
               iCol2 = iCol - iDistUnit
               If iCol2 >= 0 Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iCol2 = iCol + iDistUnit
               If iCol2 < iCols Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow - iDistUnit
               If iRow2 >= 0 Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow + iDistUnit
               If iRow2 < iRows Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               
               If dDist < MAX_SINGLE Then
                  iDistUnit1 = iDistUnit + 1
                  iDistUnit2 = Int(dDist) + 1 ' Int(Sqr(2) * iDistUnit) + 1
                  For iDistUnit = iDistUnit1 To iDistUnit2
                     iCol2 = iCol - iDistUnit
                     If iCol2 >= 0 Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iCol2 = iCol + iDistUnit
                     If iCol2 < iCols Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow - iDistUnit
                     If iRow2 >= 0 Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow + iDistUnit
                     If iRow2 < iRows Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                  Next
               End If
            Loop Until (dDist < MAX_SINGLE) 'Or (iDistUnit0 >= nRows And iDistUnit0 >= nCols)
            pDist2Valley.Cell(iCol, iRow) = dDist * dCellSize
         End If
      Next
   Next
   DoEvents
   
   ' compute the distance from the every pixel to the nearest Ridge pixel
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then
            pDist2Ridge.Cell(iCol, iRow) = 0#
         Else
            dDist = MAX_SINGLE
            iDistUnit0 = 0
            Do
               iDistUnit0 = iDistUnit0 + 1
               iDistUnit = iDistUnit0
               iCol2 = iCol - iDistUnit
               If iCol2 >= 0 Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iCol2 = iCol + iDistUnit
               If iCol2 < iCols Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow - iDistUnit
               If iRow2 >= 0 Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow + iDistUnit
               If iRow2 < iRows Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               
               If dDist < MAX_SINGLE Then
                  iDistUnit1 = iDistUnit + 1
                  iDistUnit2 = Int(dDist) + 1 ' Int(Sqr(2) * iDistUnit) + 1
                  For iDistUnit = iDistUnit1 To iDistUnit2
                     iCol2 = iCol - iDistUnit
                     If iCol2 >= 0 Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iCol2 = iCol + iDistUnit
                     If iCol2 < iCols Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow - iDistUnit
                     If iRow2 >= 0 Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow + iDistUnit
                     If iRow2 < iRows Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                  Next
               End If
            Loop Until (dDist < MAX_SINGLE) 'Or (iDistUnit0 >= nRows And iDistUnit0 >= nCols)
            pDist2Ridge.Cell(iCol, iRow) = dDist * dCellSize
         End If
      Next
   Next
   DoEvents
   
   ' Relative position = (Euclidean distance to the nearest valley) / (Euclidean distance to the nearest valley + Euclidean distance to nearest ridge)
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pDist2Ridge.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 1#
         ElseIf pDist2Valley.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 0#
'         ElseIf vDist2Ridge(iCol, iRow) = MAX_SINGLE Or vDist2Valley(iCol, iRow) = MAX_SINGLE Then
'            pRPI.Cell(iCol, iRow) = pRPI.NoData_Value
         Else
            pRPI.Cell(iCol, iRow) = pDist2Valley.Cell(iCol, iRow) / (pDist2Valley.Cell(iCol, iRow) + pDist2Ridge.Cell(iCol, iRow))
         End If
      Next
   Next
   
   RelativePosition = True
ErrH:
   ' Release memeory
   'vDist2Ridge = Empty:   vDist2Valley = Empty
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
' Revised Relative position index:
'        Pij = (Euclidean distance to the nearest valley)
'              / (Euclidean distance to the nearest valley + Euclidean distance to the nearest ridge)
' P.S.   the nearest valley: not higher than the interest cell;
'        the nearest ridge: not lower than the interest cell.
' Pij<0.1 → valley; 0.1≤Pij<0.4 → lower mid-slope; 0.4≤Pij<0.6 → mid-slope; 0.6≤Pij<0.8 → an upper mid-slope; Pij≥0.8 → ridge.
'
Public Function RelativePositionIndex_KeepRelief(pDEM As clsGrid, pRidge As clsGrid, pChannel As clsGrid, _
      pRPI As clsGrid, pDist2Ridge As clsGrid, pDist2Valley As clsGrid, _
      Optional iRidgeTag As Integer = 1, Optional iChannelTag As Integer = 1) As Boolean
On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dNoData As Double
   'Dim vDist2Ridge As Variant, vDist2Valley As Variant   'vTemp As Variant
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer
   Dim dCellSize As Double, dDist As Double, iDistUnit As Integer, dTemp As Double, iDistUnit1 As Integer, iDistUnit2 As Integer, iDistUnit0 As Integer
   Dim dElev As Double
   
   RelativePositionIndex_KeepRelief = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   If iCols <> pRidge.nCols Or iRows <> pRidge.nRows Or dCellSize <> pRidge.CellSize _
         Or iCols <> pChannel.nCols Or iRows <> pChannel.nRows Or dCellSize <> pChannel.CellSize Then
      '   Or pRidge.xllcorner <> pChannel.xllcorner Or pRidge.yllcorner <> pChannel.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID DEM, Ridge and Channel should be with same position and same size."
   End If
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pChannel.Cell(iCol, iRow) = iChannelTag Then GoTo HasValley
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Valley-cell in given Valley GRID"
HasValley:
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then GoTo HasRidge
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Ridge-cell in given Ridge GRID"
HasRidge:
   DoEvents
   
   ' initialize   ' pSizeIn==pSize
   'ReDim vDist2Ridge(0 To iCols - 1, 0 To iRows - 1)
   'ReDim vDist2Valley(0 To iCols - 1, 0 To iRows - 1)
      
   ' compute the distance from the every pixel to the nearest valley pixel
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         dDist = MAX_SINGLE
         If pChannel.Cell(iCol, iRow) = iChannelTag Then
            pDist2Valley.Cell(iCol, iRow) = 0# 'vDist2Valley(iCol, iRow) = 0#
            dDist = 0#
         ElseIf pDEM.Cell(iCol, iRow) = dNoData Then
            pDist2Valley.Cell(iCol, iRow) = pDist2Valley.NoData_Value
         Else
            dElev = pDEM.Cell(iCol, iRow)
            iDistUnit0 = 0
            Do
               iDistUnit0 = iDistUnit0 + 1
               iDistUnit = iDistUnit0
               iCol2 = iCol - iDistUnit
               If iCol2 >= 0 Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iCol2 = iCol + iDistUnit
               If iCol2 < iCols Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow - iDistUnit
               If iRow2 >= 0 Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow + iDistUnit
               If iRow2 < iRows Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               
               If dDist < MAX_SINGLE Then
                  iDistUnit1 = iDistUnit + 1
                  iDistUnit2 = Int(dDist) + 1 ' Int(Sqr(2) * iDistUnit) + 1
                  For iDistUnit = iDistUnit1 To iDistUnit2
                     iCol2 = iCol - iDistUnit
                     If iCol2 >= 0 Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iCol2 = iCol + iDistUnit
                     If iCol2 < iCols Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow - iDistUnit
                     If iRow2 >= 0 Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow + iDistUnit
                     If iRow2 < iRows Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                  Next
               End If
            Loop Until (dDist < MAX_SINGLE) Or (iDistUnit0 >= iRows And iDistUnit0 >= iCols)
            
            If dDist = MAX_SINGLE Then
               pDist2Valley.Cell(iCol, iRow) = pDist2Valley.NoData_Value
            Else
               pDist2Valley.Cell(iCol, iRow) = dDist * dCellSize
            End If
         End If
      Next
   Next
   DoEvents
   
   ' compute the distance from the every pixel to the nearest Ridge pixel
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         dDist = MAX_SINGLE
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then
            dDist = 0#
            pDist2Ridge.Cell(iCol, iRow) = 0#
         ElseIf pDEM.Cell(iCol, iRow) = dNoData Then
            pDist2Ridge.Cell(iCol, iRow) = pDist2Ridge.NoData_Value
         Else
            dElev = pDEM.Cell(iCol, iRow)
            iDistUnit0 = 0
            Do
               iDistUnit0 = iDistUnit0 + 1
               iDistUnit = iDistUnit0
               iCol2 = iCol - iDistUnit
               If iCol2 >= 0 Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iCol2 = iCol + iDistUnit
               If iCol2 < iCols Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow - iDistUnit
               If iRow2 >= 0 Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow + iDistUnit
               If iRow2 < iRows Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then dDist = dTemp
                        End If
                     End If
                  Next
               End If
               
               If dDist < MAX_SINGLE Then
                  iDistUnit1 = iDistUnit + 1
                  iDistUnit2 = Int(dDist) + 1 ' Int(Sqr(2) * iDistUnit) + 1
                  For iDistUnit = iDistUnit1 To iDistUnit2
                     iCol2 = iCol - iDistUnit
                     If iCol2 >= 0 Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iCol2 = iCol + iDistUnit
                     If iCol2 < iCols Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow - iDistUnit
                     If iRow2 >= 0 Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow + iDistUnit
                     If iRow2 < iRows Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then dDist = dTemp
                              End If
                           End If
                        Next
                     End If
                  Next
               End If
            Loop Until (dDist < MAX_SINGLE) Or (iDistUnit0 >= iRows And iDistUnit0 >= iCols)
            
            If dDist < MAX_SINGLE Then
               pDist2Ridge.Cell(iCol, iRow) = dDist * dCellSize
            Else
               pDist2Ridge.Cell(iCol, iRow) = pDist2Ridge.NoData_Value
            End If
         End If
      Next
   Next
   DoEvents
   
   ' Relative position = (Euclidean distance to the nearest valley) / (Euclidean distance to the nearest valley + Euclidean distance to nearest ridge)
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pDist2Ridge.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 1#
         ElseIf pDist2Valley.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 0#
'         ElseIf vDist2Ridge(iCol, iRow) = MAX_SINGLE Or vDist2Valley(iCol, iRow) = MAX_SINGLE Then
'            pRPI.Cell(iCol, iRow) = pRPI.NoData_Value
         ElseIf pDist2Ridge.Cell(iCol, iRow) = pDist2Ridge.NoData_Value Or pDist2Valley.Cell(iCol, iRow) = pDist2Valley.NoData_Value Then
            pRPI.Cell(iCol, iRow) = pRPI.NoData_Value
         ElseIf pDist2Valley.Cell(iCol, iRow) + pDist2Ridge.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 1#
         Else
            pRPI.Cell(iCol, iRow) = pDist2Valley.Cell(iCol, iRow) / (pDist2Valley.Cell(iCol, iRow) + pDist2Ridge.Cell(iCol, iRow))
         End If
      Next
   Next
   
   RelativePositionIndex_KeepRelief = True
ErrH:
   ' Release memeory
   'vDist2Ridge = Empty:   vDist2Valley = Empty
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
' Revised Relative position index:
'        Pij = (Euclidean distance to the nearest valley)
'              / (Euclidean distance to the nearest valley + Euclidean distance to the nearest ridge)
' P.S.   the nearest valley: routed from the interest cell;
'        the nearest ridge: routed to the interest cell.
' Pij<0.1 → valley; 0.1≤Pij<0.4 → lower mid-slope; 0.4≤Pij<0.6 → mid-slope; 0.6≤Pij<0.8 → an upper mid-slope; Pij≥0.8 → ridge.
'
Public Function RelativePositionIndex_KeepRouting0(pDEM As clsGrid, pRidge As clsGrid, pChannel As clsGrid, _
      pRPI As clsGrid, pDist2Ridge As clsGrid, pDist2Valley As clsGrid, _
      Optional iRidgeTag As Integer = 1, Optional iChannelTag As Integer = 1) As Boolean
On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dNoData As Double, dCellSize As Double
   Dim iCol As Integer, iRow As Integer, k As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer
   Dim dDist As Double, dTemp As Double
   Dim dElev As Double, dValue As Double, boolTemp As Boolean, lTemp As Long
   
   Dim queRow As clsQueue, queCol As clsQueue
   
   RelativePositionIndex_KeepRouting0 = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   If iCols <> pRidge.nCols Or iRows <> pRidge.nRows Or dCellSize <> pRidge.CellSize _
         Or iCols <> pChannel.nCols Or iRows <> pChannel.nRows Or dCellSize <> pChannel.CellSize Then
      '   Or pRidge.xllcorner <> pChannel.xllcorner Or pRidge.yllcorner <> pChannel.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID DEM, Ridge and Channel should be with same position and same size."
   End If
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pChannel.Cell(iCol, iRow) = iChannelTag Then GoTo HasValley
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Valley-cell in given Valley GRID"
HasValley:
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then GoTo HasRidge
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Ridge-cell in given Ridge GRID"
HasRidge:
   DoEvents
   
   Set queRow = New clsQueue: Set queCol = New clsQueue
      
   ' compute the distance from the every pixel to the nearest valley pixel
   'initial
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pChannel.Cell(iCol, iRow) = iChannelTag Then
            pDist2Valley.Cell(iCol, iRow) = 0#
         Else
            pDist2Valley.Cell(iCol, iRow) = MAX_SINGLE
         End If
      Next
   Next
   
   'update NearestDist2Vly along upstream from valley cell
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pChannel.Cell(iCol, iRow) <> iChannelTag And pDEM.IsValidCellValue(iCol, iRow, dElev) Then
            queRow.Clear:  queCol.Clear
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                  If pChannel.Cell(iCol1, iRow1) <> iChannelTag Then
                     If dValue > dElev Then
                        dDist = Sqr((iCol1 - iCol) ^ 2 + (iRow1 - iRow) ^ 2)
                        If pDist2Valley.Cell(iCol1, iRow1) > dDist Then
                           pDist2Valley.Cell(iCol1, iRow1) = dDist
                        End If
                        If pRidge.Cell(iCol1, iRow1) <> iRidgeTag Then
                           queCol.Add iCol1
                           queRow.Add iRow1
                        End If
                     End If
                  End If
               End If
            Next
            
            Do While queCol.ItemCount > 0
               queCol.Popup (iCol2):   queRow.Popup (iRow2)
               dElev = pDEM.Cell(iCol2, iRow2)
               For k = 1 To DIRNUM8
                  iCol1 = iCol2 + ArrDir8X(k): iRow1 = iRow2 + ArrDir8Y(k)
                  If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                     If pChannel.Cell(iCol1, iRow1) <> iChannelTag Then
                        If dValue > dElev Then
                           boolTemp = True
                           For lTemp = 1 To queCol.ItemCount
                              If queCol.EqualItem(lTemp, iCol1) And queRow.EqualItem(lTemp, iRow1) Then
                                 boolTemp = False
                                 Exit For
                              End If
                           Next
                           
                           If boolTemp Then
                              dDist = Sqr((iCol1 - iCol) ^ 2 + (iRow1 - iRow) ^ 2)
                              If pDist2Valley.Cell(iCol1, iRow1) > dDist Then
                                 pDist2Valley.Cell(iCol1, iRow1) = dDist
                              End If
                              If pRidge.Cell(iCol1, iRow1) <> iRidgeTag Then
                                 queCol.Add iCol1
                                 queRow.Add iRow1
                              End If
                           End If
                        End If
                     End If
                  End If
               Next
            Loop
            DoEvents
         End If
      Next
   Next
   
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pDist2Valley.Cell(iCol, iRow) = MAX_SINGLE Then
            pDist2Valley.Cell(iCol, iRow) = pDist2Valley.NoData_Value
         Else
            pDist2Valley.Cell(iCol, iRow) = pDist2Valley.Cell(iCol, iRow) * dCellSize
         End If
      Next
   Next
   DoEvents
   
      
   ' compute the distance from the every pixel to the nearest ridge pixel
   'initial
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then
            pDist2Ridge.Cell(iCol, iRow) = 0#
         Else
            pDist2Ridge.Cell(iCol, iRow) = MAX_SINGLE
         End If
      Next
   Next
   
   'update NearestDist2Rdg along downstream from ridge cell
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pRidge.Cell(iCol, iRow) <> iRidgeTag And pDEM.IsValidCellValue(iCol, iRow, dElev) Then
            queRow.Clear:  queCol.Clear
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                  If pRidge.Cell(iCol1, iRow1) <> iRidgeTag Then
                     If dValue < dElev Then
                        dDist = Sqr((iCol1 - iCol) ^ 2 + (iRow1 - iRow) ^ 2)
                        If pDist2Ridge.Cell(iCol1, iRow1) > dDist Then
                           pDist2Ridge.Cell(iCol1, iRow1) = dDist
                        End If
                        If pChannel.Cell(iCol1, iRow1) <> iChannelTag Then
                           queCol.Add iCol1
                           queRow.Add iRow1
                        End If
                     End If
                  End If
               End If
            Next
            
            Do While queCol.ItemCount > 0
               queCol.Popup (iCol2):   queRow.Popup (iRow2)
               dElev = pDEM.Cell(iCol2, iRow2)
               For k = 1 To DIRNUM8
                  iCol1 = iCol2 + ArrDir8X(k): iRow1 = iRow2 + ArrDir8Y(k)
                  If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                     If pRidge.Cell(iCol1, iRow1) <> iRidgeTag Then
                        If dValue < dElev Then
                           boolTemp = True
                           For lTemp = 1 To queCol.ItemCount
                              If queCol.EqualItem(lTemp, iCol1) And queRow.EqualItem(lTemp, iRow1) Then
                                 boolTemp = False
                                 Exit For
                              End If
                           Next
                           
                           If boolTemp Then
                              dDist = Sqr((iCol1 - iCol) ^ 2 + (iRow1 - iRow) ^ 2)
                              If pDist2Ridge.Cell(iCol1, iRow1) > dDist Then
                                 pDist2Ridge.Cell(iCol1, iRow1) = dDist
                              End If
                              If pChannel.Cell(iCol1, iRow1) <> iChannelTag Then
                                 queCol.Add iCol1
                                 queRow.Add iRow1
                              End If
                           End If
                        End If
                     End If
                  End If
               Next
            Loop
            DoEvents
         End If
      Next
   Next
   
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pDist2Ridge.Cell(iCol, iRow) = MAX_SINGLE Then
            pDist2Ridge.Cell(iCol, iRow) = pDist2Ridge.NoData_Value
         Else
            pDist2Ridge.Cell(iCol, iRow) = pDist2Ridge.Cell(iCol, iRow) * dCellSize
         End If
      Next
   Next
   DoEvents
   
   ' Relative position = (Euclidean distance to the nearest valley) / (Euclidean distance to the nearest valley + Euclidean distance to nearest ridge)
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pDist2Ridge.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 1#
         ElseIf pDist2Valley.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 0#
'         ElseIf vDist2Ridge(iCol, iRow) = MAX_SINGLE Or vDist2Valley(iCol, iRow) = MAX_SINGLE Then
'            pRPI.Cell(iCol, iRow) = pRPI.NoData_Value
         ElseIf pDist2Ridge.Cell(iCol, iRow) = pDist2Ridge.NoData_Value Or pDist2Valley.Cell(iCol, iRow) = pDist2Valley.NoData_Value Then
            pRPI.Cell(iCol, iRow) = pRPI.NoData_Value
         ElseIf pDist2Valley.Cell(iCol, iRow) + pDist2Ridge.Cell(iCol, iRow) = 0# Then
            pRPI.Cell(iCol, iRow) = 1#
         Else
            pRPI.Cell(iCol, iRow) = pDist2Valley.Cell(iCol, iRow) / (pDist2Valley.Cell(iCol, iRow) + pDist2Ridge.Cell(iCol, iRow))
         End If
      Next
   Next
   
   RelativePositionIndex_KeepRouting0 = True
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

''
' Relative relief:
'        Pij = (Relief to the nearest valley) / (Relief to the nearest valley + Relief to the nearest ridge)
' P.S.   the nearest valley: not higher than the interest cell;
'        the nearest ridge: not lower than the interest cell.
' '
Public Function RelativeRelief(pDEM As clsGrid, pRidge As clsGrid, pChannel As clsGrid, pRRI As clsGrid, _
      pRlf2Ridge As clsGrid, pRlf2Valley As clsGrid, _
      Optional iRidgeTag As Integer = 1, Optional iChannelTag As Integer = 1) As Boolean
On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dNoData As Double
   'Dim vDist2Ridge As Variant, vDist2Valley As Variant   'vTemp As Variant
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, iCol2 As Integer, iRow2 As Integer
   Dim dCellSize As Double, dDist As Double, iDistUnit As Integer, dTemp As Double, iDistUnit1 As Integer, iDistUnit2 As Integer, iDistUnit0 As Integer
   Dim iVlyRow As Integer, iVlyCol As Integer, iRdgRow As Integer, iRdgCol As Integer
   Dim dElev As Double
         
   RelativeRelief = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   If iCols <> pRidge.nCols Or iRows <> pRidge.nRows Or dCellSize <> pRidge.CellSize _
         Or iCols <> pChannel.nCols Or iRows <> pChannel.nRows Or dCellSize <> pChannel.CellSize Then
      '   Or pRidge.xllcorner <> pChannel.xllcorner Or pRidge.yllcorner <> pChannel.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID DEM, Ridge and Channel should be with same position and same size."
   End If
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pChannel.Cell(iCol, iRow) = iChannelTag Then GoTo HasValley
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Valley-cell in given Valley GRID"
HasValley:
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then GoTo HasRidge
      Next
   Next
   Err.Raise Number:=vbObjectError + 513, Description:="No Ridge-cell in given Ridge GRID"
HasRidge:
   DoEvents
   
   ' initialize   ' pSizeIn==pSize
   'ReDim vDist2Ridge(0 To iCols - 1, 0 To iRows - 1)
   'ReDim vDist2Valley(0 To iCols - 1, 0 To iRows - 1)
      
   ' compute the distance from the every pixel to the nearest valley pixel
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         dDist = MAX_SINGLE
         If pChannel.Cell(iCol, iRow) = iChannelTag Then
            pRlf2Valley.Cell(iCol, iRow) = 0# 'vDist2Valley(iCol, iRow) = 0#
            dDist = 0#
         ElseIf pDEM.Cell(iCol, iRow) = dNoData Then
            pRlf2Valley.Cell(iCol, iRow) = pRlf2Valley.NoData_Value
         Else
            dElev = pDEM.Cell(iCol, iRow)
            iVlyRow = iRow: iVlyCol = iCol
            iDistUnit0 = 0
            Do
               iDistUnit0 = iDistUnit0 + 1
               iDistUnit = iDistUnit0
               iCol2 = iCol - iDistUnit
               If iCol2 >= 0 Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iVlyRow = iRow2: iVlyCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               iCol2 = iCol + iDistUnit
               If iCol2 < iCols Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iVlyRow = iRow2: iVlyCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow - iDistUnit
               If iRow2 >= 0 Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iVlyRow = iRow2: iVlyCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow + iDistUnit
               If iRow2 < iRows Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iVlyRow = iRow2: iVlyCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               
               If dDist < MAX_SINGLE Then
                  iDistUnit1 = iDistUnit + 1
                  iDistUnit2 = Int(dDist) + 1 ' Int(Sqr(2) * iDistUnit) + 1
                  For iDistUnit = iDistUnit1 To iDistUnit2
                     iCol2 = iCol - iDistUnit
                     If iCol2 >= 0 Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iVlyRow = iRow2: iVlyCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iCol2 = iCol + iDistUnit
                     If iCol2 < iCols Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iVlyRow = iRow2: iVlyCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow - iDistUnit
                     If iRow2 >= 0 Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iVlyRow = iRow2: iVlyCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow + iDistUnit
                     If iRow2 < iRows Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pChannel.Cell(iCol2, iRow2) = iChannelTag) And pDEM.Cell(iCol2, iRow2) <= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iVlyRow = iRow2: iVlyCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                  Next
               End If
            Loop Until (dDist < MAX_SINGLE) Or (iDistUnit0 >= iRows And iDistUnit0 >= iCols)
            
            If dDist = MAX_SINGLE Then
               pRlf2Valley.Cell(iCol, iRow) = pRlf2Valley.NoData_Value
            ElseIf dDist = 0# Then
               pRlf2Valley.Cell(iCol, iRow) = 0#
            Else
               pRlf2Valley.Cell(iCol, iRow) = pDEM.Cell(iCol, iRow) - pDEM.Cell(iVlyCol, iVlyRow) 'dDist * dCellSize
            End If
         End If
      Next
   Next
   DoEvents
   
   ' compute the distance from the every pixel to the nearest Ridge pixel
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         dDist = MAX_SINGLE
         If pRidge.Cell(iCol, iRow) = iRidgeTag Then
            pRlf2Ridge.Cell(iCol, iRow) = 0#
            dDist = 0#
         ElseIf pDEM.Cell(iCol, iRow) = dNoData Then
            pRlf2Ridge.Cell(iCol, iRow) = pRlf2Ridge.NoData_Value
         Else
            dElev = pDEM.Cell(iCol, iRow)
            iRdgRow = iRow: iRdgCol = iCol
            iDistUnit0 = 0
            Do
               iDistUnit0 = iDistUnit0 + 1
               iDistUnit = iDistUnit0
               iCol2 = iCol - iDistUnit
               If iCol2 >= 0 Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iRdgRow = iRow2: iRdgCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               iCol2 = iCol + iDistUnit
               If iCol2 < iCols Then
                  For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                     If iRow2 >= 0 And iRow2 < iRows Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iRdgRow = iRow2: iRdgCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow - iDistUnit
               If iRow2 >= 0 Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iRdgRow = iRow2: iRdgCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               iRow2 = iRow + iDistUnit
               If iRow2 < iRows Then
                  For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                     If iCol2 >= 0 And iCol2 < iCols Then
                        If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                           dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                           If dTemp < dDist Then
                              dDist = dTemp
                              iRdgRow = iRow2: iRdgCol = iCol2
                           End If
                        End If
                     End If
                  Next
               End If
               
               If dDist < MAX_SINGLE Then
                  iDistUnit1 = iDistUnit + 1
                  iDistUnit2 = Int(dDist) + 1 ' Int(Sqr(2) * iDistUnit) + 1
                  For iDistUnit = iDistUnit1 To iDistUnit2
                     iCol2 = iCol - iDistUnit
                     If iCol2 >= 0 Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iRdgRow = iRow2: iRdgCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iCol2 = iCol + iDistUnit
                     If iCol2 < iCols Then
                        For iRow2 = iRow - iDistUnit To iRow + iDistUnit
                           If iRow2 >= 0 And iRow2 < iRows Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iRow2 - iRow) ^ 2 + (iDistUnit) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iRdgRow = iRow2: iRdgCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow - iDistUnit
                     If iRow2 >= 0 Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iRdgRow = iRow2: iRdgCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                     iRow2 = iRow + iDistUnit
                     If iRow2 < iRows Then
                        For iCol2 = iCol - (iDistUnit - 1) To iCol + (iDistUnit - 1)
                           If iCol2 >= 0 And iCol2 < iCols Then
                              If (pRidge.Cell(iCol2, iRow2) = iRidgeTag) And pDEM.Cell(iCol2, iRow2) >= dElev Then
                                 dTemp = Sqr((iDistUnit) ^ 2 + (iCol2 - iCol) ^ 2)
                                 If dTemp < dDist Then
                                    dDist = dTemp
                                    iRdgRow = iRow2: iRdgCol = iCol2
                                 End If
                              End If
                           End If
                        Next
                     End If
                  Next
               End If
            Loop Until (dDist < MAX_SINGLE) Or (iDistUnit0 >= iRows And iDistUnit0 >= iCols)
            
            If dDist = MAX_SINGLE Then
               pRlf2Ridge.Cell(iCol, iRow) = pRlf2Ridge.NoData_Value
            ElseIf dDist = 0# Then
               pRlf2Ridge.Cell(iCol, iRow) = 0#
            Else
               pRlf2Ridge.Cell(iCol, iRow) = pDEM.Cell(iRdgCol, iRdgRow) - pDEM.Cell(iCol, iRow) 'dDist * dCellSize
            End If
         End If
      Next
   Next
   DoEvents
   
   ' RelativeRelief = (Relief to the nearest valley) / (Relief to the nearest valley + Relief to nearest ridge)
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pRlf2Ridge.Cell(iCol, iRow) = 0# Then
            pRRI.Cell(iCol, iRow) = 1#
         ElseIf pRlf2Valley.Cell(iCol, iRow) = 0# Then
            pRRI.Cell(iCol, iRow) = 0#
'         ElseIf vDist2Ridge(iCol, iRow) = MAX_SINGLE Or vDist2Valley(iCol, iRow) = MAX_SINGLE Then
'            pRRI.Cell(iCol, iRow) = pRRI.NoData_Value
         ElseIf pRlf2Ridge.Cell(iCol, iRow) = pRlf2Ridge.NoData_Value Or pRlf2Valley.Cell(iCol, iRow) = pRlf2Valley.NoData_Value Then
            pRRI.Cell(iCol, iRow) = pRRI.NoData_Value
         ElseIf pRlf2Valley.Cell(iCol, iRow) + pRlf2Ridge.Cell(iCol, iRow) = 0 Then
            pRRI.Cell(iCol, iRow) = 1#
         Else
            pRRI.Cell(iCol, iRow) = pRlf2Valley.Cell(iCol, iRow) / (pRlf2Valley.Cell(iCol, iRow) + pRlf2Ridge.Cell(iCol, iRow))
         End If
      Next
   Next
   
   RelativeRelief = True
ErrH:
   ' Release memeory
   'vDist2Ridge = Empty:   vDist2Valley = Empty
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Schmidt & Hewitt, 2004; Rodriguez et al, 2002
' function: compute Hill index, Valley index, and Hillslope index (membership grid) by TOP HAT approach
'--Original implementation: Very slow
'
Public Function Terrain_TOPHAT(pDEM As clsGrid, iHalfWinCells As Integer, dElevThreshold As Double, _
      pHillI As clsGrid, pHillslopeI As clsGrid, pValleyI As clsGrid) As Boolean
On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim vMin As Variant, vMax As Variant, vHH As Variant, vVD As Variant
   Dim dMin As Double, dMax As Double, dElev As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, iDeltaCol As Integer, iDeltaRow As Integer
   
   Terrain_TOPHAT = False
   'iHalfWinCells = (Int(135 / dCellX) - 1) / 2
   'dElevThreshold = 0.05
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   ' initialize
   ReDim vHH(0 To iCols - 1, 0 To iRows - 1)
   ReDim vVD(0 To iCols - 1, 0 To iRows - 1)
   ReDim vMin(0 To iCols - 1, 0 To iRows - 1)
   ReDim vMax(0 To iCols - 1, 0 To iRows - 1)
   
   'Opening: max(min(DEM)); Closing: min(max(DEM))
   ' peak cut = DTM - Opening; valleys filled= Closing  - DTM
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (pDEM.Cell(iCol, iRow) = dNoData) Then
            vMin(iCol, iRow) = dNoData:              vMax(iCol, iRow) = dNoData
         Else
            dMin = MAX_SINGLE:   dMax = MIN_SINGLE
            For iDeltaCol = -iHalfWinCells To iHalfWinCells
               iCol1 = iCol + iDeltaCol
               For iDeltaRow = -iHalfWinCells To iHalfWinCells
                  iRow1 = iRow + iDeltaRow
                  If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If dElev > dMax Then dMax = dElev
                     If dElev < dMin Then dMin = dElev
                  End If
               Next
            Next
            vMin(iCol, iRow) = dMin: vMax(iCol, iRow) = dMax
         End If
      Next
   Next
   DoEvents
   
   ' peaks cut
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (vMin(iCol, iRow) = dNoData) Then
            vHH(iCol, iRow) = dNoData
         Else
            dMax = MIN_SINGLE
            For iDeltaCol = -iHalfWinCells To iHalfWinCells
               iCol1 = iCol + iDeltaCol
               For iDeltaRow = -iHalfWinCells To iHalfWinCells
                  iRow1 = iRow + iDeltaRow
                  If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If vMin(iCol1, iRow1) > dMax Then dMax = vMin(iCol1, iRow1)
                  End If
               Next
            Next
            vHH(iCol, iRow) = pDEM.Cell(iCol, iRow) - dMax
         End If
      Next
   Next
   DoEvents
   
   ' valleys filled
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (vMax(iCol, iRow) = dNoData) Then
            vVD(iCol, iRow) = dNoData
         Else
            dMin = MAX_SINGLE
            For iDeltaCol = -iHalfWinCells To iHalfWinCells
               iCol1 = iCol + iDeltaCol
               For iDeltaRow = -iHalfWinCells To iHalfWinCells
                  iRow1 = iRow + iDeltaRow
                  If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If vMax(iCol1, iRow1) < dMin Then dMin = vMax(iCol1, iRow1)
                  End If
               Next
            Next
            vVD(iCol, iRow) = dMin - pDEM.Cell(iCol, iRow)
         End If
      Next
   Next
   DoEvents
   
   ' adjust by delevthreshold
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If vHH(iCol, iRow) <> dNoData Then
            If vHH(iCol, iRow) < dElevThreshold Then vHH(iCol, iRow) = 0#
         End If
         If vVD(iCol, iRow) <> dNoData Then
            If vVD(iCol, iRow) < dElevThreshold Then vVD(iCol, iRow) = 0#
         End If
      Next
   Next
   vMin = Empty: vMax = Empty
   
   ' hill index
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If vHH(iCol, iRow) = dNoData Then
            pHillI.Cell(iCol, iRow) = pHillI.NoData_Value
         Else
            If vHH(iCol, iRow) = 0 Then
               pHillI.Cell(iCol, iRow) = 0#
            Else
               If vVD(iCol, iRow) = 0 Then
                     pHillI.Cell(iCol, iRow) = 1#
                  Else
                     pHillI.Cell(iCol, iRow) = vHH(iCol, iRow) / (vHH(iCol, iRow) + vVD(iCol, iRow))
                  End If
            End If
         End If
      Next
   Next
     
   ' valley index
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If vVD(iCol, iRow) = dNoData Then
            pValleyI.Cell(iCol, iRow) = pValleyI.NoData_Value
         Else
            If vVD(iCol, iRow) = 0 Then
               pValleyI.Cell(iCol, iRow) = 0#
            Else
               If vHH(iCol, iRow) = 0 Then
                  pValleyI.Cell(iCol, iRow) = 1#
               Else
                  pValleyI.Cell(iCol, iRow) = vVD(iCol, iRow) / (vHH(iCol, iRow) + vVD(iCol, iRow))
               End If
            End If
         End If
      Next
   Next
   
   ' hillslope index
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If vHH(iCol, iRow) = dNoData Then
            pHillslopeI.Cell(iCol, iRow) = pHillslopeI.NoData_Value
         Else
            If vHH(iCol, iRow) = 0 And vVD(iCol, iRow) = 0 Then
               pHillslopeI.Cell(iCol, iRow) = 1#
            Else
               If vHH(iCol, iRow) = 0 Or vVD(iCol, iRow) = 0 Then
                  pHillslopeI.Cell(iCol, iRow) = 0#
               Else
                  dMin = IIf(vHH(iCol, iRow) < vVD(iCol, iRow), vHH(iCol, iRow), vVD(iCol, iRow))
                  pHillslopeI.Cell(iCol, iRow) = 2 * dMin / (vHH(iCol, iRow) + vVD(iCol, iRow))
               End If
            End If
         End If
      Next
   Next
      
   Terrain_TOPHAT = True
ErrH:
   ' Release memeory
   vMin = Empty: vMax = Empty
   vHH = Empty:  vVD = Empty: vMin = Empty: vMax = Empty
   If Err.Number > 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
'  Terrain Char. Index  (TCI= Cs * lg(As) )(Park and van de Giesen, 2004)
'
Public Function TerrainCharI(pSCA As clsGrid, pCs As clsGrid, pTCI As clsGrid) As Boolean
   Dim iCol As Integer, iRow As Integer
On Error GoTo ErrH
   TerrainCharI = False
   If pSCA.nCols <> pCs.nCols Or pSCA.nRows <> pCs.nRows Or pSCA.CellSize <> pCs.CellSize _
         Or pSCA.xllcorner <> pCs.xllcorner Or pSCA.yllcorner <> pCs.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID SCA and Cs should be with same position and same size."
   End If
   
   For iRow = 0 To pSCA.nRows - 1
      For iCol = 0 To pSCA.nCols - 1
         If pSCA.Cell(iCol, iRow) = pSCA.NoData_Value Or pCs.Cell(iCol, iRow) = pCs.NoData_Value Then
            pTCI.Cell(iCol, iRow) = pTCI.NoData_Value
         ElseIf pSCA.Cell(iCol, iRow) <= 0# Then
            pTCI.Cell(iCol, iRow) = pTCI.NoData_Value
         Else
            pTCI.Cell(iCol, iRow) = pCs.Cell(iCol, iRow) * Log10(pSCA.Cell(iCol, iRow))
         End If
      Next
   Next
   TerrainCharI = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


