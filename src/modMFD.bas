Attribute VB_Name = "modMFD"
Option Explicit
Option Base 1  ' index of array begins from 1

Dim marrPara As Variant, marrMFFractSum As Variant
Dim mpAccumCells As clsGrid, mpDEM As clsGrid
Dim P0 As Double, P_range As Double

Dim mpFlowLen As clsGrid, mpTag As clsGrid
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Planchon & Darboux, 2001
' function: fill depressions in DEM and replace it with a surface either strictly horizontal (used for calculation of depression storage capacity)
'                 , or slightly sloping (used for drainage network extraction)
'  ways: first inundate the surface with a thick layer of water, then remove the excess water
'
Public Function FillDep_RemoveExcessWater_Planchon01(pDEM0 As clsGrid, pNewDEM As clsGrid, Optional dDeltaElev As Double = 0.01) As Boolean
   Const Value4NoData = -1000#
On Error GoTo ErrH
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer, iScan As Integer
   Dim RowFrom As Variant, ColFrom As Variant, dRow As Variant, dCol As Variant, RowTo As Variant, ColTo As Variant
   Dim boolSomeDone As Boolean, next_cell As Boolean
   Dim dElevEps As Double, dElevEps0 As Double, dElevEps1 As Double
   Dim pDEM As New clsGrid
   
   FillDep_RemoveExcessWater_Planchon01 = False
   
   ' initialize
   With pDEM0
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   RowFrom = Array(0, iRows - 1, 0, iRows - 1, 0, iRows - 1, 0, iRows - 1)
   ColFrom = Array(0, iCols - 1, iCols - 1, 0, iCols - 1, 0, 0, iCols - 1)
   dRow = Array(0, 0, 1, -1, 0, 0, 1, -1)
   dCol = Array(1, -1, 0, 0, -1, 1, 0, 0)
   RowTo = Array(1, -1, -iRows + 1, iRows - 1, 1, -1, -iRows + 1, iRows - 1)
   ColTo = Array(-iCols + 1, iCols - 1, -1, 1, iCols - 1, -iCols + 1, 1, -1)
   dElevEps0 = dDeltaElev
   dElevEps1 = dDeltaElev * Sqr(2#)
      
   ' Process NoData cells in original DEM
   pDEM.NewGrid iCols, iRows, pDEM0.xllcorner, pDEM0.yllcorner, dCellSize, dNoData
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         pDEM.Cell(iCol, iRow) = IIf(pDEM0.Cell(iCol, iRow) = dNoData, Value4NoData, pDEM0.Cell(iCol, iRow))
      Next
   Next
   
   'Stage 1. Initialisation of the surface to infinite altitude
   For iRow = 0 To iRows - 1
      pNewDEM.Cell(0, iRow) = pDEM.Cell(0, iRow): pNewDEM.Cell(iCols - 1, iRow) = pDEM.Cell(iCols - 1, iRow)
   Next
   For iCol = 1 To iCols - 2
      pNewDEM.Cell(iCol, 0) = pDEM.Cell(iCol, 0): pNewDEM.Cell(iCol, iRows - 1) = pDEM.Cell(iCol, iRows - 1)
      For iRow = 1 To iRows - 2
         If (pDEM.Cell(iCol, iRow) = dNoData) Then
            pNewDEM.Cell(iCol, iRow) = pNewDEM.NoData_Value
         Else
            pNewDEM.Cell(iCol, iRow) = MAX_SINGLE
         End If
      Next
   Next
   
   ' Stage 2. Removal of excess water
   '                Operarion (1)   Z(c)>=W(n)+eps(c,n) ==> W(c)=Z(c)
   '                Operarion (2)   W(c)>W(n)+eps(c,n) ==> W(c)=W(n)+eps(c,n)
   '=============================
   '  for each cell c of DEM
   '     for each neighbour n of c
   '        determine eps for the pair (c,n)
   '        if possible, apply operation (1)
   '        else try to apply operation (2)
   '     next
   '  next
   '  if W was modified during this iScan, then go on loop
   '
   For iRow = 0 To iRows - 1
      If (pDEM.Cell(0, iRow) = dNoData) Then
         Dry_Upward_Cell 0, iRow, pDEM, pNewDEM, dElevEps0, dElevEps1
      End If
   Next
   For iRow = 0 To iRows - 1
      If (pDEM.Cell(iCols - 1, iRow) = dNoData) Then
         Dry_Upward_Cell iCols - 1, iRow, pDEM, pNewDEM, dElevEps0, dElevEps1
      End If
   Next
   For iCol = 1 To iCols - 2
      If (pDEM.Cell(iCol, 0) = dNoData) Then
         Dry_Upward_Cell iCol, 0, pDEM, pNewDEM, dElevEps0, dElevEps1
      End If
   Next
   For iCol = 1 To iCols - 2
      If (pDEM.Cell(iCol, iRows - 1) = dNoData) Then
         Dry_Upward_Cell iCol, iRows - 1, pDEM, pNewDEM, dElevEps0, dElevEps1
      End If
   Next
   
IterativeScan:
   For iScan = 1 To 8
      iCol = ColFrom(iScan): iRow = RowFrom(iScan)
      boolSomeDone = False
      If (pNewDEM.Cell(iCol, iRow) <> dNoData) Then
         Do
            If pNewDEM.Cell(iCol, iRow) > pDEM.Cell(iCol, iRow) Then
               For k = 1 To DIRNUM8
                  iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                  dElevEps = IIf(k Mod 2 = 1, dElevEps1, dElevEps0)
                  If pDEM.IsValidCellValue(iCol1, iRow1) Then
                     ' c(iCol,iRow)  n(iCol1,iRow1)
                     If pDEM.Cell(iCol, iRow) >= pNewDEM.Cell(iCol1, iRow1) + dElevEps Then
                        pNewDEM.Cell(iCol, iRow) = pDEM.Cell(iCol, iRow)
                        boolSomeDone = True
                        Dry_Upward_Cell iCol, iRow, pDEM, pNewDEM, dElevEps0, dElevEps1
                        GoTo ProcNextCell
                     End If
                     If pNewDEM.Cell(iCol, iRow) > pNewDEM.Cell(iCol1, iRow1) + dElevEps Then
                        pNewDEM.Cell(iCol, iRow) = pNewDEM.Cell(iCol1, iRow1) + dElevEps
                        boolSomeDone = True
                     End If
                  End If
               Next
            End If
ProcNextCell:
            next_cell = True
            iRow = iRow + dRow(iScan): iCol = iCol + dCol(iScan)
            If iRow < 0 Or iCol < 0 Or iRow >= iRows Or iCol >= iCols Then
               iRow = iRow + RowTo(iScan): iCol = iCol + ColTo(iScan)
               If iRow < 0 Or iCol < 0 Or iRow >= iRows Or iCol >= iCols Then
                  next_cell = False
               End If
            End If
         Loop While next_cell
      
         If Not boolSomeDone Then Exit For
      End If
   Next
   If boolSomeDone Then GoTo IterativeScan
   're-assign Nodata to cells which are NoData in original DEM
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If pDEM0.Cell(iCol, iRow) = dNoData Then
            pNewDEM.Cell(iCol, iRow) = pNewDEM.NoData_Value
         End If
      Next
   Next
   
   FillDep_RemoveExcessWater_Planchon01 = True
ErrH:
   Set pDEM = Nothing
   RowFrom = Empty: ColFrom = Empty: dRow = Empty: dCol = Empty: RowTo = Empty: ColTo = Empty
   If Err.Number > 0 Then MsgBox Err.Description, vbExclamation
End Function

Private Sub Dry_Upward_Cell(iCol As Integer, iRow As Integer, pDEM As clsGrid, pNewDEM As clsGrid, dElevEps0 As Double, dElevEps1 As Double)
   Const MAX_DEPTH = 4000
   Static iDepth As Integer    ' initialized as 0
   Dim iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim dValue As Double, dElevEps As Double
   
   iDepth = iDepth + 1
   If iDepth < MAX_DEPTH Then
      For k = 1 To DIRNUM8
         iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
         If pDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
            If pNewDEM.Cell(iCol1, iRow1) = MAX_SINGLE Then
               dElevEps = IIf(k Mod 2 = 1, dElevEps1, dElevEps0)
               If dValue >= pNewDEM.Cell(iCol, iRow) + dElevEps Then
                  pNewDEM.Cell(iCol1, iRow1) = dValue
                  Dry_Upward_Cell iCol1, iRow1, pDEM, pNewDEM, dElevEps0, dElevEps1
               End If
            End If
         End If
      Next
   End If
   iDepth = iDepth - 1
End Sub

''
'  Multiple flow direction flow length algorithm (Bogaart and Troch, 2006):
'     FlowDist_central = Sum(MulDirFlow_Fract_neighbor * (FlowDist_DownslpNeighbor + Dist_between_Cells))
'
Public Function FlowLen_by_MFD(pDEM As clsGrid, pTag As clsGrid, pFlowLen As clsGrid, sMFDAlg As String, _
                              Optional dP0 As Double = 1#, Optional dP_range As Double = 8.9) As Boolean

   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim dValue As Double, dElev As Double, dSum As Double, dSlope As Double, dMax As Double
   Dim tanb_LB As Double, tanb_UB As Double, a As Double, b As Double
   
On Error GoTo ErrH
   FlowLen_by_MFD = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   Set mpDEM = pDEM
   Set mpFlowLen = pFlowLen
   Set mpTag = pTag
   
   If sMFDAlg = FUNC_TYPE_MFD_QUINN91 Then
      ReDim marrMFFractSum(0 To iCols - 1, 0 To iRows - 1)
      P0 = dP0
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            If (mpDEM.Cell(iCol, iRow) = dNoData) Then
               marrMFFractSum(iCol, iRow) = 0#
            Else
               dSum = 0#
               For k = 1 To DIRNUM8
                  iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                  If mpDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                     If mpDEM.Cell(iCol, iRow) > dValue Then
                        'dSlope = 0.5 * (mpDEM.Cell(iCol, iRow) - dValue) / dCellSize
                        dSlope = (mpDEM.Cell(iCol, iRow) - dValue) / dCellSize
                        If k Mod 2 = 1 Then
                           dSlope = dSlope / SQRT2
                           dSum = dSum + (dSlope ^ P0) * SQRT2 / 4
                        Else
                           dSum = dSum + (dSlope ^ P0) / 2
                        End If
                     End If
                  End If
               Next
               marrMFFractSum(iCol, iRow) = IIf(dSum = 0#, MAX_SINGLE, dSum)
            End If
         Next
      Next
      
      '''''''
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            If (mpDEM.Cell(iCol, iRow) <> dNoData) And Not mpTag.Cell(iCol, iRow) Then
               FlowLen_CheckNeighbor iCol, iRow, True
            End If
         Next
      Next
      
   Else  ' FUNC_TYPE_MFD_QIN07 ' based on MFD-md
      tanb_UB = 1#:  tanb_LB = 0#
      P0 = dP0:      P_range = dP_range   '10 - P0
      a = P_range / (tanb_UB - tanb_LB)     'P_range / (Atn(tanb_UB) - Atn(tanb_LB)) '
      b = P0 - P_range * tanb_LB / (tanb_UB - tanb_LB)      'P0 - P_range * Atn(tanb_LB) / (Atn(tanb_UB) - Atn(tanb_LB))   '
      
      ReDim marrPara(0 To iCols - 1, 0 To iRows - 1)
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            If (mpDEM.Cell(iCol, iRow) <> dNoData) Then
               dMax = 0#
               For k = 1 To DIRNUM8
                  iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                  If mpDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If mpDEM.Cell(iCol, iRow) > dElev Then
                        dSlope = (mpDEM.Cell(iCol, iRow) - dElev) / dCellSize    ' tan(SlopeDegree): (icol, irow) -> k
                        If k Mod 2 = 1 Then dSlope = dSlope / SQRT2
                        If dSlope > dMax Then dMax = dSlope
                     End If
                  End If
               Next
               marrPara(iCol, iRow) = dMax
            Else
               marrPara(iCol, iRow) = 0#
            End If
         Next
      Next
      
      ''''''''''''''''''''''''''''''
      ReDim marrMFFractSum(0 To iCols - 1, 0 To iRows - 1)
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            If (mpDEM.Cell(iCol, iRow) = dNoData) Then
               marrMFFractSum(iCol, iRow) = 0#
            Else
               If marrPara(iCol, iRow) <= tanb_LB Then
                  marrPara(iCol, iRow) = P0
               ElseIf marrPara(iCol, iRow) >= tanb_UB Then
                  marrPara(iCol, iRow) = P0 + P_range
               Else
                  marrPara(iCol, iRow) = a * marrPara(iCol, iRow) + b 'a * Atn(marrPara(icol, irow)) + b '
               End If
               dSum = 0#
               For k = 1 To DIRNUM8
                  iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                  If mpDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                     If mpDEM.Cell(iCol, iRow) > dElev Then
                        dSlope = (mpDEM.Cell(iCol, iRow) - dElev) / dCellSize  ' tan(SlopeDegree): (icol, irow) -> k
                        If k Mod 2 = 1 Then
                           dSlope = dSlope / SQRT2
                           dSum = dSum + (dSlope ^ marrPara(iCol, iRow)) * SQRT2 / 4
                        Else
                           dSum = dSum + (dSlope ^ marrPara(iCol, iRow)) / 2
                        End If
                     End If
                  End If
               Next
               marrMFFractSum(iCol, iRow) = IIf(dSum = 0#, MAX_SINGLE, dSum) ' set as MAX_SINGLE when none downslope
            End If
         Next
      Next
      
      '''''''
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            If (mpDEM.Cell(iCol, iRow) <> dNoData) And Not mpTag.Cell(iCol, iRow) Then
               FlowLen_CheckNeighbor iCol, iRow, False
            End If
         Next
      Next
   End If
      
   FlowLen_by_MFD = True
ErrH:
   marrMFFractSum = Empty: marrPara = Empty
   Set mpFlowLen = Nothing:   Set mpTag = Nothing
   Set mpDEM = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
' Recursion function, call by FlowLen_by_MFD()
'
Private Function FlowLen_CheckNeighbor(iCol As Integer, iRow As Integer, Optional bMFD0_P0 As Boolean = True) As Double
   Dim iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim temp As Double, dSlope As Double, dElev As Double
   Dim dCellSize As Double
   
   On Error GoTo ErrH
   If mpFlowLen.Cell(iCol, iRow) = mpFlowLen.NoData_Value Then
      FlowLen_CheckNeighbor = 0#
      Exit Function
   End If
   
   If Not mpTag.Cell(iCol, iRow) Then
      dCellSize = mpDEM.CellSize
      mpFlowLen.Cell(iCol, iRow) = 0#
      For k = 1 To DIRNUM8
         iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
         If mpDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
            If dElev < mpDEM.Cell(iCol, iRow) Then
               If bMFD0_P0 Then
                  If marrMFFractSum(iCol, iRow) = MAX_SINGLE Then
                     temp = 0#
                  Else
                     dSlope = (mpDEM.Cell(iCol, iRow) - dElev) / dCellSize    ' tan(SlopeDegree): (iCol1, iRow1) -> k
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        temp = (dSlope ^ P0) * SQRT2 / (4 * marrMFFractSum(iCol, iRow))
                     Else
                        temp = (dSlope ^ P0) / (2 * marrMFFractSum(iCol, iRow))
                     End If
                  End If
               Else
                  If marrMFFractSum(iCol, iRow) = MAX_SINGLE Then
                     temp = 0#
                  Else
                     dSlope = (mpDEM.Cell(iCol, iRow) - dElev) / dCellSize ' tan(SlopeDegree): (iCol1, iRow1) -> k
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        temp = (dSlope ^ marrPara(iCol, iRow)) * SQRT2 / (4 * marrMFFractSum(iCol, iRow))
                     Else
                        temp = (dSlope ^ marrPara(iCol, iRow)) / (2 * marrMFFractSum(iCol, iRow))
                     End If
                  End If
               End If
               
               If temp > 0 Then
                  mpFlowLen.Cell(iCol, iRow) = mpFlowLen.Cell(iCol, iRow) _
                        + temp * (FlowLen_CheckNeighbor(iCol1, iRow1, bMFD0_P0) + IIf(k Mod 2 = 1, SQRT2 * dCellSize, dCellSize))
               End If
            End If
         End If
      Next
      
      mpTag.Cell(iCol, iRow) = True
   End If
   
   FlowLen_CheckNeighbor = mpFlowLen.Cell(iCol, iRow)
   Exit Function
ErrH:
   If Err.Number > 0 Then MsgBox Err.Description, vbExclamation, "modMFD.FlowLen_CheckNeighbor"
End Function


'
'  MFD-Quinn (Quinn et al., 1991)
'
Public Function FlowAccumulation_MFD_Quinn(pDEM As clsGrid, pAccumCells As clsGrid, Optional dP0 As Double = 1#) As Boolean
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim dValue As Double, dSum As Double, dSlope As Double
On Error GoTo ErrH
   FlowAccumulation_MFD_Quinn = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   Set mpDEM = pDEM
   Set mpAccumCells = pAccumCells
      
   ReDim marrMFFractSum(0 To iCols - 1, 0 To iRows - 1)
   P0 = dP0 ' 1#
   ' Case I_MFD_Freeman, I_MFD_Quinn
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (mpDEM.Cell(iCol, iRow) = dNoData) Then
            marrMFFractSum(iCol, iRow) = 0#
         Else
            dSum = 0#
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If mpDEM.IsValidCellValue(iCol1, iRow1, dValue) Then
                  If mpDEM.Cell(iCol, iRow) > dValue Then
                     'dSlope = 0.5 * (mpDEM.Cell(iCol, iRow) - dValue) / dCellSize
                     dSlope = (mpDEM.Cell(iCol, iRow) - dValue) / dCellSize
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        dSum = dSum + (dSlope ^ P0) * SQRT2 / 4
                     Else
                        dSum = dSum + (dSlope ^ P0) / 2
                     End If
                  End If
               End If
            Next
            marrMFFractSum(iCol, iRow) = IIf(dSum = 0#, MAX_SINGLE, dSum)
         End If
      Next
   Next
   '''''''
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         mpAccumCells.Cell(iCol, iRow) = IIf(mpDEM.Cell(iCol, iRow) = dNoData, mpAccumCells.NoData_Value, 0#)
      Next
   Next
             
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (mpDEM.Cell(iCol, iRow) <> dNoData) Then
            CheckNeighbor iCol, iRow, True
         End If
      Next
   Next
   
   FlowAccumulation_MFD_Quinn = True
ErrH:
   marrMFFractSum = Empty
   Set mpAccumCells = Nothing
   Set mpDEM = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'
'  MFD-md (Qin et al., 2007)
'
Public Function FlowAccumulation_MFD_md(pDEM As clsGrid, pAccumCells As clsGrid, _
            Optional dP0 As Double = 1.1, Optional dP_range As Double = 8.9, _
            Optional tanb_LB As Double = 0#, Optional tanb_UB = 1#) As Boolean
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim dElev As Double, dSum As Double, dSlope As Double, dMax As Double
   Dim a As Double, b As Double
   
On Error GoTo ErrH
   FlowAccumulation_MFD_md = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   Set mpDEM = pDEM
   Set mpAccumCells = pAccumCells
   P0 = dP0
   P_range = dP_range
   a = P_range / (tanb_UB - tanb_LB)   'P_range / (Atn(tanb_UB) - Atn(tanb_LB)) '
   b = P0 - P_range * tanb_LB / (tanb_UB - tanb_LB)   'P0 - P_range * Atn(tanb_LB) / (Atn(tanb_UB) - Atn(tanb_LB))   '
         
   '''''''''''''''''''''''''
   ReDim marrPara(0 To iCols - 1, 0 To iRows - 1)
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (mpDEM.Cell(iCol, iRow) <> dNoData) Then
            dMax = 0#
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If mpDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                  If mpDEM.Cell(iCol, iRow) > dElev Then
                     dSlope = (mpDEM.Cell(iCol, iRow) - dElev) / dCellSize    ' tan(SlopeDegree): (icol, irow) -> k
                     If k Mod 2 = 1 Then dSlope = dSlope / SQRT2
                     If dSlope > dMax Then dMax = dSlope
                  End If
               End If
            Next
            marrPara(iCol, iRow) = dMax
         Else
            marrPara(iCol, iRow) = 0#
         End If
      Next
   Next
   
   ''''''''''''''''''''''''''''''
   ReDim marrMFFractSum(0 To iCols - 1, 0 To iRows - 1)
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (mpDEM.Cell(iCol, iRow) = dNoData) Then
            marrMFFractSum(iCol, iRow) = 0#
         Else
            If marrPara(iCol, iRow) <= tanb_LB Then
               marrPara(iCol, iRow) = P0
            ElseIf marrPara(iCol, iRow) >= tanb_UB Then
               marrPara(iCol, iRow) = P0 + P_range
            Else
               marrPara(iCol, iRow) = a * marrPara(iCol, iRow) + b 'a * Atn(marrPara(icol, irow)) + b '
            End If
            dSum = 0#
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If mpDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                  If mpDEM.Cell(iCol, iRow) > dElev Then
                     dSlope = (mpDEM.Cell(iCol, iRow) - dElev) / dCellSize  ' tan(SlopeDegree): (icol, irow) -> k
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        dSum = dSum + (dSlope ^ marrPara(iCol, iRow)) * SQRT2 / 4
                     Else
                        dSum = dSum + (dSlope ^ marrPara(iCol, iRow)) / 2
                     End If
                  End If
               End If
            Next
            marrMFFractSum(iCol, iRow) = IIf(dSum = 0#, MAX_SINGLE, dSum) ' set as MAX_SINGLE when none downslope
         End If
      Next
   Next
   
   ''''''''''''''''''''''
   For iCol = 0 To iCols - 1
      For iRow = 0 To iRows - 1
         mpAccumCells.Cell(iCol, iRow) = IIf(mpDEM.Cell(iCol, iRow) = dNoData, mpAccumCells.NoData_Value, 0#)
      Next
   Next
             
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (mpDEM.Cell(iCol, iRow) <> dNoData) Then
            CheckNeighbor iCol, iRow, False
         End If
      Next
   Next
   
   FlowAccumulation_MFD_md = True
ErrH:
   marrMFFractSum = Empty
   marrPara = Empty
   Set mpAccumCells = Nothing
   Set mpDEM = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
' Recursion function, to do depth-first search for computing catchment from multiple flow direction algorithm
' global variate m_FlowDistribType determined which flow distribution type is used (i.e. call which Fract_Flow_BySlope function)
'
Private Function CheckNeighbor(iCol As Integer, iRow As Integer, Optional bMFD0_P0 As Boolean = True) As Double
   Dim iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim temp As Double, dSlope As Double, dElev As Double
   
   On Error GoTo ErrH
   If mpAccumCells.Cell(iCol, iRow) = mpAccumCells.NoData_Value Then
      CheckNeighbor = 0#
      Exit Function
   End If
   
   If mpAccumCells.Cell(iCol, iRow) <= 0# Then
      mpAccumCells.Cell(iCol, iRow) = 1#
      For k = 1 To DIRNUM8
         iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
         If mpDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
            If dElev > mpDEM.Cell(iCol, iRow) Then
               If bMFD0_P0 Then
                  If marrMFFractSum(iCol1, iRow1) = MAX_SINGLE Then
                     temp = 0#
                  Else
                     dSlope = (dElev - mpDEM.Cell(iCol, iRow)) / mpDEM.CellSize    ' tan(SlopeDegree): (iCol1, iRow1) -> k
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        temp = (dSlope ^ P0) * SQRT2 / (4 * marrMFFractSum(iCol1, iRow1))
                     Else
                        temp = (dSlope ^ P0) / (2 * marrMFFractSum(iCol1, iRow1))
                     End If
                  End If
               Else
                  If marrMFFractSum(iCol1, iRow1) = MAX_SINGLE Then
                     temp = 0#
                  Else
                     dSlope = (dElev - mpDEM.Cell(iCol, iRow)) / mpDEM.CellSize   ' tan(SlopeDegree): (iCol1, iRow1) -> k
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        temp = (dSlope ^ marrPara(iCol1, iRow1)) * SQRT2 / (4 * marrMFFractSum(iCol1, iRow1))
                     Else
                        temp = (dSlope ^ marrPara(iCol1, iRow1)) / (2 * marrMFFractSum(iCol1, iRow1))
                     End If
                  End If
               End If
               
               If temp > 0 Then
                  mpAccumCells.Cell(iCol, iRow) = mpAccumCells.Cell(iCol, iRow) + temp * CheckNeighbor(iCol1, iRow1, bMFD0_P0)
               End If
            End If
         End If
      Next
   End If
   CheckNeighbor = mpAccumCells.Cell(iCol, iRow)
   Exit Function
ErrH:
   If Err.Number > 0 Then MsgBox Err.Description, vbExclamation, "modMFD.CheckNeighbor"
End Function

'
'  Specific Catchment Area = Flow Accumulation in cells * Cellsize
' i.e. flow accumulation area /contour length (i.e., grid size, horizontal resolution)
'
Public Function SpecificCatchmentArea(pAccumCells As clsGrid, pSCA As clsGrid, _
                                       Optional sContourLen As String = FUNC_TYPE_EffectContourLen_Cell, _
                                       Optional pDEM As clsGrid = Nothing) As Boolean
                                       
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer
   Dim k As Integer, iCol1 As Integer, iRow1 As Integer, dSumEffectContour As Double, dElev As Double, dElev1 As Double
   
On Error GoTo ErrH
   SpecificCatchmentArea = False
   With pAccumCells
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   Select Case sContourLen
   Case FUNC_TYPE_EffectContourLen_Cell
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            If pAccumCells.Cell(iCol, iRow) = dNoData Then
               pSCA.Cell(iCol, iRow) = pSCA.NoData_Value
            Else
               pSCA.Cell(iCol, iRow) = pAccumCells.Cell(iCol, iRow) * dCellSize
            End If
         Next
      Next
      
   Case FUNC_TYPE_EffectContourLen_UpslopeWeighted
      For iRow = 0 To iRows - 1
         For iCol = 0 To iCols - 1
            If pAccumCells.Cell(iCol, iRow) = dNoData Or pDEM.Cell(iCol, iRow) = pDEM.NoData_Value Then
               pSCA.Cell(iCol, iRow) = pSCA.NoData_Value
            ElseIf pAccumCells.Cell(iCol, iRow) <= 1# Then
               pSCA.Cell(iCol, iRow) = 0#
            Else
               dElev = pDEM.Cell(iCol, iRow)
               dSumEffectContour = 0#
               For k = 1 To DIRNUM8
                  iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
                  If pDEM.IsValidCellValue(iCol1, iRow1, dElev1) Then
                     If dElev1 > dElev Then
                        If k Mod 2 = 1 Then
                           dSumEffectContour = dSumEffectContour + SQRT2 / 4
                        Else
                           dSumEffectContour = dSumEffectContour + 0.5
                        End If
                     End If
                  End If
               Next
               
               If dSumEffectContour = 0# Then
                  pSCA.Cell(iCol, iRow) = pSCA.NoData_Value
               Else
                  pSCA.Cell(iCol, iRow) = (pAccumCells.Cell(iCol, iRow) - 1) * dCellSize / dSumEffectContour
               End If
            End If
         Next
      Next
   
   End Select
   
   SpecificCatchmentArea = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'
'  StreamPowerIndex = tan(slope) * a
'
Public Function StreamPowerIndex(pSCA As clsGrid, pTanb As clsGrid, pSPI As clsGrid, Optional dInvalidSPI As Double = -9999#) As Boolean
   Dim iCol As Integer, iRow As Integer
On Error GoTo ErrH
   StreamPowerIndex = False
   If pSCA.nCols <> pTanb.nCols Or pSCA.nRows <> pTanb.nRows Or pSCA.CellSize <> pTanb.CellSize _
         Or pSCA.xllcorner <> pTanb.xllcorner Or pSCA.yllcorner <> pTanb.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID SCA and Tan(Slope) should be with same position and same size."
   End If
   
   For iRow = 0 To pSCA.nRows - 1
      For iCol = 0 To pSCA.nCols - 1
         If pSCA.Cell(iCol, iRow) = pSCA.NoData_Value Then
            pSPI.Cell(iCol, iRow) = pSPI.NoData_Value
         ElseIf pTanb.Cell(iCol, iRow) = pTanb.NoData_Value Then
            pSPI.Cell(iCol, iRow) = pSPI.NoData_Value
         Else
            pSPI.Cell(iCol, iRow) = pSCA.Cell(iCol, iRow) * pTanb.Cell(iCol, iRow)
         End If
      Next
   Next
   StreamPowerIndex = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
'  Topographic Wetness Index = -ln(tan(slope)/a)
'
Public Function TWI_OriginForm(pSCA As clsGrid, pTanb As clsGrid, pTWI As clsGrid, Optional dInvalidTWI As Double = -9999#) As Boolean
   Dim iCol As Integer, iRow As Integer
On Error GoTo ErrH
   TWI_OriginForm = False
   If pSCA.nCols <> pTanb.nCols Or pSCA.nRows <> pTanb.nRows Or pSCA.CellSize <> pTanb.CellSize _
         Or pSCA.xllcorner <> pTanb.xllcorner Or pSCA.yllcorner <> pTanb.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID SCA and Tan(Slope) should be with same position and same size."
   End If
   
   For iRow = 0 To pSCA.nRows - 1
      For iCol = 0 To pSCA.nCols - 1
         If pSCA.Cell(iCol, iRow) = pSCA.NoData_Value Then
            pTWI.Cell(iCol, iRow) = pTWI.NoData_Value
         ElseIf pTanb.Cell(iCol, iRow) = pTanb.NoData_Value Then
            pTWI.Cell(iCol, iRow) = pTWI.NoData_Value
         ElseIf pTanb.Cell(iCol, iRow) <= 0# Or pSCA.Cell(iCol, iRow) <= 0# Then
            pTWI.Cell(iCol, iRow) = IIf(dInvalidTWI = -9999, pTWI.NoData_Value, dInvalidTWI)
         Else
            pTWI.Cell(iCol, iRow) = Log(pSCA.Cell(iCol, iRow) / pTanb.Cell(iCol, iRow))
         End If
      Next
   Next
   TWI_OriginForm = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
' Topographic Wetness Index = ln(a/sum(Lj*tan(bj))) : Original type for MFD-Quinn
' Lj: contour weight (0.5 or sqr(2)/4 (about 0.35)) to downslope neighboring cell
'
Public Function TWI_in_MFD_Quinn(pSCA As clsGrid, pDEM As clsGrid, pTWI As clsGrid, Optional dInvalidTWI As Double = -9999#) As Boolean
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim dSumLj As Double, dSumKi As Double, dSumLjTanbj As Double, dElev As Double, dSlope As Double, dSCA As Double
   
On Error GoTo ErrH
   TWI_in_MFD_Quinn = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (pDEM.Cell(iCol, iRow) <> dNoData) And (pSCA.Cell(iCol, iRow) <> pSCA.NoData_Value) Then
            dSumLj = 0#: dSumKi = 0#: dSumLjTanbj = 0#
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                  If pDEM.Cell(iCol, iRow) > dElev Then   'downslope neighboring cell
                     dSlope = (pDEM.Cell(iCol, iRow) - dElev) / dCellSize    ' tan(SlopeDegree): (icol, irow) -> k
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        dSumLj = dSumLj + SQRT2 / 4      '0.35
                        dSumLjTanbj = dSumLjTanbj + dSlope * SQRT2 / 4
                     Else
                        dSumLj = dSumLj + 0.5
                        dSumLjTanbj = dSumLjTanbj + dSlope * 0.5
                     End If
                  ElseIf pDEM.Cell(iCol, iRow) < dElev Then  'upslope neighboring cell
                     If k Mod 2 = 1 Then
                        dSumKi = dSumKi + SQRT2 / 4      '0.35
                     Else
                        dSumKi = dSumKi + 0.5
                     End If
                  End If
               End If
            Next
            
            'If tanb < tanb_threshold Then tanb = tanb_threshold
            dSCA = pSCA.Cell(iCol, iRow)
            If dSCA <= 0# Or dSumLjTanbj = 0# Then
               pTWI.Cell(iCol, iRow) = IIf(dInvalidTWI = -9999, pTWI.NoData_Value, dInvalidTWI)
            Else
               pTWI.Cell(iCol, iRow) = Log(dSCA / dSumLjTanbj)
            End If
         Else
            pTWI.Cell(iCol, iRow) = pTWI.NoData_Value  ' nodatavalue
         End If
      Next
   Next
   
   TWI_in_MFD_Quinn = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'
' Topographic Wetness Index = ln(a/sum(Lj*tan(bj))) + ln(sum(Lj)/sum(Ki)) (Kong and Rui, 2003)
' Ki: contour weight (0.5 or 0.35) to upslope neighboring cell
' Lj: contour weight (0.5 or 0.35) to downslope neighboring cell
'
Public Function TWI_KongRui2003(pSCA As clsGrid, pDEM As clsGrid, pTWI As clsGrid, Optional dInvalidTWI As Double = -9999#) As Boolean
   Dim iCols As Integer, iRows As Integer, dCellSize As Double, dNoData As Double
   Dim iCol As Integer, iRow As Integer, iCol1 As Integer, iRow1 As Integer, k As Integer
   Dim dSumLj As Double, dSumKi As Double, dSumLjTanbj As Double, dElev As Double, dSlope As Double, dSCA As Double
   
On Error GoTo ErrH
   TWI_KongRui2003 = False
   With pDEM
      iCols = .nCols: iRows = .nRows
      dCellSize = .CellSize
      dNoData = .NoData_Value
   End With
   
   For iRow = 0 To iRows - 1
      For iCol = 0 To iCols - 1
         If (pDEM.Cell(iCol, iRow) <> dNoData) And (pSCA.Cell(iCol, iRow) <> pSCA.NoData_Value) Then
            dSumLj = 0#: dSumKi = 0#: dSumLjTanbj = 0#
            For k = 1 To DIRNUM8
               iCol1 = iCol + ArrDir8X(k): iRow1 = iRow + ArrDir8Y(k)
               If pDEM.IsValidCellValue(iCol1, iRow1, dElev) Then
                  If pDEM.Cell(iCol, iRow) > dElev Then   'downslope neighboring cell
                     dSlope = (pDEM.Cell(iCol, iRow) - dElev) / dCellSize    ' tan(SlopeDegree): (icol, irow) -> k
                     If k Mod 2 = 1 Then
                        dSlope = dSlope / SQRT2
                        dSumLj = dSumLj + SQRT2 / 4      '0.35
                        dSumLjTanbj = dSumLjTanbj + dSlope * SQRT2 / 4
                     Else
                        dSumLj = dSumLj + 0.5
                        dSumLjTanbj = dSumLjTanbj + dSlope * 0.5
                     End If
                  ElseIf pDEM.Cell(iCol, iRow) < dElev Then  'upslope neighboring cell
                     If k Mod 2 = 1 Then
                        dSumKi = dSumKi + SQRT2 / 4      '0.35
                     Else
                        dSumKi = dSumKi + 0.5
                     End If
                  End If
               End If
            Next
            
            'If tanb < tanb_threshold Then tanb = tanb_threshold
            dSCA = pSCA.Cell(iCol, iRow)
            If dSCA <= 0# Or dSumLjTanbj = 0# Then  'also (dSumLj = 0#)
               pTWI.Cell(iCol, iRow) = IIf(dInvalidTWI = -9999, pTWI.NoData_Value, dInvalidTWI)
            ElseIf dSumKi = 0# Then
               pTWI.Cell(iCol, iRow) = Log(dSCA / dSumLjTanbj)
            Else
               pTWI.Cell(iCol, iRow) = Log(dSCA * dSumLj / (dSumLjTanbj * dSumKi))
            End If
         Else
            pTWI.Cell(iCol, iRow) = pTWI.NoData_Value  ' nodatavalue
         End If
      Next
   Next
   
   TWI_KongRui2003 = True
   Exit Function
ErrH:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

