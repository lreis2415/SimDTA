Attribute VB_Name = "modSlopeShape"
Option Explicit
Option Base 1  ' index of array begins from 1

Const C_RIDGE = 1
Const C_VALLEY = 1

Const C_SLPSHAPE_CONCAVE = 1
Const C_SLPSHAPE_STRAIGHT = 3
Const C_SLPSHAPE_CONVEX = 4
Const C_SLPSHAPE_UPCONCAVE_DOWNCONVEX = 2
Const C_SLPSHAPE_UPCONVEX_DOWNCONCAVE = 5

Dim mpDEM As clsGrid, mpFlowDir As clsGrid, mpRidge As clsGrid, mpValley As clsGrid
Dim mpUpslpCells As clsGrid, mpUpslpRlf As clsGrid, mpUpslpDir As clsGrid
Dim mpDownslpCells As clsGrid, mpDownslpRlf As clsGrid
Dim mpRlfDiffer As clsGrid
Dim mpSlpShape As clsGrid
Dim mpUpRelRlfMax As clsGrid, mpUpRelRlfMin As clsGrid, mpUpSlpShape As clsGrid
Dim mpDRelRlfMax As clsGrid, mpDRelRlfMin As clsGrid, mpDSlpShape As clsGrid

Dim miRows As Integer, miCols As Integer, mdCell As Double
Dim mpTag As clsGrid
Dim strPath As String

''''''''''''''''
' for RPI, RRI
Const C_RPITag_ReachStop = 1  'Reach user-assigned valley/channel tag
Const C_RPITag_ForceStop = 2  'Reach outside/pit of area during searching downslope, or peak/outside of area during searching upslope

Dim miRidge As Integer, miValley As Integer  'miRows, miCols, mdCell
Dim mp2Vly As clsGrid, mp2Rdg As clsGrid     'mpDEM, mpRidge, mpValley
Dim mpPosCol As clsGrid, mpPosRow As clsGrid ' mpTag

'''''''''''''''''''''''''''''''
' for RRI
'' 若无法下游找到最近的沟谷，则找下游到达最近的pit
Private Function RRI_SearchDownslope(iCol0 As Integer, iRow0 As Integer) As Boolean
   Dim iCol As Integer, iRow As Integer, k As Integer, iDir_NearDist As Integer, iDir_ForceStop As Integer
   Dim dElev0 As Double, dElev As Double, dDist As Double, dDist2ForceStop As Double, dLowest As Double
   Dim boolReachStop As Boolean
   
   RRI_SearchDownslope = False
   
   If mpValley.Cell(iCol0, iRow0) = miValley Then
      mp2Vly.Cell(iCol0, iRow0) = 0#
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
   ElseIf Not mpDEM.IsValidCellValue(iCol0, iRow0, dElev0) Then
      mp2Vly.Cell(iCol0, iRow0) = mp2Vly.NoData_Value
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
   Else
      boolReachStop = False
      dDist = MAX_SINGLE
      dDist2ForceStop = MAX_SINGLE
      dLowest = MAX_SINGLE
      For k = 1 To DIRNUM8
         iCol = iCol0 + ArrDir8X(k): iRow = iRow0 + ArrDir8Y(k)
         If mpDEM.IsValidCellValue(iCol, iRow, dElev) Then
            If dElev < dElev0 Then
               If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
                  RRI_SearchDownslope iCol, iRow
               End If
               
               If mpTag.Cell(iCol, iRow) = C_RPITag_ReachStop Then
                  If dDist > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_NearDist = k
                     boolReachStop = True
                  End If
               ElseIf mpTag.Cell(iCol, iRow) = C_RPITag_ForceStop Then
'                  If dLowest > mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow)) Then
'                     dLowest = mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow))
                  If dDist2ForceStop > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist2ForceStop = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_ForceStop = k
                  End If
               End If
            End If
         End If
         DoEvents
      Next
      
      If boolReachStop Then
         iCol = iCol0 + ArrDir8X(iDir_NearDist): iRow = iRow0 + ArrDir8Y(iDir_NearDist)
         mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
         mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
         mp2Vly.Cell(iCol0, iRow0) = dElev0 - mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow))
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
      Else
         If dDist2ForceStop = MAX_SINGLE Then
         ' at bottom of a pit
            mpPosCol.Cell(iCol0, iRow0) = iCol0
            mpPosRow.Cell(iCol0, iRow0) = iRow0
            mp2Vly.Cell(iCol0, iRow0) = 0#
         Else
         ' can not drainage into valley in area
            iCol = iCol0 + ArrDir8X(iDir_ForceStop): iRow = iRow0 + ArrDir8Y(iDir_ForceStop)
            mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
            mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
            mp2Vly.Cell(iCol0, iRow0) = dElev0 - mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow))
         End If
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
      End If
   End If
      
   RRI_SearchDownslope = True
Err:
   
End Function

' 若无法上溯找到最近的山脊，则找上溯到达最近的peak
Private Function RRI_SearchUpslope(iCol0 As Integer, iRow0 As Integer) As Boolean
   Dim iCol As Integer, iRow As Integer, k As Integer, iDir_NearDist As Integer, iDir_ForceStop As Integer
   Dim dElev0 As Double, dElev As Double, dDist As Double, dDist2ForceStop As Double, dHighest As Double
   Dim boolReachStop As Boolean
   
   RRI_SearchUpslope = False
   
   If mpRidge.Cell(iCol0, iRow0) = miRidge Then
      mp2Rdg.Cell(iCol0, iRow0) = 0#
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
   ElseIf Not mpDEM.IsValidCellValue(iCol0, iRow0, dElev0) Then
      mp2Rdg.Cell(iCol0, iRow0) = mp2Rdg.NoData_Value
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
   Else
      boolReachStop = False
      dDist = MAX_SINGLE
      dDist2ForceStop = MAX_SINGLE
      dHighest = MIN_SINGLE
      For k = 1 To DIRNUM8
         iCol = iCol0 + ArrDir8X(k): iRow = iRow0 + ArrDir8Y(k)
         If mpDEM.IsValidCellValue(iCol, iRow, dElev) Then
            If dElev > dElev0 Then
               If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
                  RRI_SearchUpslope iCol, iRow
               End If
               
               If mpTag.Cell(iCol, iRow) = C_RPITag_ReachStop Then
                  If dDist > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_NearDist = k
                     boolReachStop = True
                  End If
               ElseIf mpTag.Cell(iCol, iRow) = C_RPITag_ForceStop Then
                  'If dHighest < mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow)) Then
                  '   dHighest = mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow))
                  If dDist2ForceStop > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist2ForceStop = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_ForceStop = k
                  End If
               End If
            End If
         End If
         DoEvents
      Next
      
      If boolReachStop Then
         iCol = iCol0 + ArrDir8X(iDir_NearDist): iRow = iRow0 + ArrDir8Y(iDir_NearDist)
         mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
         mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
         mp2Rdg.Cell(iCol0, iRow0) = mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow)) - dElev0
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
      Else
         If dDist2ForceStop = MAX_SINGLE Then
         ' at top of a peak
            mpPosCol.Cell(iCol0, iRow0) = iCol0
            mpPosRow.Cell(iCol0, iRow0) = iRow0
            mp2Rdg.Cell(iCol0, iRow0) = 0#
         Else
         ' can not reach hardened ridge in area
            iCol = iCol0 + ArrDir8X(iDir_ForceStop): iRow = iRow0 + ArrDir8Y(iDir_ForceStop)
            mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
            mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
            mp2Rdg.Cell(iCol0, iRow0) = mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow)) - dElev0
         End If
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
      End If
   End If
      
   RRI_SearchUpslope = True
Err:
   
End Function

'
' Revised Relative relief index:
'        Pij = (relief to the nearest valley)
'              / (relief to the nearest valley + relief to the nearest ridge)
' P.S.   the nearest valley: routed from the interest cell;
'        the nearest ridge: routed to the interest cell.
' Pij<0.1 → valley; 0.1≤Pij<0.4 → lower mid-slope; 0.4≤Pij<0.6 → mid-slope; 0.6≤Pij<0.8 → an upper mid-slope; Pij≥0.8 → ridge.
'
Public Function RelativeReliefIndex_KeepRouting(pDEM As clsGrid, pRidge As clsGrid, pChannel As clsGrid, _
      pRRI As clsGrid, pRlf2Ridge As clsGrid, pRlf2Valley As clsGrid, _
      Optional iRidgeTag As Integer = 1, Optional iChannelTag As Integer = 1) As Boolean
      
On Error GoTo ErrH
   Dim iCol As Integer, iRow As Integer
   
   RelativeReliefIndex_KeepRouting = False
   
   ' verify parameters
   With pDEM
      miCols = .nCols: miRows = .nRows
      mdCell = .CellSize
   End With
   If miCols <> pRidge.nCols Or miRows <> pRidge.nRows Or mdCell <> pRidge.CellSize _
         Or miCols <> pChannel.nCols Or miRows <> pChannel.nRows Or mdCell <> pChannel.CellSize Then
      '   Or pRidge.xllcorner <> pChannel.xllcorner Or pRidge.yllcorner <> pChannel.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID DEM, Ridge and Channel should be with same position and same size."
   End If
'   For iCol = 0 To miCols - 1
'      For iRow = 0 To miRows - 1
'         If pChannel.Cell(iCol, iRow) = iChannelTag Then GoTo HasValley
'      Next
'   Next
'   Err.Raise Number:=vbObjectError + 513, Description:="No Valley-cell in given Valley GRID"
'HasValley:
'   For iCol = 0 To miCols - 1
'      For iRow = 0 To miRows - 1
'         If pRidge.Cell(iCol, iRow) = iRidgeTag Then GoTo HasRidge
'      Next
'   Next
'   Err.Raise Number:=vbObjectError + 513, Description:="No Ridge-cell in given Ridge GRID"
'HasRidge:
   DoEvents
   
   ' initial paremeters
   miRidge = iRidgeTag: miValley = iChannelTag
   Set mpDEM = pDEM: Set mpRidge = pRidge: Set mpValley = pChannel
   Set mp2Vly = pRlf2Valley: Set mp2Rdg = pRlf2Ridge
   With mpDEM
      Set mpTag = New clsGrid
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      Set mpPosCol = New clsGrid
      If Not mpPosCol.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      Set mpPosRow = New clsGrid
      If Not mpPosRow.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
      
   'search downslope for nearest valley
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
            RRI_SearchDownslope iCol, iRow
         End If
      Next
      DoEvents
   Next
   
   'search upslope for nearest ridge
   With mpDEM
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      If Not mpPosCol.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      If Not mpPosRow.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
            RRI_SearchUpslope iCol, iRow
         End If
      Next
      DoEvents
   Next
   
   ' Relative relief = (relief to the nearest valley) / (relief to the nearest valley + relief to nearest ridge)
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If pRlf2Ridge.Cell(iCol, iRow) = 0# Then
            pRRI.Cell(iCol, iRow) = 1#
         ElseIf pRlf2Valley.Cell(iCol, iRow) = 0# Then
            pRRI.Cell(iCol, iRow) = 0#
'         ElseIf vDist2Ridge(iCol, iRow) = MAX_SINGLE Or vDist2Valley(iCol, iRow) = MAX_SINGLE Then
'            pRRI.Cell(iCol, iRow) = pRRI.NoData_Value
         ElseIf pRlf2Ridge.Cell(iCol, iRow) = pRlf2Ridge.NoData_Value Or pRlf2Valley.Cell(iCol, iRow) = pRlf2Valley.NoData_Value Then
            pRRI.Cell(iCol, iRow) = pRRI.NoData_Value
         ElseIf pRlf2Valley.Cell(iCol, iRow) + pRlf2Ridge.Cell(iCol, iRow) = 0# Then
            pRRI.Cell(iCol, iRow) = 1#
         Else
            pRRI.Cell(iCol, iRow) = pRlf2Valley.Cell(iCol, iRow) / (pRlf2Valley.Cell(iCol, iRow) + pRlf2Ridge.Cell(iCol, iRow))
         End If
      Next
   Next
   
   RelativeReliefIndex_KeepRouting = True
ErrH:
   On Error Resume Next
   Set mpDEM = Nothing: Set mpRidge = Nothing: Set mpValley = Nothing
   Set mp2Vly = Nothing: Set mp2Rdg = Nothing
   Set mpPosCol = Nothing: Set mpPosRow = Nothing
   
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function


'''''''''''''''''''''''''''''''
' for RPI
'' 若无法下游找到最近的沟谷，则找下游到达最近的pit
Private Function RPI_SearchDownslope(iCol0 As Integer, iRow0 As Integer) As Boolean
   Dim iCol As Integer, iRow As Integer, k As Integer, iDir_NearDist As Integer, iDir_ForceStop As Integer
   Dim dElev0 As Double, dElev As Double, dDist As Double, dDist2ForceStop As Double, dLowest As Double
   Dim boolReachStop As Boolean
   
   RPI_SearchDownslope = False
   
   If mpValley.Cell(iCol0, iRow0) = miValley Then
      mp2Vly.Cell(iCol0, iRow0) = 0#
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
   ElseIf Not mpDEM.IsValidCellValue(iCol0, iRow0, dElev0) Then
      mp2Vly.Cell(iCol0, iRow0) = mp2Vly.NoData_Value
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
   Else
      boolReachStop = False
      dDist = MAX_SINGLE
      dDist2ForceStop = MAX_SINGLE
      dLowest = MAX_SINGLE
      For k = 1 To DIRNUM8
         iCol = iCol0 + ArrDir8X(k): iRow = iRow0 + ArrDir8Y(k)
         If mpDEM.IsValidCellValue(iCol, iRow, dElev) Then
            If dElev < dElev0 Then
               If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
                  RPI_SearchDownslope iCol, iRow
               End If
               
               If mpTag.Cell(iCol, iRow) = C_RPITag_ReachStop Then
                  If dDist > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_NearDist = k
                     boolReachStop = True
                  End If
               ElseIf mpTag.Cell(iCol, iRow) = C_RPITag_ForceStop Then
'                  If dLowest > mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow)) Then
'                     dLowest = mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow))
                  If dDist2ForceStop > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist2ForceStop = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_ForceStop = k
                  End If
               End If
            End If
         End If
         DoEvents
      Next
      
      If boolReachStop Then
         iCol = iCol0 + ArrDir8X(iDir_NearDist): iRow = iRow0 + ArrDir8Y(iDir_NearDist)
         mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
         mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
         mp2Vly.Cell(iCol0, iRow0) = Sqr(dDist) * mdCell
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
      Else
         If dDist2ForceStop = MAX_SINGLE Then
         ' at bottom of a pit
            mpPosCol.Cell(iCol0, iRow0) = iCol0
            mpPosRow.Cell(iCol0, iRow0) = iRow0
            mp2Vly.Cell(iCol0, iRow0) = 0#
         Else
         ' can not drainage into valley in area
            iCol = iCol0 + ArrDir8X(iDir_ForceStop): iRow = iRow0 + ArrDir8Y(iDir_ForceStop)
            mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
            mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
            mp2Vly.Cell(iCol0, iRow0) = Sqr(dDist2ForceStop) * mdCell
         End If
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
      End If
   End If
      
   RPI_SearchDownslope = True
Err:
   
End Function

' 若无法上溯找到最近的山脊，则找上溯到达最近的peak
Private Function RPI_SearchUpslope(iCol0 As Integer, iRow0 As Integer) As Boolean
   Dim iCol As Integer, iRow As Integer, k As Integer, iDir_NearDist As Integer, iDir_ForceStop As Integer
   Dim dElev0 As Double, dElev As Double, dDist As Double, dDist2ForceStop As Double, dHighest As Double
   Dim boolReachStop As Boolean
   
   RPI_SearchUpslope = False
   
   If mpRidge.Cell(iCol0, iRow0) = miRidge Then
      mp2Rdg.Cell(iCol0, iRow0) = 0#
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
   ElseIf Not mpDEM.IsValidCellValue(iCol0, iRow0, dElev0) Then
      mp2Rdg.Cell(iCol0, iRow0) = mp2Rdg.NoData_Value
      mpPosCol.Cell(iCol0, iRow0) = iCol0:   mpPosRow.Cell(iCol0, iRow0) = iRow0
      mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
   Else
      boolReachStop = False
      dDist = MAX_SINGLE
      dDist2ForceStop = MAX_SINGLE
      dHighest = MIN_SINGLE
      For k = 1 To DIRNUM8
         iCol = iCol0 + ArrDir8X(k): iRow = iRow0 + ArrDir8Y(k)
         If mpDEM.IsValidCellValue(iCol, iRow, dElev) Then
            If dElev > dElev0 Then
               If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
                  RPI_SearchUpslope iCol, iRow
               End If
               
               If mpTag.Cell(iCol, iRow) = C_RPITag_ReachStop Then
                  If dDist > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_NearDist = k
                     boolReachStop = True
                  End If
               ElseIf mpTag.Cell(iCol, iRow) = C_RPITag_ForceStop Then
                  'If dHighest < mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow)) Then
                  '   dHighest = mpDEM.Cell(mpPosCol.Cell(iCol, iRow), mpPosRow.Cell(iCol, iRow))
                  If dDist2ForceStop > (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2 Then
                     dDist2ForceStop = (mpPosCol.Cell(iCol, iRow) - iCol0) ^ 2 + (mpPosRow.Cell(iCol, iRow) - iRow0) ^ 2
                     iDir_ForceStop = k
                  End If
               End If
            End If
         End If
         DoEvents
      Next
      
      If boolReachStop Then
         iCol = iCol0 + ArrDir8X(iDir_NearDist): iRow = iRow0 + ArrDir8Y(iDir_NearDist)
         mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
         mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
         mp2Rdg.Cell(iCol0, iRow0) = Sqr(dDist) * mdCell
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ReachStop
      Else
         If dDist2ForceStop = MAX_SINGLE Then
         ' at top of a peak
            mpPosCol.Cell(iCol0, iRow0) = iCol0
            mpPosRow.Cell(iCol0, iRow0) = iRow0
            mp2Rdg.Cell(iCol0, iRow0) = 0#
         Else
         ' can not reach hardened ridge in area
            iCol = iCol0 + ArrDir8X(iDir_ForceStop): iRow = iRow0 + ArrDir8Y(iDir_ForceStop)
            mpPosCol.Cell(iCol0, iRow0) = mpPosCol.Cell(iCol, iRow)
            mpPosRow.Cell(iCol0, iRow0) = mpPosRow.Cell(iCol, iRow)
            mp2Rdg.Cell(iCol0, iRow0) = Sqr(dDist2ForceStop) * mdCell
         End If
         mpTag.Cell(iCol0, iRow0) = C_RPITag_ForceStop
      End If
   End If
      
   RPI_SearchUpslope = True
Err:
   
End Function

'
' Revised Relative position index:
'        Pij = (Euclidean distance to the nearest valley)
'              / (Euclidean distance to the nearest valley + Euclidean distance to the nearest ridge)
' P.S.   the nearest valley: routed from the interest cell;
'        the nearest ridge: routed to the interest cell.
' Pij<0.1 → valley; 0.1≤Pij<0.4 → lower mid-slope; 0.4≤Pij<0.6 → mid-slope; 0.6≤Pij<0.8 → an upper mid-slope; Pij≥0.8 → ridge.
'
Public Function RelativePositionIndex_KeepRouting(pDEM As clsGrid, pRidge As clsGrid, pChannel As clsGrid, _
      pRPI As clsGrid, pDist2Ridge As clsGrid, pDist2Valley As clsGrid, _
      Optional iRidgeTag As Integer = 1, Optional iChannelTag As Integer = 1) As Boolean
      
On Error GoTo ErrH
   Dim iCol As Integer, iRow As Integer
   
   RelativePositionIndex_KeepRouting = False
   
   ' verify parameters
   With pDEM
      miCols = .nCols: miRows = .nRows
      mdCell = .CellSize
   End With
   If miCols <> pRidge.nCols Or miRows <> pRidge.nRows Or mdCell <> pRidge.CellSize _
         Or miCols <> pChannel.nCols Or miRows <> pChannel.nRows Or mdCell <> pChannel.CellSize Then
      '   Or pRidge.xllcorner <> pChannel.xllcorner Or pRidge.yllcorner <> pChannel.yllcorner Then
      Err.Raise Number:=vbObjectError + 513, Description:="GRID DEM, Ridge and Channel should be with same position and same size."
   End If
'   For iCol = 0 To miCols - 1
'      For iRow = 0 To miRows - 1
'         If pChannel.Cell(iCol, iRow) = iChannelTag Then GoTo HasValley
'      Next
'   Next
'   Err.Raise Number:=vbObjectError + 513, Description:="No Valley-cell in given Valley GRID"
'HasValley:
'   For iCol = 0 To miCols - 1
'      For iRow = 0 To miRows - 1
'         If pRidge.Cell(iCol, iRow) = iRidgeTag Then GoTo HasRidge
'      Next
'   Next
'   Err.Raise Number:=vbObjectError + 513, Description:="No Ridge-cell in given Ridge GRID"
'HasRidge:
   DoEvents
   
   ' initial paremeters
   miRidge = iRidgeTag: miValley = iChannelTag
   Set mpDEM = pDEM: Set mpRidge = pRidge: Set mpValley = pChannel
   Set mp2Vly = pDist2Valley: Set mp2Rdg = pDist2Ridge
   With mpDEM
      Set mpTag = New clsGrid
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      Set mpPosCol = New clsGrid
      If Not mpPosCol.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      Set mpPosRow = New clsGrid
      If Not mpPosRow.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
      
   'search downslope for nearest valley
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
            RPI_SearchDownslope iCol, iRow
         End If
      Next
      DoEvents
   Next
   
   'search upslope for nearest ridge
   With mpDEM
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      If Not mpPosCol.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      If Not mpPosRow.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, -9999, -9999, True) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = mpTag.NoData_Value Then
            RPI_SearchUpslope iCol, iRow
         End If
      Next
      DoEvents
   Next
   
   ' Relative position = (Euclidean distance to the nearest valley) / (Euclidean distance to the nearest valley + Euclidean distance to nearest ridge)
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
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
   
   RelativePositionIndex_KeepRouting = True
ErrH:
   On Error Resume Next
   Set mpDEM = Nothing: Set mpRidge = Nothing: Set mpValley = Nothing
   Set mp2Vly = Nothing: Set mp2Rdg = Nothing
   Set mpPosCol = Nothing: Set mpPosRow = Nothing
   
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, APP_TITLE
End Function

'
''''''''''''''''''''''''
'==========begin of functions for frmSlopeShape================
'
Public Function SlopeDescrib() As Boolean
   Dim iRow As Integer, iCol As Integer, iProcRow As Integer, iProcCol As Integer
   'Dim iTopRow As Integer, iTopCol As Integer, iBottomRow As Integer, iBottomCol As Integer
   'Dim dTop As Double, dBottom As Double, dHeight As Double, dRelief As Double
   'Dim iTempRow As Integer, iTempCol As Integer
   Dim iDir As Integer
   'Dim lCells As Long
   Dim blnProc As Boolean
   Dim strFile As String
   
   strPath = "f:\DataMining\LandProject\FuncPrg\LandslideInfer-Prog\Data\PleasentVly\"
   
   If Not PrepareInputData() Then Err.Raise vbObjectError + 513, , "Failed in DTA function"
   
   'If Not PrepareOutputData() Then Err.Raise vbObjectError + 513, , "Failed in DTA function"
   'create para Grid
   With mpDEM
      Set mpUpslpCells = New clsGrid
      If Not mpUpslpCells.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpUpslpRlf = New clsGrid
      If Not mpUpslpRlf.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpUpslpDir = New clsGrid
      If Not mpUpslpDir.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   
   With mpDEM
      Set mpTag = New clsGrid
      If Not mpTag.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, 0) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   
   '''''''''''''''''''''''''
   ' search upslope
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = 0 Then SearchUpslope iCol, iRow
      Next
   Next
   
   ''''''''''''''''''''
   ' output para
   strFile = strPath & "SimDTA\UpslpCells.asc"
   If mpUpslpCells.SaveAscGrid(strFile, , 0) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If

   strFile = strPath & "SimDTA\UpslpRlf.asc"
   If mpUpslpRlf.SaveAscGrid(strFile, , 2) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If

   strFile = strPath & "SimDTA\UpslpDir.asc"
   If mpUpslpDir.SaveAscGrid(strFile, , 0) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
   'Set mpUpslpCells = Nothing:   Set mpUpslpRlf = Nothing:   Set mpUpslpDir = Nothing
         
   ''''''''''''''''''''''''''
   'Set mpTag = Nothing
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         mpTag.Cell(iCol, iRow) = 0
      Next
   Next
   
   'create para Grid
   With mpDEM
      Set mpDownslpCells = New clsGrid
      If Not mpDownslpCells.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
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
   End With
   
   '''''''''''''''''''''''''
   ' search downslope
   For iRow = 0 To miRows - 1
      For iCol = 0 To miCols - 1
         If mpTag.Cell(iCol, iRow) = 0 Then SearchDownslope iCol, iRow
      Next
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
   Next
   
   ''''''''''''''''''''''''
   ' output para
   strFile = strPath & "SimDTA\DslpCells.asc"
   If mpDownslpCells.SaveAscGrid(strFile, , 0) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If

   strFile = strPath & "SimDTA\DslpRlf.asc"
   If mpDownslpRlf.SaveAscGrid(strFile, , 2) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If

   strFile = strPath & "SimDTA\RlfDiffer.asc"   'convexHigh.asc"
   If mpRlfDiffer.SaveAscGrid(strFile, , 2) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
'   Set mpDownslpCells = Nothing:   Set mpDownslpRlf = Nothing
'   Set mpRlfDiffer = Nothing

   '''''''''''''''''''''''
   'create para Grid
   With mpDEM
      Set mpSlpShape = New clsGrid
      If Not mpSlpShape.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      ' upslope part
      Set mpUpRelRlfMax = New clsGrid
      If Not mpUpRelRlfMax.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpUpRelRlfMin = New clsGrid
      If Not mpUpRelRlfMin.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpUpSlpShape = New clsGrid
      If Not mpUpSlpShape.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      'downslope part
      Set mpDRelRlfMax = New clsGrid
      If Not mpDRelRlfMax.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpDRelRlfMin = New clsGrid
      If Not mpDRelRlfMin.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
      
      Set mpDSlpShape = New clsGrid
      If Not mpDSlpShape.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
      End If
   End With
   
   ' calc slopeShape
   Call SlopeShape
   
   ' output slope-shape result
   strFile = strPath & "SimDTA\slpShape.asc"
   If mpSlpShape.SaveAscGrid(strFile, , 0) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
   strFile = strPath & "SimDTA\UpRelRlfMax.asc"
   If mpUpRelRlfMax.SaveAscGrid(strFile, , 2) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
   strFile = strPath & "SimDTA\UpRelRlfMin.asc"
   If mpUpRelRlfMin.SaveAscGrid(strFile, , 2) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
   strFile = strPath & "SimDTA\UpslpShape.asc"
   If mpUpSlpShape.SaveAscGrid(strFile, , 0) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
   strFile = strPath & "SimDTA\DRelRlfMax.asc"
   If mpDRelRlfMax.SaveAscGrid(strFile, , 2) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
   strFile = strPath & "SimDTA\DRelRlfMin.asc"
   If mpDRelRlfMin.SaveAscGrid(strFile, , 2) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If
   
   strFile = strPath & "SimDTA\DSlpShape.asc"
   If mpDSlpShape.SaveAscGrid(strFile, , 0) Then
      'MsgBox "Completed. Save result GRID: " & vbCrLf & strFile
   Else
      Err.Raise vbObjectError + 513, , "Failed to save result GRID: " & strFile
   End If

   ''''''''''''''''''''''
   CleanUpMemory
   
   MsgBox "Done!"
   
End Function

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
      If mpFlowDir.IsValidCellValue(iProcCol, iProcRow, dVal) And mpValley.Cell(iProcCol, iProcRow) <> C_VALLEY Then
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

Private Function PrepareInputData() As Boolean
   Dim sDEMFile As String, sFlowDirFile As String, sRidgeFile As String, sValleyFile As String
   
   PrepareInputData = False
   
   sDEMFile = strPath & "elev_3dr.asc"
   sFlowDirFile = strPath & "d0fs_d8.asc"
   'sRidgeFile = strPath & "d0pd_pdrdg950.asc"
   sValleyFile = strPath & "d0pdmmdvly500.asc"
   
   Set mpDEM = New clsGrid
   mpDEM.LoadAscGrid sDEMFile
   Set mpFlowDir = New clsGrid
   mpFlowDir.LoadAscGrid sFlowDirFile
   Set mpValley = New clsGrid
   mpValley.LoadAscGrid sValleyFile
   
   With mpDEM
      miRows = .nRows: miCols = .nCols: mdCell = .CellSize
   End With
   
   PrepareInputData = True
End Function

'Private Function PrepareOutputData()
'   With mpDEM
'      Set mpUpslpCells = New clsGrid
'      If Not mpUpslpCells.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
'         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
'      End If
'
'      Set mpUpslpRlf = New clsGrid
'      If Not mpUpslpRlf.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
'         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
'      End If
'
'      Set mpUpslpDir = New clsGrid
'      If Not mpUpslpDir.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
'         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
'      End If
'
'      Set mpDownslpCells = New clsGrid
'      If Not mpDownslpCells.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
'         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
'      End If
'
'      Set mpDownslpRlf = New clsGrid
'      If Not mpDownslpRlf.NewGrid(miCols, miRows, .xllcorner, .yllcorner, mdCell, .NoData_Value, .NoData_Value) Then
'         Err.Raise Number:=vbObjectError + 513, Description:="Error in NewGRID"
'      End If
'   End With
'End Function

Private Function CleanUpMemory() As Boolean
   Set mpTag = Nothing
   Set mpDEM = Nothing:   Set mpFlowDir = Nothing:   Set mpRidge = Nothing: Set mpValley = Nothing
   Set mpUpslpCells = Nothing:   Set mpUpslpRlf = Nothing:   Set mpUpslpDir = Nothing
   Set mpDownslpCells = Nothing:   Set mpDownslpRlf = Nothing
   Set mpRlfDiffer = Nothing
   Set mpSlpShape = Nothing
   Set mpUpRelRlfMax = Nothing:  Set mpUpRelRlfMin = Nothing:   Set mpUpSlpShape = Nothing
   Set mpDRelRlfMax = Nothing:   Set mpDRelRlfMin = Nothing:   Set mpDSlpShape = Nothing
End Function

'
' search along both upslope and downslope flow-path
' find both position and value with max and min of RlfDiffer
'
Public Function SlopeShape() As Boolean
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
            mpSlpShape.Cell(iCol, iRow) = mpSlpShape.NoData_Value
            
            mpUpSlpShape.Cell(iCol, iRow) = mpUpSlpShape.NoData_Value
            mpUpRelRlfMax.Cell(iCol, iRow) = mpUpRelRlfMax.NoData_Value
            mpUpRelRlfMin.Cell(iCol, iRow) = mpUpRelRlfMin.NoData_Value
            
            mpDSlpShape.Cell(iCol, iRow) = mpDSlpShape.NoData_Value
            mpDRelRlfMax.Cell(iCol, iRow) = mpDRelRlfMax.NoData_Value
            mpDRelRlfMin.Cell(iCol, iRow) = mpDRelRlfMin.NoData_Value
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
            mpUpRelRlfMax.Cell(iCol, iRow) = dUpRelRlfMax
            mpUpRelRlfMin.Cell(iCol, iRow) = dUpRelRlfMin
            
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
            mpDRelRlfMax.Cell(iCol, iRow) = dDRelRlfMax
            mpDRelRlfMin.Cell(iCol, iRow) = dDRelRlfMin
                 
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
      Next
   Next
   
End Function

'
'==========end of functions for frmSlopeShape================
'''''''''''''''''''''''''
