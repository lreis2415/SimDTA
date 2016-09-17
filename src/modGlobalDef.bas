Attribute VB_Name = "modGlobalDef"
Option Explicit
Option Base 1

#If Win16 Then
    Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#Else
    Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#End If

Public Const C_INNER_VERSION = True
Public Const C_Version = "Version: 1.0.1"
Public Const C_LastModify = "  (last modification: 2009/1/20)"

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
'SetWindowPos(form.hwnd, -1, 0, 0, 0, 0, 3) : form始终在最前
'SetWindowPos(form.hwnd, -2, 0, 0, 0, 0, 3) : form
'Const wFlags = SWP_NOMOVE Or SWP_NOSIZE
'SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags
                  
'''''''''''''''''''''''''''
Public Const APP_TITLE = "SimDTA"

Public Const FUNC_FILLDEP = 0
Public Const FUNC_SLOPE = 1
Public Const FUNC_ASPECT = 2
Public Const FUNC_CURVATURE = 3
Public Const FUNC_SurfaceCurvature = 4
Public Const FUNC_ElevPercentile = 5
Public Const FUNC_TOPHAT = 6
Public Const FUNC_TopoRugI = 7
Public Const FUNC_LandPosI = 8
Public Const FUNC_UPNESS = 9
Public Const FUNC_RelaPosI = 10
Public Const FUNC_DownslopeIndex = 11
Public Const FUNC_RIDGE_Peucker = 12
Public Const FUNC_DRAINAGE_Peucker = 13
Public Const FUNC_MFD = 14
Public Const FUNC_TWI = 15
Public Const FUNC_TerrainCharI = 16
Public Const FUNC_StreamPowerI = 17
Public Const FUNC_ElevReliefRatio = 18
Public Const FUNC_Relief = 19
Public Const FUNC_TopoPosIndex = 20
Public Const FUNC_SurfaceArea = 21
Public Const FUNC_Openness = 22
'Public Const FUNC_HypsomIntegral = 23
Public Const FUNC_RelaRlfI = 23

Public Const FUNC_TYPE_SLOPE = "Slope in ArcInfo"
Public Const FUNC_TYPE_MAXDOWNSLOPE = "Max Downslope"
Public Const FUNC_TYPE_LOCALDOWNSLOPE = "Local downslope (Quinn et al., 1991))"

Public Const FUNC_TYPE_MFD_QUINN91 = "MFD (Quinn et al., 1991)"
Public Const FUNC_TYPE_MFD_QIN07 = "MFD-md (Qin et al., 2007)"
Public Const FUNC_TYPE_EffectContourLen_Cell = "Cell size (Original algorithm)"
Public Const FUNC_TYPE_EffectContourLen_UpslopeWeighted = "Weighted Upslope Contour (Yong et al., 2008)"

Public Const FUNC_TYPE_RPI_Skidmore90 = "RPI (Skidmore, 1990)"
Public Const FUNC_TYPE_RPI_relief = "RPI (maintain relief) (Qin et al., preparing)"
Public Const FUNC_TYPE_RPI_routing = "RPI (maintain routing) (Qin et al., preparing)"

Public Const FUNC_TYPE_RRI_relief = "RRI (maintain relief) (Qin et al., preparing)"
Public Const FUNC_TYPE_RRI_routing = "RRI (maintain routing) (Qin et al., preparing)"

'''''''''''''''''''''''''''

Public Const MAX_SINGLE = 3.402823E+38
Public Const MIN_SINGLE = -1.402823E+38   ' -3.402823E+38
Public Const SQRT2 = 1.414213562373
Public Const PI = 3.14159265358979
Public Const COEF_2AngleDegree = 180 / PI

Public Const ESRI_DIR_E = 1
Public Const ESRI_DIR_SE = 2
Public Const ESRI_DIR_S = 4
Public Const ESRI_DIR_SW = 8
Public Const ESRI_DIR_W = 16
Public Const ESRI_DIR_NW = 32
Public Const ESRI_DIR_N = 64
Public Const ESRI_DIR_NE = 128
Public Const ESRI_DIR_UNDEF = 0
Public Const DIRNUM8 = 8

Public C_VersionInfo As String

' deltaX, Y of 8 neighbor, begins from NE by clockwise
Public ArrDir8X As Variant    '= (1, 1, 1, 0, -1, -1, -1, 0)     'index from 1; unaffected by Option Base
Public ArrDir8Y As Variant   ' = Array(1, 0, -1, -1, -1, 0, 1, 1)
    ' Left-up corner is the (0,0) of grid
Public ESRIDir As Variant  ' = Array(ESRI_DIR_SE, ESRI_DIR_E, ESRI_DIR_NE, ESRI_DIR_N, ESRI_DIR_NW, ESRI_DIR_W, ESRI_DIR_SW, ESRI_DIR_S)
'''''''''''''''''''''''''

Public g_Statusbar As StatusBar

Public Sub InitializeGlobalVar()
   
   If C_INNER_VERSION Then
      C_VersionInfo = C_Version & ".beta" & C_LastModify
   Else
      C_VersionInfo = C_Version & C_LastModify
   End If

   ' don't change this sequence!! will affect max downslope & MFD & ridge & valley!!
   ArrDir8X = Array(1, 1, 1, 0, -1, -1, -1, 0)   'index from 1; unaffected by Option Base
   ArrDir8Y = Array(1, 0, -1, -1, -1, 0, 1, 1)
    ' Left-up corner is the (0,0) of grid
   ESRIDir = Array(ESRI_DIR_SE, ESRI_DIR_E, ESRI_DIR_NE, ESRI_DIR_N, ESRI_DIR_NW, ESRI_DIR_W, ESRI_DIR_SW, ESRI_DIR_S)
End Sub


Public Function ESRIDir_Reverse(FlowDir As Integer) As Integer
   Select Case FlowDir
   Case ESRI_DIR_E
      ESRIDir_Reverse = ESRI_DIR_W
   Case ESRI_DIR_SE
      ESRIDir_Reverse = ESRI_DIR_NW
   Case ESRI_DIR_S
      ESRIDir_Reverse = ESRI_DIR_N
   Case ESRI_DIR_SW
      ESRIDir_Reverse = ESRI_DIR_NE
   Case ESRI_DIR_W
      ESRIDir_Reverse = ESRI_DIR_E
   Case ESRI_DIR_NW
      ESRIDir_Reverse = ESRI_DIR_SE
   Case ESRI_DIR_N
      ESRIDir_Reverse = ESRI_DIR_S
   Case ESRI_DIR_NE
      ESRIDir_Reverse = ESRI_DIR_SW
   Case Else
      ESRIDir_Reverse = ESRI_DIR_UNDEF
   End Select
End Function

Public Function GetESRIDir_ArrayIndex(FlowDir As Integer) As Integer
   Select Case FlowDir
   Case ESRI_DIR_E
      GetESRIDir_ArrayIndex = 2
   Case ESRI_DIR_SE
      GetESRIDir_ArrayIndex = 1
   Case ESRI_DIR_S
      GetESRIDir_ArrayIndex = 8
   Case ESRI_DIR_SW
      GetESRIDir_ArrayIndex = 7
   Case ESRI_DIR_W
      GetESRIDir_ArrayIndex = 6
   Case ESRI_DIR_NW
      GetESRIDir_ArrayIndex = 5
   Case ESRI_DIR_N
      GetESRIDir_ArrayIndex = 4
   Case ESRI_DIR_NE
      GetESRIDir_ArrayIndex = 3
   Case Else
      GetESRIDir_ArrayIndex = -1
   End Select
End Function

Public Sub ReleaseMemory()
   ArrDir8X = Empty
   ArrDir8Y = Empty
   ESRIDir = Empty
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'设置窗口是否总在最前
Public Sub SetFormStayOnTop(mForm As Form, StayOnTop As Boolean)
    Dim stat
    If StayOnTop Then
        stat = SetWindowPos(mForm.hwnd, -1, 0, 0, 0, 0, 3)    'form始终在最前
    Else
        stat = SetWindowPos(mForm.hwnd, -2, 0, 0, 0, 0, 3)
    End If
End Sub

Public Sub OutputProgress(Percent As Double)
    If Percent >= 0 Then
        g_Statusbar.Panels(1).Text = "进度: " & CInt(Percent * 100) & "%"
    Else
        g_Statusbar.Panels(1).Text = ""
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'最小二乘法求斜率
'
'Y=a+bX, {(Xi,Yi)}(i=1..n)
'resolve : n*a+b*sum(Xi)=sum(Yi);  a*sum(Xi)+b*sum(Xi^2)=sum(Xi*Yi)
'
'answer: b=(sum(Yi)*sum(Xi)-n*sum(Xi*Yi))/((sum(Xi))^2-n*sum(Xi^2))
'        when sumX=0 and sumY=0, b=sum(Xi*Yi)/sum(Xi^2)
'
Public Function MinSquare_K(arrX() As Double, arrY() As Double, iFrom As Integer, iTo As Integer, _
                            k As Double, Degree90 As Boolean) As Boolean

    Dim i As Long, n As Long, lTo As Long, lFrom As Long
    Dim aveX As Double, aveY As Double, sumY As Double, sumX As Double
    Dim b1 As Double, b2 As Double
    On Error GoTo ErrHandler
    
    lFrom = iFrom:    lTo = iTo
    If lFrom >= lTo Or lTo > UBound(arrY) Or lFrom < LBound(arrY) Then
        MsgBox "arrX can not match arrY!"
        GoTo ErrHandler
    End If
    
    sumX = 0#: sumY = 0#
    n = lTo - lFrom + 1
    For i = lFrom To lTo
        If arrX(i) = -9999 Or arrY(i) = -9999 Then
            Degree90 = True
            k = 0#
            MinSquare_K = False
            Exit Function
        End If
        sumX = sumX + arrX(i)
        sumY = sumY + arrY(i)
    Next
    aveX = sumX / n:    aveY = sumY / n
    
    b1 = 0#:    b2 = 0#
    For i = lFrom To lTo
        b1 = b1 + (arrX(i) - aveX) * (arrY(i) - aveY)
        b2 = b2 + (arrX(i) - aveX) ^ 2
    Next
    
    If b2 = 0 Then
        Degree90 = True
        k = 0#
    Else
        Degree90 = False
        k = b1 / b2
    End If
    
    MinSquare_K = True
    Exit Function
ErrHandler:
    'MsgBox "最小二乘法求斜率时发生错误！", vbExclamation, "GeoVisor"
    MinSquare_K = False
End Function

Public Function Log10(X)
    Log10 = Log(X) / Log(10#)
End Function

Public Function ArcSin(X As Double) As Double
    If X = 1# Then
        ArcSin = PI / 2
    Else
        ArcSin = Atn(X / Sqr(-X ^ 2 + 1))
    End If
End Function

' Insert sort algorithm
Public Function InsertSort(arrElem As Variant, iLowBound As Integer, iUpBound As Integer) As Boolean
   Dim i As Integer, j As Integer, temp As Double
   Dim iLB As Integer, iUB As Integer
   
On Error GoTo ErrH
   InsertSort = False
   iLB = iLowBound ' LBound(arrElem)
   iUB = iUpBound 'UBound(arrElem)
      
   For i = iLB + 1 To iUB
      For j = i To iLB + 1 Step -1
         If arrElem(j) < arrElem(j - 1) Then
            temp = arrElem(j)
            arrElem(j) = arrElem(j - 1)
            arrElem(j - 1) = temp
         End If
      Next
   Next
   
   InsertSort = True
ErrH:

End Function

' non-recursive Quick Sort algorithm,
' sub-array in algorithm which length is smaller than THRESHOD will be leave for final InsertSort algorithm to be finished
Public Function QuickSort_NonRecursive(arrElem As Variant, iLowBound As Integer, iUpBound As Integer, Optional THRESHOLD As Integer = 9) As Boolean
'   Const THRESHOLD = 1  '9
   Dim stack() As Integer
   Dim listsize As Integer, iLB As Integer, iUB As Integer
   Dim top As Integer, pivotindex As Integer, iLeft As Integer, iRight As Integer, i As Integer, j As Integer
   Dim pivot As Double, temp As Double
   
On Error GoTo ErrH
   QuickSort_NonRecursive = False
   iLB = iLowBound ' LBound(arrElem)
   iUB = iUpBound ' UBound(arrElem)
   listsize = iUB - iLB + 1
   'initial stack
   ReDim stack(0 To listsize - 1)
   top = 0
   stack(top) = iLB
   top = top + 1
   stack(top) = iUB
   
   While (top > 0)   ' while there are unprocessed subarrays
      'pop stack
      j = stack(top)
      top = top - 1
      i = stack(top)
      top = top - 1
      
      'find pivot
      pivotindex = Int((i + j) / 2)
      pivot = arrElem(pivotindex)
      'stick pivot at end
      temp = arrElem(pivotindex)
      arrElem(pivotindex) = arrElem(j)
      arrElem(j) = temp
      
      'partition
      iLeft = i - 1
      iRight = j
      Do
         Do While iLeft < iUB
            iLeft = iLeft + 1
            If arrElem(iLeft) >= pivot Then Exit Do
         Loop 'While iLeft < iUB And arrElem(iLeft) < pivot
         Do While iRight > iLB
            iRight = iRight - 1
            If arrElem(iRight) <= pivot Then Exit Do
         Loop 'While iRight > iLB And arrElem(iRight) > pivot
'         If iLeft <= iUB And iRight > 0 Then
            temp = arrElem(iLeft)
            arrElem(iLeft) = arrElem(iRight)
            arrElem(iRight) = temp
'         End If
      Loop While iLeft < iRight
      'undo final swap
      temp = arrElem(iLeft)
      arrElem(iLeft) = arrElem(iRight)
      arrElem(iRight) = temp
      'put pivot value in place
      temp = arrElem(iLeft)
      arrElem(iLeft) = arrElem(j)
      arrElem(j) = temp
      
      'put new subarrays onto stack if they are small
      If iLeft - i > THRESHOLD Then  'left partition
         top = top + 1
         stack(top) = i
         top = top + 1
         stack(top) = iLeft - 1
      End If
      If j - iLeft > THRESHOLD Then  'right partition
         top = top + 1
         stack(top) = iLeft + 1
         top = top + 1
         stack(top) = j
      End If
   Wend
   
   If InsertSort(arrElem, iLB, iUB) Then
      QuickSort_NonRecursive = True
   End If
   Exit Function
ErrH:
   Debug.Print ""

End Function
