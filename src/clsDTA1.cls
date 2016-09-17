VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDTA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim iCols As Integer, iRows As Integer
Dim dCellSize As Double
Dim vDataOut As Variant
Dim vDataDEM As Variant
Dim dNoData As Double
Dim demLyr As mapobjects2.ImageLayer

Public Function ComputeDTA1(strTerrainFeat As String) As Boolean
    Select Case strTerrainFeat
    Case ConstDTA1Cs
    
    End Select
    
    ComputeDTA1 = True
    Exit Function
ErrH:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "DTA_MO"
    End If
    ComputeDTA1 = False
End Function

' Cs=(Sum((Zi-Zaverage)/din))/n
Private Function SurfaceCurvatureIndex() As Boolean
    On Error GoTo ErrH
    Dim x As Integer, y As Integer, x1 As Integer, y1 As Integer, i As Integer, j As Integer
    Dim iHalfWinCells As Integer
    Dim dSum As Double, dAve As Double, n As Integer
        
    For x = 0 To iCols - 1
        For y = 0 To iRows - 1
            If vDataDEM(x, y) <> dNoData Then
                dSum = 0#
                n = 0
                For i = -iHalfWinCells To iHalfWinCells
                    x1 = x + i
                    For j = -iHalfWinCells To iHalfWinCells
                        y1 = y + j
                        If IsValidCell(x1, y1) And (x <> x1 Or y <> y1) Then
                            If vDataDEM(x1, y1) <> dNoData Then
                                dSum = dSum + vDataDEM(x1, y1)
                                n = n + 1
                            End If
                        End If
                    Next
                Next
                If n = 0 Then
                    vDataOut(x, y) = dNoData
                Else
                    dAve = dSum / n
                    dSum = 0
                    For i = -iHalfWinCells To iHalfWinCells
                        x1 = x + i
                        For j = -iHalfWinCells To iHalfWinCells
                            y1 = y + j
                            If IsValidCell(x1, y1) And (x <> x1 Or y <> y1) Then
                                If vDataDEM(x1, y1) <> dNoData Then
                                    dSum = dSum + (vDataDEM(x1, y1) - dAve) / (dCellSize * Sqr((x1 - x) ^ 2 + (y1 - y) ^ 2))
                                End If
                            End If
                        Next
                    Next
                    vDataOut(x, y) = dSum / n
                End If
            Else
                vDataOut(x, y) = dNoData
            End If
        Next
    Next
    SurfaceCurvatureIndex = True
    Exit Function
ErrH:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "DTA_MO"
    End If
    SurfaceCurvatureIndex = False
End Function

Private Function IsValidCell(x As Integer, y As Integer) As Boolean
    If x >= 0 And x < iCols And y >= 0 And y < iRows Then
        IsValidCell = True
    Else
        IsValidCell = False
    End If
End Function

Private Function PrepareOutput() As Boolean
    '
End Function

Private Function PrepareDEMData() As Boolean
    '
    
    iCols = demLyr.Columns: iRows = demLyr.Rows
End Function


Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    '
End Sub
