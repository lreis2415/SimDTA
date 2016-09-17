Attribute VB_Name = "modCommonDialog"
Global AppPath As String

Public Function TrimNulls(strString As String) As String
   
   Dim l As Long
   
   l = InStr(1, strString, Chr(0))
   
   If l = 1 Then
      TrimNulls = ""
   ElseIf l > 0 Then
      TrimNulls = Left$(strString, l - 1)
   Else
      TrimNulls = strString
   End If
   
End Function

Public Sub InitSysVar()
    AppPath = App.Path
End Sub


Public Function GetFileName(dlgFile As CommonDialog, _
                            Optional boolOpen = True, _
                            Optional boolAppPath = False, _
                            Optional strFileType = ".txt") As String
                            
    dlgFile.CancelError = True
    On Error GoTo ErrHandler
    'On Error Resume Next
    
    dlgFile.CancelError = True
    dlgFile.Flags = cdlOFNHideReadOnly
    
    Select Case LCase(strFileType)
        Case ".txt"
            dlgFile.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        Case ".mxd"
            dlgFile.Filter = _
                "ArcMap Files (*.mxd)|*.mxd|All Files (*.*)|*.*"
        Case ".jpg", ".bmp"
            dlgFile.Filter = _
                "Picture Files (*.bmp)|*.bmp|Picture Files (*.jpg)|*.jpg|All Files (*.*)|*.*"
        Case "layer"  'shapefile, or coverage, or image
            dlgFile.Filter = _
                "Grid Coverage (hdr.adf)|hdr.adf" _
                & "|Shapefile Layers (*.shp)|*.shp" _
                & "|Polygon/Point Coverage (.pat)|pat.adf" _
                & "|Line Coverage (.aat)|aat.adf" _
                & "|基础注记类 Coverage (.txt)|txt.adf" _
                & "|图像 (*.tif; *.bmp; *.jpg)|*.tif; *.bmp; *.jpg" _
                & "|All Files (*.*)|*.*"
        Case Else
            dlgFile.Filter = "file (*" & strFileType & ")|*" & strFileType & "|All Files (*.*)|*.*"
    End Select
    dlgFile.FilterIndex = 1
    If boolAppPath Then dlgFile.InitDir = App.Path
    If boolOpen Then
        dlgFile.ShowOpen
    Else
        dlgFile.ShowSave
    End If
    'If Err <> vbCancel Then
        GetFileName = dlgFile.FileName
    'Else
    '    GetFileName = ""
    'End If
    Exit Function
ErrHandler:
    ' 用户按了“取消”按钮
    GetFileName = ""
End Function

' return -1 means NO Choice
Public Function GetColorFromPannel(dlgColor As CommonDialog, Optional strTitle = "", _
                                    Optional initcolor As OLE_COLOR = -1) As OLE_COLOR
    dlgColor.CancelError = True
    On Error GoTo ErrHandler
    
    dlgColor.Flags = cdlCCRGBInit
    If initcolor <> -1 Then dlgColor.color = initcolor
    If strTitle <> "" Then
        dlgColor.DialogTitle = strTitle
    End If
    dlgColor.ShowColor
    GetColorFromPannel = dlgColor.color
    Exit Function
ErrHandler:
    GetColorFromPannel = -1
End Function

'if CANCEL, then return font as nothing, color as -1
'Call example:
'    Dim font As StdFont
'    Set font = GetFontFromDialog(dlgFile, , color)
'    If Not (font Is Nothing) Then
'        MsgBox font
'    End If
'
Public Function GetFontFromDialog(dlgFont As CommonDialog, Optional strTitle = "", _
                                  Optional color As OLE_COLOR) As StdFont
    Dim font As New StdFont
    
    dlgFont.CancelError = True
    On Error GoTo ErrHandler
    dlgFont.Flags = cdlCFBoth Or cdlCFEffects  ' choose Screen&Printer font AND choose color
    If strTitle <> "" Then
        dlgFont.DialogTitle = strTitle
    End If
    dlgFont.ShowFont
    
    font.Name = dlgFont.FontName
    font.Size = dlgFont.FontSize
    font.Bold = dlgFont.FontBold
    font.Italic = dlgFont.FontItalic
    font.Underline = dlgFont.FontItalic
    font.Strikethrough = dlgFont.FontStrikethru
    color = dlgFont.color
    
    Set GetFontFromDialog = font
    Exit Function
ErrHandler:
    Set GetFontFromDialog = Nothing
    color = -1
End Function

Public Sub GetHelpDialod(dlgHelp As CommonDialog, strHelpFile As String)
    dlgHelp.CancelError = True
    On Error Resume Next
    dlgHelp.HelpCommand = cdlHelpForceFile
    dlgHelp.HelpFile = strHelpFile
    dlgHelp.ShowHelp
    Exit Sub
End Sub


'
' Usage:
'       for i=1 to iCopies
'           ...
'       next i
'
Public Function GetPrintPara(dlgPrint As CommonDialog, _
                iFromPage As Integer, iToPage As Integer, Optional strTitle = "", _
                Optional iCopies = 1, Optional Orientation = cdlPortrait) As Boolean
    dlgPrint.CancelError = True
    On Error GoTo ErrHandler
    
    dlgPrint.Copies = iCopies
    dlgPrint.Orientation = Orientation
    dlgPrint.ShowPrinter
    
    iFromPage = dlgPrint.FromPage
    iToPage = dlgPrint.ToPage
    iCopies = dlgPrint.Copies
    Orientation = dlgPrint.Orientation
    
    GetPrintPara = True
    Exit Function
ErrHandler:
    GetPrintPara = False
End Function


