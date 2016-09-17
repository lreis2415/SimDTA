VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SimDTA"
   ClientHeight    =   8085
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6450
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580.411
   ScaleMode       =   0  'User
   ScaleWidth      =   6056.883
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   1560
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1053.5
      ScaleMode       =   0  'User
      ScaleWidth      =   1053.5
      TabIndex        =   1
      Top             =   120
      Width           =   1560
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&OK)"
      Default         =   -1  'True
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Top             =   7620
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "系统信息(&S)..."
      Height          =   345
      Left            =   3780
      TabIndex        =   2
      Top             =   7620
      Width           =   1485
   End
   Begin VB.Label lblMailAddr 
      Caption         =   "Label1"
      Height          =   2775
      Left            =   180
      TabIndex        =   8
      Top             =   4620
      Width           =   6015
   End
   Begin VB.Label lblAcknowledgement 
      Caption         =   "Label1"
      Height          =   2235
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   5955
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   6
      X1              =   -281.715
      X2              =   6676.657
      Y1              =   5135.221
      Y2              =   5135.221
   End
   Begin VB.Label lblDescription 
      Caption         =   "秦承志 / QIN Cheng-Zhi  (email: qincz@lreis.ac.cn)"
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1920
      TabIndex        =   3
      Top             =   1260
      Width           =   4245
   End
   Begin VB.Label lblTitle 
      Caption         =   "SimDTA: Simple/Simulation Digital Terrain Analysis (developed with VB6.0)"
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   1920
      TabIndex        =   5
      Top             =   180
      Width           =   4065
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本: 0.94  (released on 2008/7/1)"
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   3765
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "这是一个明信片软件 / This is a postcard software."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   4320
      Width           =   5565
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   -169.029
      X2              =   6789.343
      Y1              =   1283.805
      Y2              =   1283.805
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -140.858
      X2              =   6817.515
      Y1              =   2898.915
      Y2              =   2898.915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   -154.944
      X2              =   6733
      Y1              =   1283.805
      Y2              =   1283.805
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -126.772
      X2              =   6817.515
      Y1              =   2898.915
      Y2              =   2898.915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   -436.659
      X2              =   6507.627
      Y1              =   5135.221
      Y2              =   5135.221
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   -450.745
      X2              =   6507.627
      Y1              =   5135.221
      Y2              =   5135.221
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字 ROOT 类型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' 独立的空的终结字符串
Const REG_DWORD = 4                      ' 32位数字

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    'lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
    lblDescription.Caption = "秦承志 / QIN Cheng-Zhi" + Chr(10) + "(email: qincz@lreis.ac.cn)"
    
    lblVersion.Caption = C_VersionInfo
        
    lblAcknowledgement.Caption = "致谢 / Acknowledgements" + Chr(10) + Chr(10) _
         + "国家自然科学基金青年基金项目（40501056）；" _
         + "中国科学院知识创新工程重要方向项目群项目子课题（KZCX2-YW-Q10-1-5）；" _
         + "资源与环境信息系统国家重点实验室自主创新资助" + Chr(10) + Chr(10) _
         & "Supported by the National Natural Science Foundation of China (Project Number: 40501056), " _
         + "Knowledge Innovation Program of the Chinese Academy of Sciences (KZCX2-YW-Q10-1-5)," _
         + "Innovation from the State Key Laboratory of Resources and Environmental Information Systems"
    
    lblMailAddr.Caption = "给我寄张明信片告诉这个软件对你有用就可以了。" + Chr(10) + Chr(10) _
         & "北京 安定门外 大屯路 甲11号" & vbCrLf & "中国科学院地理科学与资源研究所" + Chr(10) _
         & "资源与环境信息系统国家重点实验室" + Chr(10) & "邮编：100101" & vbCrLf & vbCrLf _
         & "State Key Laboratory of Resources & Environmental Information System," & vbCrLf _
         & "Institute of Geographical Sciences & Natural Resources Research," & vbCrLf & "Chinese Academy of Sciences" & vbCrLf _
         & "11A Datun Road, Anwai, Beijing 100101, PR China" & vbCrLf & "PO Number：9719"
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 试图从注册表中获得系统信息程序的路径及名称...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' 试图仅从注册表中获得系统信息程序的路径...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 已知32位文件版本的有效位置
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' 错误 - 文件不能被找到...
        Else
            GoTo SysInfoErr
        End If
    ' 错误 - 注册表相应条目不能被找到...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "此时系统信息不可用", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 打开的注册表关键字句柄
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字值的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键自变量的尺寸
    '------------------------------------------------------------
    ' 打开 {HKEY_LOCAL_MACHINE...} 下的 RegKey
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 外接程序空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null 被找到,从字符串中分离出来
    Else                                                    ' WinNT 没有空终结字符串...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null 没有被找到, 分离字符串
    End If
    '------------------------------------------------------------
    ' 决定转换的关键字的值类型...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 搜索数据类型...
    Case REG_SZ                                             ' 字符串注册关键字数据类型
        KeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节的注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 将每位进行转换
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 生成值字符。 By Char。
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 转换四字节的字符为字符串
    End Select
    
    GetKeyValue = True                                      ' 返回成功
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:      ' 错误发生后将其清除...
    KeyVal = ""                                             ' 设置返回值到空字符串
    GetKeyValue = False                                     ' 返回失败
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function
