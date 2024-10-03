Attribute VB_Name = "Module1"
Rem 显式定义全部变量
Option Explicit

Rem 定义常量
Private Const G_PASSWORD As String = "yangcai666"
Private Const G_IS_CLICK_COMMAND_BUTTON As String = "ClickCommandButton"
Private Const EVENT_ALL_ACCESS As Long = &H1F0003

Rem 定义模块级全局常量
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = &HFFFFFFFF
Public Const HWND_NOTOPMOST = &HFFFFFFFE

Rem 声明变量
Dim g_szFullPath As String
Dim g_hWindowHandle As Long
Dim g_nCreateProcessRet As Long

Rem 声明模块级全局变量
Public g_nPin As Long                '关键变量 决定窗口是否置顶
Public g_hEvent As Long             '关键变量 决定什么时候对缓冲区进行读写
Public g_nSelectedWindow As Long    '关键变量 决定将参数送往哪一个函数处理

Rem 声明模块级函数SetWindowPos
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal uFlags As Long) As Long

Rem 声明函数Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Rem 声明函数CreateEventA
Private Declare Function CreateEventA Lib "Kernel32.dll" _
(ByVal lpEventAttributes As Long, _
ByVal bManualReset As Boolean, _
ByVal bInitialState As Boolean, _
ByVal lpName As String) As Long

Rem 声明函数OpenEventA
Private Declare Function OpenEventA Lib "kernel32" _
(ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long

Rem 声明函数FindWindowA
Public Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Rem 声明函数SwitchToThisWindow
Public Declare Sub SwitchToThisWindow Lib "user32" _
(ByVal hWnd As Long, _
ByVal fAltTab As Long)

Rem 声明函数CreateProcessAndSendParameters
' 此函数为Dll内封装的函数
' 第一个参数为:密码,默认为"yangcai666"
' 第二个参数为Python的EXE路径(相对路径)
' 返回值为True or False 4字节
Public Declare Function CreateProcessHideAndSendParameters Lib "Bin\Connector.dll" _
(ByVal PASSWORD As String, _
ByVal PyExePate As String) As Long

Rem 进度条过程
Private Sub RunProgressBar()
    Rem 显示在窗体底部
    Form0.ProgressBar1.Align = vbAlignBottom
    Rem 进度条最小值
    Form0.ProgressBar1.Value = Form0.ProgressBar1.Min
    Dim i As Long
    For i = 0 To 100
        If i < 35 Then
            Sleep (30)
        ElseIf i > 75 Then
            Sleep (50)
        Else
            Sleep (5)
        End If
        Form0.ProgressBar1.Value = Form0.ProgressBar1.Min + i
    Next i
    Rem 进度条最大值
    Form0.ProgressBar1.Value = Form0.ProgressBar1.Max
End Sub

Rem 创建HCGraphiCoreLite64.exe
Private Sub CreateHCGraphiCoreLite64Module()
    g_nCreateProcessRet = CreateProcessHideAndSendParameters(G_PASSWORD, g_szFullPath)
    If g_nCreateProcessRet = -1 Then
        Call MsgBox("空指针异常或其它异常", vbOKOnly Or vbExclamation, "模块创建失败")
        End
    ElseIf g_nCreateProcessRet = -2 Then
        Call MsgBox("模块不存在或不是有效的可执行程序", vbOKOnly Or vbExclamation, "模块创建失败")
        End
    End If
End Sub

Rem 防多开
Private Sub Initialize()
Rem 判断是否有另一个实例
    If App.PrevInstance = True Then
        Rem 关闭窗口前把已存在的程序窗口切换到顶层
        g_hWindowHandle = FindWindowA("ThunderRT6FormDC", "选择函数")
        Call SwitchToThisWindow(g_hWindowHandle, False)
        End
    End If
End Sub

Rem 程序用户入口
Sub Main()
    Rem 防多开
    Call Initialize
    Rem 加载Form0
    Load Form0
    Form0.Show
    Rem 防止Form0界面被阻塞
    DoEvents
    Rem 单线程启动HCGraphiCoreLite64.exe
    Rem 创建事件对象
    g_hEvent = CreateEventA(0, False, False, G_IS_CLICK_COMMAND_BUTTON)
    If g_hEvent = 0 Then
        Call MsgBox("创建事件失败", vbOKOnly Or vbExclamation, "失败")
        End
    End If
    g_szFullPath = App.Path & "\Bin\HCGraphiCoreLite64.exe"
    Call CreateHCGraphiCoreLite64Module
    Rem 进度条动画
    Call RunProgressBar
    DoEvents
    Rem 卸载Form0并显示主窗口
    Unload Form0
    Load FormMain
    FormMain.Show
End Sub
