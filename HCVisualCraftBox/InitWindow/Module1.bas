Attribute VB_Name = "Module1"
Rem ��ʽ����ȫ������
Option Explicit

Rem ���峣��
Private Const G_PASSWORD As String = "yangcai666"
Private Const G_IS_CLICK_COMMAND_BUTTON As String = "ClickCommandButton"
Private Const EVENT_ALL_ACCESS As Long = &H1F0003

Rem ����ģ�鼶ȫ�ֳ���
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = &HFFFFFFFF
Public Const HWND_NOTOPMOST = &HFFFFFFFE

Rem ��������
Dim g_szFullPath As String
Dim g_hWindowHandle As Long
Dim g_nCreateProcessRet As Long

Rem ����ģ�鼶ȫ�ֱ���
Public g_nPin As Long                '�ؼ����� ���������Ƿ��ö�
Public g_hEvent As Long             '�ؼ����� ����ʲôʱ��Ի��������ж�д
Public g_nSelectedWindow As Long    '�ؼ����� ����������������һ����������

Rem ����ģ�鼶����SetWindowPos
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal uFlags As Long) As Long

Rem ��������Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Rem ��������CreateEventA
Private Declare Function CreateEventA Lib "Kernel32.dll" _
(ByVal lpEventAttributes As Long, _
ByVal bManualReset As Boolean, _
ByVal bInitialState As Boolean, _
ByVal lpName As String) As Long

Rem ��������OpenEventA
Private Declare Function OpenEventA Lib "kernel32" _
(ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long

Rem ��������FindWindowA
Public Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Rem ��������SwitchToThisWindow
Public Declare Sub SwitchToThisWindow Lib "user32" _
(ByVal hWnd As Long, _
ByVal fAltTab As Long)

Rem ��������CreateProcessAndSendParameters
' �˺���ΪDll�ڷ�װ�ĺ���
' ��һ������Ϊ:����,Ĭ��Ϊ"yangcai666"
' �ڶ�������ΪPython��EXE·��(���·��)
' ����ֵΪTrue or False 4�ֽ�
Public Declare Function CreateProcessHideAndSendParameters Lib "Bin\Connector.dll" _
(ByVal PASSWORD As String, _
ByVal PyExePate As String) As Long

Rem ����������
Private Sub RunProgressBar()
    Rem ��ʾ�ڴ���ײ�
    Form0.ProgressBar1.Align = vbAlignBottom
    Rem ��������Сֵ
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
    Rem ���������ֵ
    Form0.ProgressBar1.Value = Form0.ProgressBar1.Max
End Sub

Rem ����HCGraphiCoreLite64.exe
Private Sub CreateHCGraphiCoreLite64Module()
    g_nCreateProcessRet = CreateProcessHideAndSendParameters(G_PASSWORD, g_szFullPath)
    If g_nCreateProcessRet = -1 Then
        Call MsgBox("��ָ���쳣�������쳣", vbOKOnly Or vbExclamation, "ģ�鴴��ʧ��")
        End
    ElseIf g_nCreateProcessRet = -2 Then
        Call MsgBox("ģ�鲻���ڻ�����Ч�Ŀ�ִ�г���", vbOKOnly Or vbExclamation, "ģ�鴴��ʧ��")
        End
    End If
End Sub

Rem ���࿪
Private Sub Initialize()
Rem �ж��Ƿ�����һ��ʵ��
    If App.PrevInstance = True Then
        Rem �رմ���ǰ���Ѵ��ڵĳ��򴰿��л�������
        g_hWindowHandle = FindWindowA("ThunderRT6FormDC", "ѡ����")
        Call SwitchToThisWindow(g_hWindowHandle, False)
        End
    End If
End Sub

Rem �����û����
Sub Main()
    Rem ���࿪
    Call Initialize
    Rem ����Form0
    Load Form0
    Form0.Show
    Rem ��ֹForm0���汻����
    DoEvents
    Rem ���߳�����HCGraphiCoreLite64.exe
    Rem �����¼�����
    g_hEvent = CreateEventA(0, False, False, G_IS_CLICK_COMMAND_BUTTON)
    If g_hEvent = 0 Then
        Call MsgBox("�����¼�ʧ��", vbOKOnly Or vbExclamation, "ʧ��")
        End
    End If
    g_szFullPath = App.Path & "\Bin\HCGraphiCoreLite64.exe"
    Call CreateHCGraphiCoreLite64Module
    Rem ����������
    Call RunProgressBar
    DoEvents
    Rem ж��Form0����ʾ������
    Unload Form0
    Load FormMain
    FormMain.Show
End Sub
