VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Name"
   ClientHeight    =   8055
   ClientLeft      =   6615
   ClientTop       =   2835
   ClientWidth     =   8280
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000B&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleMode       =   0  'User
   ScaleWidth      =   8280
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   0
      Picture         =   "Form1.frx":81CF
      ScaleHeight     =   3510
      ScaleMode       =   0  'User
      ScaleWidth      =   8475
      TabIndex        =   17
      Top             =   4680
      Width           =   8475
   End
   Begin VB.TextBox Text90 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3167
      TabIndex        =   16
      Top             =   3780
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   489
      TabIndex        =   10
      Top             =   3105
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5927
      TabIndex        =   9
      Top             =   1905
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3167
      TabIndex        =   8
      Top             =   1905
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   489
      TabIndex        =   7
      Top             =   1905
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5927
      TabIndex        =   6
      Top             =   825
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6158
      MaskColor       =   &H00808080&
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":69051
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "生成GDS文件"
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3167
      TabIndex        =   1
      Top             =   825
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   527
      TabIndex        =   0
      Top             =   825
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label0 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3285
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   4995
      Left            =   0
      Picture         =   "Form1.frx":7741D
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   8595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem 显式定义全部变量
Option Explicit

Rem 声明函数SetEvent
Private Declare Function SetEvent Lib "Kernel32.dll" (ByVal hEvent As Long) As Long

Rem 声明函数SendDataByFileMapping
' 此函数为Dll内封装的函数
' 第一个参数为: 需要打开的全局符号标志
' 第二个参数为: 字符串缓冲区指针
' 返回值:       函数执行成功返回0，失败返回非0
Private Declare Function SendDataByFileMapping Lib "Bin\Connector.dll" _
(ByVal pszSymbol As String, _
ByVal pszBuffer As String) As Long

Rem Text90被按下时
Private Sub Text90_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Text90.text = "-请输入文件名！！！" Then Text90.text = ""
End Sub

Rem Text90失去焦点时
Private Sub Text90_LostFocus()
    If Text90.text = "" Then Text90.text = "-请输入文件名！！！"
End Sub

Rem 按钮过程
Private Sub Command1_Click()
    Rem 通过标志g_nSelectedWindow隐藏多余的TextBox
    Rem 用于拼接字符串的缓冲区
    Dim szText As String
    Rem 声明需要存储的字符串常量
    Dim m_Text1 As String, m_Text2 As String, m_Text3 As String, m_Text4 As String, _
    m_Text5 As String, m_Text6 As String, m_Text7 As String, m_Text90 As String
    Rem 获取文本框字符串
    m_Text1 = Text1.text
    m_Text2 = Text2.text
    m_Text3 = Text3.text
    m_Text4 = Text4.text
    m_Text5 = Text5.text
    m_Text6 = Text6.text
    m_Text7 = Text7.text
    m_Text90 = Text90.text
    If m_Text90 = "" Or m_Text90 = "-请输入文件名！！！" Then
        Label0.Caption = "文件名不能为空！"
        Exit Sub
    End If

    Rem 根据用户选择的不同函数进行操作
    Select Case g_nSelectedWindow
        Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Case 0
            Rem REC 矩形
            If m_Text1 = "" Or m_Text2 = "" Then
                Label0.Caption = "REC 不允许有空值"
                Exit Sub
            End If
            szText = "0" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text90
        Case 1
            Rem MRR 微环
            If m_Text1 = "" Or m_Text2 = "" Or m_Text3 = "" _
            Or m_Text4 = "" Or m_Text5 = "" Then
                Label0.Caption = "MRR 不允许有空值"
                Exit Sub
            End If
            szText = "1" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text3 _
            & "|" & m_Text4 & "|" & m_Text5 & "|" & m_Text90
        Case 2
            Rem MZI 马赫曾德干涉仪
            If m_Text1 = "" Or m_Text2 = "" Or m_Text3 = "" _
            Or m_Text4 = "" Or m_Text5 = "" Or m_Text6 = "" _
            Or m_Text7 = "" Then
                Label0.Caption = "MZI 不允许有空值"
                Exit Sub
            End If
            szText = "2" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text3 & _
            "|" & m_Text4 & "|" & m_Text5 & "|" & m_Text6 & "|" & m_Text7 & "|" & m_Text90
        Case 3
            Rem WG  矩形光栅
            If m_Text1 = "" Or m_Text2 = "" Or m_Text3 = "" _
            Or m_Text4 = "" Or m_Text5 = "" Then
                Label0.Caption = "WG 不允许有空值"
                Exit Sub
            End If
            szText = "3" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text3 _
            & "|" & m_Text4 & "|" & m_Text5 & "|" & m_Text90
        Rem ——————————————————————————————————————————————
    End Select
    Rem 写入共享内存
    Call SendDataByFileMapping("MemShared", szText)
    Rem 通知HCGraphiCoreLite64.exe
    Call SetEvent(g_hEvent)
End Sub

Rem 格式化用户输入事件Text1~z
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Call HandleTextBoxEdit(Text1, KeyAscii)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Call HandleTextBoxEdit(Text2, KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Call HandleTextBoxEdit(Text3, KeyAscii)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    Call HandleTextBoxEdit(Text4, KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    Call HandleTextBoxEdit(Text5, KeyAscii)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    Call HandleTextBoxEdit(Text6, KeyAscii)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    Call HandleTextBoxEdit(Text7, KeyAscii)
End Sub

Rem 格式化用户输入事件Text90~99
Private Sub Text90_KeyPress(KeyAscii As Integer)
    Rem 如果输入的是Backspace则退出过程
    If KeyAscii = 8 Then Exit Sub
    Rem 如果大于10个字符且没有选择其它字符则输入无效
    If Len(Text90) > &HA And Text90.SelLength = 0 Then KeyAscii = 0
    Rem 只允许输入ascii码
    If KeyAscii < 0 Or KeyAscii > 127 Then KeyAscii = 0
End Sub

Rem 对TextBox编辑框处理函数的封装
Private Sub HandleTextBoxEdit(ByVal text As TextBox, ByRef KeyAscii As Integer)
    Rem 如果小数点在第一位则输入无效
    If Len(text) = 0 And KeyAscii = 46 Then KeyAscii = 0
    Rem 如果输入的是小数点并且没有其它小数点则退出过程
    If KeyAscii = 46 And Not CBool(InStr(text, ".")) Then Exit Sub
    Rem 如果输入的是Backspace则退出过程
    If KeyAscii = 8 Then Exit Sub
    Rem 如果输入的值不在0-9之间则输入无效
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    Rem 将输入的值限制到最大99999
    If Len(text) = 5 And Not CBool(InStr(text, ".")) And text.SelLength = 0 Then KeyAscii = 0
    Rem 小数点后最多3位
    If InStr(text, ".") <> 0 And Len(text) - InStr(text, ".") = 3 And text.SelLength = 0 Then KeyAscii = 0
    Rem 不允许在开始时连续输入2个0
    If Len(text) = 1 And text.text = "0" And KeyAscii = 48 Then KeyAscii = 0
End Sub

Rem 窗体加载过程
Private Sub Form_load()
    Rem 窗口置顶控制
    If g_nPin = 1 Then Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Rem Text90编辑框提示语
    Text90.text = "-请输入文件名！！！"
    Rem 窗口样式设置
    Select Case g_nSelectedWindow
        Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Rem REC 矩形
        Case 0
            Rem 设置编辑框是否可见
            Text3.Visible = False
            Text4.Visible = False
            Text5.Visible = False
            Text6.Visible = False
            Text7.Visible = False
            Rem 设置Label是否可见
            Label3.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Rem 设置Label的值
            Label0.Caption = "矩形"
            Label1.Caption = "Length"
            Label2.Caption = "Width"
            Rem 设置窗口标题
            Me.Caption = "Rectangular"
            Rem 加载位图
            Picture1.Picture = LoadResPicture(101, vbResBitmap)
        Case 1
            Rem MRR 微环
            Text6.Visible = False
            Text7.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label0.Caption = "微环"
            Label1.Caption = "gap"
            Label2.Caption = "radius"
            Label3.Caption = "length_x"
            Label4.Caption = "length_y"
            Label5.Caption = "width"
            Me.Caption = "Micro-Ring"
            Picture1.Picture = LoadResPicture(102, vbResBitmap)
        Case 2
            Rem MZI 马赫曾德干涉仪
            Label1.Caption = "width1"
            Label2.Caption = "width2"
            Label3.Caption = "width_taper"
            Label4.Caption = "width_mmi"
            Label5.Caption = "delta_length"
            Label6.Caption = "length_y"
            Label7.Caption = "length_x"
            Label0.Caption = "马赫曾德干涉仪"
            Me.Caption = "Mach-Zehnder"
            Picture1.Picture = LoadResPicture(103, vbResBitmap)
        Case 3
            Rem RCG  矩形光栅
            Text6.Visible = False
            Text7.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label1.Caption = "n_periods"
            Label2.Caption = "period"
            Label3.Caption = "fill_factor"
            Label4.Caption = "width_grating"
            Label5.Caption = "length_taper"
            Label0.Caption = "矩形光栅"
            Me.Caption = "waveguide"
            Picture1.Picture = LoadResPicture(104, vbResBitmap)
        Rem ——————————————————————————————————————————————
    End Select
End Sub

Rem 窗体卸载过程
Private Sub Form_Unload(Cancel As Integer)
    FormMain.Show
End Sub
