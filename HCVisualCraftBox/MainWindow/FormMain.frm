VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择函数"
   ClientHeight    =   5475
   ClientLeft      =   7635
   ClientTop       =   4560
   ClientWidth     =   6015
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6015
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5400
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":81CF
            Key             =   "IcoS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":84E9
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7858
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "置顶"
            Object.ToolTipText     =   "置顶"
            ImageKey        =   "IcoS"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "帮助"
            Object.ToolTipText     =   "帮助"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Picture         =   "FormMain.frx":85FB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem 显式定义全部变量
Option Explicit

Rem 声明变量
Dim g_hWindowHandle As Long

Rem 工具栏回调函数
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "置顶"
            Rem 窗口置顶/取消置顶控制
            Call AlwaysOnTop
        Case "帮助"
            Rem 打开帮助文档
            Shell "Bin\hh.exe HELP\HCVisualCraftBoxHelp.chm"
            Dim hHWND As Long
            hHWND = FindWindowA("HH Parent", "HCVisualCraftBox Help")
            If hHWND <> 0 Then
                Call SwitchToThisWindow(hHWND, True)
            End If
    End Select
End Sub

Rem 窗口置顶/取消置顶处理函数
Private Sub AlwaysOnTop()
    If g_nPin = 0 Then
        Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or &H4000)
        g_nPin = 1
        Exit Sub
    End If
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE Or SWP_NOSIZE)
    g_nPin = 0
End Sub

Rem 用户双击ListView某行的事件
Private Sub ListView1_DblClick()
    Dim clickedRow As Integer
    Rem 获取双击的索引
    clickedRow = ListView1.SelectedItem.Index
    Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Rem 执行双击行后的操作
    Select Case clickedRow - 1
        Case 0
            Rem REC 矩形
            g_nSelectedWindow = 0
        Case 1
            Rem MRR 微环
            g_nSelectedWindow = 1
            
        Case 2
            Rem MZI 马赫曾德干涉仪
            g_nSelectedWindow = 2

        Case 3
            Rem WG  矩形光栅
            g_nSelectedWindow = 3
    Rem ——————————————————————————————————————————————
    End Select
    Call LoadWindow(Form1)
End Sub

Rem 根据参数选择某行
Private Sub SelectRow(ByVal rowIndex As Integer)
    Rem 如果选择了有效的一行
    If rowIndex >= 0 And rowIndex < ListView1.ListItems.Count Then
        Rem 将其变为整行选中
        ListView1.ListItems(rowIndex).Selected = True
    End If
End Sub

Rem 对添加一行数据的封装
Private Sub AddData(ByVal Itm As ListItem, ByVal Identifier As String, ByVal Name As String, ByVal FullName As String)
    Rem 添加column1的名称
    Set Itm = ListView1.ListItems.Add(1, Identifier, Name)
    Rem 使用SubItemIndex将SubItem与正确的ColumnHeader关联
    Itm.SubItems(ListView1.ColumnHeaders("Description").SubItemIndex) = FullName
End Sub

Rem 设置ListView
Private Sub SetListView()
    Rem 选中整行
    ListView1.FullRowSelect = True
    Rem 设置字体大小为 12
    ListView1.Font.Size = 12
    Rem 设置字体大小为微软雅黑
    ListView1.Font.Name = "Microsoft YaHei"
    Rem ListView设置为报表视图
    ListView1.View = lvwReport
    Rem 添加两列
    ListView1.ColumnHeaders.Add , "Function", "名称"
    ListView1.ColumnHeaders.Add , "Description", "描述"
    Rem 通过设置字体的大小改变行高度、设置指定行宽度
    ListView1.ColumnHeaders("Function").Width = 1000
    ListView1.ColumnHeaders("Description").Width = 6000
    Rem 向控件添加ListItem对象
    Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Rem 声明item对象
    Dim Item1 As ListItem
    Dim Item2 As ListItem
    Dim Item3 As ListItem
    Dim Item4 As ListItem
    Rem 添加数据 AddData(item对象,标识符,名称,描述)
    Call AddData(Item4, "waveguide", "WG", "waveguide grating(矩形光栅)")
    Call AddData(Item3, "Mach-Zehnder", "MZI", "Mach-Zehnder Interferometer(马赫曾德干涉仪)")
    Call AddData(Item2, "Micro-Ring", "MRR", "Micro-ring Resonator(微环)")
    Call AddData(Item1, "Rectangular", "RGC", "Rectangular Grating Coupler(矩形)")
    Rem ——————————————————————————————————————————————
    Rem 默认选择第一行 至少从1开始
    Call SelectRow(1)
End Sub

Rem Form_Main窗口加载过程函数
Private Sub Form_load()
    Rem 初始化置顶变量
    g_nPin = 0
    Rem 设置ListView
    Call SetListView
End Sub

Rem 加载传入的窗口
Private Sub LoadWindow(ByVal VBFormObject As Form)
    Rem 加载窗口
    Load VBFormObject
    Rem 显示窗口
    VBFormObject.Show
    Rem 隐藏主窗口
    FormMain.Hide
End Sub

Rem Form_Main窗口卸载过程函数
Private Sub Form_Unload(Cancel As Integer)
    Rem 杀死HCGraphiCoreLite64.exe
    Shell "cmd /c taskkill /F /IM " & "HCGraphiCoreLite64.exe", vbHide
    Cancel = 0
    Exit Sub
End Sub
