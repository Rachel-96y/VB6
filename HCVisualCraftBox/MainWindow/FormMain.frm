VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ����"
   ClientHeight    =   5475
   ClientLeft      =   7635
   ClientTop       =   4560
   ClientWidth     =   6015
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6015
   StartUpPosition =   2  '��Ļ����
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
            Key             =   "�ö�"
            Object.ToolTipText     =   "�ö�"
            ImageKey        =   "IcoS"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
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
Rem ��ʽ����ȫ������
Option Explicit

Rem ��������
Dim g_hWindowHandle As Long

Rem �������ص�����
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "�ö�"
            Rem �����ö�/ȡ���ö�����
            Call AlwaysOnTop
        Case "����"
            Rem �򿪰����ĵ�
            Shell "Bin\hh.exe HELP\HCVisualCraftBoxHelp.chm"
            Dim hHWND As Long
            hHWND = FindWindowA("HH Parent", "HCVisualCraftBox Help")
            If hHWND <> 0 Then
                Call SwitchToThisWindow(hHWND, True)
            End If
    End Select
End Sub

Rem �����ö�/ȡ���ö�������
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

Rem �û�˫��ListViewĳ�е��¼�
Private Sub ListView1_DblClick()
    Dim clickedRow As Integer
    Rem ��ȡ˫��������
    clickedRow = ListView1.SelectedItem.Index
    Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Rem ִ��˫���к�Ĳ���
    Select Case clickedRow - 1
        Case 0
            Rem REC ����
            g_nSelectedWindow = 0
        Case 1
            Rem MRR ΢��
            g_nSelectedWindow = 1
            
        Case 2
            Rem MZI ������¸�����
            g_nSelectedWindow = 2

        Case 3
            Rem WG  ���ι�դ
            g_nSelectedWindow = 3
    Rem ��������������������������������������������������������������������������������������������
    End Select
    Call LoadWindow(Form1)
End Sub

Rem ���ݲ���ѡ��ĳ��
Private Sub SelectRow(ByVal rowIndex As Integer)
    Rem ���ѡ������Ч��һ��
    If rowIndex >= 0 And rowIndex < ListView1.ListItems.Count Then
        Rem �����Ϊ����ѡ��
        ListView1.ListItems(rowIndex).Selected = True
    End If
End Sub

Rem �����һ�����ݵķ�װ
Private Sub AddData(ByVal Itm As ListItem, ByVal Identifier As String, ByVal Name As String, ByVal FullName As String)
    Rem ���column1������
    Set Itm = ListView1.ListItems.Add(1, Identifier, Name)
    Rem ʹ��SubItemIndex��SubItem����ȷ��ColumnHeader����
    Itm.SubItems(ListView1.ColumnHeaders("Description").SubItemIndex) = FullName
End Sub

Rem ����ListView
Private Sub SetListView()
    Rem ѡ������
    ListView1.FullRowSelect = True
    Rem ���������СΪ 12
    ListView1.Font.Size = 12
    Rem ���������СΪ΢���ź�
    ListView1.Font.Name = "Microsoft YaHei"
    Rem ListView����Ϊ������ͼ
    ListView1.View = lvwReport
    Rem �������
    ListView1.ColumnHeaders.Add , "Function", "����"
    ListView1.ColumnHeaders.Add , "Description", "����"
    Rem ͨ����������Ĵ�С�ı��и߶ȡ�����ָ���п��
    ListView1.ColumnHeaders("Function").Width = 1000
    ListView1.ColumnHeaders("Description").Width = 6000
    Rem ��ؼ����ListItem����
    Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Rem ����item����
    Dim Item1 As ListItem
    Dim Item2 As ListItem
    Dim Item3 As ListItem
    Dim Item4 As ListItem
    Rem ������� AddData(item����,��ʶ��,����,����)
    Call AddData(Item4, "waveguide", "WG", "waveguide grating(���ι�դ)")
    Call AddData(Item3, "Mach-Zehnder", "MZI", "Mach-Zehnder Interferometer(������¸�����)")
    Call AddData(Item2, "Micro-Ring", "MRR", "Micro-ring Resonator(΢��)")
    Call AddData(Item1, "Rectangular", "RGC", "Rectangular Grating Coupler(����)")
    Rem ��������������������������������������������������������������������������������������������
    Rem Ĭ��ѡ���һ�� ���ٴ�1��ʼ
    Call SelectRow(1)
End Sub

Rem Form_Main���ڼ��ع��̺���
Private Sub Form_load()
    Rem ��ʼ���ö�����
    g_nPin = 0
    Rem ����ListView
    Call SetListView
End Sub

Rem ���ش���Ĵ���
Private Sub LoadWindow(ByVal VBFormObject As Form)
    Rem ���ش���
    Load VBFormObject
    Rem ��ʾ����
    VBFormObject.Show
    Rem ����������
    FormMain.Hide
End Sub

Rem Form_Main����ж�ع��̺���
Private Sub Form_Unload(Cancel As Integer)
    Rem ɱ��HCGraphiCoreLite64.exe
    Shell "cmd /c taskkill /F /IM " & "HCGraphiCoreLite64.exe", vbHide
    Cancel = 0
    Exit Sub
End Sub
