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
   StartUpPosition =   2  '��Ļ����
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      ToolTipText     =   "����GDS�ļ�"
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
Rem ��ʽ����ȫ������
Option Explicit

Rem ��������SetEvent
Private Declare Function SetEvent Lib "Kernel32.dll" (ByVal hEvent As Long) As Long

Rem ��������SendDataByFileMapping
' �˺���ΪDll�ڷ�װ�ĺ���
' ��һ������Ϊ: ��Ҫ�򿪵�ȫ�ַ��ű�־
' �ڶ�������Ϊ: �ַ���������ָ��
' ����ֵ:       ����ִ�гɹ�����0��ʧ�ܷ��ط�0
Private Declare Function SendDataByFileMapping Lib "Bin\Connector.dll" _
(ByVal pszSymbol As String, _
ByVal pszBuffer As String) As Long

Rem Text90������ʱ
Private Sub Text90_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Text90.text = "-�������ļ���������" Then Text90.text = ""
End Sub

Rem Text90ʧȥ����ʱ
Private Sub Text90_LostFocus()
    If Text90.text = "" Then Text90.text = "-�������ļ���������"
End Sub

Rem ��ť����
Private Sub Command1_Click()
    Rem ͨ����־g_nSelectedWindow���ض����TextBox
    Rem ����ƴ���ַ����Ļ�����
    Dim szText As String
    Rem ������Ҫ�洢���ַ�������
    Dim m_Text1 As String, m_Text2 As String, m_Text3 As String, m_Text4 As String, _
    m_Text5 As String, m_Text6 As String, m_Text7 As String, m_Text90 As String
    Rem ��ȡ�ı����ַ���
    m_Text1 = Text1.text
    m_Text2 = Text2.text
    m_Text3 = Text3.text
    m_Text4 = Text4.text
    m_Text5 = Text5.text
    m_Text6 = Text6.text
    m_Text7 = Text7.text
    m_Text90 = Text90.text
    If m_Text90 = "" Or m_Text90 = "-�������ļ���������" Then
        Label0.Caption = "�ļ�������Ϊ�գ�"
        Exit Sub
    End If

    Rem �����û�ѡ��Ĳ�ͬ�������в���
    Select Case g_nSelectedWindow
        Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Case 0
            Rem REC ����
            If m_Text1 = "" Or m_Text2 = "" Then
                Label0.Caption = "REC �������п�ֵ"
                Exit Sub
            End If
            szText = "0" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text90
        Case 1
            Rem MRR ΢��
            If m_Text1 = "" Or m_Text2 = "" Or m_Text3 = "" _
            Or m_Text4 = "" Or m_Text5 = "" Then
                Label0.Caption = "MRR �������п�ֵ"
                Exit Sub
            End If
            szText = "1" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text3 _
            & "|" & m_Text4 & "|" & m_Text5 & "|" & m_Text90
        Case 2
            Rem MZI ������¸�����
            If m_Text1 = "" Or m_Text2 = "" Or m_Text3 = "" _
            Or m_Text4 = "" Or m_Text5 = "" Or m_Text6 = "" _
            Or m_Text7 = "" Then
                Label0.Caption = "MZI �������п�ֵ"
                Exit Sub
            End If
            szText = "2" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text3 & _
            "|" & m_Text4 & "|" & m_Text5 & "|" & m_Text6 & "|" & m_Text7 & "|" & m_Text90
        Case 3
            Rem WG  ���ι�դ
            If m_Text1 = "" Or m_Text2 = "" Or m_Text3 = "" _
            Or m_Text4 = "" Or m_Text5 = "" Then
                Label0.Caption = "WG �������п�ֵ"
                Exit Sub
            End If
            szText = "3" & "|" & m_Text1 & "|" & m_Text2 & "|" & m_Text3 _
            & "|" & m_Text4 & "|" & m_Text5 & "|" & m_Text90
        Rem ��������������������������������������������������������������������������������������������
    End Select
    Rem д�빲���ڴ�
    Call SendDataByFileMapping("MemShared", szText)
    Rem ֪ͨHCGraphiCoreLite64.exe
    Call SetEvent(g_hEvent)
End Sub

Rem ��ʽ���û������¼�Text1~z
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

Rem ��ʽ���û������¼�Text90~99
Private Sub Text90_KeyPress(KeyAscii As Integer)
    Rem ����������Backspace���˳�����
    If KeyAscii = 8 Then Exit Sub
    Rem �������10���ַ���û��ѡ�������ַ���������Ч
    If Len(Text90) > &HA And Text90.SelLength = 0 Then KeyAscii = 0
    Rem ֻ��������ascii��
    If KeyAscii < 0 Or KeyAscii > 127 Then KeyAscii = 0
End Sub

Rem ��TextBox�༭�������ķ�װ
Private Sub HandleTextBoxEdit(ByVal text As TextBox, ByRef KeyAscii As Integer)
    Rem ���С�����ڵ�һλ��������Ч
    If Len(text) = 0 And KeyAscii = 46 Then KeyAscii = 0
    Rem ����������С���㲢��û������С�������˳�����
    If KeyAscii = 46 And Not CBool(InStr(text, ".")) Then Exit Sub
    Rem ����������Backspace���˳�����
    If KeyAscii = 8 Then Exit Sub
    Rem ��������ֵ����0-9֮����������Ч
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    Rem �������ֵ���Ƶ����99999
    If Len(text) = 5 And Not CBool(InStr(text, ".")) And text.SelLength = 0 Then KeyAscii = 0
    Rem С��������3λ
    If InStr(text, ".") <> 0 And Len(text) - InStr(text, ".") = 3 And text.SelLength = 0 Then KeyAscii = 0
    Rem �������ڿ�ʼʱ��������2��0
    If Len(text) = 1 And text.text = "0" And KeyAscii = 48 Then KeyAscii = 0
End Sub

Rem ������ع���
Private Sub Form_load()
    Rem �����ö�����
    If g_nPin = 1 Then Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Rem Text90�༭����ʾ��
    Text90.text = "-�������ļ���������"
    Rem ������ʽ����
    Select Case g_nSelectedWindow
        Rem ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Rem REC ����
        Case 0
            Rem ���ñ༭���Ƿ�ɼ�
            Text3.Visible = False
            Text4.Visible = False
            Text5.Visible = False
            Text6.Visible = False
            Text7.Visible = False
            Rem ����Label�Ƿ�ɼ�
            Label3.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Rem ����Label��ֵ
            Label0.Caption = "����"
            Label1.Caption = "Length"
            Label2.Caption = "Width"
            Rem ���ô��ڱ���
            Me.Caption = "Rectangular"
            Rem ����λͼ
            Picture1.Picture = LoadResPicture(101, vbResBitmap)
        Case 1
            Rem MRR ΢��
            Text6.Visible = False
            Text7.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label0.Caption = "΢��"
            Label1.Caption = "gap"
            Label2.Caption = "radius"
            Label3.Caption = "length_x"
            Label4.Caption = "length_y"
            Label5.Caption = "width"
            Me.Caption = "Micro-Ring"
            Picture1.Picture = LoadResPicture(102, vbResBitmap)
        Case 2
            Rem MZI ������¸�����
            Label1.Caption = "width1"
            Label2.Caption = "width2"
            Label3.Caption = "width_taper"
            Label4.Caption = "width_mmi"
            Label5.Caption = "delta_length"
            Label6.Caption = "length_y"
            Label7.Caption = "length_x"
            Label0.Caption = "������¸�����"
            Me.Caption = "Mach-Zehnder"
            Picture1.Picture = LoadResPicture(103, vbResBitmap)
        Case 3
            Rem RCG  ���ι�դ
            Text6.Visible = False
            Text7.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label1.Caption = "n_periods"
            Label2.Caption = "period"
            Label3.Caption = "fill_factor"
            Label4.Caption = "width_grating"
            Label5.Caption = "length_taper"
            Label0.Caption = "���ι�դ"
            Me.Caption = "waveguide"
            Picture1.Picture = LoadResPicture(104, vbResBitmap)
        Rem ��������������������������������������������������������������������������������������������
    End Select
End Sub

Rem ����ж�ع���
Private Sub Form_Unload(Cancel As Integer)
    FormMain.Show
End Sub
