VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form Form0 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "启动"
   ClientHeight    =   4575
   ClientLeft      =   6825
   ClientTop       =   3180
   ClientWidth     =   7425
   Icon            =   "Form0.frx":0000
   LinkTopic       =   "Form0"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   4575
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   4695
      Left            =   0
      MousePointer    =   11  'Hourglass
      Picture         =   "Form0.frx":81CF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem 显式定义全部变量
Option Explicit
