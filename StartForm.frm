VERSION 5.00
Begin VB.Form StartForm 
   Appearance      =   0  '���
   BackColor       =   &H00000000&
   BorderStyle     =   0  '����
   Caption         =   "���"
   ClientHeight    =   13080
   ClientLeft      =   -3360
   ClientTop       =   -2775
   ClientWidth     =   14700
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   13080
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Image BackGroundImage 
      Height          =   495
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   BackGroundImage.Picture = LoadPicture("���.jpg")
   BackGroundImage.Left = 0
   BackGroundImage.Top = 0
   BackGroundImage.Width = StartForm.Width
   BackGroundImage.Height = StartForm.Height
   MainForm.Show 1
   End
End Sub
