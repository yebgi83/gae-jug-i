VERSION 5.00
Begin VB.Form WaitForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  '����
   Caption         =   "������Դϴ�."
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   Icon            =   "WaitForm.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin VB.Label Label�����ð� 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�� ����"
      BeginProperty Font 
         Name            =   "HY����B"
         Size            =   26.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1395
      TabIndex        =   2
      Top             =   960
      Width           =   1785
   End
   Begin VB.Shape BorderShape 
      DrawMode        =   1  '����
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "?�����Ŀ� ��������� ����մϴ�."
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��� ����� �Ͻ������� ������ ��ȣ�� �����մϴ�."
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4380
   End
End
Attribute VB_Name = "WaitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ���ð� = 3

Dim �������ð� As Long

Private Sub Form_Load()
   BorderShape.Top = Me.Top
   BorderShape.Left = Me.Left
   BorderShape.Width = Me.Width
   BorderShape.Height = Me.Height
   Label1(1).Caption = ���ð� & "�����Ŀ� ��������� ����մϴ�."
   �������ð� = GetTickCount
End Sub

Private Sub Timer1_Timer()
   Dim �����ð� As Long
   �����ð� = Int((GetTickCount - �������ð�) / 1000)
   
   Label�����ð� = ���ð� - �����ð� & "�� ����"
   
   If �����ð� = ���ð� Then
      Unload Me
   End If
End Sub
