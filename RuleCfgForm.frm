VERSION 5.00
Begin VB.Form RuleCfgForm 
   BorderStyle     =   1  '���� ����
   Caption         =   "��Ģ ����"
   ClientHeight    =   4815
   ClientLeft      =   7830
   ClientTop       =   5880
   ClientWidth     =   3600
   Icon            =   "RuleCfgForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.Frame Frame4 
      Caption         =   "���ð� ����"
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   3375
      Begin VB.TextBox Txt�������� 
         Height          =   270
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(�и���)"
         Height          =   180
         Index           =   5
         Left            =   2520
         TabIndex        =   23
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "�����˹�Ģ"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   3375
      Begin VB.TextBox Txt�����̷�� 
         Height          =   270
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "�����̷��"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(�и���)"
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   18
         Top             =   285
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "���ʽ�"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3375
      Begin VB.TextBox Txt2�����ʽ� 
         Height          =   270
         Left            =   1440
         TabIndex        =   13
         Top             =   675
         Width           =   975
      End
      Begin VB.TextBox Txt�������ʽ� 
         Height          =   270
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "2�����ʽ�"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(�и���)"
         Height          =   180
         Index           =   3
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "�������ʽ�"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(�и���)"
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   285
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�⺻ ��Ģ"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Txt���ѽð� 
         Height          =   270
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Txt�ִ�����Ƚ�� 
         Height          =   270
         Left            =   1440
         TabIndex        =   3
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "���ѽð�"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(��)"
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   7
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "�ִ�����Ƚ��"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(ȸ)"
         Height          =   180
         Index           =   1
         Left            =   2760
         TabIndex        =   4
         Top             =   720
         Width           =   330
      End
   End
   Begin VB.CommandButton Cmd��� 
      Caption         =   "���"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "RuleCfgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd����_Click()
   Set���ѽð� Val(Txt���ѽð�.Text)
   Set�ִ�����Ƚ�� Val(Txt�ִ�����Ƚ��.Text)
   Set�������ʽ�����ġ Val(Txt�������ʽ�.Text)
   Set2�����ʽ�����ġ Val(Txt2�����ʽ�.Text)
   Set�������Ⱑ��ġ Val(Txt��������.Text)
   Set�����̷�Ⱑ��ġ Val(Txt�����̷��.Text)
   Unload Me
End Sub

Private Sub Cmd���_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Txt���ѽð�.Text = Get���ѽð�
   Txt�ִ�����Ƚ��.Text = Get�ִ�����Ƚ��
   Txt�������ʽ�.Text = Get�������ʽ�����ġ
   Txt2�����ʽ�.Text = Get2�����ʽ�����ġ
   Txt��������.Text = Get�������Ⱑ��ġ
   Txt�����̷��.Text = Get�����̷�Ⱑ��ġ
End Sub

Sub ��������(Text As String, ��ȣ As Integer)
   If Text <> "0" And Text <> "" Then
      Text = Val(Text)
   End If
   If ��ȣ = -1 Then
      If Val(Text) > 0 Then
         Text = RTrim$(Str(Val(Text) * -1))
      End If
   End If
End Sub

Private Sub Txt2�����ʽ�_Change()
   �������� Txt2�����ʽ�, -1
End Sub

Private Sub Txt��������_Change()
   �������� Txt��������, 1
End Sub

Private Sub Txt�����̷��_Change()
   �������� Txt�����̷��, 1
End Sub

Private Sub Txt�������ʽ�_Change()
   �������� Txt�������ʽ�, -1
End Sub

Private Sub Txt���ѽð�_Change()
   �������� Txt���ѽð�, 1
End Sub

Private Sub Txt�ִ�����Ƚ��_Change()
   �������� Txt�ִ�����Ƚ��, 1
End Sub
