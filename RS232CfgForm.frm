VERSION 5.00
Begin VB.Form RS232CfgForm 
   BorderStyle     =   1  '���� ����
   Caption         =   "��� ȯ�� ����"
   ClientHeight    =   3195
   ClientLeft      =   7830
   ClientTop       =   7050
   ClientWidth     =   3450
   Icon            =   "RS232CfgForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3450
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton Cmd��� 
      Caption         =   "���"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox Combo�帧���� 
      Height          =   300
      Left            =   1560
      TabIndex        =   11
      Top             =   1980
      Width           =   1575
   End
   Begin VB.ComboBox Combo�и�Ƽ��Ʈ 
      Height          =   300
      Left            =   1560
      TabIndex        =   10
      Top             =   1620
      Width           =   1575
   End
   Begin VB.ComboBox Combo������Ʈ 
      Height          =   300
      Left            =   1560
      TabIndex        =   9
      Top             =   1260
      Width           =   1575
   End
   Begin VB.ComboBox Combo�����ͺ�Ʈ 
      Height          =   300
      Left            =   1560
      TabIndex        =   8
      Top             =   900
      Width           =   1575
   End
   Begin VB.ComboBox Combo��żӵ� 
      Height          =   300
      Left            =   1560
      TabIndex        =   7
      Top             =   540
      Width           =   1575
   End
   Begin VB.ComboBox Combo��Ʈ��ȣ 
      Height          =   300
      Left            =   1560
      TabIndex        =   6
      Top             =   200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "�帧 ����"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�и�Ƽ ��Ʈ"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "���� ��Ʈ"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "������ ��Ʈ"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��� �ӵ�"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��Ʈ ��ȣ"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "RS232CfgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd����_Click()
   Set��Ʈ Combo��Ʈ��ȣ.ListIndex
   Set�ӵ� Combo��żӵ�.ListIndex
   Set�����ͺ�Ʈ Combo�����ͺ�Ʈ.ListIndex
   Set������Ʈ Combo������Ʈ.ListIndex
   Set�и�Ƽ Combo�и�Ƽ��Ʈ.ListIndex
   Set�帧���� Combo�帧����.ListIndex
   Unload Me
End Sub

Private Sub Cmd���_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Combo��Ʈ��ȣ.AddItem "COM1"
   Combo��Ʈ��ȣ.AddItem "COM2"
   Combo��Ʈ��ȣ.AddItem "COM3"
   Combo��Ʈ��ȣ.AddItem "COM4"
   Combo��Ʈ��ȣ.AddItem "COM5"
   Combo��Ʈ��ȣ.AddItem "COM6"
   Combo��Ʈ��ȣ.AddItem "COM7"
   Combo��Ʈ��ȣ.AddItem "COM8"
   Combo��Ʈ��ȣ.AddItem "COM9"
   Combo��Ʈ��ȣ.AddItem "COM10"
   
   Combo��żӵ�.AddItem "4800"
   Combo��żӵ�.AddItem "7200"
   Combo��żӵ�.AddItem "9600"
   Combo��żӵ�.AddItem "14400"
   Combo��żӵ�.AddItem "19200"
   Combo��żӵ�.AddItem "38400"
   Combo��żӵ�.AddItem "57600"
   Combo��żӵ�.AddItem "115200"
   Combo��żӵ�.AddItem "128000"
   
   Combo�����ͺ�Ʈ.AddItem "7"
   Combo�����ͺ�Ʈ.AddItem "8"
   
   Combo������Ʈ.AddItem "1"
   Combo������Ʈ.AddItem "1.5"
   Combo������Ʈ.AddItem "2"
   
   Combo�и�Ƽ��Ʈ.AddItem "����"
   Combo�и�Ƽ��Ʈ.AddItem "Ȧ��"
   Combo�и�Ƽ��Ʈ.AddItem "¦��"
   
   Combo�帧����.AddItem "Xon / Xoff"
   Combo�帧����.AddItem "�ϵ����"
   Combo�帧����.AddItem "����"
   
   Combo��Ʈ��ȣ.ListIndex = Get��Ʈ
   Combo��żӵ�.ListIndex = Get�ӵ�
   Combo�����ͺ�Ʈ.ListIndex = Get�����ͺ�Ʈ
   Combo������Ʈ.ListIndex = Get������Ʈ
   Combo�и�Ƽ��Ʈ.ListIndex = Get�и�Ƽ
   Combo�帧����.ListIndex = Get�帧����
End Sub
