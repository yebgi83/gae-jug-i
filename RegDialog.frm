VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RegDialog 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "������ ���"
   ClientHeight    =   6090
   ClientLeft      =   6120
   ClientTop       =   5550
   ClientWidth     =   8805
   Icon            =   "RegDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton EditBtn 
      Caption         =   "����"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton MoveFirstBtn 
      Caption         =   "�� ���� �̵�"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton DeleteBtn 
      Caption         =   "����"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtRobotName 
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   200
      Width           =   2295
   End
   Begin VB.TextBox TxtName 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   200
      Width           =   2295
   End
   Begin VB.TextBox TxtUni 
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   5415
   End
   Begin VB.CommandButton RegisterBtn 
      Caption         =   "���"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�̸�"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�б�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�κ���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ְ���"
         Object.Width           =   2170
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "����Ƚ��"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "���ð�"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "�κ���"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "�̸�"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "�б�"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   645
      Width           =   615
   End
End
Attribute VB_Name = "RegDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ÷�� As Integer
Dim ���� As Integer

Sub �����ں��ⰻ��()
    ������DB���� ListView1
    
    If ÷�� > 0 And ListView1.ListItems.Count > 0 Then
       ListView1_ItemClick ListView1.ListItems.Item(÷��)
       ��ũ����ġ���� ListView1.hwnd, ÷��
    Else
       TxtName.Text = ""
       TxtUni.Text = ""
       TxtRobotName.Text = ""
    End If
End Sub

Private Sub DeleteBtn_Click()
    Dim �̸� As String
    Dim �κ��� As String
    Dim ����� As String
    
    If (ListView1.ListItems.Count = 0) Then Exit Sub
    
    �̸� = TxtName.Text
    �κ��� = TxtRobotName.Text
    
    If ÷�� > 0 Then
       ����� = "���õ� [" & ���� & "�� " & �̸� & "] �����ڸ� �����Ͻðڽ��ϱ�?"
       If MsgBox(�����, vbOKCancel, "���") = vbOK Then
          �����ڻ��� ÷��
          MsgBox "�����Ͽ����ϴ�.", vbOKOnly, "�޼���"
          
          �������������� �����ں�������
          �����ں��ⰻ��
       End If
    End If
End Sub

Private Sub EditBtn_Click()
    If (ListView1.ListItems.Count = 0) Then Exit Sub
        
    If ÷�� > 0 Then
       �����ڹ迭���� ÷��, TxtName.Text, TxtRobotName.Text, TxtUni.Text
       
       MsgBox ���� & "�� �������� ������ �Ϸ��߽��ϴ�.", vbOKOnly, "�޼���"
       
       �������������� �����ں�������
       �����ں��ⰻ��
    End If
End Sub

Private Sub Form_Load()
    ���������Ϻθ��� �����ں�������
    ������DB���� ListView1
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TxtName.Text = Item.SubItems(1)
    TxtUni.Text = Item.SubItems(2)
    TxtRobotName.Text = Item.SubItems(3)
    ���� = Item.Text
    ÷�� = Item.Index
End Sub

Private Sub MoveFirstBtn_Click()
    Dim �̸� As String
    Dim �κ��� As String
    Dim ����� As String
    
    If (ListView1.ListItems.Count = 0) Then Exit Sub
    
    �̸� = TxtName.Text
    �κ��� = TxtRobotName.Text
    
    If ÷�� > 0 Then
       ����� = "���õ� [" & ���� & "�� " & �̸� & "] �����ڸ� �� ���� �ű�ڽ��ϱ�?"
       If MsgBox(�����, vbOKCancel, "���") = vbOK Then
          �������̵� ÷��
          MsgBox "�̵��� �Ϸ��߽��ϴ�.", vbOKOnly, "�޼���"
          
          �������������� �����ں�������
          �����ں��ⰻ��
       End If
    End If
End Sub

Private Sub RegisterBtn_Click()
    Dim �̸� As String
    Dim �κ��� As String
    Dim �б� As String
    
    �̸� = TxtName.Text
    �κ��� = TxtRobotName.Text
    �б� = TxtUni.Text
    �����ڵ�� �̸�, �κ���, �б�
    
    �������������� �����ں�������
    �����ں��ⰻ��
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       TxtRobotName.SetFocus
    End If
End Sub

Private Sub TxtRobotName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       TxtUni.SetFocus
    End If
End Sub

Private Sub TxtUni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       RegisterBtn_Click
       TxtName.SetFocus
    End If
End Sub
