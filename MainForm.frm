VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  '���� ����
   Caption         =   "������"
   ClientHeight    =   10890
   ClientLeft      =   2715
   ClientTop       =   3855
   ClientWidth     =   14730
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   14730
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame3 
      Caption         =   "���ʽ�"
      Height          =   1455
      Left            =   6360
      TabIndex        =   32
      Top             =   2520
      Width           =   2895
      Begin VB.CommandButton Cmd�������ʽ� 
         Caption         =   "�������ʽ�"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Cmd2�����ʽ� 
         Caption         =   "2�����ʽ�"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2�����ʽ� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   300
         Left            =   1920
         TabIndex        =   36
         Top             =   1000
         Width           =   360
      End
      Begin VB.Label Label�������ʽ� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1920
         TabIndex        =   35
         Top             =   390
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   10515
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6112
            Text            =   "��� ȯ��"
            TextSave        =   "��� ȯ��"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12700
            MinWidth        =   12700
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "��� ������ üũ"
            TextSave        =   "��� ������ üũ"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   1455
      Left            =   3240
      TabIndex        =   20
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton Cmd��ϻ��� 
         Caption         =   "��ϻ���"
         Height          =   495
         Left            =   1560
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Cmd�������� 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Cmd������� 
         Caption         =   "�������"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Cmd�������� 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������ ����"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton Cmd�����̷�� 
         Caption         =   "�����̷��"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1560
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Cmd�������� 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Cmd���ʽ��� 
         Caption         =   "���ʽ���"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Cmd�������� 
         Caption         =   "��������"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "���� ����"
      Height          =   7935
      Left            =   9360
      TabIndex        =   6
      Top             =   2520
      Width           =   5295
      Begin VB.CommandButton Cmd���ñ�ϻ��� 
         Caption         =   "������ ��� ��ȿ"
         Height          =   495
         Left            =   2520
         TabIndex        =   31
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TxtRank 
         Height          =   270
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   195
         Width           =   735
      End
      Begin VB.TextBox TxtNum 
         Height          =   270
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox TxtNumOfRace 
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1000
         Width           =   855
      End
      Begin VB.TextBox TxtRobotName 
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   200
         Width           =   1815
      End
      Begin VB.TextBox TxtName 
         Height          =   270
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   620
         Width           =   1935
      End
      Begin VB.TextBox TxtUni 
         Height          =   270
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   11033
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����Ƚ��"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "����ð�"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "���"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "���ʽ�"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "���"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Lab���� 
         Caption         =   "����:"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   25
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Lab������Ƚ�� 
         Caption         =   "������Ƚ��:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Lab�̸� 
         Caption         =   "�κ���:"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Lab�̸� 
         Caption         =   "�б�:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Lab�̸� 
         Caption         =   "�̸�:"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   8
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Lab�̸� 
         Caption         =   "����:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   720
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11245
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�б�"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�κ���"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ְ���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "����Ƚ��"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "���ð�"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4680
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label LabMsg 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '���� ����
      Caption         =   "�� ���α׷��� ����Ʈ���̼� ��� ����� ""������"" �Դϴ�."
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   21.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   14535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "�����ð�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Lab���ð� 
      BackColor       =   &H00000000&
      Caption         =   "00:00.000"
      BeginProperty Font 
         Name            =   "HY����L"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   960
      Left            =   5880
      TabIndex        =   3
      Top             =   1395
      Width           =   4050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "����ð�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Lab���ֽð� 
      BackColor       =   &H00000000&
      Caption         =   "00:00.000"
      BeginProperty Font 
         Name            =   "HY����L"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   960
      Left            =   600
      TabIndex        =   1
      Top             =   1395
      Width           =   4050
   End
   Begin VB.Menu Menu��� 
      Caption         =   "���"
      Begin VB.Menu Menu�����ڵ�� 
         Caption         =   "������ ���"
      End
      Begin VB.Menu Menu�ʱ�ȭ 
         Caption         =   "������ �ʱ�ȭ"
      End
      Begin VB.Menu Cmd����������� 
         Caption         =   "���� ��� ����"
      End
      Begin VB.Menu Menu���� 
         Caption         =   "����"
      End
   End
   Begin VB.Menu Menu���� 
      Caption         =   "����"
      Begin VB.Menu MenuRS232 
         Caption         =   "RS232 ����"
      End
      Begin VB.Menu Menu��Ģ���� 
         Caption         =   "��Ģ ����"
      End
   End
   Begin VB.Menu Menu������ 
      Caption         =   "������"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ���� As Integer
Dim �������� As Integer

Function ��������(���� As String, ���� As String) As Boolean
   Dim ���� As Integer
   Dim ������ As Boolean
   Dim ����� As Boolean
   
   ������ = Get������
   ����� = Get�����
   If ������ = True Then ���ֽð�����
   If ����� = True Then ���ð�����
   
   ���� = MsgBox(����, vbApplicationModal Or vbYesNo, ����)
   If ���� = vbYes Then
      �������� = True
   Else
      �������� = False
      If ������ = True Then ���ֽð�����
      If ����� = True Then ���ð�����
   End If
End Function

Sub ����ǥ��(Msg As String)
   StatusBar1.Panels(2).Text = Msg
   StatusBar1.Refresh
End Sub

Sub ���3�ʴ��()
   On Error GoTo NotOpened
   '3�ʰ� ����Ѵ�.
   MSComm1.PortOpen = False
   Timer1.Enabled = False
   ���ð�����
   WaitForm.Show 1
   ���ð�����
   Timer1.Enabled = True
   MSComm1.PortOpen = True
   Exit Sub

NotOpened:
   Timer1.Enabled = False
   ���ð�����
   WaitForm.Show 1
   ���ð�����
   Timer1.Enabled = True
End Sub

Sub �������()
   If Get����� = False Then ���ð�����
   
   Cmd��������.Enabled = False
   Cmd�������.Enabled = False
   Cmd��ϻ���.Enabled = False
   Cmd��������.Enabled = True
   Cmd��������.Enabled = True
   Set�÷��� 0
   
   '������ ������������ �ڵ�����
   LabMsg.Caption = TxtNum.Text & "�� [" & TxtRobotName.Text & "] �� " + Str(Val(TxtNumOfRace.Text) + 1) + "��° ������ ���۵Ǿ����ϴ�."
      
   ���ֽð�����
   
   '����
   PlaySound ��߽�ȣ����, 0, 1
End Sub

Sub ��������()
   Cmd��������.Enabled = True
   Cmd�������.Enabled = True
   Cmd��ϻ���.Enabled = True
   Cmd��������.Enabled = False
   Cmd��������.Enabled = False
   
   If Get���ֽð� > 0 Then
      ���ֽð�����
      ����߰� ����, Get���ֽð�, Get�÷���
      �����ں��ⰻ��
            
      '��� ǥ��
      LabMsg.Caption = TxtNum.Text & "�� [" & TxtRobotName.Text & "] ���� " + Str(Val(TxtNumOfRace.Text)) + "��° ����� " + �ð����·κ�ȯ(Get���ֽð�) + " �Դϴ�."
      ���ֽð��ʱ�ȭ
            
      '����
      PlaySound �ڼ��Ҹ�����, 0, 1
      If Val(TxtNumOfRace.Text) = Get�ִ�����Ƚ�� Then
         MsgBox "�ִ�����Ƚ���� " & Str(Get�ִ�����Ƚ��) & "ȸ�� ä�����ϴ�.", vbApplicationModal And vbOKOnly
         Cmd��������_Click
      Else
         ���3�ʴ��
      End If
    End If
End Sub

Sub �����ں��ⰻ��()
   ����ǥ�� "������ ������ �������Դϴ�."
   ������DB���� ListView1
   
   If ���� = 0 And ListView1.ListItems.Count > 0 Then ���� = 1
   If ���� > ListView1.ListItems.Count Then ���� = ListView1.ListItems.Count
   If ���� > 0 Then
      ListView1_ItemClick ListView1.ListItems.Item(����)
   Else
      TxtNum.Text = ""
      TxtName.Text = ""
      TxtUni.Text = ""
      TxtRank.Text = ""
      TxtRobotName.Text = ""
      TxtNumOfRace.Text = ""
   End If
   
   �����Ϻ��� ����, ListView2
   ����ǥ�� ""
End Sub

Private Sub Cmd��ϻ���_Click()
   If ���� = 0 Then
      MsgBox "���õ� �����ڰ� �����ϴ�.", vbApplicationModal Or vbOKOnly
      Exit Sub
   Else
      If ��������("���õ� [" & TxtNum.Text & "�� " & TxtName.Text & "] �������� ����� ��� �����մϴ�. �ε����� ��Ȳ���� ����ϴ� ����Դϴ�. �׷��� �Ͻðڽ��ϱ�?", "��Ȯ��") = True Then
         ��ϻ��� ����
         
         ����ǥ�� "������ ������ �������Դϴ�."
         �������������� �����ں�������
         ����ǥ�� ""

         �����ں��ⰻ��
         
         '�޽��� ������� ������
         MsgBox "��ϻ����� �Ϸ�Ǿ����ϴ�.", vbApplicationModal Or vbOKOnly
      End If
   End If
End Sub

Private Sub Cmd��������_Click()
   If (���� < ListView1.ListItems.Count) Then
       ListView1_ItemClick ListView1.ListItems.Item(���� + 1)
   Else
       LabMsg.Caption = "��� ���ʰ� �������ϴ�. �����ϼ̽��ϴ�."
   End If
End Sub

Private Sub Cmd���ñ�ϻ���_Click()
   Dim ÷�� As Integer
   Dim ���� As Integer
            
   ÷�� = Val(TxtNum.Text)
   If ListView2.ListItems.Count = 0 Then
      MsgBox "���� ����� �����ϴ�", vbApplicationModal Or vbOKOnly
      Exit Sub
   End If
   If ListView2.SelectedItem <> 0 Then
      ���� = MsgBox("���õ�" & ListView2.SelectedItem & "�� �������� �����˴ϴ�. �׷��� �Ͻðڽ��ϱ�?", vbApplicationModal Or vbYesNo, "���")
      If ���� = vbYes Then
         ���ñ�ϻ��� ����, Val(ListView2.SelectedItem)
         �����ں��ⰻ��
      End If
   Else
      MsgBox "������ ����� �����ϴ�", vbSystemModal
   End If
End Sub

Private Sub Cmd�����̷��_Click()
   If ���� = 0 Then
      MsgBox "���õ� �����ڰ� �����ϴ�.", vbApplicationModal Or vbOKOnly
      Exit Sub
   Else
      If ��������("���õ� [" & TxtNum.Text & "�� " & TxtName.Text & "] �������� ������ ���� �ڷ� �̷����ϴ�. �׸��� ���ð��� �����Ǵ� �������� �ۿ��մϴ�. �׷��� �Ͻðڽ��ϱ�?", "��Ȯ��") = True Then
         '����ġ �ο��Ѵ��� ���ʸ� �����ϸ�, ����� �ȴ�.
         ���ֽð��ʱ�ȭ
         ���ð��߰� Get�����̷�Ⱑ��ġ
         Cmd��������_Click
         
         '����� ������ ��ġ�� �ڷ� �̷��, �� ���¸� �����ش�.
         �����̷�� ����
         �����ں��ⰻ��
         
         '�޽��� ������� ������
         MsgBox "�����̷�� �۾��� �Ϸ�Ǿ����ϴ�.", vbApplicationModal Or vbOKOnly
      End If
   End If
End Sub

Private Sub Cmd�����������_Click()
   ������ϸ���� �����������
   Call Shell("C:\Windows\NOTEPAD.EXE " & �����������, vbMaximizedFocus)
End Sub

Private Sub Cmd�������ʽ�_Click()
   Dim �÷��� As Integer
   �÷��� = ��Ͽ����÷��׾��(����, ��������)
   
   If ((�÷��� And �������ʽ�) <> 0) Then
      �÷��� = �÷��� And (255 Xor �������ʽ�)
   Else
      �÷��� = �÷��� Or �������ʽ�
   End If
   ��Ͽ��÷��׵����� ����, ��������, �÷���
   
   'ȭ�鿡 ���̴� ������ ����
   ������DB���� ListView1
   �����Ϻ��� ����, ListView2
   
   'ListView1, �� ������ ���� ����� �����ϸ� �ڵ����� ���� ���õ� ������ ������ ��Ȳ(�����ʺκ�)�� ����ǰ� �س���.
   ListView1_ItemClick ListView1.ListItems.Item(����)
End Sub

Private Sub Cmd2�����ʽ�_Click()
   Dim �÷��� As Integer
   �÷��� = ��Ͽ����÷��׾��(����, ��������)
   
   If ((�÷��� And �������ʽ�) <> 0) Then
      �÷��� = �÷��� And (255 Xor �������ʽ�)
   Else
      �÷��� = �÷��� Or �������ʽ�
   End If
   ��Ͽ��÷��׵����� ����, ��������, �÷���
   
   'ȭ�鿡 ���̴� ������ ����
   ������DB���� ListView1
   �����Ϻ��� ����, ListView2
   
   'ListView1, �� ������ ���� ����� �����ϸ� �ڵ����� ���� ���õ� ������ ������ ��Ȳ(�����ʺκ�)�� ����ǰ� �س���.
   ListView1_ItemClick ListView1.ListItems.Item(����)
End Sub

Private Sub Cmd�������_Click()
   �������
End Sub

Private Sub Cmd��������_Click()
   ��������
End Sub

Private Sub Cmd��������_Click()
   ���ֽð�����
   
   Cmd��������.Enabled = True
   Cmd�������.Enabled = True
   Cmd��ϻ���.Enabled = True
   Cmd��������.Enabled = False
   Cmd��������.Enabled = False
      
   ������������
      
   ���ð��߰� Get�������Ⱑ��ġ
   
   ����߰� ����, Get���ֽð�, 0
   �����ں��ⰻ��
   
   LabMsg.Caption = TxtNum.Text & "�� [" & TxtRobotName.Text & "] ���� " + Str(Val(TxtNumOfRace.Text)) + "��° ������ �����ϼ̽��ϴ�."
   
   ���ֽð��ʱ�ȭ
   Timer1_Timer
   
   If Val(TxtNumOfRace.Text) = Get�ִ�����Ƚ�� Then
      MsgBox "�ִ�����Ƚ���� " & Str(Get�ִ�����Ƚ��) & "ȸ�� ä�����ϴ�.", vbApplicationModal Or vbOKOnly
      Cmd��������_Click
   Else
      ���3�ʴ��
   End If
End Sub

Private Sub Cmd���ʽ���_Click()
   If ���� = 0 Then
      MsgBox "�غ��� �����ڸ� �����ϼ���"
   Else
      ����Ƚ�� = Val(TxtNumOfRace.Text)
      If ����Ƚ�� >= Get�ִ�����Ƚ�� Then
         MsgBox "�̹� �ִ�����Ƚ���� " & Str(Get�ִ�����Ƚ��) & "ȸ�� ä�����ϴ�.", vbApplicationModal Or vbOKOnly
      Else
         ListView1.Enabled = False
         Cmd���ʽ���.Enabled = False
         Cmd��������.Enabled = True
         Cmd��������.Enabled = False
         Cmd�������.Enabled = True
         Cmd�����̷��.Enabled = True
         
         LabMsg.Caption = TxtNum.Text & "�� [" & TxtRobotName.Text & "] ����԰� ���ÿ� ���ʰ� ���۵˴ϴ�."
         ���ð�����
         ���ð�����
      End If
   End If
End Sub

Private Sub Cmd��������_Click()
   LabMsg.Caption = TxtNum.Text & "�� [" & TxtRobotName.Text & "] " & TxtRank.Text & "���� ���ʰ� �������ϴ�."
   Set�����ڻ��ð� ����, Get���ð�
   
   ListView1.Enabled = True
   
   Cmd���ʽ���.Enabled = True
   Cmd��������.Enabled = False
   Cmd��������.Enabled = True
   Cmd��ϻ���.Enabled = True
   Cmd�������.Enabled = False
   Cmd��������.Enabled = False
   Cmd��������.Enabled = False
   Cmd�����̷��.Enabled = False
   
   ���ð�����
   
   ����ǥ�� "������ ������ �������Դϴ�."
   �������������� �����ں�������
   ����ǥ�� ""
   
   �����ں��ⰻ��
End Sub

Private Sub Form_Load()
   MainForm.Caption = MainForm.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
   
   ��Ģ���Ϻθ��� ��Ģ��������
   ���ȯ��θ��� ���ȯ�溸������
   ��ż������� MSComm1
   
   If MSComm1.PortOpen = True Then
      StatusBar1.Panels(3).Text = "��ſ��� ����"
   Else
      StatusBar1.Panels(3).Text = "��ſ��� �ȵ�"
   End If
   
   On Error GoTo LabelFileNotFound
   ���������Ϻθ��� �����ں�������
   �����ں��ⰻ��
   
   Exit Sub
LabelFileNotFound:
End Sub

Private Sub Form_Paint()
   StatusBar1.Panels(1).Text = ��Ż���
   �����ں��ⰻ��
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ����ǥ�� "�������� ������ ������ �������Դϴ�. (1/3)"
   �������������� �����ں�������
   ����ǥ�� "�������� ���ȯ�漳���� �������Դϴ�. (2/3)"
   ���ȯ������ ���ȯ�溸������
   ����ǥ�� "�������� ��Ģȯ�漳���� �������Դϴ�. (3/3)"
   ��Ģ�������� ��Ģ��������
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
   TxtNum.Text = Val(Item.Text)
   TxtName.Text = Item.SubItems(1)
   TxtUni.Text = Item.SubItems(2)
   TxtRobotName.Text = Item.SubItems(3)
   TxtNumOfRace.Text = Item.SubItems(6)
      
   If (Val(Item.SubItems(4)) = 0) Then '������ ���ٸ�
      TxtRank.Text = "---"
   Else
      TxtRank.Text = Item.SubItems(4)
   End If
  
   '���� ���� ǥ�� ����
   ListView1.ListItems.Item(����).ForeColor = &H0
   ListView1.ListItems.Item(����).Bold = False
   
   ���� = Item.Index
   ��ũ����ġ���� ListView1.hwnd, ����

   'ǥ��
   ListView1.ListItems.Item(����).ForeColor = &HFF
   ListView1.ListItems.Item(����).Bold = True
   
   Set���ð� Get�����ڻ��ð�(����)
   
   �����Ϻ��� ����, ListView2
   
   If (ListView2.ListItems.Count > 0) Then
      �������� = ListView2.ListItems.Count
   Else
      �������� = 0
      Frame3.Caption = "���ʽ�"
      Label�������ʽ�.ForeColor = 0
      Label�������ʽ�.Caption = "����"
      Label2�����ʽ�.ForeColor = 0
      Label2�����ʽ�.Caption = "����"
   End If
      
   If �������� > 0 Then ListView2_ItemClick ListView2.ListItems.Item(��������)
      
   If (Val(TxtNumOfRace.Text) > 0) Then
      Cmd�������ʽ�.Enabled = True
      Cmd2�����ʽ�.Enabled = True
   Else
      Cmd�������ʽ�.Enabled = False
      Cmd2�����ʽ�.Enabled = False
   End If
   
   '�޽��� �˸�
   LabMsg.Caption = TxtNum.Text & "�� [" & TxtRobotName.Text & "] �����Դϴ�. �غ����ֽʽÿ�."
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim �÷��� As Integer
   
   If Item.SubItems(2) <> "��������" Then
      Cmd2�����ʽ�.Enabled = True
      Cmd�������ʽ�.Enabled = True
      Frame3.Caption = �������� & "��° ���� ���ʽ� "
   Else
      Cmd2�����ʽ�.Enabled = False
      Cmd�������ʽ�.Enabled = False
      Label�������ʽ�.ForeColor = 0
      Label�������ʽ�.Caption = "����"
      Label2�����ʽ�.ForeColor = 0
      Label2�����ʽ�.Caption = "����"
      Frame3.Caption = �������� & "��° ���ʽ� ����"
   End If
   
   'ǥ��
   ListView2.ListItems.Item(��������).ForeColor = &H0
   ListView2.ListItems.Item(��������).Bold = False
   
   �������� = Val(Item.Index)
   
   'ǥ��
   ListView2.ListItems.Item(��������).ForeColor = &HFF0000
   ListView2.ListItems.Item(��������).Bold = True
   If Item.SubItems(2) = "��������" Then Exit Sub
         
   �÷��� = ��Ͽ����÷��׾��(����, Item.Index)
   
   If (�÷��� And �������ʽ�) <> 0 Then
      Label�������ʽ�.ForeColor = &HFF0000
      Label�������ʽ�.Caption = "ȹ��"
   Else
      Label�������ʽ�.ForeColor = 0
      Label�������ʽ�.Caption = "����"
   End If
   
   If (�÷��� And �������ʽ�) <> 0 Then
      Label2�����ʽ�.ForeColor = &HFF0000
      Label2�����ʽ�.Caption = "ȹ��"
   Else
      Label2�����ʽ�.ForeColor = 0
      Label2�����ʽ�.Caption = "����"
   End If
End Sub

Private Sub Menu������_Click()
   frmAbout.Show 1
End Sub

Private Sub Menu��Ģ����_Click()
   RuleCfgForm.Show 1
   ��Ģ�������� ��Ģ��������
End Sub

Private Sub Menu����_Click()
   Unload Me
End Sub

Private Sub Menu�����ڵ��_Click()
   ���ð�����
   RegDialog.Show 1
   MainForm.Refresh
   �����ں��ⰻ��
End Sub

Private Sub Menu�ʱ�ȭ_Click()
   Dim ���� As Integer
   ���� = MsgBox("�����ڿ� ���� ��� �����Ͱ� �����˴ϴ�. �׷��� �Ͻðڽ��ϱ�?", vbYesNo, "���")
      
   If ���� = vbYes Then
      '���ϳ��� �ʱ�ȭ
      Open �����ں������� For Output As #1: Close #1
      ���� = 0
      ������DB�ʱ�ȭ
      ���ð�����
      MsgBox "��� �����Ͱ� �����Ǿ����ϴ�.", vbApplicationModal And vbOKOnly, "�޼���"
      ListView1.ListItems.Clear
      ListView2.ListItems.Clear
   End If
End Sub

Private Sub MenuRS232_Click()
   RS232CfgForm.Show 1
   
   ��ż������� MSComm1
   If MSComm1.PortOpen = True Then
      StatusBar1.Panels(3).Text = "��ſ��� ����"
   Else
      StatusBar1.Panels(3).Text = "��ſ��� �ȵ�"
   End If
   
   ���ȯ������ ���ȯ�溸������
   MainForm.Refresh
End Sub

Private Sub MSComm1_OnComm()
   Dim RvChar As String * 1
   Select Case MSComm1.CommEvent
      '�޾��� ���
      Case comEvReceive
         RvChar = MSComm1.Input
         
         '������ Status Bar�� ǥ�� (�����)
         StatusBar1.Panels(3).Text = "���� ������ : " & RvChar & " (" & Hex$(Asc(RvChar)) & ")"
         
         If Asc(RvChar) = ��߽�ȣ Then
            StatusBar1.Panels(3).Text = StatusBar1.Panels(3).Text & " �� ��߽�ȣȮ��"
            If Cmd���ʽ���.Enabled = False And Get������ = False Then
               �������
            End If
         End If
         If Asc(RvChar) = ������ȣ Then
            StatusBar1.Panels(3).Text = StatusBar1.Panels(3).Text & " �� ������ȣȮ��"
            If Get������ = True Then
               ��������
            End If
         End If
   End Select
End Sub

Private Sub Timer1_Timer()
   �ð��ݹ��Լ�
   'StatusBar1.Panels(2) = ���� & "   " & Get���ð�
   
   StatusBar1.Panels(2) = DateTime.Now
   If (Get���ѽð� * 1000# - Get���ð� > 0) Then
      'Get���ѽð��� �⺻������ �ʴ���, Get���ð��� 1MilliSec�� �⺻������ ���ѽð��� 100�� ����
      Lab���ð�.Caption = �ð����·κ�ȯ(Get���ѽð� * 1000# - Get���ð�)
      Lab���ֽð�.Caption = �ð����·κ�ȯ(Get���ֽð�)
   Else
      Lab���ð�.Caption = �ð����·κ�ȯ(0)
      Lab���ֽð�.Caption = �ð����·κ�ȯ(0)
      If (Get����� = True) Then
         If (Get������ = True) Then
            ���ֽð�����
            ���ֽð��ʱ�ȭ
         End If
         
         ���ð����� Get���ѽð� * 1000#
         Cmd��������_Click
         LabMsg.Caption = TxtNum.Text & "�� [" & TxtRobotName.Text & "]�� ���ð��� �� �Ǿ����ϴ�."
      End If
   End If
End Sub

