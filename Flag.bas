Attribute VB_Name = "Module6"
Public Const �������ʽ� = 1
Public Const �������ʽ� = 2
Public Const ���� = 4
Public Const �����˹�Ģ = 8

Dim �÷��� As Integer
Sub Set�÷���(Flag As Integer)
   �÷��� = Flag
End Sub

Sub Set�������ʽ��÷���(Bit As Boolean)
   If (Bit = True) Then
      �÷��� = �÷��� Or �������ʽ�
   Else
      �÷��� = �÷��� And (255 Xor �������ʽ�)
   End If
End Sub

Sub Set2�����ʽ��÷���(Bit As Boolean)
   If (Bit = True) Then
      �÷��� = �÷��� Or �������ʽ�
   Else
      �÷��� = �÷��� And (255 Xor �������ʽ�)
   End If
End Sub

Sub Set�����÷���(Bit As Boolean)
   If (Bit = True) Then
      �÷��� = �÷��� Or ����
   Else
      �÷��� = �÷��� And (255 Xor ����)
   End If
End Sub

Sub Set�����˹�Ģ�÷���(Bit As Boolean)
   If (Bit = True) Then
      �÷��� = �÷��� Or �����˹�Ģ
   Else
      �÷��� = �÷��� And (255 Xor �����˹�Ģ)
   End If
End Sub

Function Get�÷���() As Integer
   Get�÷��� = �÷���
End Function

Function Get�������ʽ��÷���() As Boolean
   If (�÷��� And �������ʽ�) <> 0 Then
      Get�������ʽ��÷��� = True
   Else
      Get�������ʽ��÷��� = False
   End If
End Function

Function Get2�����ʽ��÷���() As Boolean
   If (�÷��� And �������ʽ�) <> 0 Then
      Get�������ʽ��÷��� = True
   Else
      Get�������ʽ��÷��� = False
   End If
End Function
