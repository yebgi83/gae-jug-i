Attribute VB_Name = "Module2"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const ����� = 1
Public Const ������ = 2
Public Const �������� = 250000000#

Dim ó�����ð� As Long
Dim ó�����ֽð� As Long

Dim ���ð� As Long
Dim ���ֽð� As Long

Dim �ð����� As Integer

Sub ���ð�����()
   �ð����� = �ð����� Or �����
   ó�����ð� = GetTickCount - ���ð�
End Sub

Sub ���ֽð�����()
   If ((�ð����� And �����) = 0) Then Exit Sub
   �ð����� = �ð����� Or ������
   ó�����ֽð� = GetTickCount - ���ֽð�
End Sub

Sub ���ֽð��ʱ�ȭ()
   ���ֽð� = 0
End Sub

Sub �ð������ʱ�ȭ()
   ���ð� = 0
   ���ֽð� = 0
End Sub

Sub ���ð��߰�(�߰��ð� As Long)
   ó�����ð� = GetTickCount - (���ð� + �߰��ð�)
   ���ð� = GetTickCount - ó�����ð�
   If (Get���ѽð� * 1000# - ���ð� = 0) Then
      ���ð� = Get���ѽð� * 1000#
   End If
End Sub

Sub ���ð�����(�ð� As Long)
   If ((�ð����� And �����) = 0) Then Exit Sub
   ó�����ð� = GetTickCount - �ð�
   ���ð� = �ð�
End Sub

Sub ���ð�����()
   �ð����� = 0
End Sub

Sub ���ֽð�����()
   �ð����� = �ð����� And �����
End Sub

Sub ������������()
   ���ֽð� = ��������
End Sub

Sub Set���ð�(�ð� As Long)
   ���ð� = �ð�
End Sub

Sub �ð��ݹ��Լ�()
   If (�ð����� And �����) <> 0 Then ���ð� = GetTickCount - ó�����ð�
   If (�ð����� And ������) <> 0 Then ���ֽð� = GetTickCount - ó�����ֽð�
End Sub

Function Get��������() As Boolean
   If ���ð� = �������� Then
      Get�������� = True
   Else
      Get�������� = False
   End If
End Function

Function Get���ð�() As Long
   Get���ð� = ���ð�
End Function

Function Get���ֽð�() As Long
   Get���ֽð� = ���ֽð�
End Function

Function Get�����() As Boolean
   If (�ð����� And �����) > 0 Then
      Get����� = True
   Else
      Get����� = False
   End If
End Function

Function Get������() As Boolean
   If (�ð����� And ������) > 0 Then
      Get������ = True
   Else
      Get������ = False
   End If
End Function

