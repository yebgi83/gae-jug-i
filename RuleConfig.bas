Attribute VB_Name = "Module4"
Dim ���ѽð� As Long
Dim �ִ�����Ƚ�� As Integer
Dim �������ʽ�����ġ As Integer
Dim �������ʽ�����ġ As Integer
Dim �����̷�Ⱑ��ġ As Integer
Dim �������Ⱑ��ġ As Integer

Sub Set���ѽð�(�ð� As Long)
   ���ѽð� = �ð�
End Sub

Sub Set�ִ�����Ƚ��(����Ƚ�� As Integer)
   �ִ�����Ƚ�� = ����Ƚ��
End Sub

Sub Set�������ʽ�����ġ(����ġ As Integer)
   �������ʽ�����ġ = ����ġ
End Sub

Sub Set2�����ʽ�����ġ(����ġ As Integer)
   �������ʽ�����ġ = ����ġ
End Sub

Sub Set�����̷�Ⱑ��ġ(����ġ As Integer)
   �����̷�Ⱑ��ġ = ����ġ
End Sub

Sub Set�������Ⱑ��ġ(����ġ As Integer)
   �������Ⱑ��ġ = ����ġ
End Sub

Function Get���ѽð�() As Long
   Get���ѽð� = ���ѽð�
End Function

Function Get�ִ�����Ƚ��() As Integer
   Get�ִ�����Ƚ�� = �ִ�����Ƚ��
End Function

Function Get�������ʽ�����ġ() As Integer
   Get�������ʽ�����ġ = �������ʽ�����ġ
End Function

Function Get2�����ʽ�����ġ() As Integer
   Get2�����ʽ�����ġ = �������ʽ�����ġ
End Function

Function Get�����̷�Ⱑ��ġ() As Integer
   Get�����̷�Ⱑ��ġ = �����̷�Ⱑ��ġ
End Function

Function Get�������Ⱑ��ġ() As Integer
   Get�������Ⱑ��ġ = �������Ⱑ��ġ
End Function

Sub �⺻��Ģ����()
   ���ѽð� = 60 * 20
   �ִ�����Ƚ�� = 10
   �������ʽ�����ġ = -10
   �������ʽ�����ġ = -10
   �����̷�Ⱑ��ġ = 30
   �������Ⱑ��ġ = 30
End Sub

Sub ��Ģ��������(���ϸ� As String)
   Open ���ϸ� For Output As #1
      Print #1, ���ѽð�
      Print #1, �ִ�����Ƚ��
      Print #1, �������ʽ�����ġ
      Print #1, �������ʽ�����ġ
      Print #1, �����̷�Ⱑ��ġ
      Print #1, �������Ⱑ��ġ
   Close #1
End Sub

Sub ��Ģ���Ϻθ���(���ϸ� As String)
   Open ���ϸ� For Input As #1
      Input #1, ���ѽð�
      Input #1, �ִ�����Ƚ��
      Input #1, �������ʽ�����ġ
      Input #1, �������ʽ�����ġ
      Input #1, �����̷�Ⱑ��ġ
      Input #1, �������Ⱑ��ġ
   Close #1
End Sub

Function �÷��׿���������ġ(�÷��� As Integer)
   Dim ���ʽ� As Integer
   If (�÷��� And �������ʽ�) <> 0 Then ���ʽ� = ���ʽ� + �������ʽ�����ġ
   If (�÷��� And �������ʽ�) <> 0 Then ���ʽ� = ���ʽ� + �������ʽ�����ġ
   �÷��׿���������ġ = ���ʽ�
End Function


