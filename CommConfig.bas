Attribute VB_Name = "Module3"
Type ���ȯ�汸��ü
   ��Ʈ As Integer
   �ӵ� As Long
   �����ͺ�Ʈ As Integer
   ������Ʈ As Integer
   �и�Ƽ As Integer
   �帧���� As Integer
   ȯ�漳�����ڿ� As String
End Type

Dim ���ȯ�� As ���ȯ�汸��ü
Function Get��Ʈ() As Integer
   Get��Ʈ = ���ȯ��.��Ʈ - 1
End Function
Function Get�ӵ�() As Long
   Select Case ���ȯ��.�ӵ�
      Case 4800
        Get�ӵ� = 0
      Case 7200
        Get�ӵ� = 1
      Case 9600
        Get�ӵ� = 2
      Case 14400
        Get�ӵ� = 3
      Case 19200
        Get�ӵ� = 4
      Case 38400
        Get�ӵ� = 5
      Case 57600
        Get�ӵ� = 6
      Case 115200
        Get�ӵ� = 7
      Case 128000
        Get�ӵ� = 8
   End Select
End Function
Function Get������Ʈ() As Integer
   Select Case ���ȯ��.������Ʈ
      Case 1
         Get������Ʈ = 0
      Case 1.5
         Get������Ʈ = 1
      Case 2
         Get������Ʈ = 2
   End Select
End Function
Function Get�и�Ƽ() As Integer
   Get�и�Ƽ = ���ȯ��.�и�Ƽ
End Function
Function Get�����ͺ�Ʈ() As Integer
   Select Case ���ȯ��.�����ͺ�Ʈ
      Case 7
        Get�����ͺ�Ʈ = 0
      Case 8
        Get�����ͺ�Ʈ = 1
   End Select
End Function
Function Get�帧����() As Integer
   Get�帧���� = ���ȯ��.�帧����
End Function
Sub Set��Ʈ(��Ʈ���� As Integer)
   ���ȯ��.��Ʈ = ��Ʈ���� + 1
End Sub
Sub Set�ӵ�(�ӵ� As Integer)
   Select Case �ӵ�
      Case 0
        ���ȯ��.�ӵ� = 4800
      Case 1
        ���ȯ��.�ӵ� = 7200
      Case 2
        ���ȯ��.�ӵ� = 9600
      Case 3
        ���ȯ��.�ӵ� = 14400
      Case 4
        ���ȯ��.�ӵ� = 19200
      Case 5
        ���ȯ��.�ӵ� = 38400
      Case 6
        ���ȯ��.�ӵ� = 57600
      Case 7
        ���ȯ��.�ӵ� = 115200
      Case 8
        ���ȯ��.�ӵ� = 128000
   End Select
End Sub
Sub Set������Ʈ(������Ʈ As Integer)
   Select Case ������Ʈ
      Case 1
        ���ȯ��.������Ʈ = 0
      Case 1.5
        ���ȯ��.������Ʈ = 1
      Case 2
        ���ȯ��.������Ʈ = 2
   End Select
End Sub
Sub Set�и�Ƽ(�и�Ƽ As Integer)
   ���ȯ��.�и�Ƽ = �и�Ƽ
End Sub
Sub Set�����ͺ�Ʈ(�����ͺ�Ʈ As Integer)
    Select Case �����ͺ�Ʈ
      Case 0
        ���ȯ��.�����ͺ�Ʈ = 7
      Case 1
        ���ȯ��.�����ͺ�Ʈ = 8
   End Select
End Sub
Sub Set�帧����(�帧���� As Integer)
   ���ȯ��.�帧���� = �帧����
End Sub
Sub ��ű⺻ȯ��(��Ű�ü As MSComm)
   ���ȯ��.��Ʈ = 3
   ���ȯ��.�ӵ� = 57600
   ���ȯ��.�����ͺ�Ʈ = 8
   ���ȯ��.������Ʈ = 1
   ���ȯ��.�и�Ƽ = 0
   ���ȯ��.ȯ�漳�����ڿ� = "57600,n,8"
   ���ȯ��.�帧���� = 0
   ��ż������� ��Ű�ü
End Sub
Sub ��ż�������(��Ű�ü As MSComm)
   If (��Ű�ü.PortOpen = True) Then
      ��Ű�ü.PortOpen = False
      For c = 1 To 1000: Next c
   End If
   
   ��Ű�ü.CommPort = ���ȯ��.��Ʈ
   ��Ű�ü.Handshaking = ���ȯ��.�帧����
   ��Ű�ü.RThreshold = 1
   ��Ű�ü.RTSEnable = True
   
   ���ȯ��.ȯ�漳�����ڿ� = LTrim$(Str$(���ȯ��.�ӵ�)) & ","
   Select Case ���ȯ��.�и�Ƽ
      Case 0 'None
         ���ȯ��.ȯ�漳�����ڿ� = ���ȯ��.ȯ�漳�����ڿ� & "n"
      Case 1 'Odd
         ���ȯ��.ȯ�漳�����ڿ� = ���ȯ��.ȯ�漳�����ڿ� & "o"
      Case 2 'Even
         ���ȯ��.ȯ�漳�����ڿ� = ���ȯ��.ȯ�漳�����ڿ� & "e"
   End Select
   
   ���ȯ��.ȯ�漳�����ڿ� = ���ȯ��.ȯ�漳�����ڿ� & "," & LTrim$(Str$(���ȯ��.�����ͺ�Ʈ))
   ���ȯ��.ȯ�漳�����ڿ� = ���ȯ��.ȯ�漳�����ڿ� & "," & LTrim$(Str$(���ȯ��.������Ʈ))
      
   ��Ű�ü.Settings = ���ȯ��.ȯ�漳�����ڿ�
   ��Ű�ü.SThreshold = 1
   ��Ű�ü.InputLen = 1
   
   On Error GoTo ErrorOpenPort
   ��Ű�ü.PortOpen = True
ErrorOpenPort:
End Sub

Function ��Ż���() As String
   ��Ż��� = "COM" & LTrim$(���ȯ��.��Ʈ) & ":" & ���ȯ��.ȯ�漳�����ڿ�
End Function

Sub ���ȯ������(���ϸ� As String)
   Open ���ϸ� For Output As #1
      Print #1, ���ȯ��.��Ʈ
      Print #1, ���ȯ��.�ӵ�
      Print #1, ���ȯ��.�����ͺ�Ʈ
      Print #1, ���ȯ��.������Ʈ
      Print #1, ���ȯ��.�и�Ƽ
      Print #1, ���ȯ��.�帧����
   Close #1
End Sub
Sub ���ȯ��θ���(���ϸ� As String)
   Open ���ϸ� For Input As #1
      Input #1, ���ȯ��.��Ʈ
      Input #1, ���ȯ��.�ӵ�
      Input #1, ���ȯ��.�����ͺ�Ʈ
      Input #1, ���ȯ��.������Ʈ
      Input #1, ���ȯ��.�и�Ƽ
      Input #1, ���ȯ��.�帧����
   Close #1
End Sub

