Attribute VB_Name = "Module1"
Const �ִ������� = 1000
Const NOTTHING = 0
Const MINIMUM = 200000000#
Const ���ڿ����� = 64
Const �����ϱ��� = 1000
Const ��ϵ�����ũ�� = 5

Type �����ڼ���
    ����     As Integer
    �̸�     As String * ���ڿ�����
    �κ���   As String * ���ڿ�����
    �б�     As String * ���ڿ�����
    ����     As Integer
    ����Ƚ�� As Integer
    �ְ��� As Long
    ������ As String * �����ϱ���
    ���ð� As Long
End Type

Dim ������(�ִ�������) As �����ڼ���
Dim ��������           As Integer
Dim �����ο�           As Integer

Sub �����ڹ迭����(÷�� As Integer, �̸� As String, �κ��� As String, �б� As String)
    n = ÷��
    ������(n).�̸� = �̸�
    ������(n).�κ��� = �κ���
    ������(n).�б� = �б�
End Sub

'�����ڸ� Ư�� �迭 ÷�ڿ� �Է��Ѵ�
Sub �����ڹ迭�Է�(÷�� As Integer, ���� As Integer, �̸� As String, �κ��� As String, �б� As String, ���� As Integer, ����Ƚ�� As Integer, �ְ��� As Integer, ������ As String, ���ð� As Long)
    n = ÷��
    ������(n).���� = ����
    ������(n).�̸� = �̸�
    ������(n).�κ��� = �κ���
    ������(n).�б� = �б�
    ������(n).���� = ����
    ������(n).����Ƚ�� = ����Ƚ��
    ������(n).�ְ��� = �ְ���
    ������(n).������ = ������
    ������(n).���ð� = ���ð�
End Sub

'������DB�� �ʱ�ȭ�Ѵ�.
Sub ������DB�ʱ�ȭ()
    Dim i As Integer
    
    �����ο� = 0
    For i = 1 To �ִ�������
        �����ڹ迭�Է� i, i, "", "", "", 0, 0, 0, "", 0
    Next i
End Sub

'�����ڸ� ������DB�� �߰��Ѵ�.
Sub �����ڵ��(�̸� As String, �κ��� As String, �б� As String)
    �������� = �������� + 1
    �����ο� = �����ο� + 1 '�����ο� �Ѹ� �߰�
    
    '������ ���
    �����ڹ迭�Է� �����ο�, ��������, �̸�, �κ���, �б�, 0, 0, 0, "", 0
End Sub

Sub �����ڻ���(÷�� As Integer)
    Dim i As Integer
    
    Dim ����     As Integer
    Dim �̸�     As String * ���ڿ�����
    Dim �κ���   As String * ���ڿ�����
    Dim �б�     As String * ���ڿ�����
    Dim ����     As Integer
    Dim ����Ƚ�� As Integer
    Dim �ְ��� As Integer
    Dim ������ As String * �����ϱ���
    Dim ���ð� As Long
    
    For i = ÷�� To �����ο� - 1
        ���� = ������(i + 1).����
        �̸� = ������(i + 1).�̸�
        �κ��� = ������(i + 1).�κ���
        �б� = ������(i + 1).�б�
        ���� = ������(i + 1).����
        ����Ƚ�� = ������(i + 1).����Ƚ��
        �ְ��� = ������(i + 1).�ְ���
        ������ = ������(i + 1).������
        ���ð� = ������(i + 1).���ð�
        �����ڹ迭�Է� i, ����, �̸�, �κ���, �б�, ����, ����Ƚ��, �ְ���, ������, ���ð�
    Next i
    
    �����ڹ迭�Է� �����ο�, 0, "", "", "", 0, 0, 0, "", 0
    �����ο� = �����ο� - 1
    
    '���� ����
    For i = 1 To �����ο�
        ������� i
    Next i
End Sub

Sub ��������������(���ϸ� As String)
    '���� �ʱ�ȭ
    Open ���ϸ� For Output As #1: Close #1
    
    '���� ����
    Open ���ϸ� For Binary As #1
        Put #1, , ��������
        For i = 1 To �����ο�
            Put #1, , ������(i)
        Next i
    Close #1
End Sub

Sub ���������Ϻθ���(���ϸ� As String)
    On Error GoTo FileNotFound
    
    '������ ��� �ʱ�ȭ
    ������DB�ʱ�ȭ
    
    
    '���� �θ���.
    Open ���ϸ� For Binary As #1
    Get #1, , ��������
    Do
        If LOF(1) <= Loc(1) Then Exit Do
        �����ο� = �����ο� + 1
        Get #1, , ������(�����ο�)
    Loop
    Close #1
FileNotFound:
End Sub

Function �ð����·κ�ȯ(Count As Long) As String
    Dim Min As Integer
    Dim Sec As Integer
    Dim MSec As Integer
    Dim TempCount As Long
    
    'Call by referenceȿ���� ����
    TempCount = Count
    
    If TempCount = �������� Then
       �ð����·κ�ȯ = "---------"
       Exit Function
    End If
    
    '��.��.�и��������� �и�
    Min = Int(TempCount / 60000#)
    TempCount = TempCount Mod (60000#)
    
    Sec = Int(TempCount / 1000)
    TempCount = TempCount Mod 1000
    
    MSec = TempCount
    
    '���ڿ��� ��ȯ ��ü
    StrMin$ = LTrim$(Min)
    If Len(StrMin$) < 2 Then StrMin$ = "0" + StrMin$
    
    StrSec$ = LTrim$(Sec)
    If Len(StrSec$) < 2 Then StrSec$ = "0" + StrSec$
    
    StrMSec$ = LTrim$(MSec)
    If Len(StrMSec$) < 3 Then StrMSec$ = String$(3 - Len(StrMSec$), "0") + StrMSec$
    
    '������� ������
    �ð����·κ�ȯ = StrMin$ + ":" + StrSec$ + "." + StrMSec$
End Function

Sub ������DB����(lv As ListView)
    '����Ʈ�� ���� ����
    lv.ListItems.Clear
    '����Ʈ�信 ����
    For i = 1 To �����ο�
        Set ItmX = lv.ListItems.Add(, , ������(i).����)
        ItmX.SubItems(1) = RTrim$(������(i).�̸�)
        ItmX.SubItems(2) = RTrim$(������(i).�б�)
        ItmX.SubItems(3) = RTrim$(������(i).�κ���)
        
        If ������(i).���� = 0 Then
           ItmX.SubItems(4) = "---"
        Else
           ItmX.SubItems(4) = Str(������(i).����)
        End If
        
        If ������(i).�ְ��� = 0 Then
           ItmX.SubItems(5) = "--------"
        Else
           ItmX.SubItems(5) = �ð����·κ�ȯ(������(i).�ְ���)
        End If
        
        ItmX.SubItems(6) = Str(������(i).����Ƚ��)
        ItmX.SubItems(7) = �ð����·κ�ȯ(������(i).���ð�)
    Next i
    lv.Refresh
End Sub

Sub ���ڸ�4����Ʈ�κ���(���� As Long, Ch1 As Integer, Ch2 As Integer, Ch3 As Integer, Ch4 As Integer)
    
End Sub

Sub �����Ϻ���(÷�� As Integer, lv As ListView)
    '�������� ����Ʈ�信 ������
    Dim TempCount As Long
    Dim Bonus As Long
    Dim Record As String
    
    '����Ʈ�� ���� ����
    lv.ListItems.Clear
    
    '����Ʈ�信 ����
    Record = ������(÷��).������
    For i = 1 To ������(÷��).����Ƚ��
        HighOffset = ��ϵ�����ũ�� * (i - 1) + 1 ' ��������Ʈ
        MidOffset1 = ��ϵ�����ũ�� * (i - 1) + 2 ' �߾ӹ���Ʈ1
        MidOffset2 = ��ϵ�����ũ�� * (i - 1) + 3 ' �߾ӹ���Ʈ2
        LowOffset = ��ϵ�����ũ�� * (i - 1) + 4 ' ��������Ʈ
        FlagOffset = ��ϵ�����ũ�� * (i - 1) + 5     '�÷���
        
        H# = Asc(Mid$(Record, HighOffset, 1))
        M1# = Asc(Mid$(Record, MidOffset1, 1))
        M2# = Asc(Mid$(Record, MidOffset2, 1))
        L# = Asc(Mid$(Record, LowOffset, 1))
        Flag% = Asc(Mid$(Record, FlagOffset, 1))
        TempCount = H# * 2091752# + M1# * 16384# + M2# * 128# + L#
        
        Set ItmX = lv.ListItems.Add(, , Str(i))
        ItmX.SubItems(1) = �ð����·κ�ȯ(TempCount)
        
        ItmX.SubItems(2) = ""
        
        If TempCount <> �������� Then
           If (Flag And �������ʽ�) <> 0 Then ItmX.SubItems(2) = ItmX.SubItems(2) & "���� "
           If (Flag And �������ʽ�) <> 0 Then ItmX.SubItems(2) = ItmX.SubItems(2) & "2�� "
           If (Flag And ����) <> 0 Then ItmX.SubItems(2) = ItmX.SubItems(2) & "���� "
           If (Flag And �����˹�Ģ) <> 0 Then ItmX.SubItems(2) = ItmX.SubItems(2) & "���� "
        
           Bonus = �÷��׿���������ġ(Flag%)
           TempCount = TempCount + Bonus
           If TempCount < 1 Then TempCount = 1
                 
           If Bonus > 0 Then
              ItmX.SubItems(3) = "+" & Format(Bonus / 1000, "##.####")
           Else
              ItmX.SubItems(3) = Format(Bonus / 1000, "##.####")
           End If
           
           ItmX.SubItems(4) = �ð����·κ�ȯ(TempCount)
        End If
        
        Select Case TempCount
           Case ��������
              ItmX.SubItems(2) = "��������"
           Case ������(÷��).�ְ���
              ItmX.SubItems(2) = "�ְ���"
        End Select
        
        If ItmX.SubItems(2) = "" Then ItmX.SubItems(2) = "."
    Next i
    
    lv.Refresh
End Sub

Sub �����̷��(÷�� As Integer)
    Dim c As Integer
        
    '�ӽú���
    Dim ����     As Integer
    Dim �̸�     As String * ���ڿ�����
    Dim �κ���   As String * ���ڿ�����
    Dim �б�     As String * ���ڿ�����
    Dim ����     As Integer
    Dim ����Ƚ�� As Integer
    Dim �ְ��� As Integer
    Dim ������ As String * �����ϱ���
    Dim ���ð� As Long
    
    ���� = ������(÷��).����
    �̸� = ������(÷��).�̸�
    �κ��� = ������(÷��).�κ���
    �б� = ������(÷��).�б�
    ���� = ������(÷��).����
    ����Ƚ�� = ������(÷��).����Ƚ��
    �ְ��� = ������(÷��).�ְ���
    ������ = ������(÷��).������
    ���ð� = ������(÷��).���ð�
    
    '÷�ڿ� �����Ǵ� ������
    �����ڻ��� ÷��
    �����ο� = �����ο� + 1
    �����ڹ迭�Է� �����ο�, ����, �̸�, �κ���, �б�, ����, ����Ƚ��, �ְ���, ������, ���ð�
    
    '��������� ���� ���� �κ� ����
    For c = 1 To �����ο�
        ������� c
    Next c
    
    '�����ٲٱ� �Ϸ�
End Sub

Sub ��ϻ���(÷�� As Integer)
    Dim c As Integer
    
    ������(÷��).���� = 0
    ������(÷��).�ְ��� = 0
    ������(÷��).����Ƚ�� = 0
    ������(÷��).������ = ""
    ������(÷��).���ð� = 0
    
    For c = 1 To �����ο�
        ������� c
    Next c
End Sub

Sub �������(÷�� As Integer)
    Dim ���ϼ��� As Integer
    ���ϼ��� = �����ο�
    For i = 1 To �����ο�
       If ������(i).�ְ��� = 0 Then ���ϼ��� = ���ϼ��� - 1
    Next i
    For i = 1 To �����ο�
       Do
          If ������(i).�ְ��� <> 0 Then
             ������(i).���� = ���ϼ���
             Exit Do
          Else
             ������(i).���� = 0
             i = i + 1
             If i > �����ο� Then Exit For
          End If
       Loop
       For j = 1 To �����ο�
           If i <> j And ������(j).�ְ��� <> 0 Then
              If ������(i).�ְ��� <= ������(j).�ְ��� Then ������(i).���� = ������(i).���� - 1
           End If
       Next j
    Next i
End Sub

Sub ��Ͽ������ð���ȯ(÷�� As Integer, ���ð� As Long)
    ������(÷��).���ð� = ���ð�
End Sub
Sub ����߰�(÷�� As Integer, ��� As Long, �÷��� As Integer)
    Dim Temp��� As Long
    Dim c As Integer
    
    Temp��� = ���
    
    n = ������(÷��).����Ƚ��
    
    HighOffset = ��ϵ�����ũ�� * n + 1  ' ��������Ʈ
    MidOffset1 = ��ϵ�����ũ�� * n + 2  ' �߰�����Ʈ
    MidOffset2 = ��ϵ�����ũ�� * n + 3  ' �߰�����Ʈ
    LowOffset = ��ϵ�����ũ�� * n + 4   ' ��������Ʈ
    FlagOffset = ��ϵ�����ũ�� * n + 5  ' �÷��׹���Ʈ
    
    '�� �ƽ�Ű �ڵ� ���� ��������� ȯ���� �븩, �� ����Ʈ�� ���� ������ ������ 0 ~ 127�� ����
    H# = Int(Temp��� / 2091752#)
    Temp��� = Temp��� Mod 2091752#
    M1# = Int(Temp��� / 16384#)
    Temp��� = Temp��� Mod 16384#
    M2# = Int(Temp��� / 128#)
    L# = Temp��� Mod 128#
    
    '����� �����Ͽ� �߰�
    Mid$(������(÷��).������, HighOffset, 1) = Chr$(H#)
    Mid$(������(÷��).������, MidOffset1, 1) = Chr$(M1#)
    Mid$(������(÷��).������, MidOffset2, 1) = Chr$(M2#)
    Mid$(������(÷��).������, LowOffset, 1) = Chr$(L#)
    Mid$(������(÷��).������, FlagOffset, 1) = Chr$(�÷���)
    
    ������(÷��).����Ƚ�� = ������(÷��).����Ƚ�� + 1
    ������(÷��).���ð� = Get���ð�
    
    '��Ͽ� ���ʽ��� ������ ���� ����ġ �ο�( ���� ����� �����Ͽ� ����� )
    ��� = ��� + �÷��׿���������ġ(�÷���)
    
    '�ְ��Ϻ��� ������ �ְ��� ����( ��������� ��� ���� ���� )
    If (������(÷��).�ְ��� > ��� Or ������(÷��).�ְ��� = 0) And ��� <> �������� Then ������(÷��).�ְ��� = ���
    
    '��������� ���� ���� �κ� ����
    For c = 1 To �����ο�
        ������� c
    Next c
End Sub

Sub �������ְ��Ϲ׼�������(÷�� As Integer)
    Dim Record As String
    Dim c As Integer
    Dim TempCount As Long
    
    Record = ������(÷��).������
    If ������(÷��).�ְ��� <> 0 Then
       ������(÷��).�ְ��� = MINIMUM
       For c = 1 To ������(÷��).����Ƚ��
           HighOffset = ��ϵ�����ũ�� * (c - 1) + 1 ' ��������Ʈ
           MidOffset1 = ��ϵ�����ũ�� * (c - 1) + 2 ' �߾ӹ���Ʈ1
           MidOffset2 = ��ϵ�����ũ�� * (c - 1) + 3 ' �߾ӹ���Ʈ2
           LowOffset = ��ϵ�����ũ�� * (c - 1) + 4  ' ��������Ʈ
           FlagOffset = ��ϵ�����ũ�� * (c - 1) + 5 ' �÷��׹���Ʈ
           
           H# = Asc(Mid$(Record, HighOffset, 1))
           M1# = Asc(Mid$(Record, MidOffset1, 1))
           M2# = Asc(Mid$(Record, MidOffset2, 1))
           L# = Asc(Mid$(Record, LowOffset, 1))
           f% = Asc(Mid$(Record, FlagOffset, 1))
           TempCount = H# * 2091752# + M1# * 16384# + M2# * 128# + L#
        
           'TempCount�� ��� �����͸� ������ �ִ�. �ϴ� TempCount�� ������ ���� �߻��Ǵ� ����ġ�� �ο�.
           TempCount = TempCount + �÷��׿���������ġ(f%)
           
           '����ġ �ο��� ������� �ְ��ϰ� ��
           If TempCount < 0 Then TempCount = 1
           If ������(÷��).�ְ��� > TempCount Then ������(÷��).�ְ��� = TempCount
       Next c
       If ������(÷��).�ְ��� = MINIMUM Then ������(÷��).�ְ��� = 0
       
       '��������� ���� ���� �κ� ����
       For c = 1 To �����ο�
           ������� c
       Next c
    End If
End Sub
Sub ���ñ�ϻ���(÷�� As Integer, ��Ϲ�ȣ As Integer)
    Temp$ = ������(÷��).������
    If ��Ϲ�ȣ = 1 Then
       Result$ = ""
    Else
       Result$ = Left$(Temp$, ��ϵ�����ũ�� * (��Ϲ�ȣ - 1))
    End If
    
    HighOffset = ��ϵ�����ũ�� * (��Ϲ�ȣ - 1) + 1 ' ��������Ʈ
    MidOffset1 = ��ϵ�����ũ�� * (��Ϲ�ȣ - 1) + 2 ' �߾ӹ���Ʈ1
    MidOffset2 = ��ϵ�����ũ�� * (��Ϲ�ȣ - 1) + 3 ' �߾ӹ���Ʈ2
    LowOffset = ��ϵ�����ũ�� * (��Ϲ�ȣ - 1) + 4  ' ��������Ʈ
    FlagOffset = ��ϵ�����ũ�� * (��Ϲ�ȣ - 1) + 5 ' �÷��׹���Ʈ
        
    H# = Asc(Mid$(Temp$, HighOffset, 1))
    M1# = Asc(Mid$(Temp$, MidOffset1, 1))
    M2# = Asc(Mid$(Temp$, MidOffset2, 1))
    L# = Asc(Mid$(Temp$, LowOffset, 1))
    f% = Asc(Mid$(Temp$, FlagOffset, 1))
    TempCount = H# * 2091752# + M1# * 16384# + M2# * 128# + L#
    
    If TempCount <> �������� Then
       ������(÷��).���ð� = ������(÷��).���ð� - TempCount
    End If
        
    '�ʿ���� �κ��� �߶󳽴�.
    Result$ = Result$ + Mid$(Temp$, 1 + (��ϵ�����ũ�� * ��Ϲ�ȣ), Len(Temp$) - (1 + (��ϵ�����ũ�� * ��Ϲ�ȣ)))
    
    '�ʿ���� �κ��� ���ְ� ���� ������ �ƽ�Ű�ڵ� 0�� ���ڷ� ä���.
    Result$ = Result$ + String(�����ϱ��� - Len(Result$), Chr$(0))
        
    ������(÷��).������ = Result$
    ������(÷��).����Ƚ�� = ������(÷��).����Ƚ�� - 1
    
    '�ְ��� ����
    �������ְ��Ϲ׼������� ÷��
End Sub

Sub �������̵�(÷�� As Integer)
    Dim �ű������� As �����ڼ���
    Dim �ӽð���(�ִ�������) As �����ڼ���
    Dim SwapRec As �����ڼ���
    Dim i As Integer, j As Integer
    Dim OutStr As String
    
    �ű�������.���� = ������(÷��).����
    �ű�������.�̸� = ������(÷��).�̸�
    �ű�������.�κ��� = ������(÷��).�κ���
    �ű�������.�б� = ������(÷��).�б�
    �ű�������.���� = ������(÷��).����
    �ű�������.����Ƚ�� = ������(÷��).����Ƚ��
    �ű�������.�ְ��� = ������(÷��).�ְ���
    �ű�������.������ = ������(÷��).������
    �ű�������.���ð� = ������(÷��).���ð�
    
    For i = ÷�� To 2 Step -1
       ������(i).���� = ������(i - 1).����
       ������(i).�̸� = ������(i - 1).�̸�
       ������(i).�κ��� = ������(i - 1).�κ���
       ������(i).�б� = ������(i - 1).�б�
       ������(i).���� = ������(i - 1).����
       ������(i).����Ƚ�� = ������(i - 1).����Ƚ��
       ������(i).�ְ��� = ������(i - 1).�ְ���
       ������(i).������ = ������(i - 1).������
       ������(i).���ð� = ������(i - 1).���ð�
    Next i
        
    ������(1).���� = �ű�������.����
    ������(1).�̸� = �ű�������.�̸�
    ������(1).�κ��� = �ű�������.�κ���
    ������(1).�б� = �ű�������.�б�
    ������(1).���� = �ű�������.����
    ������(1).����Ƚ�� = �ű�������.����Ƚ��
    ������(1).�ְ��� = �ű�������.�ְ���
    ������(1).������ = �ű�������.������
    ������(1).���ð� = �ű�������.���ð�
End Sub
Sub ������ϸ����(���ϸ� As String)
    Dim �ӽð���(�ִ�������) As �����ڼ���
    Dim SwapRec As �����ڼ���
    Dim i As Integer, j As Integer
    Dim OutStr As String
    
    For i = 1 To �����ο�
       �ӽð���(i).�κ��� = ������(i).�κ���
       �ӽð���(i).���ð� = ������(i).���ð�
       �ӽð���(i).���� = ������(i).����
       �ӽð���(i).���� = ������(i).����
       �ӽð���(i).�̸� = ������(i).�̸�
       �ӽð���(i).�ְ��� = ������(i).�ְ���
       �ӽð���(i).�б� = ������(i).�б�
       
       If �ӽð���(i).���� = 0 Then �ӽð���(i).���� = 32767
    Next i
        
    For i = 1 To �����ο�
       For j = 1 To �����ο�
          If �ӽð���(i).���� < �ӽð���(j).���� Then
             'Swap
             SwapRec.�κ��� = �ӽð���(i).�κ���
             SwapRec.���ð� = �ӽð���(i).���ð�
             SwapRec.���� = �ӽð���(i).����
             SwapRec.���� = �ӽð���(i).����
             SwapRec.�̸� = �ӽð���(i).�̸�
             SwapRec.�ְ��� = �ӽð���(i).�ְ���
             SwapRec.�б� = �ӽð���(i).�б�
                
             �ӽð���(i).�κ��� = �ӽð���(j).�κ���
             �ӽð���(i).���ð� = �ӽð���(j).���ð�
             �ӽð���(i).���� = �ӽð���(j).����
             �ӽð���(i).���� = �ӽð���(j).����
             �ӽð���(i).�̸� = �ӽð���(j).�̸�
             �ӽð���(i).�ְ��� = �ӽð���(j).�ְ���
             �ӽð���(i).�б� = �ӽð���(j).�б�
                
             �ӽð���(j).�κ��� = SwapRec.�κ���
             �ӽð���(j).���ð� = SwapRec.���ð�
             �ӽð���(j).���� = SwapRec.����
             �ӽð���(j).���� = SwapRec.����
             �ӽð���(j).�̸� = SwapRec.�̸�
             �ӽð���(j).�ְ��� = SwapRec.�ְ���
             �ӽð���(j).�б� = SwapRec.�б�
          End If
       Next j
    Next i
    
    Open ���ϸ� For Output As #1
       LineStr$ = String$(6 + 32 + 32 + 32 + 15, "-")
       Print #1, LineStr$
       OutStr = "| ���� |          �б� �� �Ҽ�          |              �̸�              |              �κ���            |�ְ��� |"
       Print #1, OutStr
       Print #1, String$(6 + 32 + 32 + 32 + 15, "-")
       For i = 1 To �����ο�
          If (�ӽð���(i).���� < 32767) Then
             s$ = LTrim$(Str$(�ӽð���(i).����))
             s$ = Space$(6 - Len(s$)) + s$
             OutStr = "|" & s$ & "|"
          Else
             OutStr = "|------|"
          End If
          s$ = ""
          n�б�$ = LeftB$(RTrim$(�ӽð���(i).�б�), 32)
          n�б�$ = n�б�$ + Space$(32 - LenB(StrConv(n�б�$, vbFromUnicode)))
          OutStr = OutStr & n�б�$ & "|"
          n�̸�$ = LeftB$(RTrim$(�ӽð���(i).�̸�), 32)
          n�̸�$ = n�̸�$ + Space$(32 - LenB(StrConv(n�̸�$, vbFromUnicode)))
          OutStr = OutStr & n�̸�$ & "|"
          n�κ���$ = LeftB$(RTrim$(�ӽð���(i).�κ���), 32)
          n�κ���$ = n�κ���$ + Space$(32 - LenB(StrConv(n�κ���$, vbFromUnicode)))
          OutStr = OutStr & n�κ���$ & "|"
          
          If (�ӽð���(i).�ְ��� > 0) Then
             OutStr = OutStr & �ð����·κ�ȯ(�ӽð���(i).�ְ���) & "|"
          Else
             OutStr = OutStr & "---------" & "|"
          End If
          Print #1, OutStr
       Next i
       Print #1, LineStr$
    Close #1
End Sub

Function ��Ͽ����÷��׾��(÷�� As Integer, ȸ�� As Integer) As Integer
    ��Ͽ����÷��׾�� = Asc(Mid$(������(÷��).������, ��ϵ�����ũ�� * (ȸ�� - 1) + 5))
End Function

Sub ��Ͽ��÷��׵�����(÷�� As Integer, ȸ�� As Integer, �÷��� As Integer)
    Mid$(������(÷��).������, ��ϵ�����ũ�� * (ȸ�� - 1) + 5) = Chr$(�÷���)
    �������ְ��Ϲ׼������� ÷��
End Sub

Function Set�����ڻ��ð�(÷�� As Integer, ���ð� As Long)
    ������(÷��).���ð� = ���ð�
End Function

Function Get�����ڻ��ð�(÷�� As Integer) As Long
    Get�����ڻ��ð� = ������(÷��).���ð�
End Function

Function Get�����ο�() As Integer
    Get�����ο� = �����ο�
End Function
