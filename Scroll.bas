Attribute VB_Name = "Module8"
Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
   
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long

Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3

Sub ��ũ����ġ����(Handle As OLE_HANDLE, ��ġ As Integer)
    Dim ScrInfo As SCROLLINFO
    Dim nPos As Integer
    
    '���� ��ũ�� ��ġ�� ������ ����� ����.
    'SB_PAGEUP�̳� SB_PAGEDOWN�� �̿��Ͽ� �����Ѵ�.
    ScrInfo.cbSize = Len(ScrInfo)
    ScrInfo.fMask = SIF_ALL
    
    nPos = ��ġ - 1
    
    Do
        GetScrollInfo Handle, SB_VERT, ScrInfo
        If ScrInfo.nTrackPos <= nPos And ScrInfo.nTrackPos + ScrInfo.nPage > nPos Then Exit Do
        
        '��ũ�� ������ �Ѿ�� ���
        If nPos >= ScrInfo.nTrackPos + ScrInfo.nPage Then
           SendMessage Handle, WM_VSCROLL, SB_PAGEDOWN, 0
        End If
        
        If nPos < ScrInfo.nTrackPos Then
           SendMessage Handle, WM_VSCROLL, SB_PAGEUP, 0
        End If
    Loop
        
End Sub

