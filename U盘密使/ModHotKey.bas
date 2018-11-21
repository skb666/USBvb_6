Attribute VB_Name = "ModHotKey"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fskey_Modifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Const WM_HOTKEY = &H312
Const MOD_ALT = &H1
Const MOD_CONTROL = &H2
Const MOD_SHIFT = &H4
Const GWL_WNDPROC = (-4)  'ע�ͣ����ں����ĵ�ַ
Const KEYEVENTF_KEYUP = &H2 'ע�ͣ����ں����ĵ�ַ
'Download by http://www.codefans.net
Dim key_preWinProc As Long 'ע�ͣ��������洰����Ϣ
Dim key_Modifiers As Long, key_uVirtKey As Long, key_idHotKey As Long
Dim key_IsWinAddress  As Boolean 'ע�ͣ��Ƿ�ȡ�ô�����Ϣ���ж�
Public isend As Long 'ע�ͣ��������洰����Ϣ
Const KEY_A = 65
Const KEY_S = 83
Const KEY_P = 80
Const KEY_I = 73
Const KEY_R = 82
Const KEY_N = 78
Const KEY_E = 69
Const KEY_C = 67
Const KEY_O = 79
Const KEY_U = 85
Const KEY_T = 84
Const KEY_V = 86
Const KEY_M = 77
Const KEY_L = 76
Const KEY_Shift = 16
Const KEY_Ctrl = 17
Const KEY_9 = 57

Function keyWndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    If Msg = WM_HOTKEY Then
        Select Case wParam 'ע�ͣ�wParam ֵ���� key_idHotKey
            Case 1 'ע�ͣ����� 3 ���ȼ���,3 ���ȼ�����Ӧ�Ĳ���,����������ĳ����У�ֻҪ�޸Ĵ˴��Ϳ�����
            Frmmain.WindowState = 0
            Frmmain.Show
            Frmmain.Visible = True
            Case 2
            GoTo aso2
            Case 3
            StopCopy = 1
            Case 4
            NoCopys = False
            Frmmain.Command5.Caption = "��ͣ���"
            Frmmain.Label3.Caption = "��ǰ�����ڿ�����!"
            Case 5
            NoCopys = True
            Frmmain.Command5.Caption = "��ʼ���"
            Frmmain.Label3.Caption = "����Ѿ��رգ������²����U�̽������и��ƣ�"
        End Select
    End If
    
'ע��:     ����Ϣ���͸�ָ���Ĵ���
    keyWndproc = CallWindowProc(key_preWinProc, hwnd, Msg, wParam, lParam)
Exit Function
aso2:
keyWndproc = CallWindowProc(key_preWinProc, hwnd, Msg, wParam, lParam)
ProcUse = False
Unload Frmmain
End Function

Function SetHotkey(ByVal KeyId As Long, ByVal KeyAss0 As String, ByVal Action As String)
    Dim KeyAss1 As Long
    Dim KeyAss2 As String
    Dim i As Long
    i = InStr(1, KeyAss0, ",")
    If i = 0 Then
        KeyAss1 = Val(KeyAss0)
        KeyAss2 = ""
    Else
        KeyAss1 = Right(KeyAss0, Len(KeyAss0) - i)
        KeyAss2 = Left(KeyAss0, i - 1)
    End If
    key_idHotKey = 0
    key_Modifiers = 0
    key_uVirtKey = 0
    If key_IsWinAddress = False Then  'ע�ͣ��ж��Ƿ���Ҫȡ�ô�����Ϣ������ظ�ȡ��,�����ָ�����ʱ��������ɳ�������
'ע��:         ��¼ԭ����window�����ַ
        key_preWinProc = GetWindowLong(Frmmain.hwnd, GWL_WNDPROC)
'ע��:         ���Զ���������ԭ����window����
        SetWindowLong Frmmain.hwnd, GWL_WNDPROC, AddressOf keyWndproc
    End If
    key_idHotKey = KeyId
    Select Case Action
        Case "Add"
            If KeyAss2 = "Ctrl" Then key_Modifiers = MOD_CONTROL
            If KeyAss2 = "Alt" Then key_Modifiers = MOD_ALT
            If KeyAss2 = "Shift" Then key_Modifiers = MOD_SHIFT
            If KeyAss2 = "Ctrl+Alt" Then key_Modifiers = MOD_CONTROL + MOD_ALT
            If KeyAss2 = "Ctrl+Shift" Then key_Modifiers = MOD_CONTROL + MOD_SHIFT
            If KeyAss2 = "Ctrl+Alt+Shift" Then key_Modifiers = MOD_CONTROL + MOD_ALT + MOD_SHIFT
            If KeyAss2 = "Shift+Alt" Then key_Modifiers = MOD_SHIFT + MOD_ALT
            key_uVirtKey = Val(KeyAss1)
            RegisterHotKey Frmmain.hwnd, key_idHotKey, key_Modifiers, key_uVirtKey 'ע�ͣ��򴰿�ע��ϵͳ�ȼ�
            key_IsWinAddress = True 'ע�ͣ�����Ҫ��ȡ�ô�����Ϣ
            
        Case "Del"
            SetWindowLong Frmmain.hwnd, GWL_WNDPROC, key_preWinProc 'ע�ͣ��ָ�������Ϣ
            UnregisterHotKey Frmmain.hwnd, key_uVirtKey 'ע�ͣ�ȡ��ϵͳ�ȼ�
            key_IsWinAddress = False 'ע�ͣ������ٴ�ȡ�ô�����Ϣ
    End Select
End Function
