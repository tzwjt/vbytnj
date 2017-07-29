Attribute VB_Name = "basBarCode"

'*************************************************************************
'**ģ �� ����basBarCode
'**��    ������ȡ����������
'**��    ����V1.0.0
'*************************************************************************

Option Explicit
Private Type KeyboardBytes
    kbByte(0 To 255) As Byte
End Type
Dim kbArray As KeyboardBytes
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As KeyboardBytes) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As KeyboardBytes, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Type EVENTMSG
    message As Long
    paramL As Long
    paramH As Long
    Time As Long
    hwnd As Long
End Type
Private Type BARCODES
    VirtKey As Long         '������
    ScanCode As Long           'ɨ����
    KeyName As String       '��������
    AscII As Long           'AscII
    Chr As String           '�ַ�
   
    BarCode As String      'ɨ������Ϣ
    Time As Date            'ɨ��ʱ��
    bGetFlag As Boolean     '�Ƿ��ѻ�ȡɨ����
End Type
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Private Const WH_KEYBOARD_LL = 13
Private m_lHook As Long
Public g_BarCode As BARCODES
'*************************************************************************
'**�� �� ����SetHook / UnHook
'**��    �룺��
'**��    ������
'**����������װж����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    ����V1.0.0
'*************************************************************************
Public Sub SetHook()
    m_lHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf CallHookProc, App.hInstance, 0)
End Sub
Public Sub UnHook()
    If m_lHook <> 0 Then
        UnhookWindowsHookEx m_lHook
    End If
End Sub
'*************************************************************************
'**�� �� ����GetBarCode
'**��    �룺��
'**��    ����(String) -
'**������������ȡɨ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    ����V1.0.0
'*************************************************************************
Public Function GetBarCode() As String
    If g_BarCode.bGetFlag = True Then
        g_BarCode.bGetFlag = False
        GetBarCode = g_BarCode.BarCode
    Else
        GetBarCode = ""
    End If
End Function
'*************************************************************************
'**�� �� ����CallHookProc
'**��    �룺ByVal code(Long)   -
'**        ��ByVal wParam(Long) -
'**        ��ByVal lParam(Long) -
'**��    ����(Long) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    ����V1.0.0
'*************************************************************************
Private Function CallHookProc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim msg As EVENTMSG
    Dim strKeyName As String
    Dim lngKey As Long
    Static lngTime As Long
    Static strBarCode As String
If code = 0 Then
        CopyMemory msg, lParam, LenB(msg)
        If wParam = &H100 Then   'WM_KEYDOWN
            g_BarCode.VirtKey = msg.message And &HFF           '������
            g_BarCode.ScanCode = msg.paramL And &HFF              'ɨ����
           
            strKeyName = Space(255)
            If GetKeyNameText(g_BarCode.ScanCode * 65536, strKeyName, 255) > 0 Then  '����
                g_BarCode.KeyName = Trim(strKeyName)
            Else
                g_BarCode.KeyName = ""
            End If
'---------------------------------------
            Call GetKeyboardState(kbArray)
            If ToAscii(g_BarCode.VirtKey, g_BarCode.ScanCode, kbArray, lngKey, 0) > 0 Then
                g_BarCode.AscII = lngKey
                g_BarCode.Chr = Chr(lngKey)
            End If
'--------------------
            If Abs(GetCurrentTime - lngTime) > 50 Then   '
                strBarCode = g_BarCode.Chr
            Else
                If (msg.message And &HFF) = 13 And Len(strBarCode) > 3 Then '�س�
                    g_BarCode.BarCode = strBarCode
                   ' MsgBox strBarCode
                    
                    g_BarCode.Time = Now
                    g_BarCode.bGetFlag = True
                End If
                strBarCode = strBarCode & g_BarCode.Chr
            End If
            lngTime = GetCurrentTime
            '---------------------------------------
            '���Դ���
            'Call ShowKeyInfo
            '---------------------------------------
        End If
End If
CallHookProc = CallNextHookEx(m_lHook, code, wParam, lParam)
End Function
'��ʾ������Ϣ
Public Sub ShowKeyInfo()
   ' frmDemo.txtKey(0) = g_BarCode.KeyName
   ' frmDemo.txtKey(1) = g_BarCode.VirtKey
  '  frmDemo.txtKey(2) = g_BarCode.ScanCode
'frmDemo.txtKey(3) = g_BarCode.AscII
 '   frmDemo.txtKey(4) = g_BarCode.Chr
 '   frmDemo.txtBarCode = g_BarCode.BarCode
   
 '   frmDemo.lblTime = g_BarCode.Time
End Sub

