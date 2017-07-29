VERSION 5.00
Begin VB.Form frmAutoScan 
   Caption         =   "����ɨ��"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   15825
   Icon            =   "frmAutoScan.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   15825
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrScan 
      Interval        =   10
      Left            =   960
      Top             =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Ϊ�˱�֤����ȷ��ȡɨ��,��������һ�´˽���,����֤��겻�뿪�˽���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   6000
      Width           =   9975
   End
   Begin VB.Label Label1 
      Caption         =   "�� �� ɨ ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   5040
      TabIndex        =   0
      Top             =   3360
      Width           =   7215
   End
   Begin VB.Menu mnuExit 
      Caption         =   "�˳�����ɨ��"
   End
End
Attribute VB_Name = "frmAutoScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objVoice As Object

Dim fProductCode As String
Dim fProcessNo As String
Dim fEmpNo As String
Dim fTime As Date


 
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, _
ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)




Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_Load()
'***********************************************
      If SysPass() = False Then  ''Ȩ��
        End
        Exit Sub
       End If
'***********************************************
   
    Set objVoice = CreateObject("SAPI.SpVoice")
    SetHook
End Sub



Private Sub Form_Resize()
    Label1.Left = (Me.Width - Label1.Width) / 2
    Label2.Left = (Me.Width - Label2.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   UnHook
   Set objVoice = Nothing
    
   
End Sub



Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tmrScan_Timer()
    Dim strBarCode As String
    Dim processNo As Integer
    Dim productCode As String
    Dim strSpeak As String
    Dim pos As Integer
    Dim xPos As Integer
    Dim strMsg As String
    Dim prefix As String
    
    pos = 0
    xPos = 0
    prefix = "\\\"
    strSpeak = ""
    strMsg = ""
   
    DoEvents
    strBarCode = GetBarCode
    
    
    If Len(strBarCode) < 1 Then
        Exit Sub
    End If
    
    pos = InStr(1, strBarCode, " ")
    
    If pos < 1 Then
        strSpeak = "�ղ�ɨ�����������ɨ��!"
        strMsg = strBarCode & " ɨ�����������ɨ��!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "����ɨ��"
        Exit Sub
    End If
         
    
    strBarCode = getValidBarCode(strBarCode)
    strBarCode = Trim(strBarCode)
    
     xPos = InStr(1, strBarCode, prefix)
        
    If xPos > 0 Then
        strMsg = strBarCode & "ɨ��ǹ����ȷ���������ü���Ա��ɨ��ǹɨ��!"
        strSpeak = "ɨ��ǹ����ȷ���������ü���Ա��ɨ��ǹɨ��!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
      
    
     pos = InStr(1, strBarCode, " ")
     If pos <> 3 Then
        strSpeak = "�ղ�ɨ�����������ɨ��!"
        strMsg = strBarCode & " ɨ�����������ɨ��!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "����ɨ��"
        Exit Sub
    End If
                
        
    processNo = Val(Left(strBarCode, pos - 1))
    productCode = Mid(strBarCode, pos + 1)
    
    If saveProduceScan(processNo, productCode) = True Then
        
        strSpeak = processNo & "���� ɨ��ɹ�!"
        strMsg = strBarCode & "����ɹ�!"
           
    Else
        strSpeak = processNo & "����ɨ�����������ɨ��!"
        strMsg = strBarCode & " ����ʧ��!"
    
    End If

    objVoice.Speak strSpeak, 1
        
  '  MsgBox strMsg
    
End Sub




Private Function getValidBarCode(strBarCode As String) As String
    Dim pos As Integer
    Dim strLen As Integer
    Dim i As Integer
    Dim J As Integer
    Dim str As String
    Dim s As Integer
    Dim xPos As Integer
    
   ' MsgBox strBarCode
    xPos = 0
    pos = InStr(1, strBarCode, " ")
        
    For i = pos + 1 To Len(strBarCode)
        xPos = InStr(i, strBarCode, " ")
        If xPos = 0 Then Exit For
    Next
    xPos = i
   
       ' If (pos > 3) Then
    strLen = Len(strBarCode)
    str = ""
        '    s = (pos - 1) / 2
    s = xPos - pos
    For i = 1 To strLen Step s
        str = str + Mid(strBarCode, i, 1)
    Next
       ' Else
          '  str = strBarCode
       ' End If
        'getValidBarCode = str
   
    getValidBarCode = str
End Function
'�ö��̴߳���֮







Public Function doAddBarCode(ByVal xCode As String, ByVal xPrefix As String) As Boolean
     Dim recordsAffected As Long
     Dim rsResult As ADODB.Recordset
     Dim pos As Integer
     Dim xPos As Integer
     Dim msgStr As String
     Dim empNo As String
     Dim empName As String
     Dim BarCode As String
     Dim sql As String
    
     pos = InStr(1, xCode, " ")
     If pos = 0 Then
         msgStr = "ɨ�����������ɨ��!"
         doAddBarCode = False
         objVoice.Speak msgStr, 1
         MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
         Exit Function
     End If
     
     xPos = InStr(1, xCode, xPrefix)
      If xPos = 0 Then
         msgStr = "ɨ�����������ɨ��!"
         doAddBarCode = False
         objVoice.Speak msgStr, 1
         MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
         Exit Function
     End If
         
     empNo = Val(Left(xCode, pos - 1))
    
     BarCode = Right(xCode, Len(xCode) - (xPos + Len(xPrefix) + 1))
     
     If Len(BarCode) > 3 Then  'ɨ��������
     
        If fProductCode <> "" And fProductCode <> BarCode And fEmpNo <> empNo And DateDiff("n", fTime, Now) < 1 Then
            msgStr = "ǰһ��Ա����ɨ��Ʒ���뻹δ�������,���Ժ�"
            doAddBarCode = False
            objVoice.Speak msgStr, 1
            MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
            Exit Function
        End If
            
        
        fProductCode = BarCode
        fTime = Now
        fEmpNo = empNo
        fProcessNo = ""
        msgStr = "�������빤������!"
        doAddBarCode = True
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
        Exit Function
     End If
     'ɨ��������
     If Trim(fProductCode) = "" Then
        msgStr = "ɨ���������ɨ��Ʒ����!"
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
        Exit Function
    End If
        
    If empNo <> fEmpNo Then
        msgStr = "����ɨ���������Ա����ǰ��ɨ��Ʒ�����Ա������ͬһ��!"
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
        Exit Function
    End If
            
    If DateDiff("n", fTime, Now) > 1 Then
        msgStr = "����ɨ������ʱ��������ɨ��Ʒ����!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
        Exit Function
    End If
        
        
    fProcessNo = BarCode
        
    If Trim(fProductCode) = "" Or Trim(fProcessNo) = "" Then
        msgStr = "ɨ�봦�����������ɨ��Ʒ���룬��ɨ��������!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
        Exit Function
    End If
            
    sql = "select emp_no, name as emp_name from yt_employee where emp_no = '" & fEmpNo & "' and status = 1"
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        msgStr = "���Ϊ " & fEmpNo & " ����ЧԱ��������,ɨ�봦��ʧ��!"
        doAddBarCode = False
        rsResult.Close
        Set rsResult = Nothing
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
        Exit Function
    End If
            
    empNo = rsResult.Fields("emp_no")
    empName = rsResult.Fields("emp_name")
            
           
    sql = "update yt_produce_scan set operator_no='" & empNo & "', operator_name='" & empName & "',scan_status=1, scan_time='" & Now & "', update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & fProductCode & "' and  yt_produce_scan.process_no = " & fProcessNo & ""
                        
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        msgStr = "���ݴ���ʧ�ܣ�������ɨ��Ʒ���룬��ɨ��������!"
        fProductCode = ""
        fProcessNo = ""
        fEmpNo = ""
        Set rsResult = Nothing
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
        Exit Function
    End If
        
    'fProductCode��Ӧ�����й����Ƿ���ȫɨ��
    sql = "select count(*) as recordCount from yt_produce_scan " & _
                        " where yt_produce_scan.product_code = '" & fProductCode & "' and scan_status = 0 "
    Set rsResult = conn.Execute(sql)
    If rsResult.Fields("recordCount") < 1 Then
        rsResult.Close
        sql = "update yt_produce_scan set status=1,  update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & fProductCode & "' "
        Set rsResult = conn.Execute(sql)
    Else
        rsResult.Close
    End If
        
    Set rsResult = Nothing
        
    msgStr = "����ɨ�봦��ɹ�!"
    fProductCode = ""
    fProcessNo = ""
    fEmpNo = ""
    doAddBarCode = True
    objVoice.Speak msgStr, 1
    MsgBox msgStr, vbExclamation, "����Ա��ɨ��"
    Exit Function
        
End Function


