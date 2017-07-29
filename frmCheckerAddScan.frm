VERSION 5.00
Begin VB.Form frmCheckerAddScan 
   Caption         =   "����Ա������ɨ��"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   13995
   Icon            =   "frmCheckerAddScan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   13995
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrScan 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Caption         =   "����Ա������ɨ��"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   2640
      Width           =   10455
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
      Left            =   3000
      TabIndex        =   0
      Top             =   5280
      Width           =   9975
   End
   Begin VB.Menu mnuExit 
      Caption         =   "�˳�����Ա������ɨ��"
   End
End
Attribute VB_Name = "frmCheckerAddScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim index2 As Integer
Dim objVoice As Object
Dim fEmpNo As String
Dim fProductCode As String
Dim fProcessNo As String
Dim fTime As Date
  
'Dim WithEvents Voice As SpVoice









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
    fEmpNo = ""
    fProductCode = ""
    fProcessNo = ""
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
    Dim prefix As String
    Dim recordsAffected As Long
    Dim rsResult As ADODB.Recordset
    Dim pos As Integer
    Dim xPos As Integer
    Dim strMsg As String
    Dim strSpeak As String
    Dim empNo As String
    Dim empName As String
    Dim barCode As String
    Dim sql As String
    
    
    pos = 0
    xPos = 0
    prefix = "\\\"
   
    DoEvents
 
    strBarCode = GetBarCode
    strMsg = ""
    strSpeak = ""
    
    If Len(strBarCode) < 1 Then
         Exit Sub
    End If
    
    pos = InStr(1, strBarCode, " ")
    If pos < 1 Then
        strMsg = strBarCode & " ɨ�����������ɨ��!"
        strSpeak = "ɨ�����������ɨ��!"
        objVoice.Speak strMsg, 1
      '  MsgBox strMsg
        Exit Sub
    End If
    
   ' MsgBox strBarCode
    'Exit Sub
    
        
    strBarCode = getValidBarCode(strBarCode)
    strBarCode = Trim(strBarCode)
      
    xPos = InStr(1, strBarCode, prefix)
        
    If xPos < 1 Then
        strMsg = strBarCode & "ɨ��ǹ����ȷ�����ü���Ա��ɨ��ǹ����ɨ��!"
        strSpeak = "ɨ��ǹ����ȷ�����ü���Ա��ɨ��ǹ����ɨ��!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
    
    pos = InStr(1, strBarCode, " ")
    If pos < 1 Then
        strMsg = strBarCode & " ɨ�����������ɨ��!"
        strSpeak = "ɨ�����������ɨ��!"
        objVoice.Speak strSpeak, 1
       ' MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
    
   
    empNo = Val(Left(strBarCode, pos - 1))
    barCode = Right(strBarCode, Len(strBarCode) - (xPos + Len(prefix)) + 1)
     
    If Len(barCode) > 3 Then  'ɨ��������
     
        If fProductCode <> "" And fProductCode <> barCode And fEmpNo <> empNo And DateDiff("n", fTime, Now) < 1 Then
            strMsg = "ǰһ��������ɨ��Ʒ���뻹δ�������,���Ժ�"
            objVoice.Speak strMsg, 1
          '  MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
            Exit Sub
        End If
            
        
        fProductCode = barCode
        fTime = Now
        fEmpNo = empNo
        fProcessNo = ""
        strMsg = "��������ɨ��������!"
        objVoice.Speak strMsg, 1
      '  MsgBox strMsg
        Exit Sub
    End If
     'ɨ��������
    If Trim(fProductCode) = "" Then
        strMsg = "ɨ���������ɨ��Ʒ����!"
        objVoice.Speak strMsg, 1
       ' MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
        
    If empNo <> fEmpNo Then
        strMsg = "����ɨ���������ɨ������ǰ��ɨ��Ʒ�����ɨ��������ͬ!"
        objVoice.Speak strMsg, 1
   '     MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
            
    If DateDiff("n", fTime, Now) > 1 Then
        strMsg = "����ɨ������ʱ��������ɨ��Ʒ����!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak strMsg, 1
      '  MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
        
        
    fProcessNo = barCode
        
    If Trim(fProductCode) = "" Or Trim(fProcessNo) = "" Then
        strMsg = "ɨ�봦�����������ɨ��Ʒ���룬��ɨ��������!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak strMsg, 1
        MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
    
     '���ݿ����
    sql = "select emp_no, name as emp_name from yt_employee where emp_no = '" & fEmpNo & "' and status = 1"
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        strMsg = "���Ϊ " & fEmpNo & " ����ЧԱ��������,ɨ�봦��ʧ��!"
        rsResult.Close
        Set rsResult = Nothing
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak strMsg, 1
      '  MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
    End If
            
    empNo = rsResult.Fields("emp_no")
    empName = rsResult.Fields("emp_name")
            
           
    sql = "update yt_produce_scan set operator_no='" & empNo & "', operator_name='" & empName & "',scan_status=2, scan_time='" & Now & "', update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & fProductCode & "' and  yt_produce_scan.process_no = " & fProcessNo & ""
                        
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        strMsg = "���ݴ���ʧ�ܣ�������ɨ��Ʒ���룬��ɨ��������!"
        fProductCode = ""
        fProcessNo = ""
        fEmpNo = ""
        Set rsResult = Nothing
        objVoice.Speak strMsg, 1
   '     MsgBox strMsg, vbExclamation, "����Ա������ɨ��"
        Exit Sub
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
        
    strMsg = "������ɨ�봦��ɹ�!"
    fProductCode = ""
    fProcessNo = ""
    fEmpNo = ""
    objVoice.Speak strMsg, 1
     '  MsgBox strMsg
    Exit Sub
        
        
      
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
     
    getValidBarCode = str
End Function










Private Sub �˳�����Ա������ɨ��_Click()

End Sub
