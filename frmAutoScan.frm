VERSION 5.00
Begin VB.Form frmAutoScan 
   Caption         =   "生产扫码"
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
      Caption         =   "为了保证能正确获取扫码,请用鼠标点一下此界面,并保证光标不离开此界面"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "生 产 扫 码"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "退出生产扫码"
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
      If SysPass() = False Then  ''权限
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
        strSpeak = "刚才扫码错误，请重新扫码!"
        strMsg = strBarCode & " 扫码错误，请重新扫码!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "生产扫码"
        Exit Sub
    End If
         
    
    strBarCode = getValidBarCode(strBarCode)
    strBarCode = Trim(strBarCode)
    
     xPos = InStr(1, strBarCode, prefix)
        
    If xPos > 0 Then
        strMsg = strBarCode & "扫码枪不正确，不可以用检验员的扫码枪扫码!"
        strSpeak = "扫码枪不正确，不可以用检验员的扫码枪扫码!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
      
    
     pos = InStr(1, strBarCode, " ")
     If pos <> 3 Then
        strSpeak = "刚才扫码错误，请重新扫码!"
        strMsg = strBarCode & " 扫码错误，请重新扫码!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "生产扫码"
        Exit Sub
    End If
                
        
    processNo = Val(Left(strBarCode, pos - 1))
    productCode = Mid(strBarCode, pos + 1)
    
    If saveProduceScan(processNo, productCode) = True Then
        
        strSpeak = processNo & "工序， 扫码成功!"
        strMsg = strBarCode & "处理成功!"
           
    Else
        strSpeak = processNo & "工序，扫码错误，请重新扫码!"
        strMsg = strBarCode & " 处理失败!"
    
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
'用多线程处理之







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
         msgStr = "扫码错误，请重新扫码!"
         doAddBarCode = False
         objVoice.Speak msgStr, 1
         MsgBox msgStr, vbExclamation, "检验员补扫码"
         Exit Function
     End If
     
     xPos = InStr(1, xCode, xPrefix)
      If xPos = 0 Then
         msgStr = "扫码错误，请重新扫码!"
         doAddBarCode = False
         objVoice.Speak msgStr, 1
         MsgBox msgStr, vbExclamation, "检验员补扫码"
         Exit Function
     End If
         
     empNo = Val(Left(xCode, pos - 1))
    
     BarCode = Right(xCode, Len(xCode) - (xPos + Len(xPrefix) + 1))
     
     If Len(BarCode) > 3 Then  '扫生产条码
     
        If fProductCode <> "" And fProductCode <> BarCode And fEmpNo <> empNo And DateDiff("n", fTime, Now) < 1 Then
            msgStr = "前一个员工所扫产品条码还未处理完毕,请稍候"
            doAddBarCode = False
            objVoice.Speak msgStr, 1
            MsgBox msgStr, vbExclamation, "检验员补扫码"
            Exit Function
        End If
            
        
        fProductCode = BarCode
        fTime = Now
        fEmpNo = empNo
        fProcessNo = ""
        msgStr = "接下来请工序条码!"
        doAddBarCode = True
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "检验员补扫码"
        Exit Function
     End If
     '扫工序条码
     If Trim(fProductCode) = "" Then
        msgStr = "扫码错误，请先扫产品条码!"
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "检验员补扫码"
        Exit Function
    End If
        
    If empNo <> fEmpNo Then
        msgStr = "本次扫工序条码的员工与前次扫产品条码的员工不是同一人!"
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "检验员补扫码"
        Exit Function
    End If
            
    If DateDiff("n", fTime, Now) > 1 Then
        msgStr = "两次扫码间隔超时，请重新扫产品条码!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "检验员补扫码"
        Exit Function
    End If
        
        
    fProcessNo = BarCode
        
    If Trim(fProductCode) = "" Or Trim(fProcessNo) = "" Then
        msgStr = "扫码处理错误，请重新扫产品条码，再扫工序条码!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "检验员补扫码"
        Exit Function
    End If
            
    sql = "select emp_no, name as emp_name from yt_employee where emp_no = '" & fEmpNo & "' and status = 1"
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        msgStr = "编号为 " & fEmpNo & " 的有效员工不存在,扫码处理失败!"
        doAddBarCode = False
        rsResult.Close
        Set rsResult = Nothing
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "检验员补扫码"
        Exit Function
    End If
            
    empNo = rsResult.Fields("emp_no")
    empName = rsResult.Fields("emp_name")
            
           
    sql = "update yt_produce_scan set operator_no='" & empNo & "', operator_name='" & empName & "',scan_status=1, scan_time='" & Now & "', update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & fProductCode & "' and  yt_produce_scan.process_no = " & fProcessNo & ""
                        
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        msgStr = "数据处理失败，请重新扫产品条码，再扫工序条码!"
        fProductCode = ""
        fProcessNo = ""
        fEmpNo = ""
        Set rsResult = Nothing
        doAddBarCode = False
        objVoice.Speak msgStr, 1
        MsgBox msgStr, vbExclamation, "检验员补扫码"
        Exit Function
    End If
        
    'fProductCode对应的所有工序是否已全扫码
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
        
    msgStr = "两次扫码处理成功!"
    fProductCode = ""
    fProcessNo = ""
    fEmpNo = ""
    doAddBarCode = True
    objVoice.Speak msgStr, 1
    MsgBox msgStr, vbExclamation, "检验员补扫码"
    Exit Function
        
End Function


