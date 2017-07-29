VERSION 5.00
Begin VB.Form frmCheckerAddScan 
   Caption         =   "检验员生产补扫码"
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
      Caption         =   "检验员生产补扫码"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   2640
      Width           =   10455
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
      Left            =   3000
      TabIndex        =   0
      Top             =   5280
      Width           =   9975
   End
   Begin VB.Menu mnuExit 
      Caption         =   "退出检验员生产补扫码"
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
      If SysPass() = False Then  ''权限
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
        strMsg = strBarCode & " 扫码错误，请重新扫码!"
        strSpeak = "扫码错误，请重新扫码!"
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
        strMsg = strBarCode & "扫码枪不正确，请用检验员的扫码枪重新扫码!"
        strSpeak = "扫码枪不正确，请用检验员的扫码枪重新扫码!"
        objVoice.Speak strSpeak, 1
      '  MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
    
    pos = InStr(1, strBarCode, " ")
    If pos < 1 Then
        strMsg = strBarCode & " 扫码错误，请重新扫码!"
        strSpeak = "扫码错误，请重新扫码!"
        objVoice.Speak strSpeak, 1
       ' MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
    
   
    empNo = Val(Left(strBarCode, pos - 1))
    barCode = Right(strBarCode, Len(strBarCode) - (xPos + Len(prefix)) + 1)
     
    If Len(barCode) > 3 Then  '扫生产条码
     
        If fProductCode <> "" And fProductCode <> barCode And fEmpNo <> empNo And DateDiff("n", fTime, Now) < 1 Then
            strMsg = "前一个检验所扫产品条码还未处理完毕,请稍候"
            objVoice.Speak strMsg, 1
          '  MsgBox strMsg, vbExclamation, "检验员生产补扫码"
            Exit Sub
        End If
            
        
        fProductCode = barCode
        fTime = Now
        fEmpNo = empNo
        fProcessNo = ""
        strMsg = "接下来请扫工序条码!"
        objVoice.Speak strMsg, 1
      '  MsgBox strMsg
        Exit Sub
    End If
     '扫工序条码
    If Trim(fProductCode) = "" Then
        strMsg = "扫码错误，请先扫产品条码!"
        objVoice.Speak strMsg, 1
       ' MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
        
    If empNo <> fEmpNo Then
        strMsg = "本次扫工序条码的扫码抢与前次扫产品条码的扫码抢不相同!"
        objVoice.Speak strMsg, 1
   '     MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
            
    If DateDiff("n", fTime, Now) > 1 Then
        strMsg = "两次扫码间隔超时，请重新扫产品条码!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak strMsg, 1
      '  MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
        
        
    fProcessNo = barCode
        
    If Trim(fProductCode) = "" Or Trim(fProcessNo) = "" Then
        strMsg = "扫码处理错误，请重新扫产品条码，再扫工序条码!"
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak strMsg, 1
        MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
    
     '数据库操作
    sql = "select emp_no, name as emp_name from yt_employee where emp_no = '" & fEmpNo & "' and status = 1"
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        strMsg = "编号为 " & fEmpNo & " 的有效员工不存在,扫码处理失败!"
        rsResult.Close
        Set rsResult = Nothing
        fProductCode = ""
        fEmpNo = ""
        fProcessNo = ""
        objVoice.Speak strMsg, 1
      '  MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
    End If
            
    empNo = rsResult.Fields("emp_no")
    empName = rsResult.Fields("emp_name")
            
           
    sql = "update yt_produce_scan set operator_no='" & empNo & "', operator_name='" & empName & "',scan_status=2, scan_time='" & Now & "', update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & fProductCode & "' and  yt_produce_scan.process_no = " & fProcessNo & ""
                        
    Set rsResult = conn.Execute(sql, recordsAffected)
    If recordsAffected = 0 Then
        strMsg = "数据处理失败，请重新扫产品条码，再扫工序条码!"
        fProductCode = ""
        fProcessNo = ""
        fEmpNo = ""
        Set rsResult = Nothing
        objVoice.Speak strMsg, 1
   '     MsgBox strMsg, vbExclamation, "检验员生产补扫码"
        Exit Sub
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
        
    strMsg = "生产补扫码处理成功!"
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










Private Sub 退出检验员生产补扫码_Click()

End Sub
