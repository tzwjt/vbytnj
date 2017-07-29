VERSION 5.00
Begin VB.Form frmUpdatePwd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改密码"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5040
   Icon            =   "frmUpdatePwd.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5040
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtReNewPwd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtNewPwd 
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtOldPwd 
         Height          =   375
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "请再次输入新密码"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "请输入新密码"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "请输入现密码"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmUpdatePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim oldPwd, newPwd, reNewPwd As String
    oldPwd = Trim(txtOldPwd.Text)
    newPwd = Trim(txtNewPwd.Text)
    reNewPwd = Trim(txtReNewPwd.Text)
    If checkPwd(oldPwd, newPwd, reNewPwd) = False Then
        Exit Sub
    End If
    
    If updatePwd(oldPwd, newPwd) = False Then
        Exit Sub
    End If
    
    
    
    
End Sub

Private Sub Form_Load()
 If conn Is Nothing Then
     '连接数据库
        If dbConnOpen() = False Then
            MsgBox "连接数据库失败!"
            End
        End If
    End If
  '  MsgBox loginUser

End Sub

Private Function checkPwd(ByVal oldPwd As String, ByVal newPwd As String, ByVal reNewPwd As String) As Boolean
    
    
    If oldPwd = "" Then
        MsgBox "请输入现密码!", , "修改密码"
        txtOldPwd.SetFocus
        checkPwd = False
        Exit Function
    End If
    
     If newPwd = "" Then
        MsgBox "请输入新密码!", , "修改密码"
        txtNewPwd.SetFocus
        checkPwd = False
        Exit Function
    End If
    
    If reNewPwd = "" Then
        MsgBox "请再次输入新密码!", , "修改密码"
        txtReNewPwd.SetFocus
        checkPwd = False
        Exit Function
    End If
    
    If Len(newPwd) < 6 Then
       MsgBox "新密码的长度不能少于6位!", , "修改密码"
       txtNewPwd.SetFocus
       txtNewPwd.SelStart = 0
       txtNewPwd.SelLength = Len(reNewPwd)
       checkPwd = False
        Exit Function
    End If
    
    If reNewPwd <> newPwd Then
        MsgBox "两次输入的新密码不一致!", , "修改密码"
        txtReNewPwd.SetFocus
        txtReNewPwd.SelStart = 0
        txtReNewPwd.SelLength = Len(reNewPwd)
        checkPwd = False
        Exit Function
    End If
    
    checkPwd = True
    
    
End Function

Private Function updatePwd(ByVal oldPwd As String, ByVal newPwd As String) As Boolean
    Dim sql As String
    Dim userRs As ADODB.Recordset
    Dim oldPwdMD5 As String
    Dim newPwdMD5 As String
    
   ' On Error GoTo errhandler
    
    oldPwdMD5 = MD5(oldPwd, 32)
    newPwdMD5 = MD5(newPwd, 32)
    
    sql = "select password from yt_user where user_name = '" & loginUser & "'"
     Set userRs = conn.Execute(sql)
     
     If UCase(userRs.Fields("password")) <> UCase(oldPwdMD5) Then
        MsgBox "输入的现密码不正确,请重新输入!", , "修改密码"
        txtOldPwd.SetFocus
        txtOldPwd.SelStart = 0
        txtOldPwd.SelLength = Len(oldPwd)
        updatePwd = False
        userRs.Close
        Set userRs = Nothing
        updatePwd = False
        Exit Function
    End If
     sql = "update yt_user set password = '" & newPwdMD5 & "'where yt_user.user_name = '" & loginUser & "' "
    Set produceRs = conn.Execute(sql)
     MsgBox "密码修改成功!", , "修改密码"
     Set userRs = Nothing
    updatePwd = True
    Exit Function
errhandler:
    MsgBox "密码修改失败", vbOKOnly, "修改密码"
End Function

