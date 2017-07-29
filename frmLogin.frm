VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "樱田农机生产线扫码系统--登录"
   ClientHeight    =   2760
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4770
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1630.699
   ScaleMode       =   0  'User
   ScaleWidth      =   4478.771
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "系统登录"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox CB_AdminName 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H80000018&
         DataField       =   "password"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "用户名(&U):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1530
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "密 码(&P):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   420
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   1020
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim chkpassword As String '检查密码是否正确
Public LoginSucceeded As Boolean '全局变量表示登陆是否成功

Private Sub CB_AdminName_Click()
   Dim rsload As ADODB.Recordset
  Set rsload = New ADODB.Recordset
 ' Call check_condatabase
  If rsload.State = 1 Then rsload.Close
  rsload.Open "select password from yt_user where role='PRODUCE' and user_name = '" & Me.CB_AdminName.Text & "'", conn, adOpenStatic, adLockPessimistic
  Me.txtPassword.SetFocus
  chkpassword = rsload.Fields("password") '根据下拉框中的用户名得到该用户的正确密码信息
  Set rsload = Nothing
End Sub

Private Sub cmdCancel_Click()
  LoginSucceeded = False
  Me.Hide
  Unload Me
  End
End Sub

Private Sub cmdOK_Click()



 '  Dim MD5 As Object
   Dim txtPasswordMD5 As String
  
    
    
  'Call check_condatabase
  If Me.CB_AdminName.Text = "" Then
    MsgBox "别忘了用户名:)", vbOKOnly, "登录失败"
    CB_AdminName.SetFocus
    Exit Sub
  End If
  If Me.txtPassword.Text <> "" Then
     '检查正确的密码
   ' Set MD5 = New CMD5     'CMd5是新增类模块的名称
   ' txtPasswordMD5 = MD5.Md5_String_Calc(txtPassword.Text)
     txtPasswordMD5 = MD5(txtPassword.Text, 32)
   ' MsgBox txtPasswordMD5
    'MsgBox chkpassword
    If UCase(txtPasswordMD5) = UCase(chkpassword) Then
      LoginSucceeded = True
     
      loginUser = Me.CB_AdminName.Text '保存全局的登陆帐户名
     
      Unload Me
      frmMain.Show
    Else
      MsgBox "无效的密码，请重试!", vbOKOnly + vbExclamation, "登录失败"
      Me.txtPassword.SetFocus
      txtPassword.SelStart = 0
      txtPassword.SelLength = Len(txtPassword)
       
      
    '  SendKeys "{Home}+{End}}"
    End If
  Else
    MsgBox "别忘了请输入密码:)", vbOKOnly, "登录失败"
    Me.txtPassword.SetFocus
'    SendKeys "{Home}+{End}"
  End If
  Call dbDisConnect
  
End Sub

Private Sub Form_Load()

 '***********************************************
      If SysPass() = False Then  ''权限
        End
        Exit Sub
       End If
'***********************************************
  
  
  Dim rsload As ADODB.Recordset
  
  Me.txtPassword.Text = ""
   '连接数据库
  If dbConnOpen() = False Then
    MsgBox "连接数据库失败!"
    End
 End If
  
  Set rsload = New ADODB.Recordset
  rsload.Open "select * from yt_user where role='PRODUCE'", conn, adOpenStatic, adLockPessimistic
  Me.CB_AdminName.Clear
  Do Until rsload.EOF '将所有的帐户名称加入用户名下拉框中
    Me.CB_AdminName.AddItem rsload.Fields("user_name")
    rsload.MoveNext
  Loop
  Set rsload = Nothing
  
End Sub

