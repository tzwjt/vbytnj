Attribute VB_Name = "Module1"

'*************************************************************************
'**描    述：数据库相关操作

'*************************************************************************

Option Explicit

Private Const DB_HOST = "www.e-rabits.com" '数据库服务器主机
Private Const DB_USER = "root"              '数据库登录用户
Private Const DB_PASS = "Erabitsroot"       '数据库登录密码
Private Const DB_DATABASE = "ytdb"          '数据库名

Public conn As ADODB.Connection

Public Sub connOpen()
  Set conn = New ADODB.Connection
  conn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
    "SERVER=" & DB_HOST & ";" & _
    "DATABASE=" & DB_DATABASE & ";" & _
    "UID=" & DB_USER & ";" & _
    "PWD=" & DB_PASS & ";" & _
    "OPTION=3;stmt=SET NAMES GB2312"
   
  conn.Open
 
End Sub

Sub main()
 connOpen
 'Load frmMain
 frmMain.Show
 
End Sub


