Attribute VB_Name = "basMain"

'*************************************************************************
'**描    述：数据库相关操作

'*************************************************************************

Option Explicit

Public loginUser As String '保存登陆的用户名

Private Const DB_HOST = "www.e-rabits.com" '数据库服务器主机
Private Const DB_USER = "root"              '数据库登录用户
Private Const DB_PASS = "Erabitsroot"       '数据库登录密码
Private Const DB_DATABASE = "ytdb"          '数据库名

 Private IsConnect As Boolean '标记数据库是否连接

Public conn As ADODB.Connection

Public Function dbConnOpen() As Boolean
 On Error Resume Next
  Set conn = New ADODB.Connection
  conn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
    "SERVER=" & DB_HOST & ";" & _
    "DATABASE=" & DB_DATABASE & ";" & _
    "UID=" & DB_USER & ";" & _
    "PWD=" & DB_PASS & ";" & _
    "OPTION=3;stmt=SET NAMES GB2312"
   
  conn.Open
  
  If conn.State <> adStateOpen Then
    dbConnOpen = False
  Else
    dbConnOpen = True
  End If
 
End Function

'断开与数据库的连接
Public Sub dbDisConnect()
 Dim rc As Long
  If IsConnect = False Then
   Exit Sub
 End If
 '关闭连接
 conn.Close
 '释放cnn
 Set conn = Nothing
 IsConnect = False
End Sub

Public Function isDbConnect() As Boolean
    isDbConnect = IsConnect
End Function

'执行数据库查询语句
Public Function dbQueryExt(ByVal TmpSQLstmt As String) As ADODB.Recordset
   
  Dim rst As New ADODB.Recordset '创建Rescordset对象rst
 
  DB_Connect '连接数据库
 
  Set rst.ActiveConnection = conn '设置rst的ActiveConnection属性,指定与其相关的数据库的连接
 
  rst.CursorType = adOpenDynamic '设置游标类型
 
  rst.LockType = adLockOptimistic '设置锁定类型
 
  rst.Open TmpSQLstmt '打开记录集
 
  Set QueryExt = rst '返回记录集
 
  End Function

Sub main()
 'connOpen
 'Load frmMain
 'frmMain.Show
 frmLogin.Show
 
End Sub


