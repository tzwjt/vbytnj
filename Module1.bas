Attribute VB_Name = "Module1"

'*************************************************************************
'**��    �������ݿ���ز���

'*************************************************************************

Option Explicit

Private Const DB_HOST = "www.e-rabits.com" '���ݿ����������
Private Const DB_USER = "root"              '���ݿ��¼�û�
Private Const DB_PASS = "Erabitsroot"       '���ݿ��¼����
Private Const DB_DATABASE = "ytdb"          '���ݿ���

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


