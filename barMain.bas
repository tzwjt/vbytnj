Attribute VB_Name = "basMain"

'*************************************************************************
'**��    �������ݿ���ز���

'*************************************************************************

Option Explicit

Public loginUser As String '�����½���û���

Private Const DB_HOST = "www.e-rabits.com" '���ݿ����������
Private Const DB_USER = "root"              '���ݿ��¼�û�
Private Const DB_PASS = "Erabitsroot"       '���ݿ��¼����
Private Const DB_DATABASE = "ytdb"          '���ݿ���

 Private IsConnect As Boolean '������ݿ��Ƿ�����

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

'�Ͽ������ݿ������
Public Sub dbDisConnect()
 Dim rc As Long
  If IsConnect = False Then
   Exit Sub
 End If
 '�ر�����
 conn.Close
 '�ͷ�cnn
 Set conn = Nothing
 IsConnect = False
End Sub

Public Function isDbConnect() As Boolean
    isDbConnect = IsConnect
End Function

'ִ�����ݿ��ѯ���
Public Function dbQueryExt(ByVal TmpSQLstmt As String) As ADODB.Recordset
   
  Dim rst As New ADODB.Recordset '����Rescordset����rst
 
  DB_Connect '�������ݿ�
 
  Set rst.ActiveConnection = conn '����rst��ActiveConnection����,ָ��������ص����ݿ������
 
  rst.CursorType = adOpenDynamic '�����α�����
 
  rst.LockType = adLockOptimistic '������������
 
  rst.Open TmpSQLstmt '�򿪼�¼��
 
  Set QueryExt = rst '���ؼ�¼��
 
  End Function

Sub main()
 'connOpen
 'Load frmMain
 'frmMain.Show
 frmLogin.Show
 
End Sub


