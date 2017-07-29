VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5610
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Menu BarCode 
      Caption         =   "生成条码"
      Begin VB.Menu BarCodeProduct 
         Caption         =   "产品条码"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rs As ADODB.Recordset
Public fld As ADODB.Field




 

Private Sub BarCodeProduct_Click()
Load Form2

End Sub

Private Sub Command1_Click()
Call connOpen
End Sub

Public Sub connOpen()
  Set conn = New ADODB.Connection
  conn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
    "SERVER=" & DB_HOST & ";" & _
    "DATABASE=" & DB_DATABASE & ";" & _
    "UID=" & DB_USER & ";" & _
    "PWD=" & DB_PASS & ";" & _
    "OPTION=3;stmt=SET NAMES GB2312"
   
  conn.Open
 
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient   '游标位置（数据集存在服务器内存还是客户端内存）
End Sub

Public Sub connClose()
  rs.Close
  conn.Close
End Sub

