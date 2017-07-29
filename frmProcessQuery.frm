VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProcessQuery 
   Caption         =   "生产工序查询"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14925
   Icon            =   "frmProcessQuery.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   14925
   Begin VB.Frame Frame3 
      Caption         =   "生产工序信息"
      Height          =   3495
      Left            =   480
      TabIndex        =   16
      Top             =   2880
      Width           =   12255
      Begin MSDataGridLib.DataGrid dataGridProcess 
         Height          =   2295
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblProductCode 
         Caption         =   "lblProductCode"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8640
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "制造编号"
         Height          =   375
         Left            =   7560
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCmpModel 
         Caption         =   "lblCmpModel"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "企业型号"
         Height          =   255
         Left            =   5040
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblModel 
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "型号"
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblProductName 
         Caption         =   "lblProductName"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "产品名称"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "按产品制造编号查询"
      Height          =   1095
      Left            =   360
      TabIndex        =   14
      Top             =   1680
      Width           =   12375
      Begin VB.CommandButton cmdExit2 
         Caption         =   "退出(&E)"
         Height          =   495
         Left            =   9240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCodeQuery 
         Caption         =   "查询(&Q)"
         Height          =   495
         Left            =   7680
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtProductCode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "请输入“产品制造编号”"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "按产品信息查询"
      Height          =   1335
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   13695
      Begin VB.CommandButton cmdExit1 
         Caption         =   "退出(&E)"
         Height          =   495
         Left            =   12120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdProductQuery 
         Caption         =   "查询(&Q)"
         Height          =   495
         Left            =   10560
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cmbProductName 
         Height          =   300
         ItemData        =   "frmProcessQuery.frx":10CA
         Left            =   1080
         List            =   "frmProcessQuery.frx":10CC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   697
         Width           =   2415
      End
      Begin VB.ComboBox cmbModel 
         Height          =   300
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   697
         Width           =   2415
      End
      Begin VB.ComboBox cmbCmpModel 
         Height          =   300
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   697
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "请选择产品信息"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "产品名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "型号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "企业型号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmProcessQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsQuery As ADODB.Recordset
Dim sql As String

Private Sub cmbModel_Click()
    Dim productName As String
    Dim model As String
    Dim rsload As ADODB.Recordset
    
    Set rsload = New ADODB.Recordset
    productName = cmbProductName.Text
    model = cmbModel.Text
  
    
    
  
    
  rsload.Open "select company_model from yt_product where status=1 and name='" & productName & "' and model ='" & model & "'", conn, adOpenStatic, adLockPessimistic
  Me.cmbCmpModel.Clear
  Do Until rsload.EOF '将该产品的所有型号放入下拉框中
    Me.cmbCmpModel.AddItem rsload.Fields("company_model")
    rsload.MoveNext
  Loop
  rsload.Close
  Set rsload = Nothing
End Sub

Private Sub cmbProductName_Click()
    Dim productName As String
    Dim rsload As ADODB.Recordset
    
    Set rsload = New ADODB.Recordset
    productName = cmbProductName.Text
  rsload.Open "select model from yt_product where status=1 and name='" & productName & "'", conn, adOpenStatic, adLockPessimistic
  Me.cmbModel.Clear
  Me.cmbCmpModel.Clear
  Do Until rsload.EOF '将该产品的所有型号放入下拉框中
    Me.cmbModel.AddItem rsload.Fields("model")
    rsload.MoveNext
  Loop
   rsload.Close
  Set rsload = Nothing
End Sub

Private Sub cmdCodeQuery_Click()
    Dim productCode As String
    
   
    
    productCode = Trim(txtProductCode)
   
    
    If productCode = "" Then
        MsgBox "请输入制造编号", , "生产工序查询"
        txtProductCode.SetFocus
        Exit Sub
    End If
        
    sql = "select yt_process.process_no as process_no, yt_process.process_name, yt_employee.emp_no as emp_no, yt_employee.name as emp_name " & _
                    " from yt_process, yt_product, yt_employee, yt_product_code where yt_process.product_id = yt_product.id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product_code.product_id = yt_product.id and yt_product_code.product_code = '" & productCode & "' " & _
                    " order by yt_process.process_no"
                    
    

    
    rsQuery.Close
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    
    
    If rsQuery.recordCount = 0 Then
       setNullQuery
        MsgBox "没有符合条件的数据!"
        Exit Sub
    End If
    
  '  Set dataGridProcess.DataSource = rsQuery
    
     dataGridProcess.ReBind
      setDataGridColumns
     
     
    Dim rsProduct As ADODB.Recordset
     sql = "select yt_product.name as product_name, yt_product.model as model, yt_product.company_model as cmp_model, yt_product_code.product_code as product_code " & _
                    " from  yt_product, yt_product_code where  " & _
                    "  yt_product_code.product_id = yt_product.id and yt_product_code.product_code = '" & productCode & "' "
                    
    Set rsProduct = conn.Execute(sql)
    
        setLbl rsProduct.Fields("product_name"), rsProduct.Fields("model"), rsProduct.Fields("cmp_model"), rsProduct.Fields("product_code")
    rsProduct.Close
    Set rsProduct = Nothing
    
  
    
        
    
    
   
  
  
    
    
     
    
End Sub

Private Sub cmdExit1_Click()
    Unload Me
End Sub

Private Sub cmdExit2_Click()
 Unload Me
End Sub

Private Sub cmdProductQuery_Click()
    Dim productName As String
    Dim model As String
    Dim cmpModel As String
    
    txtProductCode.Text = ""
    productName = Trim(cmbProductName.Text)
    model = Trim(cmbModel.Text)
    cmpModel = Trim(cmbCmpModel.Text)
    
    If productName = "" Then
        MsgBox "请选择产品名称", , "生产工序查询"
        cmbProductName.SetFocus
        Exit Sub
    End If
        
    If model = "" Then
        MsgBox "请选择型号", , "生产工序查询"
        cmbModel.SetFocus
        setLbl "", "", "", ""
        Exit Sub
    End If
    
    If cmpModel = "" Then
        MsgBox "请选择企业型号", , "生产工序查询"
        cmbCmpModel.SetFocus
        Exit Sub
    End If
    
  
    sql = "select yt_process.process_no as process_no, yt_process.process_name, yt_employee.emp_no as emp_no, yt_employee.name as emp_name " & _
                    " from yt_process, yt_product, yt_employee where yt_process.product_id = yt_product.id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product.name = '" & productName & "' and yt_product.model = '" & model & "' and yt_product.company_model = '" & cmpModel & "' " & _
                    " order by yt_process.process_no"
                    
    

    
    rsQuery.Close
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    
    
    If rsQuery.recordCount = 0 Then
        MsgBox "没有符合条件的数据!"
        setNullQuery
        Exit Sub
    End If
    
  '  Set dataGridProcess.DataSource = rsQuery
    
     dataGridProcess.ReBind

  
    setDataGridColumns
    setLbl productName, model, cmpModel, ""
    
    
    
     
    
    
   
End Sub

Private Sub Form_Load()
     Dim rsload As ADODB.Recordset
   '  MsgBox LoginSucceeded
   
   If conn Is Nothing Then
     '连接数据库
        If dbConnOpen() = False Then
            MsgBox "连接数据库失败!"
            End
        End If
    End If
    Set rsload = New ADODB.Recordset
     rsload.Open "select distinct name from yt_product where status=1", conn, adOpenStatic, adLockPessimistic
     Me.cmbProductName.Clear
    Do Until rsload.EOF '将所有的的产品名称放入下拉框中
         Me.cmbProductName.AddItem rsload.Fields("name")
        rsload.MoveNext
    Loop
     rsload.Close
    Set rsload = Nothing
    
    Set rsQuery = New ADODB.Recordset
     sql = "select yt_process.process_no as process_no, yt_process.process_name, yt_employee.emp_no as emp_no, yt_employee.name as emp_name " & _
                    " from yt_process, yt_product, yt_employee where yt_process.product_id = yt_product.id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product.name = ''"
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    Set dataGridProcess.DataSource = rsQuery
    setDataGridColumns
    setLbl "", "", "", ""
    Me.Width = 16000
    Me.Height = 7500
    
    

End Sub

Private Sub setDataGridColumns()
    dataGridProcess.Columns("process_no").Caption = "工序号"
    dataGridProcess.Columns(0).Width = 1000
    dataGridProcess.Columns("process_name").Caption = "工序名"
    dataGridProcess.Columns(1).Width = 1800
    dataGridProcess.Columns("emp_no").Caption = "操作工编号"
    dataGridProcess.Columns(2).Width = 1500
    dataGridProcess.Columns("emp_name").Caption = "操作工姓名"
    dataGridProcess.Columns(3).Width = 1500
    
   
End Sub

Private Sub setLbl(ByVal mProductName As String, ByVal mModel As String, ByVal mCmpModel As String, ByVal mProductCode As String)
    lblProductName = mProductName
    lblModel = mModel
    lblCmpModel = mCmpModel
    lblProductCode = mProductCode
End Sub

Private Sub setNullQuery()
    sql = "select yt_process.process_no as process_no, yt_process.process_name, yt_employee.emp_no as emp_no, yt_employee.name as emp_name " & _
                    " from yt_process, yt_product, yt_employee where yt_process.product_id = yt_product.id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product.name = ''"
    rsQuery.Close
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    dataGridProcess.ReBind
    setDataGridColumns
    setLbl "", "", "", ""
End Sub

Private Sub Form_Resize()
      On Error Resume Next
     Frame1.Left = Me.Width / 2 - Frame1.Width / 2
     Frame2.Left = Frame1.Left
     Frame3.Left = Frame1.Left
     Frame2.Width = Frame1.Width
     Frame3.Width = Frame1.Width
     Frame3.Height = Me.Height - Frame1.Height - Frame2.Height - 1200
     dataGridProcess.Height = Frame3.Height - 1200
End Sub
