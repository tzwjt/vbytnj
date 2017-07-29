VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAddScan 
   Caption         =   "生产补码处理"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13140
   Icon            =   "frmAddScan.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   13140
   Begin VB.Frame Frame2 
      Caption         =   "工序列表"
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   12855
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridProduct 
         Height          =   4695
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   8281
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   135
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   238
         _Version        =   393216
      End
      Begin VB.Label Label7 
         Caption         =   "对于未扫码的工序，请双击所在行进入扫码界面"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12855
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
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "未扫码数"
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
         Left            =   10500
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Left            =   8400
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   6480
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "制造编号"
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
         Left            =   3480
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "请扫码输入“产品制造编号(条码)”"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblNoScan 
         Alignment       =   2  'Center
         Caption         =   "lblNoScan"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10440
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblStatus 
         Caption         =   "lblStatus"
         Height          =   375
         Left            =   8520
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblCompanyModel 
         Alignment       =   2  'Center
         Caption         =   "lblCompanyModel"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblModel 
         Alignment       =   2  'Center
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblProductCode 
         Alignment       =   2  'Center
         Caption         =   "lblProductCode"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2940
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         Caption         =   "lblProductName"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAddScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim rsQuery As ADODB.Recordset


Private Sub Form_Load()
     If conn Is Nothing Then
     '连接数据库
        If dbConnOpen() = False Then
            MsgBox "连接数据库失败!"
            End
        End If
    End If
   
     Set rsQuery = New ADODB.Recordset
     rsQuery.CursorLocation = adUseClient
     setLblNull
     setGrid
     'txtProductCode.SetFocus
     Me.Width = 13400
     Me.Height = 9000
     
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Left = Me.Width / 2 - Frame1.Width / 2
     Frame2.Left = Frame1.Left
     Frame2.Width = Frame1.Width
     Frame2.Height = Me.Height - Frame1.Height - 1000
     gridProduct.Height = Frame2.Height - 800
     
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
     On Error Resume Next
   ' rsQuery.Close
    Set rsQuery = Nothing
End Sub



Public Sub getProduceScan(ByVal productCode As String)
      On Error Resume Next
    Dim sql As String
    Dim produceRs As ADODB.Recordset
   
    Dim recordCount As Integer
    Dim status As Integer
    
    sql = "select  process_no, process_name, operator_no, operator_name, scan_status, id  from yt_produce_scan where product_code = '" & productCode & "' order by process_no "
    
 '  MsgBox sql
    rsQuery.Close
   '  rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    
    
    If rsQuery.recordCount < 1 Then
        MsgBox "没有对应的数据!"
        rsQuery.Close
        setLblNull
        setGridNull
        txtProductCode.SetFocus
        txtProductCode.SelStart = 0
        txtProductCode.SelLength = Len(txtProductCode.Text)
    '    Set rsQuery = Nothing
        Exit Sub
    End If
    
   ' Set dataGridProduct.DataSource = rsQuery
   Set gridProduct.DataSource = rsQuery
   
   
   
   
    
    
   
    sql = "select yt_product.name as product_name, yt_product_code.product_code as product_code, yt_product.model as model ," & _
            " yt_product.company_model as company_model,yt_product.status as status " & _
            " from yt_product, yt_product_code " & _
            " where yt_product.id = yt_product_code.product_id and yt_product_code.product_code = '" & productCode & "'"
    '  sql = "select *  from yt_produce_scan where product_code = '" & productCode & "' and scan_status = 0"
    Set produceRs = conn.Execute(sql)
    
    
        
 
    lblProductName = produceRs.Fields("product_name")
     lblProductCode = produceRs.Fields("product_code")
     lblModel = produceRs.Fields("model")
      lblCompanyModel = produceRs.Fields("company_model")
     status = produceRs.Fields("status")
     Select Case status
        Case 0
            lblStatus = "注销"
        Case 1
            lblStatus = "正常"
        Case 2
            lblStatus = "已停产"
    End Select
    
    produceRs.Close
    Set produceRs = Nothing
    
     sql = "select count(*) as recordCount from yt_produce_scan where product_code = '" & productCode & "' and scan_status = 0"
    Set produceRs = conn.Execute(sql)
    
    recordCount = produceRs.Fields("recordCount")
       lblNoScan.Caption = recordCount
    produceRs.Close
    Set produceRs = Nothing
    
  
 '
   '  Set DataGrid1.DataMember = rsQuery
  
   ' setDataGridColumns
   ' setBarCode dataGridEmployee.Columns(0).value
   ' setEmp dataGridEmployee.Columns(0).value, DataGrid1.Columns(1).value
    
    
    
    setDataGridProduct
   
   
   
   
   
  ' DataGrid1.dataM
     
     
    
        
    
    
    txtProductCode.Text = ""
    txtProductCode.SetFocus
    
    
    
    
    
End Sub

Private Sub gridProduct_DblClick()
    gridProduct.Col = 5
    If gridProduct.Text = "未扫码" Then
        frmEmpReScan.Hide
        frmEmployeeScan.Show
        frmEmployeeScan.getProductInfo Trim(lblProductCode.Caption), gridProduct.TextMatrix(gridProduct.Row, 1), gridProduct.TextMatrix(gridProduct.Row, 2), gridProduct.TextMatrix(gridProduct.Row, 6)
         frmEmployeeScan.Hide
         frmEmployeeScan.Show vbModal
         
        
    ElseIf gridProduct.Text = "已扫码" Then
        frmEmployeeScan.Hide
        frmEmpReScan.Show
        frmEmpReScan.getProductInfo Trim(lblProductCode.Caption), gridProduct.TextMatrix(gridProduct.Row, 1), gridProduct.TextMatrix(gridProduct.Row, 2), gridProduct.TextMatrix(gridProduct.Row, 6), gridProduct.TextMatrix(gridProduct.Row, 3), gridProduct.TextMatrix(gridProduct.Row, 4)
         frmEmpReScan.Hide
         frmEmpReScan.Show vbModal
    Else
        frmEmployeeScan.Hide
        frmEmpReScan.Hide
    End If
End Sub

Private Sub txtProductCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtProductCode.Text) = "" Then
            MsgBox "请扫码输入产品制造编号", vbOKOnly
            Exit Sub
        End If
         getProduceScan Trim(getValidProductCode(txtProductCode.Text))
    End If
End Sub

Private Sub setLblNull()
     lblProductName = ""
     lblProductCode = ""
     lblModel = ""
    lblCompanyModel = ""
   
    lblStatus = ""
    lblNoScan = ""
      
    
End Sub

Private Sub setDataGridProduct()
   ' For I = 1 To dataGridProduct
  ' dataGridProduct.RowBookmark
  Dim i As Integer
  Dim a
  
  For i = 1 To rsQuery.recordCount
   If gridProduct.TextMatrix(i, 5) = "0" Then
     gridProduct.Row = i
     gridProduct.Col = 5
     gridProduct.Text = "未扫码"
    gridProduct.CellFontBold = True
    gridProduct.CellForeColor = vbRed
  Else
      gridProduct.Row = i
     gridProduct.Col = 5
     gridProduct.Text = "已扫码"
    gridProduct.CellFontBold = True
    gridProduct.CellForeColor = vbGreen
End If
    

   setGrid
   ' gridProduct.Text
    
 
 
  
  
  
 'If dataGridProduct.Columns("scan_status").CellValue(i) = 1 Then
  '  MsgBox i
    'datagridproduct.
'End If
Next
    
    
    
    
End Sub

Private Sub setGrid()
    Dim i As Integer
    
    With gridProduct
   
    .Cols = 7
  '  .FixedCols = 6
    

    
    For i = 1 To .Cols - 1
        .Row = 0
        .Col = i
        .CellAlignment = 4
    Next
        
        
        .TextMatrix(0, 1) = "工序号"
       
        
        .TextMatrix(0, 2) = "工序名"
        .TextMatrix(0, 3) = "操作工编号"
         .TextMatrix(0, 4) = "操作工姓名"
        .TextMatrix(0, 5) = "扫码状态"
        
        .ColWidth(0) = 200
        .ColWidth(1) = 800
        .ColWidth(2) = 3000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1000
        .ColWidth(6) = 1
       
        
        
    End With
End Sub

Private Sub setGridNull()
With gridProduct
    .Clear
    .Rows = 2
End With
setGrid
    
    
End Sub

Private Function getValidProductCode(ByVal productCode) As String
    Dim pos As Integer
    Dim xProductCode As String
    pos = InStr(1, productCode, " ")
    If pos > 0 Then
        xProductCode = Mid(productCode, pos)
    Else
        xProductCode = productCode
    End If
    getValidProductCode = Trim(xProductCode)
End Function
