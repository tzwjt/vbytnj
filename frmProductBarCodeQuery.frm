VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProductBarCodeQuery 
   Caption         =   "产品条码查询"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10230
   Icon            =   "frmProductBarCodeQuery.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   10230
   Begin VB.Frame Frame3 
      Caption         =   "条码区"
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   4695
      Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3975
         Style           =   2
         SubStyle        =   0
         Validation      =   0
         LineWeight      =   3
         Direction       =   0
         ShowData        =   1
         Value           =   ""
         ForeColor       =   0
         BackColor       =   16777215
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   1335
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "产品条码信息"
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   9615
      Begin MSDataGridLib.DataGrid dataGridProduct 
         Height          =   4455
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   14
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
            Size            =   9
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
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   735
      Left            =   7320
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   735
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "请输入查询条件"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询(&Q)"
         Height          =   615
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtProductCode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "制造编号（条码号）"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   8280
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProductBarCodeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsQuery As ADODB.Recordset
Dim sql As String


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If rsQuery.recordCount > 0 Then
        frmProductBarCodePrint.lblProductName = dataGridProduct.Columns(1).value
        frmProductBarCodePrint.lblModel = dataGridProduct.Columns(2).value
        frmProductBarCodePrint.lblCmpModel = dataGridProduct.Columns(3).value
        frmProductBarCodePrint.BarCodeCtrl1.Style = BarCodeCtrl1.Style
        frmProductBarCodePrint.BarCodeCtrl1.value = BarCodeCtrl1.value
        setPrint dlgCommonDialog, frmProductBarCodePrint
    Else
        frmProductBarCodePrint.lblProductName = ""
        frmProductBarCodePrint.lblModel = ""
        frmProductBarCodePrint.lblCmpModel = ""
        frmProductBarCodePrint.BarCodeCtrl1.value = ""
    End If
        
End Sub

Private Sub cmdQuery_Click()
    Dim productCode As String
   
    
    productCode = Trim(txtProductCode.Text)

    
    If productCode = "" Then
        MsgBox "请输入查询条件!"
        txtProductCode.SetFocus
        Exit Sub
    End If
    
  
    sql = "select yt_product_code.product_code as product_code, yt_product.name as productName, yt_product.model as model, yt_product.company_model as company_model, yt_product.status as product_status " & _
                            "  from yt_product, yt_product_code where yt_product.id = yt_product_code.product_id and  yt_product_code.product_code like '%" & productCode & "%' order by  yt_product_code.id desc "
                            
    
    rsQuery.Close
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    
    
    If rsQuery.recordCount = 0 Then
        MsgBox "没有符合条件的数据!"
       ' rsQuery.Close
      '  Set rsQuery = Nothing
        nullQuery
        Exit Sub
    End If
    
   dataGridProduct.ReBind
  
    setDataGridColumns
    setBarCode dataGridProduct.Columns(0).value
    setProduct dataGridProduct.Columns(1).value, dataGridProduct.Columns(2).value
    
    
     
    
    
   
   
   
   
   
  ' DataGrid1.dataM
     
     
    
        
    
    
 
    
    
    
    
    
  
    
    
  
    
  
End Sub






Private Sub dataGridProduct_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     If rsQuery.recordCount > 0 Then
        setBarCode dataGridProduct.Columns(0).value
        setProduct dataGridProduct.Columns(1).value, dataGridProduct.Columns(2).value
    End If
End Sub

Private Sub Form_Load()
 Dim rsload As ADODB.Recordset
   '  MsgBox LoginSucceeded
   lblProductName.Caption = ""
   lblModel.Caption = ""
   
   If conn Is Nothing Then
     '连接数据库
        If dbConnOpen() = False Then
            MsgBox "连接数据库失败!"
            End
        End If
    End If
    
   
    
   
    Set rsQuery = New ADODB.Recordset
    
    
    
    sql = "select yt_product_code.product_code as product_Code, yt_product.name as productName, yt_product.model as model, yt_product.company_model as company_model, yt_product.status as product_status " & _
                            " from yt_product, yt_product_code where  yt_product.id = yt_product_code.product_id and yt_product.status = 1 order by  yt_product_code.id desc"
                            
    
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    Set dataGridProduct.DataSource = rsQuery
    setDataGridColumns
    
    BarCodeCtrl1.Style = 7
    BarCodeCtrl1.SubStyle = 0
    BarCodeCtrl1.Validation = 0
    If rsQuery.recordCount > 0 Then
        setBarCode dataGridProduct.Columns(0).value
        setProduct dataGridProduct.Columns(1).value, dataGridProduct.Columns(2).value
    End If
   Me.Width = 11000
    Me.Height = 8500
End Sub

Private Sub setBarCode(code As String)
    BarCodeCtrl1.value = code
End Sub

Private Sub setProduct(productName As String, productModel As String)
    lblProductName.Caption = productName
    lblModel.Caption = productModel
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Frame1.Left = Me.Width / 2 - Frame1.Width / 2
     Frame2.Left = Frame1.Left
     Frame3.Left = Frame1.Left
     cmdPrint.Left = Frame3.Left + 5150
     cmdExit.Left = cmdPrint.Left + 1950
     Frame2.Width = Frame1.Width
     Frame2.Height = Me.Height - Frame1.Height - Frame3.Height - 1000
     dataGridProduct.Height = Frame2.Height - 800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsQuery.Close
    Set rsQuery = Nothing
     Unload frmProductBarCodePrint
End Sub

Private Sub setDataGridColumns()
    dataGridProduct.RowHeight = 250
    dataGridProduct.EditActive = False
    
    dataGridProduct.Columns("product_code").Caption = "制造编号"
    dataGridProduct.Columns(0).Width = 1500
    dataGridProduct.Columns("productName").Caption = "产品名称"
    dataGridProduct.Columns(1).Width = 2500
    dataGridProduct.Columns("model").Caption = "型号"
    dataGridProduct.Columns(2).Width = 1500
    dataGridProduct.Columns("company_model").Caption = "企业型号"
    dataGridProduct.Columns(3).Width = 1500
    dataGridProduct.Columns("product_status").Caption = "产品状态"
    dataGridProduct.Columns(4).Width = 1000
   
End Sub

Private Sub nullQuery()
    sql = "select yt_product_code.product_code as product_Code, yt_product.name as productName, yt_product.model as model, yt_product.company_model as company_model, yt_product.status as product_status " & _
                            " from yt_product, yt_product_code where  yt_product.id = yt_product_code.product_id and yt_product_code.product_code = '' "
     rsQuery.Close
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    
    dataGridProduct.ReBind
  
    setDataGridColumns
    setBarCode ""
    setProduct "", ""
                            
   
End Sub
