VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProductBarCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "产品录码"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10290
   Icon            =   "frmProductBarCode.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   10290
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9480
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "录入(&I)"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ComboBox cmbCmpModel 
      Height          =   300
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cmbModel 
      Height          =   300
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cmbProductName 
      Height          =   300
      ItemData        =   "frmProductBarCode.frx":10CA
      Left            =   1440
      List            =   "frmProductBarCode.frx":10CC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame2 
         Caption         =   "条码区"
         Height          =   1335
         Left            =   4920
         TabIndex        =   10
         Top             =   1560
         Width           =   4215
         Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
            Height          =   975
            Left            =   120
            TabIndex        =   11
            Top             =   240
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   3
         Top             =   2085
         Width           =   3015
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
         Left            =   2160
         TabIndex        =   14
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
         Left            =   4920
         TabIndex        =   13
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
         Left            =   7440
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "TZYT"
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
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "请输入“制造编号”"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "请选择产品信息"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmProductBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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



Private Sub Command2_Click()
 
  
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If saveProductCode() = False Then
        Exit Sub
    End If
    
    If MsgBox("数据录入成功，接下来你需要打印此条码吗？", vbYesNo, "条码打印确认") = vbYes Then
         frmProductBarCodePrint.lblProductName = cmbProductName.Text
        frmProductBarCodePrint.lblModel = cmbModel.Text
        frmProductBarCodePrint.lblCmpModel = cmbCmpModel.Text
        frmProductBarCodePrint.BarCodeCtrl1.Style = BarCodeCtrl1.Style
        frmProductBarCodePrint.BarCodeCtrl1.value = BarCodeCtrl1.value
        setPrint dlgCommonDialog, frmProductBarCodePrint
    End If
     
  
End Sub



Private Sub Form_Load()
'***********************************************
      If SysPass() = False Then  ''权限
        End
        Exit Sub
       End If
'***********************************************
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
    
    BarCodeCtrl1.Style = 7
    BarCodeCtrl1.SubStyle = 0
    BarCodeCtrl1.Validation = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmProductBarCodePrint
End Sub

Private Sub txtProductCode_Change()
    BarCodeCtrl1.value = txtProductCode & ""
    

End Sub

Private Function saveProductCode() As Boolean
    Dim productName As String
    Dim model As String
    Dim cmpModel As String
    Dim productCode As String
    Dim sql As String
    Dim productCodeCountRs As ADODB.Recordset
    Dim productRs As ADODB.Recordset
    
    
    productName = cmbProductName.Text
    model = cmbModel.Text
    cmpModel = cmbCmpModel.Text
    productCode = txtProductCode.Text
    
     
    If productName = "" Then
        MsgBox "请选择产品名称!"
        cmbProductName.SetFocus
        saveProductCode = False
        Exit Function
    End If
    
    If model = "" Then
        MsgBox "请选择型号!"
        cmbModel.SetFocus
        saveProductCode = False
        Exit Function
    End If
    
     If cmpModel = "" Then
        MsgBox "请选择企业型号!"
        cmbCmpModel.SetFocus
        saveProductCode = False
        Exit Function
    End If
    
     If productCode = "" Then
        MsgBox "请输入制造编号!"
        txtProductCode.SetFocus
        saveProductCode = False
        Exit Function
    End If
    
    If MsgBox("产品名称:" & productName & vbCr & "型号:" & model & vbCr & "企业型号:" & cmpModel & vbCr & "制造编号:" & productCode & _
                        vbCr & "你确定录入这个条码吗?", vbOKCancel, "条码信息录入确认") = vbCancel Then
        saveProductCode = False
        Exit Function
    End If
    
    sql = "select count(*) as saveCount from yt_product_code where product_code = '" & productCode & "'"
    Set productCodeCountRs = conn.Execute(sql)
    
    sql = "select id from yt_product where status=1 and name='" & productName & "' and model ='" & model & "'and company_model = '" & cmpModel & "'"
    Set productRs = conn.Execute(sql)
    
   
   
   
   
    
    If productCodeCountRs.Fields("saveCount") = 0 Then
        sql = "insert into yt_product_code(product_id, product_code, create_time, update_time) " & _
        "values(" & productRs.Fields("id") & ", '" & productCode & "', '" & Now & "','" & Now & "')"
       
       ' MsgBox sql
        conn.Execute (sql)
    Else
        sql = "update yt_product_code set product_id= " & productRs.Fields("id") & ", update_time = '" & Now & _
                "' where  product_code= '" & productCode & "'"
                   
       ' MsgBox sql
        conn.Execute (sql)
    End If
    
    productCodeCountRs.Close
    Set productCodeCountRs = Nothing
    productRs.Close
    Set productRs = Nothing
    saveProductCode = True

End Function


Private Function saveProductCodeContinue(ByVal mProductCode As String) As Boolean
    Dim productName As String
    Dim model As String
    Dim cmpModel As String
    Dim productCode As String
    Dim sql As String
    Dim productCodeCountRs As ADODB.Recordset
    Dim productRs As ADODB.Recordset
    
    
    productName = cmbProductName.Text
    model = cmbModel.Text
    cmpModel = cmbCmpModel.Text
    productCode = mProductCode
    
     
    If productName = "" Then
        MsgBox "请选择产品名称!"
        cmbProductName.SetFocus
        saveProductCodeContinue = False
        Exit Function
    End If
    
    If model = "" Then
        MsgBox "请选择型号!"
        cmbModel.SetFocus
        saveProductCodeContinue = False
        Exit Function
    End If
    
     If cmpModel = "" Then
        MsgBox "请选择企业型号!"
        cmbCmpModel.SetFocus
        saveProductCodeContinue = False
        Exit Function
    End If
    
     If productCode = "" Then
        MsgBox "请输入制造编号!"
        txtProductCode.SetFocus
        saveProductCodeContinue = False
        Exit Function
    End If
    
    If MsgBox("产品名称:" & productName & vbCr & "型号:" & model & vbCr & "企业型号:" & cmpModel & vbCr & "制造编号:" & productCode & _
                        vbCr & "你确定录入这个条码吗?", vbOKCancel, "条码信息录入确认") = vbCancel Then
        saveProductCodeContinue = False
        Exit Function
    End If
    
    sql = "select count(*) as saveCount from yt_product_code where product_code = '" & productCode & "'"
    Set productCodeCountRs = conn.Execute(sql)
    
    sql = "select id from yt_product where status=1 and name='" & productName & "' and model ='" & model & "'and company_model = '" & cmpModel & "'"
    Set productRs = conn.Execute(sql)
    
   
   
   
   
    
    If productCodeCountRs.Fields("saveCount") = 0 Then
        sql = "insert into yt_product_code(product_id, product_code, create_time, update_time) " & _
        "values(" & productRs.Fields("id") & ", '" & productCode & "', '" & Now & "','" & Now & "')"
       
       ' MsgBox sql
        conn.Execute (sql)
    Else
        sql = "update yt_product_code set product_id= " & productRs.Fields("id") & ", update_time = '" & Now & _
                "' where  product_code= '" & productCode & "'"
                   
       ' MsgBox sql
        conn.Execute (sql)
    End If
    
    productCodeCountRs.Close
    Set productCodeCountRs = Nothing
    productRs.Close
    Set productRs = Nothing
    saveProductCodeContinue = True

End Function



