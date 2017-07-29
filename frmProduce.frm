VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProduceNo 
   Caption         =   "工序打印"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   10425
   Begin VB.CommandButton cmdInput 
      Caption         =   "录入并打印(&I)"
      Height          =   735
      Left            =   2520
      TabIndex        =   12
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   735
      Left            =   6360
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "设置打印机(&P)"
      Height          =   735
      Left            =   4440
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtProduceNoEnd 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6120
         MaxLength       =   15
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtProduceNoBegin 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "起始条码"
         Height          =   1575
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   4335
         Begin BARCODELibCtl.BarCodeCtrl beginBarCodeCtrl 
            Height          =   975
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3975
            Style           =   7
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
      Begin VB.Frame Frame3 
         Caption         =   "最后条码"
         Height          =   1575
         Left            =   5520
         TabIndex        =   1
         Top             =   1560
         Width           =   4335
         Begin BARCODELibCtl.BarCodeCtrl endBarCodeCtrl 
            Height          =   1095
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   3735
            Style           =   7
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
      Begin VB.Label Label10 
         Caption         =   "请输入“工序号”"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "起始“工序号”"
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
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "最后“工序号”"
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
         Left            =   6360
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProduceNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdInput_Click()
   Dim beginCodeStr As String
   Dim endCodeStr As String
   Dim beginCode As Double
   Dim endCode As Double
   Dim i As Double
   Dim productName As String
   Dim model As String
   Dim cmpModel As String
   
   productName = Trim(cmbProductName.Text)
   model = Trim(cmbModel.Text)
   cmpModel = Trim(cmbCmpModel.Text)
   
   If productName = "" Then
        MsgBox "请选择产品名称!", vbExclamation, "产品条码录入并打印"
        cmbProductName.SetFocus
        Exit Sub
   End If
    
   If model = "" Then
        MsgBox "请选择型号!", vbExclamation, "产品条码录入并打印"
        
        cmbModel.SetFocus
        Exit Sub
   End If
    
   If cmpModel = "" Then
        MsgBox "请选择企业型号!", vbExclamation, "产品条码录入并打印"
        cmbCmpModel.SetFocus
        Exit Sub
   End If
    
   
  beginCodeStr = Trim(txtProductCodeBegin.Text)
  endCodeStr = Trim(txtProductCodeEnd.Text)
  
  If beginCodeStr = "" Then
     MsgBox "请输入“起始制造编号”!", vbExclamation, "产品条码录入并打印"
        txtProductCodeBegin.SetFocus
        Exit Sub
  End If
  
  
  If Trim(endCodeStr) = "" Then
    If MsgBox("产品名称:" & productName & vbCr & "型号:" & model & vbCr & "企业型号:" & cmpModel & vbCr & "制造编号:" & beginCodeStr & _
                        vbCr & "你确定录入这个条码吗?", vbOKCancel + vbQuestion, "条码信息录入确认") = vbCancel Then
        Exit Sub
    End If
  
  
    If saveProductCode(beginCodeStr) = False Then
        Exit Sub
    End If
    
    If MsgBox("数据录入成功，接下来你需要打印这个条码吗？", vbYesNo + vbQuestion, "条码打印确认") = vbYes Then
        frmProductBarCodePrint.lblProductName = cmbProductName.Text
        frmProductBarCodePrint.lblModel = cmbModel.Text
        frmProductBarCodePrint.lblCmpModel = cmbCmpModel.Text
        frmProductBarCodePrint.BarCodeCtrl1.Style = beginBarCodeCtrl.Style
        frmProductBarCodePrint.BarCodeCtrl1.value = beginCodeStr
       ' frmProductBarCodePrint.Hide
        printDoc frmProductBarCodePrint
    End If
 Else
   If IsNumeric(beginCodeStr) = False Then
     MsgBox "“起始制造编号”必须为数字!", vbExclamation, "产品条码录入并打印"
     txtProductCodeBegin.SetFocus
     txtProductCodeBegin.SelStart = 0
     txtProductCodeBegin.SelLength = Len(txtProductCodeBegin)
     Exit Sub
   End If
   
   If IsNumeric(endCodeStr) = False Then
     MsgBox "“最后制造编号”必须为数字!", vbExclamation, "产品条码录入并打印"
     txtProductCodeEnd.SetFocus
     txtProductCodeEnd.SelStart = 0
     txtProductCodeEnd.SelLength = Len(txtProductCodeEnd)
     Exit Sub
   End If
   
    beginCode = beginCodeStr
    endCode = endCodeStr
    
    If beginCode >= endCode Then
        MsgBox "“最后制造编号”必须大于“起始制造编号”!", vbExclamation, "产品条码录入并打印"
        txtProductCodeEnd.SetFocus
        txtProductCodeEnd.SelStart = 0
        txtProductCodeEnd.SelLength = Len(txtProductCodeEnd)
        Exit Sub
    End If
    
    Dim count As Double
    count = endCode - beginCode
    
    
    
    
    If MsgBox("产品名称:" & productName & vbCr & "型号:" & model & vbCr & "企业型号:" & cmpModel & vbCr & "起始制造编号:" & beginCodeStr & _
                        vbCr & "最后制造编号:" & endCodeStr & vbCr & "共 " & count & " 个条码" & vbCr & "你确定录入这些条码吗?", vbOKCancel + vbQuestion, "条码信息录入确认") = vbCancel Then
        Exit Sub
    End If
   
        
    For i = beginCode To endCode
        If saveProductCode(CStr(i)) = False Then
            Exit Sub
        End If
    Next
    If MsgBox("数据录入成功，接下来你需要打印这 " & count & " 个条码吗？", vbYesNo + vbQuestion, "条码打印确认") = vbYes Then
        For i = beginCode To endCode
            frmProductBarCodePrint.lblProductName = cmbProductName.Text
            frmProductBarCodePrint.lblModel = cmbModel.Text
            frmProductBarCodePrint.lblCmpModel = cmbCmpModel.Text
            frmProductBarCodePrint.BarCodeCtrl1.Style = beginBarCodeCtrl.Style
            frmProductBarCodePrint.BarCodeCtrl1.value = CStr(i)
          '  frmProductBarCodePrint.Hide
            printDoc frmProductBarCodePrint
        Next
    End If
End If
End Sub


Private Sub cmdPrinter_Click()
    setPrinter dlgCommonDialog
End Sub

Private Sub Form_Load()
'***********************************************
      If SysPass() = False Then  ''权限
        End
        Exit Sub
       End If
'***********************************************
   
    
    beginBarCodeCtrl.Style = 7
   ' beginBarCodeCtrl.SubStyle = 0
    beginBarCodeCtrl.Validation = 0
    endBarCodeCtrl.Style = 7
   ' beginBarCodeCtrl.SubStyle = 0
    endBarCodeCtrl.Validation = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmProductBarCodePrint
End Sub



Private Function saveProductCode(ByVal mProductCode As String) As Boolean
    Dim productName As String
    Dim model As String
    Dim cmpModel As String
    Dim productCode As String
    Dim sql As String
    Dim productCodeCountRs As ADODB.Recordset
    Dim productRs As ADODB.Recordset
    
    
    productName = Trim(cmbProductName.Text)
    model = Trim(cmbModel.Text)
    cmpModel = Trim(cmbCmpModel.Text)
    productCode = mProductCode
    
     
    sql = "select count(*) as saveCount from yt_product_code where product_code = '" & productCode & "'"
    Set productCodeCountRs = conn.Execute(sql)
    
    sql = "select id from yt_product where name='" & productName & "' and model ='" & model & "'and company_model = '" & cmpModel & "'"
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

Private Sub txtProductCodeBegin_Change()
     beginBarCodeCtrl.value = txtProductCodeBegin & ""
End Sub

Private Sub txtProductCodeEnd_Change()
    endBarCodeCtrl.value = txtProductCodeEnd & ""
End Sub


