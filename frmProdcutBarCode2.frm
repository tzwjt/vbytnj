VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProdcutBarCode2 
   Caption         =   "��Ʒ����¼�벢��ӡ"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   14835
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame3 
         Caption         =   "�������"
         Height          =   1575
         Left            =   5880
         TabIndex        =   17
         Top             =   2760
         Width           =   3495
         Begin BARCODELibCtl.BarCodeCtrl endBarCodeCtrl 
            Height          =   615
            Left            =   600
            TabIndex        =   18
            Top             =   480
            Width           =   2055
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
      Begin VB.ComboBox cmbCmpModel 
         Height          =   300
         Left            =   7320
         TabIndex        =   16
         Text            =   "cmbCmpModel"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cmbModel 
         Height          =   300
         Left            =   4200
         TabIndex        =   15
         Text            =   "cmbModel"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cmbProductName 
         Height          =   300
         Left            =   1440
         TabIndex        =   14
         Text            =   "cmbProductName"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ʼ����"
         Height          =   1335
         Left            =   720
         TabIndex        =   6
         Top             =   2760
         Width           =   4215
         Begin BARCODELibCtl.BarCodeCtrl startBarCodeCtrl 
            Height          =   975
            Left            =   240
            TabIndex        =   7
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
      Begin VB.TextBox txtProductCodeBegin 
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtProductCodeEnd 
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "��������š�"
         Height          =   255
         Left            =   6720
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "��ʼ�������š�"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "TZYT"
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "��ѡ���Ʒ��Ϣ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "�����롰�����š�"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "TZYT"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "��ҵ�ͺ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "�ͺ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "��Ʒ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "���ô�ӡ��(&P)"
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "¼�벢��ӡ(&I)"
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   11760
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProdcutBarCode2"
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
  Do Until rsload.EOF '���ò�Ʒ�������ͺŷ�����������
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
  Do Until rsload.EOF '���ò�Ʒ�������ͺŷ�����������
    Me.cmbModel.AddItem rsload.Fields("model")
    rsload.MoveNext
  Loop
   rsload.Close
  Set rsload = Nothing
End Sub




Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If saveProductCode() = False Then
        Exit Sub
    End If
    
    If MsgBox("����¼��ɹ�������������Ҫ��ӡ��������", vbYesNo, "�����ӡȷ��") = vbYes Then
         frmProductBarCodePrint.lblProductName = cmbProductName.Text
        frmProductBarCodePrint.lblModel = cmbModel.Text
        frmProductBarCodePrint.lblCmpModel = cmbCmpModel.Text
        frmProductBarCodePrint.BarCodeCtrl1.Style = BarCodeCtrl1.Style
        frmProductBarCodePrint.BarCodeCtrl1.value = BarCodeCtrl1.value
        setPrint dlgCommonDialog, frmProductBarCodePrint
    End If
     
  
End Sub

Private Sub cmdInput_Click()
   Dim beginCode As Long
   Dim endCode As Long
   Dim i As Long
   Dim productName As String
   Dim model As String
   Dim cmpModel As String
   
   productName = Trim(cmbProductName.Text)
   model = Trim(cmbModel.Text)
   cmpModel = Trim(cmbCmpModel.Text)
   
   If productName = "" Then
        MsgBox "��ѡ���Ʒ����!"
        cmbProductName.SetFocus
        Exit Sub
   End If
    
   If model = "" Then
        MsgBox "��ѡ���ͺ�!"
        cmbModel.SetFocus
        Exit Sub
   End If
    
   If cmpModel = "" Then
        MsgBox "��ѡ����ҵ�ͺ�!"
        cmbCmpModel.SetFocus
        Exit Sub
   End If
    
   
  beginCode = Trim(txtProductCodeBegin)
  endCode = Trim(txtProductCodeEnd)
  
  If beginCode = "" Then
     MsgBox "�����롰��ʼ�����š�!"
        txtProductCodeBegin.SetFocus
        Exit Sub
  End If
  
  
  If Trim(endCode) = "" Then
    If MsgBox("��Ʒ����:" & productName & vbCr & "�ͺ�:" & model & vbCr & "��ҵ�ͺ�:" & cmpModel & vbCr & "������:" & beginCode & _
                        vbCr & "��ȷ��¼�����������?", vbOKCancel, "������Ϣ¼��ȷ��") = vbCancel Then
        Exit Sub
    End If
  
  
    saveProductCode beginCode
    frmProductBarCodePrint.lblProductName = cmbProductName.Text
    frmProductBarCodePrint.lblModel = cmbModel.Text
    frmProductBarCodePrint.lblCmpModel = cmbCmpModel.Text
    frmProductBarCodePrint.BarCodeCtrl1.Style = beginBarCodeCtrl.Style
    frmProductBarCodePrint.BarCodeCtrl1.value = beginCode
    printDoc frmProductBarCodePrint
 Else
    If MsgBox("��Ʒ����:" & productName & vbCr & "�ͺ�:" & model & vbCr & "��ҵ�ͺ�:" & cmpModel & vbCr & "��ʼ������:" & beginCode & _
                        vbCr & "���������:" & endCode & vbCr & "��ȷ��¼�����������?", vbOKCancel, "������Ϣ¼��ȷ��") = vbCancel Then
        Exit Sub
    End If
    For i = beginCode To endCode
        saveProductCode CStr(i)
        frmProductBarCodePrint.lblProductName = cmbProductName.Text
        frmProductBarCodePrint.lblModel = cmbModel.Text
        frmProductBarCodePrint.lblCmpModel = cmbCmpModel.Text
        frmProductBarCodePrint.BarCodeCtrl1.Style = beginBarCodeCtrl.Style
        frmProductBarCodePrint.BarCodeCtrl1.value = CStr(i)
        printDoc frmProductBarCodePrint
    Next
End If
End Sub


Private Sub cmdPrinter_Click()
    setPrinter dlgCommonDialog
End Sub

Private Sub Form_Load()
'***********************************************
      If SysPass() = False Then  ''Ȩ��
        End
        Exit Sub
       End If
'***********************************************
    Dim rsload As ADODB.Recordset
   '  MsgBox LoginSucceeded
   
   If conn Is Nothing Then
     '�������ݿ�
        If dbConnOpen() = False Then
            MsgBox "�������ݿ�ʧ��!"
            End
        End If
    End If
    Set rsload = New ADODB.Recordset
     rsload.Open "select distinct name from yt_product where status=1", conn, adOpenStatic, adLockPessimistic
     Me.cmbProductName.Clear
    Do Until rsload.EOF '�����еĵĲ�Ʒ���Ʒ�����������
         Me.cmbProductName.AddItem rsload.Fields("name")
        rsload.MoveNext
    Loop
     rsload.Close
    Set rsload = Nothing
    
    beginBarCodeCtrl.Style = 7
   ' beginBarCodeCtrl.SubStyle = 0
    beginBarCodeCtrl.Validation = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmProductBarCodePrint
End Sub

Private Sub txtProductCode_Change()
    BarCodeCtrl1.value = txtProductCode & ""
    

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
     startBarCodeCtrl.value = txtProductCodeBegin & ""
End Sub

Private Sub txtProductCodeEnd_Change()
    endBarCodeCtrl.value = txtProductCodeEnd & ""
End Sub
