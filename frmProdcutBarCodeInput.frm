VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProdcutBarCodeInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ʒ����¼�벢��ӡ"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9990
   Icon            =   "frmProdcutBarCodeInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9990
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9735
      Begin VB.Frame Frame3 
         Caption         =   "�������"
         Height          =   1575
         Left            =   5160
         TabIndex        =   16
         Top             =   2760
         Width           =   4335
         Begin BARCODELibCtl.BarCodeCtrl endBarCodeCtrl 
            Height          =   1095
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   3735
            Style           =   7
            SubStyle        =   -1
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
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbModel 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbProductName 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ʼ����"
         Height          =   1575
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   4335
         Begin BARCODELibCtl.BarCodeCtrl beginBarCodeCtrl 
            Height          =   975
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3975
            Style           =   7
            SubStyle        =   -1
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
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtProductCodeEnd 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6480
         MaxLength       =   15
         TabIndex        =   4
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "��������š�"
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
         Left            =   6600
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "��ʼ�������š�"
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
         Left            =   2040
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "TZYT"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   5760
         TabIndex        =   18
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "��ѡ���Ʒ��Ϣ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "�����롰�����š�"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "TZYT"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   2160
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   735
      Left            =   5280
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "¼�벢��ӡ(&I)"
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProdcutBarCodeInput"
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
    rsload.Open "select company_model from yt_product where status=1 and name='" & productName & "' and model ='" & model & "'", conn, adOpenStatic, adLockReadOnly
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
  rsload.Open "select model from yt_product where status=1 and name='" & productName & "'", conn, adOpenStatic, adLockReadOnly
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
        MsgBox "��ѡ���Ʒ����!", vbExclamation, "��Ʒ����¼�벢��ӡ"
        cmbProductName.SetFocus
        Exit Sub
   End If
    
   If model = "" Then
        MsgBox "��ѡ���ͺ�!", vbExclamation, "��Ʒ����¼�벢��ӡ"
        
        cmbModel.SetFocus
        Exit Sub
   End If
    
   If cmpModel = "" Then
        MsgBox "��ѡ����ҵ�ͺ�!", vbExclamation, "��Ʒ����¼�벢��ӡ"
        cmbCmpModel.SetFocus
        Exit Sub
   End If
    
   
  beginCodeStr = Trim(txtProductCodeBegin.Text)
  endCodeStr = Trim(txtProductCodeEnd.Text)
  
  If beginCodeStr = "" Then
     MsgBox "�����롰��ʼ�����š�!", vbExclamation, "��Ʒ����¼�벢��ӡ"
        txtProductCodeBegin.SetFocus
        Exit Sub
  End If
  
  If Len(beginCodeStr) <= 3 Then
     MsgBox "����ʼ�����š��ĳ��ȱ������3!", vbExclamation, "��Ʒ����¼�벢��ӡ"
     txtProductCodeBegin.SetFocus
     txtProductCodeBegin.SelStart = 0
     txtProductCodeBegin.SelLength = Len(txtProductCodeBegin)
     Exit Sub
  End If
  
  
  If Trim(endCodeStr) = "" Then
    If MsgBox("��Ʒ����:" & productName & vbCr & "�ͺ�:" & model & vbCr & "��ҵ�ͺ�:" & cmpModel & vbCr & "������:" & beginCodeStr & _
                        vbCr & "��ȷ��¼�����������?", vbOKCancel + vbQuestion, "������Ϣ¼��ȷ��") = vbCancel Then
        Exit Sub
    End If
  
  
    If saveProductCode(beginCodeStr) = False Then
        Exit Sub
    End If
    
    If MsgBox("����¼��ɹ�������������Ҫ��ӡ���������", vbYesNo + vbQuestion, "�����ӡȷ��") = vbYes Then
        If setPrinter(dlgCommonDialog) = True Then
            frmProductBarCodePrint.lblProductName = cmbProductName.Text
            frmProductBarCodePrint.lblModel = cmbModel.Text
            frmProductBarCodePrint.lblCmpModel = cmbCmpModel.Text
            frmProductBarCodePrint.BarCodeCtrl1.Style = beginBarCodeCtrl.Style
            frmProductBarCodePrint.BarCodeCtrl1.value = beginCodeStr
       ' frmProductBarCodePrint.Hide
            printDoc frmProductBarCodePrint
        End If
    End If
 Else
   If Len(beginCodeStr) <= 3 Then
     MsgBox "����ʼ�����š��ĳ��ȱ������3!", vbExclamation, "��Ʒ����¼�벢��ӡ"
     txtProductCodeBegin.SetFocus
     txtProductCodeBegin.SelStart = 0
     txtProductCodeBegin.SelLength = Len(txtProductCodeBegin)
     Exit Sub
  End If
  
   If IsNumeric(beginCodeStr) = False Then
     MsgBox "����ʼ�����š�����Ϊ����!", vbExclamation, "��Ʒ����¼�벢��ӡ"
     txtProductCodeBegin.SetFocus
     txtProductCodeBegin.SelStart = 0
     txtProductCodeBegin.SelLength = Len(txtProductCodeBegin)
     Exit Sub
   End If
   
   If Len(endCodeStr) <= 3 Then
     MsgBox "����������š��ĳ��ȱ������3!", vbExclamation, "��Ʒ����¼�벢��ӡ"
     txtProductCodeEnd.SetFocus
     txtProductCodeEnd.SelStart = 0
     txtProductCodeEnd.SelLength = Len(txtProductCodeEnd)
     Exit Sub
  End If
   
   If IsNumeric(endCodeStr) = False Then
     MsgBox "����������š�����Ϊ����!", vbExclamation, "��Ʒ����¼�벢��ӡ"
     txtProductCodeEnd.SetFocus
     txtProductCodeEnd.SelStart = 0
     txtProductCodeEnd.SelLength = Len(txtProductCodeEnd)
     Exit Sub
   End If
   
    beginCode = beginCodeStr
    endCode = endCodeStr
    
    If beginCode >= endCode Then
        MsgBox "����������š�������ڡ���ʼ�����š�!", vbExclamation, "��Ʒ����¼�벢��ӡ"
        txtProductCodeEnd.SetFocus
        txtProductCodeEnd.SelStart = 0
        txtProductCodeEnd.SelLength = Len(txtProductCodeEnd)
        Exit Sub
    End If
    
    Dim count As Double
    count = endCode - beginCode
    
    
    
    
    If MsgBox("��Ʒ����:" & productName & vbCr & "�ͺ�:" & model & vbCr & "��ҵ�ͺ�:" & cmpModel & vbCr & "��ʼ������:" & beginCodeStr & _
                        vbCr & "���������:" & endCodeStr & vbCr & "�� " & count & " ������" & vbCr & "��ȷ��¼����Щ������?", vbOKCancel + vbQuestion, "������Ϣ¼��ȷ��") = vbCancel Then
        Exit Sub
    End If
   
        
    For i = beginCode To endCode
        If saveProductCode(CStr(i)) = False Then
            Exit Sub
        End If
    Next
    If MsgBox("����¼��ɹ�������������Ҫ��ӡ�� " & count & " ��������", vbYesNo + vbQuestion, "�����ӡȷ��") = vbYes Then
        If setPrinter(dlgCommonDialog) = True Then
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
     rsload.Open "select distinct name from yt_product where status=1", conn, adOpenStatic, adLockReadOnly
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
