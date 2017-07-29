VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEmployeeBarCode 
   Caption         =   "员工条码打印"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10170
   Icon            =   "frmEmployeeBarCode.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   10170
   Begin VB.Frame Frame3 
      Caption         =   "条码区"
      Height          =   1575
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   4695
      Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
         Height          =   975
         Left            =   120
         TabIndex        =   13
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
      Begin VB.Label lblEmpName 
         Caption         =   "lblEmpName"
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
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblEmpNo 
         Caption         =   "lblEmpNo"
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
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "员工信息"
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   9135
      Begin MSDataGridLib.DataGrid dataGridEmployee 
         Height          =   3375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5953
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
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   735
      Left            =   7320
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   735
      Left            =   5400
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "请输入查询条件"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtEmpName 
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
         Left            =   4800
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询(&Q)"
         Height          =   615
         Left            =   7080
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtEmpNo 
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
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "员工姓名"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "员工编码"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   8280
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEmployeeBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsQuery As ADODB.Recordset
Dim sql As String
Private Sub Command1_Click()
   End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If rsQuery.recordCount > 0 Then
        frmEmployeeBarCodePrint.BarCodeCtrl1.Style = BarCodeCtrl1.Style
        frmEmployeeBarCodePrint.BarCodeCtrl1.value = BarCodeCtrl1.value
        frmEmployeeBarCodePrint.lblEmpNo = dataGridEmployee.Columns(0).value
        frmEmployeeBarCodePrint.lblEmpName = dataGridEmployee.Columns(1).value
        setPrint dlgCommonDialog, frmEmployeeBarCodePrint
    Else
        frmEmployeeBarCodePrint.BarCodeCtrl1.value = ""
        frmEmployeeBarCodePrint.lblEmpNo = ""
        frmEmployeeBarCodePrint.lblEmpName = ""
    End If
End Sub

Private Sub cmdQuery_Click()
    Dim empName As String
    Dim empNo As String
    
    empName = Trim(txtEmpName.Text)
    empNo = Trim(txtEmpNo.Text)
    
    If empName = "" And empNo = "" Then
        MsgBox "请输入查询条件!"
        txtEmpNo.SetFocus
        Exit Sub
    End If
    
  
    sql = "select emp_no, name, gender, work_type, department from yt_employee where status =1 "
    
    If empNo <> "" Then
        sql = sql & "and emp_no like '%" & empNo & "%'"
    End If
    
    If empName <> "" Then
        sql = sql & " and name like '%" & empName & "%'"
    End If
    
    rsQuery.Close
    rsQuery.CursorLocation = adUseClient
  '  rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
   rsQuery.Open sql, conn, adOpenStatic, adLockReadOnly
    
    
    If rsQuery.recordCount = 0 Then
        MsgBox "没有符合条件的员工!"
       nullQuery
        Exit Sub
    End If
    
   dataGridEmployee.ReBind
 ' Set DataGrid1.DataSource = rsQuery
   '  Set DataGrid1.DataMember = rsQuery
  
    setDataGridColumns
    setBarCode dataGridEmployee.Columns(0).value
    setEmp dataGridEmployee.Columns(0).value, dataGridEmployee.Columns(1).value
    
    
     
    
    
   
   
   
   
   
  ' DataGrid1.dataM
     
     
    
        
    
    
 
    
    
    
    
    
  
    
    
  
    
  
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub DataGrid1_Click()
   ' MsgBox "aaa"
    
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
      
   ' MsgBox "bbb"
End Sub

Private Sub dataGridEmployee_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsQuery.recordCount > 0 Then
        setBarCode dataGridEmployee.Columns(0).value
        setEmp dataGridEmployee.Columns(0).value, dataGridEmployee.Columns(1).value
    End If
End Sub

Private Sub Form_Load()
 Dim rsload As ADODB.Recordset
   '  MsgBox LoginSucceeded
   lblEmpNo.Caption = ""
   lblEmpName.Caption = ""
   If conn Is Nothing Then
     '连接数据库
        If dbConnOpen() = False Then
            MsgBox "连接数据库失败!"
            End
        End If
    End If
   
    Set rsQuery = New ADODB.Recordset
    
    sql = "select emp_no, name, gender, work_type, department from yt_employee where status =1 "
    rsQuery.CursorLocation = adUseClient
   ' rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    rsQuery.Open sql, conn, adOpenStatic, adLockReadOnly
    
    Set dataGridEmployee.DataSource = rsQuery
    setDataGridColumns
    
    BarCodeCtrl1.Style = 7
    BarCodeCtrl1.SubStyle = 0
    BarCodeCtrl1.Validation = 0
    If rsQuery.recordCount > 0 Then
        setBarCode dataGridEmployee.Columns(0).value
        setEmp dataGridEmployee.Columns(0).value, dataGridEmployee.Columns(1).value
    End If
    
     Me.Width = 11000
    Me.Height = 8500

End Sub

Private Sub setBarCode(code As String)
    BarCodeCtrl1.value = code
End Sub

Private Sub setEmp(empNo As String, empName As String)
    lblEmpNo.Caption = empNo
    lblEmpName.Caption = empName
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
     dataGridEmployee.Height = Frame2.Height - 800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsQuery.Close
    Set rsQuery = Nothing
    
    Unload frmEmployeeBarCodePrint
End Sub

Private Sub setDataGridColumns()
    dataGridEmployee.Columns("emp_no").Caption = "编号"
    dataGridEmployee.Columns(0).Width = 1200
    
    dataGridEmployee.Columns("name").Caption = "姓名"
    dataGridEmployee.Columns(1).Width = 1500
    dataGridEmployee.Columns("gender").Caption = "性别"
    dataGridEmployee.Columns(2).Width = 500
    dataGridEmployee.Columns("work_type").Caption = "工种"
    dataGridEmployee.Columns(3).Width = 1500
    dataGridEmployee.Columns("department").Caption = "所属部门"
    dataGridEmployee.Columns(4).Width = 1500
   
End Sub

Private Sub nullQuery()
    sql = "select emp_no, name, gender, work_type, department from yt_employee where emp_no='' "
    rsQuery.Close
    rsQuery.CursorLocation = adUseClient
   ' rsQuery.Open sql, conn, adOpenStatic, adLockPessimistic
    rsQuery.Open sql, conn, adOpenStatic, adLockReadOnly
    Set dataGridEmployee.DataSource = rsQuery
    setDataGridColumns
    setBarCode ""
    setEmp "", ""
End Sub
