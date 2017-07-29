VERSION 5.00
Begin VB.Form frmEmployeeScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "操作工补码"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmployeeScan.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6480
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtEmployeeNo 
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
         TabIndex        =   8
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "请扫码输入“员工编号”"
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
         Left            =   720
         TabIndex        =   7
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   6000
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblProcessName 
         Caption         =   "lblProcessName"
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
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "工序名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblProcessNo 
         Caption         =   "lblProcessNo"
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
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "工序号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblProductCode 
         Caption         =   "lblProductCode"
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
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "产品制造编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEmployeeScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim productCode As String
Dim processNo As Integer
Dim scanId As Integer


Public Sub getProductInfo(ByVal mProductCode As String, ByVal mProcessNo As Integer, ByVal mProcessName, ByVal mScanId As Integer)
    productCode = mProductCode
    processNo = mProcessNo
    lblProcessNo = mProcessNo
    lblProcessName = mProcessName
    lblProductCode = mProductCode
    scanId = mScanId
    
  '  MsgBox scanId & ":" & productCode & ":" & processNo
    
 
    
 
    
End Sub

Private Sub employee_scan(ByVal mEmpNo As String)
    Dim empNo As String
    Dim empName As String
    Dim sql As String
    Dim empRs As ADODB.Recordset
    Dim produceRs As ADODB.Recordset
    Dim recordCount As Integer
    

    empNo = Trim(mEmpNo)
    
    sql = "select count(*) as recordCount from yt_employee where emp_no = '" & empNo & "' and status =1"
    Set empRs = conn.Execute(sql)
    
    If empRs.Fields("recordCount") < 1 Then
        MsgBox "此编号对应的员工不存在,诱重新扫员工码输入"
        txtEmployeeNo.SetFocus
        txtEmployeeNo.SelStart = 0
        txtEmployeeNo.SelLength = Len(txtEmployeeNo)
        
        
        
        
        Exit Sub
    End If
    
     sql = "select emp_no, name as emp_name from yt_employee where emp_no = '" & empNo & "' and status =1"
    Set empRs = conn.Execute(sql)
    
    empNo = empRs.Fields("emp_no")
    empName = empRs.Fields("emp_name")
    
    sql = "update yt_produce_scan set operator_no = '" & empNo & "', operator_name = '" & empName & "'" & _
            ", scan_status =1, scan_time = '" & Now & "' where id =" & scanId & ""
     Set produceRs = conn.Execute(sql)
            
            
      'productCode对应的所有工序是否已全扫码
     sql = "select count(*) as recordCount from yt_produce_scan " & _
                        " where yt_produce_scan.product_code = '" & productCode & "' and scan_status = 0 "
    Set produceRs = conn.Execute(sql)
    recordCount = produceRs.Fields("recordCount")
    produceRs.Close
    If recordCount < 1 Then
         sql = "update yt_produce_scan set status=1,  update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & productCode & "' "
         Set produceRs = conn.Execute(sql)
    End If
    
    empRs.Close
    Set empRs = Nothing
    Set produceRs = Nothing
    frmAddScan.getProduceScan productCode
     Unload Me
     
    
    
    
    
    
    
    
    
    
    
    
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub txtEmployeeNo_KeyPress(KeyAscii As Integer)
    Dim empNo As String
    If KeyAscii = 13 Then
       ' MsgBox "aaa"
        empNo = Trim(txtEmployeeNo.Text)
        If empNo = "" Then
            MsgBox "请扫码输入员工编号"
            Exit Sub
        End If
        employee_scan getValidEmployeeCode(Trim(empNo))
    End If
        
        
End Sub

Private Function getValidEmployeeCode(ByVal employeeCode) As String
    Dim pos As Integer
    Dim xEmployeeCode As String
    pos = InStr(1, employeeCode, " ")
    If pos > 0 Then
        xEmployeeCode = Mid(employeeCode, pos)
    Else
         xEmployeeCode = employeeCode
    End If
    getValidEmployeeCode = Trim(xEmployeeCode)
End Function

