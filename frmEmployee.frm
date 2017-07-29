VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   LinkTopic       =   "Form2"
   ScaleHeight     =   4950
   ScaleWidth      =   10005
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印"
      Height          =   855
      Left            =   1680
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   615
         Left            =   5160
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label14"
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label13"
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label12"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label11"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
      End
      Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
         Height          =   615
         Left            =   6000
         TabIndex        =   13
         Top             =   1800
         Width           =   2895
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
      Begin VB.Label Label10 
         Caption         =   "条码"
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "性别"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "所属部门"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "工种"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "姓名"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "编号"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "员工信息"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "员工姓名"
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "员工编码"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "请输入查询条件"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
umload Me
End Sub
