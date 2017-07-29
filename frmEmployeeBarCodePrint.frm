VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmEmployeeBarCodePrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3510
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "员工姓名:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "员工编号:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblEmpName 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblEmpNo 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3135
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
Attribute VB_Name = "frmEmployeeBarCodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Hide
End Sub

Private Sub Form_Resize()
    Me.Hide
End Sub
