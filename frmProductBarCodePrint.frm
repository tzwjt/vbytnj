VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmProductBarCodePrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3825
   Icon            =   "frmProductBarCodePrint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3825
   Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3255
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
   Begin VB.Label lblCmpModel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "lblCmpModel"
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
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "型号:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "产品名称:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblModel 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblProductName 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "企业型号:"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmProductBarCodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
   Me.Hide
End Sub

