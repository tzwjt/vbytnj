VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmProduceNoBarCodePrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   2685
   Begin VB.Label lblProduceNo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¹¤ÐòºÅ:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
      Height          =   855
      Left            =   -120
      TabIndex        =   0
      Top             =   600
      Width           =   2655
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
Attribute VB_Name = "frmProduceNoBarCodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
  Me.Hide
End Sub
