VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmProduceNoPrint 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label lblProduceNo 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "¹¤ÐòºÅ:"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2175
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
Attribute VB_Name = "frmProduceNoPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
