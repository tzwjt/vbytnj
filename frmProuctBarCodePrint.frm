VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmProuctBarCodePrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   1635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   ScaleHeight     =   1635
   ScaleWidth      =   5445
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label2 
      Caption         =   "lblModel"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "lblProductName"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3735
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
Attribute VB_Name = "frmProuctBarCodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
