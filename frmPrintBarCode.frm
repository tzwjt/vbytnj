VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Begin VB.Form frmPrintBarCode 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "打印条码"
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "frmPrintBarCode.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   1275
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
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
End
Attribute VB_Name = "frmPrintBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
