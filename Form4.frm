VERSION 5.00
Begin VB.Form frmAutoScan 
   Caption         =   "生产扫码"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Timer tmrScan 
      Interval        =   1000
      Left            =   960
      Top             =   720
   End
End
Attribute VB_Name = "frmAutoScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
   SetHook
End Sub
Private Sub Form_Unload(Cancel As Integer)
   UnHook
End Sub
Private Sub tmrScan_Timer()
    Dim strBarCode As String
     MsgBox "abb"
    strBarCode = GetBarCode
    If Len(strBarCode) > 0 Then
        MsgBox "条形码:" & strBarCode
    End If
End Sub


