VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProduceNoBarCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������ӡ"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9945
   Icon            =   "frmProduceNoBarCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9945
   Begin VB.CommandButton cmdInput 
      Caption         =   "��ӡ(&P)"
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtProduceNoEnd 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   1
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtProduceNoBegin 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   0
         Top             =   960
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "��ʼ����"
         Height          =   1575
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   4335
         Begin BARCODELibCtl.BarCodeCtrl beginBarCodeCtrl 
            Height          =   975
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   3975
            Style           =   7
            SubStyle        =   -1
            Validation      =   0
            LineWeight      =   3
            Direction       =   0
            ShowData        =   1
            Value           =   ""
            ForeColor       =   0
            BackColor       =   16777215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "�������"
         Height          =   1575
         Left            =   5040
         TabIndex        =   5
         Top             =   1680
         Width           =   4335
         Begin BARCODELibCtl.BarCodeCtrl endBarCodeCtrl 
            Height          =   1095
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   3735
            Style           =   7
            SubStyle        =   -1
            Validation      =   0
            LineWeight      =   3
            Direction       =   0
            ShowData        =   1
            Value           =   ""
            ForeColor       =   0
            BackColor       =   16777215
         End
      End
      Begin VB.Label Label10 
         Caption         =   "�����롰����š�"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "��ʼ������š�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "��󡰹���š�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9240
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProduceNoBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdInput_Click()
   Dim beginNoStr As String
   Dim endNoStr As String
   Dim beginNo As Integer
   Dim endNo As Integer
   Dim i As Integer
   
   
  beginNoStr = Trim(txtProduceNoBegin.Text)
  endNoStr = Trim(txtProduceNoEnd.Text)
  
  If beginNoStr = "" Then
     MsgBox "�����롰��ʼ����š�!", vbExclamation, "����������ӡ"
     txtProduceNoBegin.SetFocus
     Exit Sub
  End If
  
   If IsNumeric(beginNoStr) = False Then
     MsgBox "����ʼ����š�����Ϊ����!", vbExclamation, "����������ӡ"
     txtProduceNoBegin.SetFocus
     txtProduceNoBegin.SelStart = 0
     txtProduceNoBegin.SelLength = Len(txtProduceNoBegin)
     Exit Sub
   End If
  
  
  If Trim(endNoStr) = "" Then
    If MsgBox("�����:" & beginNoStr & vbCr & vbCr & "��ȷ��Ҫ��ӡ��������������?", vbOKCancel + vbQuestion, "����������ӡȷ��") = vbCancel Then
        Exit Sub
    End If
  
  
    If setPrinter(dlgCommonDialog) = False Then
        Exit Sub
    Else
    
        frmProduceNoBarCodePrint.lblProduceNo = beginNoStr
       
        frmProduceNoBarCodePrint.BarCodeCtrl1.Style = beginBarCodeCtrl.Style
        frmProduceNoBarCodePrint.BarCodeCtrl1.value = beginNoStr
       ' frmProductBarCodePrint.Hide
        printDoc frmProduceNoBarCodePrint
    End If
 Else
   If IsNumeric(beginNoStr) = False Then
     MsgBox "����ʼ����š�����Ϊ����!", vbExclamation, "����������ӡ"
     txtProduceNoBegin.SetFocus
     txtProduceNoBegin.SelStart = 0
     txtProduceNoBegin.SelLength = Len(txtProduceNoBegin)
     Exit Sub
   End If
   
   If IsNumeric(endNoStr) = False Then
     MsgBox "�������š�����Ϊ����!", vbExclamation, "����������ӡ"
     txtProduceNoEnd.SetFocus
     txtProduceNoEnd.SelStart = 0
     txtProduceNoEnd.SelLength = Len(txtProduceNoEnd)
     Exit Sub
   End If
   
    beginNo = beginNoStr
    endNo = endNoStr
    
    If beginNo >= endNo Then
        MsgBox "�������š�������ڡ���ʼ����š�!", vbExclamation, "����������ӡ"
        txtProduceNoEnd.SetFocus
        txtProduceNoEnd.SelStart = 0
        txtProduceNoEnd.SelLength = Len(txtProduceNoEnd)
        Exit Sub
    End If
    
    Dim count As Integer
    count = endNo - beginNo
    
    
    
    
    If MsgBox("��ʼ�����:" & beginNoStr & _
                        vbCr & "������:" & endNoStr & vbCr & "�� " & count & " ������" & vbCr & "��ȷ����ӡ��Щ������?", vbOKCancel + vbQuestion, "����������ӡȷ��") = vbCancel Then
        Exit Sub
    End If
   
    
    If setPrinter(dlgCommonDialog) = False Then
        Exit Sub
    Else
        For i = beginNo To endNo
             frmProduceNoBarCodePrint.lblProduceNo = CStr(i)
       
            frmProduceNoBarCodePrint.BarCodeCtrl1.Style = beginBarCodeCtrl.Style
            frmProduceNoBarCodePrint.BarCodeCtrl1.value = CStr(i)
            
            printDoc frmProduceNoBarCodePrint
        Next
    End If
End If
End Sub


Private Sub cmdPrinter_Click()
    setPrinter dlgCommonDialog
End Sub

Private Sub Form_Load()
'***********************************************
      If SysPass() = False Then  ''Ȩ��
        End
        Exit Sub
       End If
'***********************************************
   
    
    beginBarCodeCtrl.Style = 7
   ' beginBarCodeCtrl.SubStyle = 0
    beginBarCodeCtrl.Validation = 0
    endBarCodeCtrl.Style = 7
   ' beginBarCodeCtrl.SubStyle = 0
    endBarCodeCtrl.Validation = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmProduceNoBarCodePrint
End Sub




Private Sub txtProduceNoBegin_Change()
     beginBarCodeCtrl.value = txtProduceNoBegin & ""
End Sub

Private Sub txtProduceNoEnd_Change()
    endBarCodeCtrl.value = txtProduceNoEnd & ""
End Sub


