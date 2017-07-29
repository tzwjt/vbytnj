VERSION 5.00
Object = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}#1.0#0"; "MSBCODE9.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBarCode 
   Caption         =   "生成条码"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   4560
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Text            =   "4901234567894"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印条码"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "2"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Text            =   "0"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "0"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1200
      Width           =   375
   End
   Begin BARCODELibCtl.BarCodeCtrl BarCodeCtrl1 
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   4335
      Style           =   7
      SubStyle        =   0
      Validation      =   0
      LineWeight      =   3
      Direction       =   0
      ShowData        =   1
      Value           =   ""
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.Label Label1 
      Caption         =   "式样："
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "子式样："
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "无效确认方式："
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   2880
      Width           =   735
   End
End
Attribute VB_Name = "frmBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Command3_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub Command1_Click()
   BarCodeCtrl1.Value = Text1.Text
'   条码打印.BarCodeCtrl1.Value = BarCodeCtrl1.Value
End Sub
Private Sub Command2_Click()
  Form3.BarCodeCtrl1.Value = BarCodeCtrl1.Value
     mnuFilePrint_Click Form3
  
  '  Form3.PrintForm ' 将显示窗体的内容送到打印机
   ' Printer.EndDoc ' 开始打印
End Sub
Private Sub Command3_Click()
    BarCodeCtrl1.Style = Text2.Text
    BarCodeCtrl1.SubStyle = Text3.Text
    BarCodeCtrl1.Validation = Text4.Text
    
   Form3.BarCodeCtrl1.Style = BarCodeCtrl1.Style
    Form3.BarCodeCtrl1.SubStyle = BarCodeCtrl1.SubStyle
    Form3.BarCodeCtrl1.Validation = BarCodeCtrl1.Validation
End Sub


Private Sub mnuFilePrint_Click(obj As Object)
    On Error Resume Next
  '  If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog            '打印机公用对话框
        .DialogTitle = "打印"
        .CancelError = True
        .Flags = 1
         Printer.FontSize = dlgCommonDialog.FontSize
'将打印机公用对话框设置的字体大小传递给打印机
         .ShowPrinter      '在屏幕上显示【打印】公用对话框
         If Err <> MSComDlg.cdlCancel Then
            Printer.FontTransparent = False   '初始化打印的字体为不透明
          '  SetPrinterScale Form3    '匹配打印机的缩放属性与窗体的属性
          '  PrintAnywhere Printer        '可放置用户编写的打印对象参数化例程
                                      '实现字符和图形的显示
'Printer.NewPage  W         '打印机坐标初始化
'PrintAnywhere Printer     '打印另一页的内容
'Printer.NewPage           '打印机坐标初始化
         '   Printer.EndDoc               '将该任务加入打印机任务队列
' 不打印空白页
          '  Printer.KillDoc              '取消当前的打印任务
           'Form3.PrintForm ' 将显示窗体的内容送到打印机
           obj.PrintForm
    Printer.EndDoc ' 开始打印
        End If
    End With
  End Sub
