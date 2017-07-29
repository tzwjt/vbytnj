VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0FFC0&
   Caption         =   "樱田农机生产线扫码系统"
   ClientHeight    =   7710
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17235
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuBarCode 
      Caption         =   "生产准备"
      Begin VB.Menu subMunProductBarCode 
         Caption         =   "产品条码录入并打印"
      End
      Begin VB.Menu subMunProductBarCodeQuery 
         Caption         =   "产品条码查询"
      End
      Begin VB.Menu subMunProduceBarCode 
         Caption         =   "工序号条码打印"
      End
      Begin VB.Menu subMunProduceQuery 
         Caption         =   "生产工序查询"
      End
      Begin VB.Menu subMunEmployeeBarCode 
         Caption         =   "员工条码打印"
      End
   End
   Begin VB.Menu mnuAutoScan 
      Caption         =   "生产扫码"
   End
   Begin VB.Menu mnuAddScan 
      Caption         =   "检验员补扫码"
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "系统"
      Begin VB.Menu mnuUpdatePwd 
         Caption         =   "修改密码"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "退出系统"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AutoScan_Click()
'Load Form4
End Sub

Private Sub BarCode_Click()
'Load Form2
End Sub

Private Sub exit_Click()
  Unload Me
End Sub

Private Sub MDIForm_Load()
'***********************************************
      If SysPass() = False Then  ''权限
        End
        Exit Sub
       End If
'***********************************************
'   Picture1.PaintPicture LoadPicture("background.jpg"), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub

Private Sub MDIForm_Resize()
'    Picture.Height = Me.Height
 '   Picture.Width = Me.Width
   

End Sub

Private Sub mnuAbout_Click()
     Load frmAbout
End Sub

Private Sub mnuAddScan_Click()
     Load frmCheckerAddScan
End Sub

Private Sub mnuAutoScan_Click()
    Load frmAutoScan
End Sub

Private Sub mnuBarCode_Click()
   ' Load frmBarCode
End Sub

Private Sub mnuExit_Click()
    dbDisConnect
    Unload Me
    End
End Sub

Private Sub mnuUpdatePwd_Click()
    Load frmUpdatePwd
End Sub

Private Sub subMunEmployeeBarCode_Click()
    Load frmEmployeeBarCode
End Sub

Private Sub subMunProduceBarCode_Click()
    Load frmProduceNoBarCode
End Sub

Private Sub subMunProduceQuery_Click()
    Load frmProcessQuery
End Sub

Private Sub subMunProductBarCode_Click()
    Load frmProdcutBarCodeInput
End Sub

Private Sub subMunProductBarCodeQuery_Click()
    Load frmProductBarCodeQuery
End Sub
