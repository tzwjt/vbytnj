VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0FFC0&
   Caption         =   "ӣ��ũ��������ɨ��ϵͳ"
   ClientHeight    =   7710
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17235
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuBarCode 
      Caption         =   "����׼��"
      Begin VB.Menu subMunProductBarCode 
         Caption         =   "��Ʒ����¼�벢��ӡ"
      End
      Begin VB.Menu subMunProductBarCodeQuery 
         Caption         =   "��Ʒ�����ѯ"
      End
      Begin VB.Menu subMunProduceBarCode 
         Caption         =   "����������ӡ"
      End
      Begin VB.Menu subMunProduceQuery 
         Caption         =   "���������ѯ"
      End
      Begin VB.Menu subMunEmployeeBarCode 
         Caption         =   "Ա�������ӡ"
      End
   End
   Begin VB.Menu mnuAutoScan 
      Caption         =   "����ɨ��"
   End
   Begin VB.Menu mnuAddScan 
      Caption         =   "����Ա��ɨ��"
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "ϵͳ"
      Begin VB.Menu mnuUpdatePwd 
         Caption         =   "�޸�����"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "�˳�ϵͳ"
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
      If SysPass() = False Then  ''Ȩ��
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
