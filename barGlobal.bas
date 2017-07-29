Attribute VB_Name = "basGlobal"

'*************************************************************************
'**描    述：数据库相关操作

'*************************************************************************

Option Explicit

Public loginUser As String '保存登陆的用户名

Public index As Integer

Private Const DB_HOST = "127.0.0.1" ' "we.e-rabits.com" '数据库服务器主机
Private Const DB_USER = "root"              '数据库登录用户
Private Const DB_PASS = "root2015" ''"Erabitsroot"       '数据库登录密码
Private Const DB_DATABASE = "ytdb1" '"ytdb"          '数据库名

 Private IsConnect As Boolean '标记数据库是否连接

Public conn As ADODB.Connection



Public Function dbConnOpen() As Boolean
 On Error Resume Next
 IsConnect = False
  Set conn = New ADODB.Connection
  conn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
    "SERVER=" & DB_HOST & ";" & _
    "DATABASE=" & DB_DATABASE & ";" & _
    "UID=" & DB_USER & ";" & _
    "PWD=" & DB_PASS & ";" & _
    "OPTION=3;stmt=SET NAMES GB2312"
   
  conn.Open
  
  If conn.State <> adStateOpen Then
    IsConnect = False
    dbConnOpen = False
  Else
    IsConnect = True
    dbConnOpen = True
  End If
 
End Function

'断开与数据库的连接
Public Sub dbDisConnect()
 Dim rc As Long
  If IsConnect = False Then
   Exit Sub
 End If
 '关闭连接
 conn.Close
 '释放cnn
 Set conn = Nothing
 IsConnect = False
End Sub

Public Function isDbConnect() As Boolean
    isDbConnect = IsConnect
End Function

'执行数据库查询语句
Public Function dbQueryExt(ByVal TmpSQLstmt As String) As ADODB.Recordset
   
  Dim rst As New ADODB.Recordset '创建Rescordset对象rst
 
  dbConnOpen '连接数据库
 
  Set rst.ActiveConnection = conn '设置rst的ActiveConnection属性,指定与其相关的数据库的连接
 
  rst.CursorType = adOpenDynamic '设置游标类型
 
  rst.LockType = adLockOptimistic '设置锁定类型
 
  rst.Open TmpSQLstmt '打开记录集
 
  Set dbQueryExt = rst '返回记录集
 
  End Function
  
Public Sub setPrint(dlgCommonDialog As Object, docObj As Object)
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
    docObj.PrintForm
         
    Printer.EndDoc ' 开始打印
     docObj.Hide
        End If
    End With
  End Sub
  
    
Public Function setPrinter(dlgCommonDialog As Object) As Boolean
    On Error Resume Next
  '  If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog            '打印机公用对话框
        .DialogTitle = "设置打印机"
        .CancelError = True
        .Flags = 1
         Printer.FontSize = dlgCommonDialog.FontSize
'将打印机公用对话框设置的字体大小传递给打印机
         .ShowPrinter      '在屏幕上显示【打印】公用对话框
          If Err = MSComDlg.cdlCancel Then
            setPrinter = False
          Else
            Printer.FontTransparent = False   '初始化打印的字体为不透明
            setPrinter = True
         End If
    End With
  End Function
  
  Public Sub printDoc(docObj As Object)
    On Error Resume Next

    Printer.FontTransparent = False   '初始化打印的字体为不透明
         
    docObj.PrintForm
   
    Printer.EndDoc ' 开始打印
     docObj.Hide
  End Sub

Sub Main()
 'connOpen
 'Load frmMain
 'frmMain.Show
 frmLogin.Show
 
End Sub

Public Sub FF()
   index = 0
  index = index + 1
'  MsgBox index
End Sub

Public Function saveProduceScan(ByVal processNo As Integer, ByVal productCode As String) As Boolean
     Dim sql As String
     Dim produceRs As ADODB.Recordset
     Dim processRs As ADODB.Recordset
     Dim recordsAffected As Long
     Dim rec
    
     If conn Is Nothing Then
     '连接数据库
        If dbConnOpen() = False Then
            MsgBox "连接数据库失败!"
            End
            Exit Function
        End If
     
    End If
    
    'yt_produce_scan 中是否已存在produceCode对应的产品
    sql = "select count(*) as recordCount from yt_produce_scan where product_code = '" & productCode & "'"
    Set produceRs = conn.Execute(sql)
   ' MsgBox produceRs.Fields("recordCount")
    If produceRs.Fields("recordCount") < 1 Then
        produceRs.Close
        sql = "insert into yt_produce_scan(product_code,product_name,product_model,product_company_model,process_no,process_name,operator_no,operator_name) " & _
                   " select * from (select yt_product_code.product_code as product_code, yt_product.name as product_name, yt_product.model as product_model,yt_product.company_model as product_company_model,yt_process.process_no as process_no,yt_process.process_name as process_name, yt_employee.emp_no as operator_no, yt_employee.name as operator_name " & _
                    " from yt_product_code, yt_product, yt_process, yt_employee" & _
                   " where yt_product_code.product_id = yt_product.id and yt_product.id = yt_process.product_id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product_code.product_code = '" & productCode & "' order by yt_process.process_no) as tb"
            
        Set produceRs = conn.Execute(sql)
     
       
    Else
        produceRs.Close
    End If
        
    
     'yt_produce_scan 中是否已存在produceCode、processNo对应的产品
    sql = "select count(*) as recordCount from yt_produce_scan where product_code = '" & productCode & "' and process_no=" & processNo & " "
   
    Set produceRs = conn.Execute(sql)
    
    If produceRs.Fields("recordCount") < 1 Then
        produceRs.Close
        sql = "insert into yt_produce_scan(product_code,product_name,product_model,product_company_model,process_no,process_name,operator_no,operator_name) " & _
                    " select * from (select yt_product_code.product_code as product_code, yt_product.name as product_name, yt_product.model as product_model,yt_product.company_model as product_company_model,yt_process.process_no as process_no,yt_process.process_name as process_name, yt_employee.emp_no as operator_no, yt_employee.name as operator_name " & _
                    " from yt_product_code, yt_product, yt_process, yt_employee" & _
                    " where yt_product_code.product_id = yt_product.id and yt_product.id = yt_process.product_id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product_code.product_code = '" & productCode & "' and  yt_process.process_no = " & processNo & ") as tb"
      '  MsgBox sql
        Set produceRs = conn.Execute(sql)
     '    produceRs.Close
    '    Set produceRs = Nothing
    Else
        produceRs.Close
    End If
   
   '对本次productCode、processNo对应的产品作“已扫码”处理
    sql = "update yt_produce_scan set scan_status=1, scan_time='" & Now & "', update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & productCode & "' and  yt_produce_scan.process_no = " & processNo & ""
     Set produceRs = conn.Execute(sql, recordsAffected)
     
     'productCode对应的所有工序是否已全扫码
     sql = "select count(*) as recordCount from yt_produce_scan " & _
                        " where yt_produce_scan.product_code = '" & productCode & "' and scan_status = 0 "
     Set produceRs = conn.Execute(sql)
    If produceRs.Fields("recordCount") < 1 Then
         produceRs.Close
         sql = "update yt_produce_scan set status=1,  update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & productCode & "' "
         Set produceRs = conn.Execute(sql)
    Else
        produceRs.Close
    End If
        
     
     '本次扫码是否正确
   '   sql = "select count(*) as recordCount from yt_produce_scan " & _
   '                     " where yt_produce_scan.product_code = '" & productCode & "' and  yt_produce_scan.process_no = " & processNo & " and scan_status=1 "
   '   Set produceRs = conn.Execute(sql)
  '    If produceRs.Fields("recordCount") > 0 Then
     If recordsAffected > 0 Then
        saveProduceScan = True
      '  doSpeak "扫码成功"
      Else
        saveProduceScan = False
      '  doSpeak "扫码失败"
      End If
            
  '  produceRs.Close
    Set produceRs = Nothing
    
End Function

Public Sub doBarCode2(queue As BarCodeQueue)
    Dim elements()
    Dim processNo As Integer
    Dim productCode As String
    Dim i As Integer
    Dim conn1 As ADODB.Connection
    
    
    elements = queue.ReadElements
    For i = 0 To queue.MaxLen - 1
        If Not IsEmpty(elements(i)) Then
            'MsgBox elements(i)
            processNo = Val(Left(elements(i), 2))
            productCode = Mid(elements(i), 4)
           ' MsgBox CStr(processNo) + " :" + productCode
           ' saveProduceScan processNo, productCode
           
            Dim sql As String
     Dim produceRs As ADODB.Recordset
     Dim processRs As ADODB.Recordset
     Dim rec
    
     If conn1 Is Nothing Then
     '连接数据库
       ' If dbConnOpen() = False Then
       '     MsgBox "连接数据库失败!"
       '     End
       '     Exit Sub
       ' End If
         Set conn1 = New ADODB.Connection
  conn1.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
    "SERVER=" & DB_HOST & ";" & _
    "DATABASE=" & DB_DATABASE & ";" & _
    "UID=" & DB_USER & ";" & _
    "PWD=" & DB_PASS & ";" & _
    "OPTION=3;stmt=SET NAMES GB2312"
   
  conn1.Open
       
       
       
    End If
    
    
    sql = "select count(*) as recordCount from yt_produce_scan where product_code = '" & productCode & "'"
    Set produceRs = conn1.Execute(sql)
    If produceRs.Fields("recordCount") < 1 Then
        sql = "insert into yt_produce_scan（product_code,product_name,product_model,product_company_model,process_no,process_name,operator_no,operator_name) " & _
                    " (select yt_product_code.product_code, yt_product.name, yt_product.model,yt_product.company_model,yt_process.process_no,yt_process.process_name, yt_employee.emp_no, yt_employee.name " & _
                    " from yt_product_code, yt_product, yt_process, yt_employee" & _
                    " where yt_product_code.product_id = yt_product.id and yt_product.id = yt_process.product_id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product_code.product_code = '" & productCode & "' order by yt_process.process_no)"
        produceRs.Close
        Set produceRs = Nothing
        Set produceRs = conn1.Execute(sql)
    End If
    
    sql = "select count(*) as recordCount from yt_produce_scan where product_code = '" & productCode & "' and process_no=" & processNo & " and status=1"
    produceRs.Close
    Set produceRs = Nothing
    Set produceRs = conn1.Execute(sql)
    
    If produceRs.Fields("recordCount") < 1 Then
        sql = "insert into yt_produce_scan（product_code,product_name,product_model,product_company_model,process_no,process_name,operator_no,operator_name) " & _
                    " (select yt_product_code.product_code, yt_product.name, yt_product.model,yt_product.company_model,yt_process.process_no,yt_process.process_name, yt_employee.emp_no, yt_employee.name " & _
                    " from yt_product_code, yt_product, yt_process, yt_employee" & _
                    " where yt_product_code.product_id = yt_product.id and yt_product.id = yt_process.product_id and yt_process.employee_id = yt_employee.id " & _
                    " and yt_product_code.product_code = '" & productCode & "' yt_process.process_no = " & processNo
        Set produceRs = conn1.Execute(sql)
    End If
    produceRs.Close
    Set produceRs = Nothing
    sql = "update yt_produce_scan set scan_status=1, scan_time='" & Now & "', update_time='" & Now & "'"
     Set produceRs = conn1.Execute(sql)
     
    produceRs.Close
    Set produceRs = Nothing
            
            
            
        '进行处理
        conn1.Close
        Set conn1 = Nothing
        
        End If
    Next
End Sub

'用多线程处理之
Public Sub readch1()
    Dim strText
    strText = "你好"
  '  MsgBox strText
    Dim objVoice As Object
    Dim colVoice, langCN, langEN, i, cnVoice, enVoice
    Set objVoice = CreateObject("SAPI.SpVoice")
    Set colVoice = objVoice.GetVoices() '获得语音引擎集合
    
   
    objVoice.Volume = 100 '设置音量，0到100，数字越大音量越大
    '得到所需语音引擎的编号
    langCN = "MSSimplifiedChineseVoice" '简体中文
    langEN = "MSSam" '如果安装了TTS Engines 5.1，还可以选择MSMike,MSMary
    For i = 0 To colVoice.count - 1 '选择语音引擎
   
        If Right(colVoice(i).Id, Len(langCN)) = langCN Then cnVoice = i
        If Right(colVoice(i).Id, Len(langEN)) = langEN Then enVoice = i
    Next
    Dim strSource, strCurrent, strSlice, strTemp, strSplitter As String
    Dim strArray() As String
    strSource = strText & " "
    strTemp = ""
    strSplitter = "@@"
    '把strSource中的中英文分开
    For i = 1 To Len(strSource) - 1
        strCurrent = Mid(strSource, i, 1)
        If is_hanzi1(strCurrent) = is_hanzi1(Mid(strSource, i + 1, 1)) Then '如果是中文
            strTemp = strTemp & strCurrent
        Else
            strTemp = strTemp & strCurrent & strSplitter
        End If
    Next
   
    strTemp = Replace(strTemp, "@@ @@", " ") '空字符会被识别为英文，予以纠正
  '  MsgBox strTemp
   
    strArray = Split(strTemp, strSplitter)
    For Each strSlice In strArray
        If Trim(strSlice) = "" Then
            GoTo endfor
        End If
      
     '   If is_hanzi(Mid(strSlice, 1, 1)) Then
            Set objVoice.Voice = colVoice.Item(cnVoice) '设置语音引擎为简体中文
            objVoice.Speak (strSlice)
      '  Else
       '     Set objVoice.Voice = colVoice.Item(enVoice)
        '    objVoice.Speak (strSlice)
        'End If
endfor:
    Next
End Sub
Private Function is_hanzi1(ByVal str_char As String)
    If AscW(str_char) > &H0 And AscW(str_char) < &H800 Then
        is_hanzi1 = False
    Else
        is_hanzi1 = True
    End If
End Function

Public Sub FF2()
  '  MsgBox "hello"
End Sub


Public Sub doSpeak1(ByVal strText As String)
    Dim elements()
    Dim i As Integer
    
     Dim objVoice As Object
    Dim colVoice, langCN, langEN, cnVoice, enVoice
    Set objVoice = CreateObject("SAPI.SpVoice")
  '  Set colVoice = objVoice.GetVoices() '获得语音引擎集合
   objVoice.Speak strText
    
   ' elements = queue.ReadElements
   ' For I = 0 To queue.MaxLen - 1
    '    If Not IsEmpty(elements(I)) Then
           ' MsgBox elements(I)
       '     objVoice.Speak elements(I)
          
    
    
   
      '  End If
   ' Next
   Set objVoice = Nothing
End Sub


Public Function SysPass() As Boolean
     Dim sql As String
     Dim sysRs As ADODB.Recordset
     
   ''  SysPass = True
    '' Exit Function
   
    
     If conn Is Nothing Then
     '连接数据库
        If dbConnOpen() = False Then
            MsgBox "连接数据库失败!"
            End
            Exit Function
        End If
     
    End If
     sql = "select count(*) as recordCount  from yt_sys_info where id = 1"
     Set sysRs = conn.Execute(sql)
     
     If sysRs.Fields("recordCount") < 1 Then
        MsgBox "系统数据不健全，系统无法运行，请联系开发人员", vbCritical
        SysPass = False
        sysRs.Close
        Set sysRs = Nothing
        Exit Function
    End If
    
    
    
    sql = "select *   from yt_sys_info where id = 1"
    sysRs.Close
    Set sysRs = conn.Execute(sql)
    
    
    If UCase(Trim(sysRs.Fields("sys1"))) = UCase(MD5("YES", 32)) Then
        SysPass = True
        sysRs.Close
        Set sysRs = Nothing
        Exit Function
    End If
    
     If UCase(Trim(sysRs.Fields("sys1"))) <> UCase(MD5("NO", 32)) Then
        MsgBox "系统数据不正确，系统无法运行，请联系开发人员", vbCritical
        SysPass = False
        sysRs.Close
        Set sysRs = Nothing
        Exit Function
    End If
    
    Dim firstDate As Date
    Dim curDate As Date
    Dim lastDate As Date
    
    firstDate = sysRs.Fields("sys2")
    lastDate = sysRs.Fields("sys4")
    
    If DateDiff("d", Now, lastDate) > 0 Then
         MsgBox "系统日期不正确，系统无法运行，请联系开发人员", vbCritical
        SysPass = False
        sysRs.Close
        Set sysRs = Nothing
        Exit Function
    End If
    
   
    Dim xDate As Date
    xDate = sysRs.Fields("sys6")
    
  
    
    If DateDiff("d", Now, xDate) < 0 Then
        MsgBox "由于软件开发款未结清，软件暂停服务，请联系开发人员!", vbCritical
         SysPass = False
        sysRs.Close
        Set sysRs = Nothing
        Exit Function
    End If
   
    sql = "update yt_sys_info set sys4= '" & Now & "'  where id = 1"
    sysRs.Close
    Set sysRs = conn.Execute(sql)
    
    
     Set sysRs = Nothing
     SysPass = True
End Function









