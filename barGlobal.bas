Attribute VB_Name = "basGlobal"

'*************************************************************************
'**��    �������ݿ���ز���

'*************************************************************************

Option Explicit

Public loginUser As String '�����½���û���

Public index As Integer

Private Const DB_HOST = "127.0.0.1" ' "we.e-rabits.com" '���ݿ����������
Private Const DB_USER = "root"              '���ݿ��¼�û�
Private Const DB_PASS = "root2015" ''"Erabitsroot"       '���ݿ��¼����
Private Const DB_DATABASE = "ytdb1" '"ytdb"          '���ݿ���

 Private IsConnect As Boolean '������ݿ��Ƿ�����

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

'�Ͽ������ݿ������
Public Sub dbDisConnect()
 Dim rc As Long
  If IsConnect = False Then
   Exit Sub
 End If
 '�ر�����
 conn.Close
 '�ͷ�cnn
 Set conn = Nothing
 IsConnect = False
End Sub

Public Function isDbConnect() As Boolean
    isDbConnect = IsConnect
End Function

'ִ�����ݿ��ѯ���
Public Function dbQueryExt(ByVal TmpSQLstmt As String) As ADODB.Recordset
   
  Dim rst As New ADODB.Recordset '����Rescordset����rst
 
  dbConnOpen '�������ݿ�
 
  Set rst.ActiveConnection = conn '����rst��ActiveConnection����,ָ��������ص����ݿ������
 
  rst.CursorType = adOpenDynamic '�����α�����
 
  rst.LockType = adLockOptimistic '������������
 
  rst.Open TmpSQLstmt '�򿪼�¼��
 
  Set dbQueryExt = rst '���ؼ�¼��
 
  End Function
  
Public Sub setPrint(dlgCommonDialog As Object, docObj As Object)
    On Error Resume Next
  '  If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog            '��ӡ�����öԻ���
        .DialogTitle = "��ӡ"
        .CancelError = True
        .Flags = 1
         Printer.FontSize = dlgCommonDialog.FontSize
'����ӡ�����öԻ������õ������С���ݸ���ӡ��
         .ShowPrinter      '����Ļ����ʾ����ӡ�����öԻ���
         If Err <> MSComDlg.cdlCancel Then
            Printer.FontTransparent = False   '��ʼ����ӡ������Ϊ��͸��
          '  SetPrinterScale Form3    'ƥ���ӡ�������������봰�������
          '  PrintAnywhere Printer        '�ɷ����û���д�Ĵ�ӡ�������������
                                      'ʵ���ַ���ͼ�ε���ʾ
'Printer.NewPage  W         '��ӡ�������ʼ��
'PrintAnywhere Printer     '��ӡ��һҳ������
'Printer.NewPage           '��ӡ�������ʼ��
         '   Printer.EndDoc               '������������ӡ���������
' ����ӡ�հ�ҳ
          '  Printer.KillDoc              'ȡ����ǰ�Ĵ�ӡ����
           'Form3.PrintForm ' ����ʾ����������͵���ӡ��
    docObj.PrintForm
         
    Printer.EndDoc ' ��ʼ��ӡ
     docObj.Hide
        End If
    End With
  End Sub
  
    
Public Function setPrinter(dlgCommonDialog As Object) As Boolean
    On Error Resume Next
  '  If ActiveForm Is Nothing Then Exit Sub
    With dlgCommonDialog            '��ӡ�����öԻ���
        .DialogTitle = "���ô�ӡ��"
        .CancelError = True
        .Flags = 1
         Printer.FontSize = dlgCommonDialog.FontSize
'����ӡ�����öԻ������õ������С���ݸ���ӡ��
         .ShowPrinter      '����Ļ����ʾ����ӡ�����öԻ���
          If Err = MSComDlg.cdlCancel Then
            setPrinter = False
          Else
            Printer.FontTransparent = False   '��ʼ����ӡ������Ϊ��͸��
            setPrinter = True
         End If
    End With
  End Function
  
  Public Sub printDoc(docObj As Object)
    On Error Resume Next

    Printer.FontTransparent = False   '��ʼ����ӡ������Ϊ��͸��
         
    docObj.PrintForm
   
    Printer.EndDoc ' ��ʼ��ӡ
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
     '�������ݿ�
        If dbConnOpen() = False Then
            MsgBox "�������ݿ�ʧ��!"
            End
            Exit Function
        End If
     
    End If
    
    'yt_produce_scan ���Ƿ��Ѵ���produceCode��Ӧ�Ĳ�Ʒ
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
        
    
     'yt_produce_scan ���Ƿ��Ѵ���produceCode��processNo��Ӧ�Ĳ�Ʒ
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
   
   '�Ա���productCode��processNo��Ӧ�Ĳ�Ʒ������ɨ�롱����
    sql = "update yt_produce_scan set scan_status=1, scan_time='" & Now & "', update_time='" & Now & _
                        "'where yt_produce_scan.product_code = '" & productCode & "' and  yt_produce_scan.process_no = " & processNo & ""
     Set produceRs = conn.Execute(sql, recordsAffected)
     
     'productCode��Ӧ�����й����Ƿ���ȫɨ��
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
        
     
     '����ɨ���Ƿ���ȷ
   '   sql = "select count(*) as recordCount from yt_produce_scan " & _
   '                     " where yt_produce_scan.product_code = '" & productCode & "' and  yt_produce_scan.process_no = " & processNo & " and scan_status=1 "
   '   Set produceRs = conn.Execute(sql)
  '    If produceRs.Fields("recordCount") > 0 Then
     If recordsAffected > 0 Then
        saveProduceScan = True
      '  doSpeak "ɨ��ɹ�"
      Else
        saveProduceScan = False
      '  doSpeak "ɨ��ʧ��"
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
     '�������ݿ�
       ' If dbConnOpen() = False Then
       '     MsgBox "�������ݿ�ʧ��!"
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
        sql = "insert into yt_produce_scan��product_code,product_name,product_model,product_company_model,process_no,process_name,operator_no,operator_name) " & _
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
        sql = "insert into yt_produce_scan��product_code,product_name,product_model,product_company_model,process_no,process_name,operator_no,operator_name) " & _
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
            
            
            
        '���д���
        conn1.Close
        Set conn1 = Nothing
        
        End If
    Next
End Sub

'�ö��̴߳���֮
Public Sub readch1()
    Dim strText
    strText = "���"
  '  MsgBox strText
    Dim objVoice As Object
    Dim colVoice, langCN, langEN, i, cnVoice, enVoice
    Set objVoice = CreateObject("SAPI.SpVoice")
    Set colVoice = objVoice.GetVoices() '����������漯��
    
   
    objVoice.Volume = 100 '����������0��100������Խ������Խ��
    '�õ�������������ı��
    langCN = "MSSimplifiedChineseVoice" '��������
    langEN = "MSSam" '�����װ��TTS Engines 5.1��������ѡ��MSMike,MSMary
    For i = 0 To colVoice.count - 1 'ѡ����������
   
        If Right(colVoice(i).Id, Len(langCN)) = langCN Then cnVoice = i
        If Right(colVoice(i).Id, Len(langEN)) = langEN Then enVoice = i
    Next
    Dim strSource, strCurrent, strSlice, strTemp, strSplitter As String
    Dim strArray() As String
    strSource = strText & " "
    strTemp = ""
    strSplitter = "@@"
    '��strSource�е���Ӣ�ķֿ�
    For i = 1 To Len(strSource) - 1
        strCurrent = Mid(strSource, i, 1)
        If is_hanzi1(strCurrent) = is_hanzi1(Mid(strSource, i + 1, 1)) Then '���������
            strTemp = strTemp & strCurrent
        Else
            strTemp = strTemp & strCurrent & strSplitter
        End If
    Next
   
    strTemp = Replace(strTemp, "@@ @@", " ") '���ַ��ᱻʶ��ΪӢ�ģ����Ծ���
  '  MsgBox strTemp
   
    strArray = Split(strTemp, strSplitter)
    For Each strSlice In strArray
        If Trim(strSlice) = "" Then
            GoTo endfor
        End If
      
     '   If is_hanzi(Mid(strSlice, 1, 1)) Then
            Set objVoice.Voice = colVoice.Item(cnVoice) '������������Ϊ��������
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
  '  Set colVoice = objVoice.GetVoices() '����������漯��
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
     '�������ݿ�
        If dbConnOpen() = False Then
            MsgBox "�������ݿ�ʧ��!"
            End
            Exit Function
        End If
     
    End If
     sql = "select count(*) as recordCount  from yt_sys_info where id = 1"
     Set sysRs = conn.Execute(sql)
     
     If sysRs.Fields("recordCount") < 1 Then
        MsgBox "ϵͳ���ݲ���ȫ��ϵͳ�޷����У�����ϵ������Ա", vbCritical
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
        MsgBox "ϵͳ���ݲ���ȷ��ϵͳ�޷����У�����ϵ������Ա", vbCritical
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
         MsgBox "ϵͳ���ڲ���ȷ��ϵͳ�޷����У�����ϵ������Ա", vbCritical
        SysPass = False
        sysRs.Close
        Set sysRs = Nothing
        Exit Function
    End If
    
   
    Dim xDate As Date
    xDate = sysRs.Fields("sys6")
    
  
    
    If DateDiff("d", Now, xDate) < 0 Then
        MsgBox "�������������δ���壬�����ͣ��������ϵ������Ա!", vbCritical
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









