VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarCodeQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* ************************************************************** *
'*    �������ƣ�Queue.cls
'*    �����ܣ�һ��VBѭ��������
'* ************************************************************** *
'Class Queue '��ASP��WSH��HTML��VBScript�ű���ʹ��ʱ��Ӵ˾�
    Dim m_Queue() '����
    Dim m_RetQueue() '���з���ֵ
    Dim m_ReadIndex As Integer '����ƫ������
    Dim m_WriteIndex As Integer 'Ԫ�ز���ƫ������
    Dim m_MaxLen As Integer '�������ߴ�
    
    Private Sub Class_Initialize()
        m_ReadIndex = 0
        m_WriteIndex = 0
        m_MaxLen = 100 '����Ĭ�ϴ�СΪ100
        ReDim m_Queue(m_MaxLen - 1)
    End Sub
    
    Private Sub Class_Terminate()
        Erase m_Queue
        Erase m_RetQueue
    End Sub
    
    '��ö��д�С
    Public Property Get MaxLen()
        MaxLen = m_MaxLen
    End Property
    
    '���ö��д�С
    Public Property Let MaxLen(ByVal NewValue)
        If NewValue > 0 Then
            m_MaxLen = NewValue
            ReDim m_Queue(m_MaxLen)
        End If
    End Property
    
    '�������Ԫ��
    Public Sub Insert(ByVal v)
        If m_WriteIndex = m_MaxLen Then '�����������������Ȳ����Ԫ��
            m_WriteIndex = 0
            m_ReadIndex = (m_ReadIndex + 1) Mod m_MaxLen
        End If
        m_Queue(m_WriteIndex) = v
        m_WriteIndex = m_WriteIndex + 1
    End Sub
   
    
    'ɾ������Ԫ�أ�������ԭ��ʵ������ɾ�����Ȳ����Ԫ�أ�
    Public Sub Delete()
        Dim vTemp
        m_Queue(m_ReadIndex) = vTemp '���˱�����Ϊδ��ʼ��
        m_ReadIndex = (m_ReadIndex + 1) Mod m_MaxLen
    End Sub
    
    '��ն���
    Public Sub Clear()
        ReDim m_Queue(m_MaxLen)
        m_WriteIndex = 0
        m_ReadIndex = 0
    End Sub
    
    '��ö�������
    Public Property Get GetQueue()
        Dim i, n
        
        n = 0
        ReDim m_RetQueue(m_MaxLen)
        For i = m_ReadIndex To m_MaxLen - 1
            If IsEmpty(m_Queue(i)) Then Exit For
            m_RetQueue(n) = m_Queue(i)
            n = n + 1
        Next
        For i = 0 To m_ReadIndex - 1
            If IsEmpty(m_Queue(i)) Then Exit For
            m_RetQueue(n) = m_Queue(i)
            n = n + 1
        Next
        GetQueue = m_RetQueue
    End Property
    
     'д��Ԫ��
    Public Sub WriteElement(ByVal ele)
        If m_WriteIndex >= m_MaxLen Then '��������
            m_WriteIndex = 0
        End If
        Do
            DoEvents
            If IsEmpty(m_Queue(m_WriteIndex)) Then Exit Do
        Loop
       
        m_Queue(m_WriteIndex) = ele
        m_WriteIndex = m_WriteIndex + 1
    End Sub
    
     '��ö�������
    Public Property Get ReadElements123()
        Dim i, n, x_WriteIndex
        Dim vTemp
       
        n = 0
       
        x_WriteIndex = m_WriteIndex
    '    If m_ReadIndex = x_WriteIndex Then
     '       ReDim m_RetQueue(0)
    '    Else
        ReDim m_RetQueue(m_MaxLen)
        
        If m_ReadIndex = x_WriteIndex Then
            If Not IsEmpty(m_Queue(m_ReadIndex)) Then
                 For i = m_ReadIndex To m_MaxLen - 1
                    'If IsEmpty(m_Queue(i)) Then Exit For
                    m_RetQueue(n) = m_Queue(i)
                    m_Queue(i) = vTemp '���˱�����Ϊδ��ʼ��
                    n = n + 1
                Next
                For i = 0 To m_ReadIndex - 1
                    'If IsEmpty(m_Queue(i)) Then Exit For
                    m_RetQueue(n) = m_Queue(i)
                    m_Queue(i) = vTemp '���˱�����Ϊδ��ʼ��
                    n = n + 1
                Next
            End If
                
       ElseIf m_ReadIndex < x_WriteIndex Then
            For i = m_ReadIndex To x_WriteIndex - 1
                'If IsEmpty(m_Queue(i)) Then Exit For
                m_RetQueue(n) = m_Queue(i)
                m_Queue(i) = vTemp '���˱�����Ϊδ��ʼ��
                n = n + 1
            Next
      Else
            For i = m_ReadIndex To m_MaxLen - 1
                m_RetQueue(n) = m_Queue(i)
                m_Queue(i) = vTemp '���˱�����Ϊδ��ʼ��
                n = n + 1
            Next
            For i = 0 To x_WriteIndex - 1
                m_RetQueue(n) = m_Queue(i)
                m_Queue(i) = vTemp '���˱�����Ϊδ��ʼ��
                n = n + 1
            Next
      End If
        m_ReadIndex = x_WriteIndex
        
        ReadElements = m_RetQueue
    End Property
    
    Public Property Get ReadElements()
        Dim i As Integer, n As Integer, x_WriteIndex As Integer
        Dim vTemp
       
        n = 0
       
        x_WriteIndex = m_WriteIndex
        ReDim m_RetQueue(m_MaxLen)
        For i = m_ReadIndex To m_MaxLen - 1
            m_RetQueue(n) = m_Queue(i)
            m_Queue(i) = vTemp '���˱�����Ϊδ��ʼ��
            n = n + 1
        Next
        For i = 0 To m_ReadIndex - 1
            m_RetQueue(n) = m_Queue(i)
            m_Queue(i) = vTemp '���˱�����Ϊδ��ʼ��
            n = n + 1
        Next
        
        m_ReadIndex = x_WriteIndex
        
        ReadElements = m_RetQueue
    End Property
'End Class '��ASP��WSH��HTML��VBScript�ű���ʹ��ʱ��Ӵ˾�