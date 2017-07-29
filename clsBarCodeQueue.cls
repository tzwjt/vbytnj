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
'*    程序名称：Queue.cls
'*    程序功能：一个VB循环队列类
'* ************************************************************** *
'Class Queue '在ASP、WSH、HTML的VBScript脚本中使用时需加此句
    Dim m_Queue() '队列
    Dim m_RetQueue() '队列返回值
    Dim m_ReadIndex As Integer '列首偏移索引
    Dim m_WriteIndex As Integer '元素插入偏移索引
    Dim m_MaxLen As Integer '队列最大尺寸
    
    Private Sub Class_Initialize()
        m_ReadIndex = 0
        m_WriteIndex = 0
        m_MaxLen = 100 '队列默认大小为100
        ReDim m_Queue(m_MaxLen - 1)
    End Sub
    
    Private Sub Class_Terminate()
        Erase m_Queue
        Erase m_RetQueue
    End Sub
    
    '获得队列大小
    Public Property Get MaxLen()
        MaxLen = m_MaxLen
    End Property
    
    '设置队列大小
    Public Property Let MaxLen(ByVal NewValue)
        If NewValue > 0 Then
            m_MaxLen = NewValue
            ReDim m_Queue(m_MaxLen)
        End If
    End Property
    
    '插入队列元素
    Public Sub Insert(ByVal v)
        If m_WriteIndex = m_MaxLen Then '队列已满，挤掉最先插入的元素
            m_WriteIndex = 0
            m_ReadIndex = (m_ReadIndex + 1) Mod m_MaxLen
        End If
        m_Queue(m_WriteIndex) = v
        m_WriteIndex = m_WriteIndex + 1
    End Sub
   
    
    '删除队列元素（按队列原则，实际上是删除最先插入的元素）
    Public Sub Delete()
        Dim vTemp
        m_Queue(m_ReadIndex) = vTemp '将此变量置为未初始化
        m_ReadIndex = (m_ReadIndex + 1) Mod m_MaxLen
    End Sub
    
    '清空队列
    Public Sub Clear()
        ReDim m_Queue(m_MaxLen)
        m_WriteIndex = 0
        m_ReadIndex = 0
    End Sub
    
    '获得队列数组
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
    
     '写入元素
    Public Sub WriteElement(ByVal ele)
        If m_WriteIndex >= m_MaxLen Then '队列已满
            m_WriteIndex = 0
        End If
        Do
            DoEvents
            If IsEmpty(m_Queue(m_WriteIndex)) Then Exit Do
        Loop
       
        m_Queue(m_WriteIndex) = ele
        m_WriteIndex = m_WriteIndex + 1
    End Sub
    
     '获得队列数组
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
                    m_Queue(i) = vTemp '将此变量置为未初始化
                    n = n + 1
                Next
                For i = 0 To m_ReadIndex - 1
                    'If IsEmpty(m_Queue(i)) Then Exit For
                    m_RetQueue(n) = m_Queue(i)
                    m_Queue(i) = vTemp '将此变量置为未初始化
                    n = n + 1
                Next
            End If
                
       ElseIf m_ReadIndex < x_WriteIndex Then
            For i = m_ReadIndex To x_WriteIndex - 1
                'If IsEmpty(m_Queue(i)) Then Exit For
                m_RetQueue(n) = m_Queue(i)
                m_Queue(i) = vTemp '将此变量置为未初始化
                n = n + 1
            Next
      Else
            For i = m_ReadIndex To m_MaxLen - 1
                m_RetQueue(n) = m_Queue(i)
                m_Queue(i) = vTemp '将此变量置为未初始化
                n = n + 1
            Next
            For i = 0 To x_WriteIndex - 1
                m_RetQueue(n) = m_Queue(i)
                m_Queue(i) = vTemp '将此变量置为未初始化
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
            m_Queue(i) = vTemp '将此变量置为未初始化
            n = n + 1
        Next
        For i = 0 To m_ReadIndex - 1
            m_RetQueue(n) = m_Queue(i)
            m_Queue(i) = vTemp '将此变量置为未初始化
            n = n + 1
        Next
        
        m_ReadIndex = x_WriteIndex
        
        ReadElements = m_RetQueue
    End Property
'End Class '在ASP、WSH、HTML的VBScript脚本中使用时需加此句
