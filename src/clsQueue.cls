VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim m_coll As Collection

Public Property Get ItemCount() As Long
   ItemCount = m_coll.Count
End Property

Public Sub Add(item)
   m_coll.Add item
End Sub

Public Function Popup(item) As Boolean
   If m_coll.Count = 0 Then
      Popup = False
   Else
      item = m_coll(1)
      m_coll.Remove 1
      Popup = True
   End If
End Function

Public Function EqualItem(index As Long, item) As Boolean
   If index < 1 Or index > m_coll.Count Then
      EqualItem = False
   Else
      If item = m_coll(index) Then
         EqualItem = True
      Else
         EqualItem = False
      End If
   End If
End Function

Public Function InQueue(item) As Boolean
   Dim i
   InQueue = False
   For Each i In m_coll
      If i = item Then
         InQueue = True
         Exit Function
      End If
   Next
End Function

Public Sub Clear()
   Dim i As Long
   For i = 1 To m_coll.Count
      m_coll.Remove 1
   Next
End Sub

Private Sub Class_Initialize()
   Set m_coll = New Collection
End Sub

Private Sub Class_Terminate()
   Dim i As Long
   For i = 1 To m_coll.Count
      m_coll.Remove 1
   Next
   Set m_coll = Nothing
End Sub
