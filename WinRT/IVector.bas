Attribute VB_Name = "IVector"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----

' ----==== Types ====----
Private Type tInterface
    pVTable As Long
End Type

Private Type tInterface_VTable
    VTable(0 To 17) As Long
End Type

' ----==== Vars ====----
Private m_GUIDs(0) As GUID
Private m_VectorGuid As String
Private m_VectorName As String
Private m_RefCount As Long
Private m_Interface As tInterface
Private m_Collection As Collection
Private m_Interface_VTable As tInterface_VTable

Public Function Implement(ByVal vectorName As String, _
                          ByVal vectorGuid As String) As Long
    m_VectorName = vectorName
    m_VectorGuid = vectorGuid
    m_GUIDs(0) = Str2Guid(vectorGuid)
    Set m_Collection = New Collection
    With m_Interface_VTable
        .VTable(0) = ProcPtr(AddressOf QueryInterface)
        .VTable(1) = ProcPtr(AddressOf AddRef)
        .VTable(2) = ProcPtr(AddressOf Release)
        .VTable(3) = ProcPtr(AddressOf GetIids)
        .VTable(4) = ProcPtr(AddressOf GetRuntimeClassName)
        .VTable(5) = ProcPtr(AddressOf GetTrustLevel)
        .VTable(6) = ProcPtr(AddressOf GetAt)
        .VTable(7) = ProcPtr(AddressOf get_Size)
        .VTable(8) = ProcPtr(AddressOf GetView)
        .VTable(9) = ProcPtr(AddressOf IndexOf)
        .VTable(10) = ProcPtr(AddressOf SetAt)
        .VTable(11) = ProcPtr(AddressOf InsertAt)
        .VTable(12) = ProcPtr(AddressOf RemoveAt)
        .VTable(13) = ProcPtr(AddressOf Append)
        .VTable(14) = ProcPtr(AddressOf RemoveAtEnd)
        .VTable(15) = ProcPtr(AddressOf Clear)
        .VTable(16) = ProcPtr(AddressOf GetMany)
        .VTable(17) = ProcPtr(AddressOf ReplaceAll)
    End With
    With m_Interface
        .pVTable = VarPtr(m_Interface_VTable)
    End With
    m_RefCount = 1
    Implement = VarPtr(m_Interface)
End Function

' ----==== IUnknown ====----
Private Function QueryInterface(ByVal this As Long, _
                                ByRef riid As GUID, _
                                ByRef pvObj As Long) As Long
    Dim lRet As Long
    Select Case UCase$(Guid2Str(riid))
    Case UCase$(IID_IUnknown), UCase$(m_VectorGuid)
        Call AddRef(this)
        pvObj = VarPtr(m_Interface)
        lRet = S_OK
    Case Else
        pvObj = 0&
        lRet = E_NOINTERFACE
    End Select
    QueryInterface = lRet
End Function

Private Function AddRef(ByVal this As Long) As Long
    m_RefCount = m_RefCount + 1
    AddRef = m_RefCount
End Function

Private Function Release(ByVal this As Long) As Long
    m_RefCount = m_RefCount - 1
    
    If m_RefCount = 0 Then
        Call Clear(this)
    End If
    
    Release = m_RefCount
End Function

' ----==== IInspectable ====----
Private Function GetIids(ByVal this As Long, _
                         ByRef iidCount As Long, _
                         ByRef iids As Long) As Long
    iidCount = 1
    iids = VarPtr(m_GUIDs(0))
End Function

Private Function GetRuntimeClassName(ByVal this As Long, _
                                     ByRef className As Long) As Long
    className = CreateWindowsString(m_VectorName)
End Function

Private Function GetTrustLevel(ByVal this As Long, _
                               ByRef Level As Long) As Long
    Level = TrustLevel.BaseTrust
End Function

' ----==== IVector_HSTRING ====----
Private Function GetAt(ByVal this As Long, _
                       ByVal index As Long, _
                       ByRef item As Long) As Long
    If m_Collection.count > 0& And _
       index < m_Collection.count Then
        item = m_Collection.item(index + 1)
    End If
End Function

Private Function get_Size(ByVal this As Long, _
                          ByRef sizeCount As Long) As Long
    sizeCount = m_Collection.count
End Function

Private Function GetView(ByVal this As Long, _
                         ByRef view As Long) As Long
'
End Function

Private Function IndexOf(ByVal this As Long, _
                         ByVal value As Long, _
                         ByRef index As Long, _
                         ByRef found As Long) As Long
    If m_Collection.count > 0& Then
        Dim item As Long
        For item = 1 To m_Collection.count
            If GetWindowsString(m_Collection(item)) = GetWindowsString(value) Then
                found = 1
                Exit For
            End If
        Next
        index = item - 1
    End If
End Function

Private Function SetAt(ByVal this As Long, _
                       ByVal index As Long, _
                       ByVal item As Long) As Long
    If m_Collection.count > 0& And _
       index < m_Collection.count Then
        Call DeleteWindowsString(m_Collection.item(index + 1))
        Call m_Collection.Remove(index + 1)
    End If
    If m_Collection.count > 0& And _
       index < m_Collection.count Then
        Call m_Collection.Add(item, , index + 1)
    Else
        Call m_Collection.Add(item)
    End If
End Function

Private Function InsertAt(ByVal this As Long, _
                          ByVal index As Long, _
                          ByVal item As Long) As Long
    If m_Collection.count > 0& And _
       index < m_Collection.count Then
        Call m_Collection.Add(item, , index + 1)
    Else
        Call m_Collection.Add(item)
    End If
End Function

Private Function RemoveAt(ByVal this As Long, _
                          ByVal index As Long) As Long
    If m_Collection.count > 0& And _
       index < m_Collection.count Then
        Call DeleteWindowsString(m_Collection.item(index + 1))
        Call m_Collection.Remove(index + 1)
    End If
End Function

Private Function Append(ByVal this As Long, _
                        ByVal item As Long) As Long
    Call m_Collection.Add(item)
End Function

Private Function RemoveAtEnd(ByVal this As Long) As Long
    If m_Collection.count > 0& Then
        Call DeleteWindowsString(m_Collection.count)
        Call m_Collection.Remove(m_Collection.count)
    End If
End Function

Private Function Clear(ByVal this As Long) As Long
    If m_Collection.count > 0& Then
        Do While (m_Collection.count > 0&)
            Call DeleteWindowsString(m_Collection.item(1))
            Call m_Collection.Remove(1)
        Loop
        Set m_Collection = Nothing
     End If
End Function

Private Function GetMany(ByVal this As Long, _
                         ByVal startIndex As Long, _
                         ByVal capacity As Long, _
                         ByVal value As Long, _
                         ByVal actual As Long) As Long
End Function

Private Function ReplaceAll(ByVal this As Long, _
                            ByVal count As Long, _
                            ByVal value As Long) As Long
End Function

