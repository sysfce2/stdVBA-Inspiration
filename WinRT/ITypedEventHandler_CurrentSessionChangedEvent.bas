Attribute VB_Name = "ITEH_CurrentSessionChangedEvent"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const ITypedEventHandler_Windows_Media_Control_GlobalSystemMediaTransportControlsSessionManager_CurrentSessionChangedEventArgs As String = "{228bd0ed-1fa2-5e9b-a6ec-42566173103b}"

' ----==== Types ====----
Private Type tInterface
    pVTable As Long
End Type

Private Type tInterface_VTable
    VTable(0 To 3) As Long
End Type

' ----==== Variablen ====----
Private m_RefCount As Long
Private m_Object As Object
Private m_Interface As tInterface
Private m_Interface_VTable As tInterface_VTable

Public Function Create(ByVal obj As Object) As Long
    Set m_Object = obj
    With m_Interface_VTable
        .VTable(0) = ProcPtr(AddressOf QueryInterface)
        .VTable(1) = ProcPtr(AddressOf AddRef)
        .VTable(2) = ProcPtr(AddressOf Release)
        .VTable(3) = ProcPtr(AddressOf Invoke)
    End With
    With m_Interface
        .pVTable = VarPtr(m_Interface_VTable)
    End With
    Create = VarPtr(m_Interface)
End Function

Private Function QueryInterface(ByVal this As Long, _
                                ByRef riid As GUID, _
                                ByRef pvObj As Long) As Long
    Dim lRet As Long
    Select Case UCase$(Guid2Str(riid))
    Case UCase$(IID_IUnknown), UCase$(ITypedEventHandler_Windows_Media_Control_GlobalSystemMediaTransportControlsSessionManager_CurrentSessionChangedEventArgs)
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
    Release = m_RefCount
End Function

Private Function Invoke(ByVal this As Long, _
                        ByVal sender As Long, _
                        ByVal args As Long) As Long
    On Error GoTo PROC_ERR
    Call m_Object.CurrentSessionChangedEvent(sender, args)
PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "Error: The object '" & m_Object.Name & "' must contain the following function: " & _
                "Public Sub CurrentSessionChangedEvent(ByVal sender As Long, ByVal args As Long)"
    Resume PROC_EXIT
End Function

