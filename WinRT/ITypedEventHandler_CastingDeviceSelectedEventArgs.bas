Attribute VB_Name = "ITEH_CastingDeviceSelectedEvent"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const ITypedEventHandler_Windows_Media_Casting_CastingDeviceSelectedEventArgs = "{b3655b33-c4ad-5f4c-a187-b2e4c770a16b}"

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
    Case UCase$(IID_IUnknown), UCase$(ITypedEventHandler_Windows_Media_Casting_CastingDeviceSelectedEventArgs)
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
    If args <> 0& Then
        Dim CastingDeviceSelectedEventArgs As New CastingDeviceSelectedEventArgs
        CastingDeviceSelectedEventArgs.Ifc = args
        Call m_Object.CastingDeviceSelectedEvent(sender, CastingDeviceSelectedEventArgs)
    End If
PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "Error: The object '" & m_Object.Name & "' must contain the following function: " & _
                "Public Sub CastingDeviceSelectedEvent(ByVal sender As Long, ByVal args As CastingDeviceSelectedEventArgs)"
    Resume PROC_EXIT
End Function

