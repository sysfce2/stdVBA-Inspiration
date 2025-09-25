Attribute VB_Name = "ITEH_VolumeChanged"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const ITypedEventHandler_Windows_Media_Playback_MediaPlayer_IInspectable = "{f1a6a51e-d078-5c40-ba3f-348870ba5c87}"

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

Public Sub Destroy()
    Set m_Object = Nothing
End Sub

Private Function QueryInterface(ByVal this As Long, _
                                ByRef riid As GUID, _
                                ByRef pvObj As Long) As Long
    Dim lRet As Long
    Select Case UCase$(Guid2Str(riid))
    Case UCase$(IID_IUnknown), UCase$(ITypedEventHandler_Windows_Media_Playback_MediaPlayer_IInspectable)
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
    If Not m_Object Is Nothing Then
        Call m_Object.VolumeChangedEvent(sender, args)
    End If
PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "Error: The object '" & m_Object.Name & "' must contain the following function: " & _
                "Public Sub VolumeChangedEvent(ByVal sender As Long, ByVal args As long)"
    Resume PROC_EXIT
End Function

