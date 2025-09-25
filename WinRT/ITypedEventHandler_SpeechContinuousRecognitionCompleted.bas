Attribute VB_Name = "ITEH_SpeechContinuousRecognitionCompleted"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const ITypedEventHandler_Windows_Media_SpeechRecognition_SpeechContinuousRecognitionCompletedEventArgs = "{8103c018-7952-59f9-9f41-23b17d6e452d}"

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
    Case UCase$(IID_IUnknown), UCase$(ITypedEventHandler_Windows_Media_SpeechRecognition_SpeechContinuousRecognitionCompletedEventArgs)
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
        Call m_Object.SpeechContinuousRecognitionCompletedEvent(sender, args)
    End If
PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "Error: The object '" & m_Object.Name & "' must contain the following function: " & _
                "Public Sub SpeechContinuousRecognitionCompletedEvent(ByVal sender As Long, ByVal args As long)"
    Resume PROC_EXIT
End Function

