VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognitionCompilationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognitionCompilationResult As String = "Windows.Media.SpeechRecognition.SpeechRecognitionCompilationResult"
Private Const IID_ISpeechRecognitionCompilationResult As String = "{407e6c5d-6ac7-4da4-9cc1-2fce32cf7489}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognitionCompilationResult
    ISpeechRecognitionCompilationResult_GetStatus = 6

End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognitionCompilationResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognitionCompilationResult)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognitionCompilationResult
End Property

Public Property Let Ifc(ByVal pISpeechRecognitionCompilationResult As Long)
    Call ReleaseIfc(m_pISpeechRecognitionCompilationResult)
    m_pISpeechRecognitionCompilationResult = pISpeechRecognitionCompilationResult
End Property

' ISpeechRecognitionCompilationResult
Public Property Get Status() As SpeechRecognitionResultStatus
    Dim Ret As SpeechRecognitionResultStatus
    If m_pISpeechRecognitionCompilationResult <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionCompilationResult, _
                  ISpeechRecognitionCompilationResult_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


