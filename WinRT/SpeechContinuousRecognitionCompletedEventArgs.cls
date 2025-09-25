VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechContinuousRecognitionCompletedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechContinuousRecognitionCompletedEventArgs As String = "Windows.Media.SpeechRecognition.SpeechContinuousRecognitionCompletedEventArgs"
Private Const IID_ISpeechContinuousRecognitionResultGeneratedEventArgs As String = "{e3d069bb-e30c-5e18-424b-7fbe81f8fbd0}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechContinuousRecognitionCompletedEventArgs
    ISpeechContinuousRecognitionCompletedEventArgs_GetStatus = 6

End Enum

' ----==== Variablen ====----
Private m_pISpeechContinuousRecognitionCompletedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechContinuousRecognitionCompletedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechContinuousRecognitionCompletedEventArgs
End Property

Public Property Let Ifc(ByVal pISpeechContinuousRecognitionCompletedEventArgs As Long)
    Call ReleaseIfc(m_pISpeechContinuousRecognitionCompletedEventArgs)
    m_pISpeechContinuousRecognitionCompletedEventArgs = pISpeechContinuousRecognitionCompletedEventArgs
End Property

' ISpeechContinuousRecognitionCompletedEventArgs
Public Property Get Status() As SpeechRecognitionResultStatus
    Dim Ret As SpeechRecognitionResultStatus
    If m_pISpeechContinuousRecognitionCompletedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechContinuousRecognitionCompletedEventArgs, _
                  ISpeechContinuousRecognitionCompletedEventArgs_GetStatus, _
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


