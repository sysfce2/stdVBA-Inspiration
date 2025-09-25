VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognizerStateChangedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognizerStateChangedEventArgs As String = "Windows.Media.SpeechRecognition.SpeechRecognizerStateChangedEventArgs"
Private Const IID_ISpeechContinuousRecognitionResultGeneratedEventArgs As String = "{563d4f09-ba03-4bad-ad81-ddc6c4dab0c3}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognizerStateChangedEventArgs
    ISpeechRecognizerStateChangedEventArgs_GetState = 6

End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognizerStateChangedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognizerStateChangedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognizerStateChangedEventArgs
End Property

Public Property Let Ifc(ByVal pISpeechRecognizerStateChangedEventArgs As Long)
    Call ReleaseIfc(m_pISpeechRecognizerStateChangedEventArgs)
    m_pISpeechRecognizerStateChangedEventArgs = pISpeechRecognizerStateChangedEventArgs
End Property

' ISpeechRecognizerStateChangedEventArgs
Public Property Get State() As SpeechRecognizerState
    Dim Ret As SpeechRecognizerState
    If m_pISpeechRecognizerStateChangedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognizerStateChangedEventArgs, _
                  ISpeechRecognizerStateChangedEventArgs_GetState, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    State = Ret
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


