VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechContinuousRecognitionResultGeneratedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechContinuousRecognitionResultGeneratedEventArgs As String = "Windows.Media.SpeechRecognition.SpeechContinuousRecognitionResultGeneratedEventArgs"
Private Const IID_ISpeechContinuousRecognitionResultGeneratedEventArgs As String = "{19091e1e-6e7e-5a46-40fb-76594f786504}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechContinuousRecognitionResultGeneratedEventArgs
    ISpeechContinuousRecognitionResultGeneratedEventArgs_GetResult = 6

End Enum

' ----==== Variablen ====----
Private m_pISpeechContinuousRecognitionResultGeneratedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechContinuousRecognitionResultGeneratedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechContinuousRecognitionResultGeneratedEventArgs
End Property

Public Property Let Ifc(ByVal pISpeechContinuousRecognitionResultGeneratedEventArgs As Long)
    Call ReleaseIfc(m_pISpeechContinuousRecognitionResultGeneratedEventArgs)
    m_pISpeechContinuousRecognitionResultGeneratedEventArgs = pISpeechContinuousRecognitionResultGeneratedEventArgs
End Property

' ISpeechContinuousRecognitionResultGeneratedEventArgs
Public Property Get Result() As SpeechRecognitionResult
    Dim Ret As SpeechRecognitionResult
    If m_pISpeechContinuousRecognitionResultGeneratedEventArgs <> 0& Then
        Dim pISpeechRecognitionResult As Long
        If Invoke(m_pISpeechContinuousRecognitionResultGeneratedEventArgs, _
                  ISpeechContinuousRecognitionResultGeneratedEventArgs_GetResult, _
                  VarPtr(pISpeechRecognitionResult)) = S_OK Then
            If pISpeechRecognitionResult <> 0& Then
                Set Ret = New SpeechRecognitionResult
                Ret.Ifc = pISpeechRecognitionResult
            End If
        End If
    End If
    Set Result = Ret
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


