VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechContinuousRecognitionSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechContinuousRecognitionSession As String = "Windows.Media.SpeechRecognition.SpeechContinuousRecognitionSession"
Private Const IID_ISpeechContinuousRecognitionSession As String = "{6a213c04-6614-49f8-99a2-b5e9b3a085c8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechContinuousRecognitionSession
    ISpeechContinuousRecognitionSession_GetAutoStopSilenceTimeout = 6
    ISpeechContinuousRecognitionSession_PutAutoStopSilenceTimeout = 7
    ISpeechContinuousRecognitionSession_StartAsync = 8
    ISpeechContinuousRecognitionSession_StartWithModeAsync = 9
    ISpeechContinuousRecognitionSession_StopAsync = 10
    ISpeechContinuousRecognitionSession_CancelAsync = 11
    ISpeechContinuousRecognitionSession_PauseAsync = 12
    ISpeechContinuousRecognitionSession_Resume = 13
    ISpeechContinuousRecognitionSession_AddCompleted = 14
    ISpeechContinuousRecognitionSession_RemoveCompleted = 15
    ISpeechContinuousRecognitionSession_AddResultGenerated = 16
    ISpeechContinuousRecognitionSession_RemoveResultGenerated = 17

End Enum

' ----==== Variablen ====----
Private m_pISpeechContinuousRecognitionSession As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechContinuousRecognitionSession)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechContinuousRecognitionSession
End Property

Public Property Let Ifc(ByVal pISpeechContinuousRecognitionSession As Long)
    Call ReleaseIfc(m_pISpeechContinuousRecognitionSession)
    m_pISpeechContinuousRecognitionSession = pISpeechContinuousRecognitionSession
End Property

' ISpeechContinuousRecognitionSession
Public Property Get AutoStopSilenceTimeout() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        Dim value As Currency
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_GetAutoStopSilenceTimeout, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set AutoStopSilenceTimeout = Ret
End Property

Public Property Let AutoStopSilenceTimeout(ByVal value As TimeSpan)
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_PutAutoStopSilenceTimeout, _
                  value.TimeSpanValue) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' ISpeechContinuousRecognitionSession
Public Function StartAsync() As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_StartAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    StartAsync = Ret
End Function

Public Function StartWithModeAsync(ByVal Mode As SpeechContinuousRecognitionMode) As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_StartWithModeAsync, _
                  Mode, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    StartWithModeAsync = Ret
End Function

Public Function StopAsync() As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_StopAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    StopAsync = Ret
End Function

Public Function CancelAsync() As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_CancelAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    CancelAsync = Ret
End Function

Public Function PauseAsync() As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_PauseAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    PauseAsync = Ret
End Function

Public Function ResumeSession() As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_Resume) = S_OK Then
            Ret = True
        End If
    End If
    ResumeSession = Ret
End Function

Public Function AddCompleted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISpeechContinuousRecognitionSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_AddCompleted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddCompleted = Ret
End Function

Public Function RemoveCompleted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& And token <> 0@ Then
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_RemoveCompleted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveCompleted = Ret
End Function

Public Function AddResultGenerated(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISpeechContinuousRecognitionSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_AddResultGenerated, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddResultGenerated = Ret
End Function

Public Function RemoveResultGenerated(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISpeechContinuousRecognitionSession <> 0& And token <> 0@ Then
        If Invoke(m_pISpeechContinuousRecognitionSession, _
                  ISpeechContinuousRecognitionSession_RemoveResultGenerated, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveResultGenerated = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


