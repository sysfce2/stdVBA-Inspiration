VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognizerTimeouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognizerTimeouts As String = "Windows.Media.SpeechRecognition.SpeechRecognizerTimeouts"
Private Const IID_ISpeechRecognizerTimeouts As String = "{2ef76fca-6a3c-4dca-a153-df1bc88a79af}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognizerTimeouts
    ISpeechRecognizerTimeouts_GetInitialSilenceTimeout = 6
    ISpeechRecognizerTimeouts_PutInitialSilenceTimeout = 7
    ISpeechRecognizerTimeouts_GetEndSilenceTimeout = 8
    ISpeechRecognizerTimeouts_PutEndSilenceTimeout = 9
    ISpeechRecognizerTimeouts_GetBabbleTimeout = 10
    ISpeechRecognizerTimeouts_PutBabbleTimeout = 11

End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognizerTimeouts As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognizerTimeouts)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognizerTimeouts
End Property

Public Property Let Ifc(ByVal pISpeechRecognizerTimeouts As Long)
    Call ReleaseIfc(m_pISpeechRecognizerTimeouts)
    m_pISpeechRecognizerTimeouts = pISpeechRecognizerTimeouts
End Property

' ISpeechRecognizerTimeouts
Public Property Get InitialSilenceTimeout() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISpeechRecognizerTimeouts <> 0& Then
        Dim value As Currency
        If Invoke(m_pISpeechRecognizerTimeouts, _
                  ISpeechRecognizerTimeouts_GetInitialSilenceTimeout, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set InitialSilenceTimeout = Ret
End Property

Public Property Let InitialSilenceTimeout(ByVal value As TimeSpan)
    If m_pISpeechRecognizerTimeouts <> 0& Then
        If Invoke(m_pISpeechRecognizerTimeouts, _
                  ISpeechRecognizerTimeouts_PutInitialSilenceTimeout, _
                  value.TimeSpanValue) = S_OK Then
        End If
    End If
End Property

Public Property Get EndSilenceTimeout() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISpeechRecognizerTimeouts <> 0& Then
        Dim value As Currency
        If Invoke(m_pISpeechRecognizerTimeouts, _
                  ISpeechRecognizerTimeouts_GetEndSilenceTimeout, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set EndSilenceTimeout = Ret
End Property

Public Property Let EndSilenceTimeout(ByVal value As TimeSpan)
    If m_pISpeechRecognizerTimeouts <> 0& Then
        If Invoke(m_pISpeechRecognizerTimeouts, _
                  ISpeechRecognizerTimeouts_PutEndSilenceTimeout, _
                  value.TimeSpanValue) = S_OK Then
        End If
    End If
End Property

Public Property Get BabbleTimeout() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISpeechRecognizerTimeouts <> 0& Then
        Dim value As Currency
        If Invoke(m_pISpeechRecognizerTimeouts, _
                  ISpeechRecognizerTimeouts_GetBabbleTimeout, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set BabbleTimeout = Ret
End Property

Public Property Let BabbleTimeout(ByVal value As TimeSpan)
    If m_pISpeechRecognizerTimeouts <> 0& Then
        If Invoke(m_pISpeechRecognizerTimeouts, _
                  ISpeechRecognizerTimeouts_PutBabbleTimeout, _
                  value.TimeSpanValue) = S_OK Then
        End If
    End If
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

