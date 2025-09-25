VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrepareTranscodeResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Transcoding_PrepareTranscodeResult As String = "Windows.Media.Transcoding.PrepareTranscodeResult"
Private Const IID_IPrepareTranscodeResult As String = "{05f25dce-994f-4a34-9d68-97ccce1730d6}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IPrepareTranscodeResult
    IPrepareTranscodeResult_GetCanTranscode = 6
    IPrepareTranscodeResult_GetFailureReason = 7
    IPrepareTranscodeResult_TranscodeAsync = 8

End Enum

' ----==== Variablen ====----
Private m_pIPrepareTranscodeResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPrepareTranscodeResult)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPrepareTranscodeResult
End Property

Public Property Let Ifc(ByVal pIPrepareTranscodeResult As Long)
    Call ReleaseIfc(m_pIPrepareTranscodeResult)
    m_pIPrepareTranscodeResult = pIPrepareTranscodeResult
End Property

' IPrepareTranscodeResult
Public Property Get CanTranscode() As Boolean
    Dim Ret As Boolean
    If m_pIPrepareTranscodeResult <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPrepareTranscodeResult, _
                  IPrepareTranscodeResult_GetCanTranscode, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    CanTranscode = Ret
End Property

Public Property Get FailureReason() As TranscodeFailureReason
    Dim Ret As TranscodeFailureReason
    If m_pIPrepareTranscodeResult <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPrepareTranscodeResult, _
                  IPrepareTranscodeResult_GetFailureReason, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    FailureReason = Ret
End Property

' ----==== Functions ====----
' IPrepareTranscodeResult
Public Function TranscodeAsync() As AsyncActionWithProgress_Double
    Dim Ret As AsyncActionWithProgress_Double
    If m_pIPrepareTranscodeResult <> 0& Then
        Dim pIAsyncActionWithProgress_Double As Long
        If Invoke(m_pIPrepareTranscodeResult, _
                  IPrepareTranscodeResult_TranscodeAsync, _
                  VarPtr(pIAsyncActionWithProgress_Double)) = S_OK Then
            If pIAsyncActionWithProgress_Double <> 0& Then
                Set Ret = New AsyncActionWithProgress_Double
                Ret.Ifc = pIAsyncActionWithProgress_Double
            End If
        End If
    End If
    Set TranscodeAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


