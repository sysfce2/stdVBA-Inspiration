VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioTrackSupportInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Core_AudioTrackSupportInfo As String = "Windows.Media.Core.AudioTrackSupportInfo"
Private Const IID_IAudioTrackSupportInfo As String = "{178beff7-cc39-44a6-b951-4a5653f073fa}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAudioTrackSupportInfo
    IAudioTrackSupportInfo_GetDecoderStatus = 6
    IAudioTrackSupportInfo_GetDegradation = 7
    IAudioTrackSupportInfo_GetDegradationReason = 8
    IAudioTrackSupportInfo_GetMediaSourceStatus = 9
                        
End Enum

' ----==== Vars ====----
Private m_pIAudioTrackSupportInfo As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAudioTrackSupportInfo)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIAudioTrackSupportInfo
End Property

Public Property Let Ifc(ByVal pIAudioTrackSupportInfo As Long)
    Call ReleaseIfc(m_pIAudioTrackSupportInfo)
    m_pIAudioTrackSupportInfo = pIAudioTrackSupportInfo
End Property

' IAudioTrackSupportInfo
Public Property Get DecoderStatus() As MediaDecoderStatus
    Dim Ret As MediaDecoderStatus
    If m_pIAudioTrackSupportInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioTrackSupportInfo, _
                  IAudioTrackSupportInfo_GetDecoderStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DecoderStatus = Ret
End Property

Public Property Get Degradation() As AudioDecoderDegradation
    Dim Ret As AudioDecoderDegradation
    If m_pIAudioTrackSupportInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioTrackSupportInfo, _
                  IAudioTrackSupportInfo_GetDegradation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Degradation = Ret
End Property

Public Property Get DegradationReason() As AudioDecoderDegradationReason
    Dim Ret As AudioDecoderDegradationReason
    If m_pIAudioTrackSupportInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioTrackSupportInfo, _
                  IAudioTrackSupportInfo_GetDegradationReason, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DegradationReason = Ret
End Property

Public Property Get MediaSourceStatus() As MediaSourceStatus
    Dim Ret As MediaSourceStatus
    If m_pIAudioTrackSupportInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioTrackSupportInfo, _
                  IAudioTrackSupportInfo_GetMediaSourceStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    MediaSourceStatus = Ret
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


