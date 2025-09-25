VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaTranscoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Transcoding_MediaTranscoder As String = "Windows.Media.Transcoding.MediaTranscoder"
Private Const IID_IMediaTranscoder As String = "{190c99d2-a0aa-4d34-86bc-eed1b12c2f5b}"
Private Const IID_IMediaTranscoder2 As String = "{40531d74-35e0-4f04-8574-ca8bc4e5a082}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaTranscoder
'    IMediaTranscoder_PutTrimStartTime = 6
'    IMediaTranscoder_GetTrimStartTime = 7
'    IMediaTranscoder_PutTrimStopTime = 8
'    IMediaTranscoder_GetTrimStopTime = 9
    IMediaTranscoder_PutAlwaysReencode = 10
    IMediaTranscoder_GetAlwaysReencode = 11
    IMediaTranscoder_PutHardwareAccelerationEnabled = 12
    IMediaTranscoder_GetHardwareAccelerationEnabled = 13
    IMediaTranscoder_AddAudioEffect = 14
'    IMediaTranscoder_AddAudioEffectWithSettings = 15
    IMediaTranscoder_AddVideoEffect = 16
'    IMediaTranscoder_AddVideoEffectWithSettings = 17
    IMediaTranscoder_ClearEffects = 18
    IMediaTranscoder_PrepareFileTranscodeAsync = 19
'    IMediaTranscoder_PrepareStreamTranscodeAsync = 20

    ' IMediaTranscoder2
'    IMediaTranscoder2_PrepareMediaStreamSourceTranscodeAsync = 6
    IMediaTranscoder2_PutVideoProcessingAlgorithm = 7
    IMediaTranscoder2_GetVideoProcessingAlgorithm = 8

End Enum

' ----==== Variablen ====----
Private m_pIMediaTranscoder As Long
Private m_pIMediaTranscoder2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIMediaTranscoder As Long
    If GetActivateInstance(Windows_Media_Transcoding_MediaTranscoder, _
                           IID_IMediaTranscoder, _
                           pIMediaTranscoder) Then
        Me.Ifc = pIMediaTranscoder
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaTranscoder2)
    Call ReleaseIfc(m_pIMediaTranscoder)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaTranscoder
End Property

Public Property Let Ifc(ByVal pIMediaTranscoder As Long)
    Call ReleaseIfc(m_pIMediaTranscoder2)
    Call ReleaseIfc(m_pIMediaTranscoder)
    m_pIMediaTranscoder = pIMediaTranscoder
    Call QueryIfc(m_pIMediaTranscoder, IID_IMediaTranscoder2, m_pIMediaTranscoder2)
End Property

' IMediaTranscoder
'    IMediaTranscoder_PutTrimStartTime = 6
'    IMediaTranscoder_GetTrimStartTime = 7
'    IMediaTranscoder_PutTrimStopTime = 8
'    IMediaTranscoder_GetTrimStopTime = 9

Public Property Let AlwaysReencode(ByVal Value As Boolean)
    If m_pIMediaTranscoder <> 0& Then
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_PutAlwaysReencode, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get AlwaysReencode() As Boolean
    Dim Ret As Boolean
    If m_pIMediaTranscoder <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_GetAlwaysReencode, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    AlwaysReencode = Ret
End Property

Public Property Let HardwareAccelerationEnabled(ByVal Value As Boolean)
    If m_pIMediaTranscoder <> 0& Then
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_PutHardwareAccelerationEnabled, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get HardwareAccelerationEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaTranscoder <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_GetHardwareAccelerationEnabled, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    HardwareAccelerationEnabled = Ret
End Property

' IMediaTranscoder2
Public Property Let VideoProcessingAlgorithm(ByVal Value As MediaVideoProcessingAlgorithm)
    If m_pIMediaTranscoder2 <> 0& Then
        If Invoke(m_pIMediaTranscoder2, _
                  IMediaTranscoder2_PutVideoProcessingAlgorithm, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get VideoProcessingAlgorithm() As MediaVideoProcessingAlgorithm
    Dim Ret As MediaVideoProcessingAlgorithm
    If m_pIMediaTranscoder2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaTranscoder2, _
                  IMediaTranscoder2_GetVideoProcessingAlgorithm, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    VideoProcessingAlgorithm = Ret
End Property


' ----==== Functions ====----
' IMediaTranscoder
Public Function AddAudioEffect(ByVal activatableClassId As String) As Boolean
    Dim Ret As Boolean
    If m_pIMediaTranscoder <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(activatableClassId)
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_AddAudioEffect, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    AddAudioEffect = Ret
End Function

'    IMediaTranscoder_AddAudioEffectWithSettings = 15

Public Function AddVideoEffect(ByVal activatableClassId As String) As Boolean
    Dim Ret As Boolean
    If m_pIMediaTranscoder <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(activatableClassId)
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_AddVideoEffect, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    AddVideoEffect = Ret
End Function

'    IMediaTranscoder_AddVideoEffectWithSettings = 17

Public Function ClearEffects() As Boolean
    Dim Ret As Boolean
    If m_pIMediaTranscoder <> 0& Then
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_ClearEffects) = S_OK Then
            Ret = True
        End If
    End If
    ClearEffects = Ret
End Function

Public Function PrepareFileTranscodeAsync(ByVal source As StorageFile, _
                                          ByVal destination As StorageFile, _
                                          ByVal profile As MediaEncodingProfile) As PrepareTranscodeResult
    Dim Ret As PrepareTranscodeResult
    If m_pIMediaTranscoder <> 0& And _
       IsNotNothing(source) And _
       IsNotNothing(destination) And _
       IsNotNothing(profile) Then
        Dim pIPrepareTranscodeResult As Long
        If Invoke(m_pIMediaTranscoder, _
                  IMediaTranscoder_PrepareFileTranscodeAsync, _
                  source.Ifc, _
                  destination.Ifc, _
                  profile.Ifc, _
                  VarPtr(pIPrepareTranscodeResult)) = S_OK Then
            If pIPrepareTranscodeResult <> 0& Then
                If Await(pIPrepareTranscodeResult) Then
                    Set Ret = New PrepareTranscodeResult
                    Ret.Ifc = pIPrepareTranscodeResult
                End If
            End If
        End If
    End If
    Set PrepareFileTranscodeAsync = Ret
End Function

'    IMediaTranscoder_PrepareStreamTranscodeAsync = 20

' IMediaTranscoder2
'    IMediaTranscoder2_PrepareMediaStreamSourceTranscodeAsync = 6

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


