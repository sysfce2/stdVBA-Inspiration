VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Core_AudioTrack As String = "Windows.Media.Core.AudioTrack"
Private Const IID_IMediaTrack As String = "{03e1fafc-c931-491a-b46b-c10ee8c256b7}"
Private Const IID_IAudioTrack As String = "{f23b6e77-3ef7-40de-b943-068b1321701d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaTrack
    IMediaTrack_GetId = 6
    IMediaTrack_GetLanguage = 7
    IMediaTrack_GetTrackKind = 8
    IMediaTrack_PutLabel = 9
    IMediaTrack_GetLabel = 10
                        
    ' IAudioTrack
    IAudioTrack_AddOpenFailed = 6
    IAudioTrack_RemoveOpenFailed = 7
    IAudioTrack_GetEncodingProperties = 8
    IAudioTrack_GetPlaybackItem = 9
    IAudioTrack_GetName = 10
    IAudioTrack_GetSupportInfo = 11
                        
End Enum

' ----==== Vars ====----
Private m_pIAudioTrack As Long
Private m_pIMediaTrack As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAudioTrack)
    Call ReleaseIfc(m_pIMediaTrack)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaTrack
End Property

Public Property Let Ifc(ByVal pIMediaTrack As Long)
    Call ReleaseIfc(m_pIAudioTrack)
    Call ReleaseIfc(m_pIMediaTrack)
    m_pIMediaTrack = pIMediaTrack
    Call QueryIfc(m_pIMediaTrack, _
                  IID_IAudioTrack, _
                  m_pIAudioTrack)
End Property

' IMediaTrack
Public Property Get Id() As String
    Dim Ret As String
    If m_pIMediaTrack <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaTrack, _
                  IMediaTrack_GetId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Id = Ret
End Property

Public Property Get Language() As String
    Dim Ret As String
    If m_pIMediaTrack <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaTrack, _
                  IMediaTrack_GetLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Language = Ret
End Property

Public Property Get TrackKind() As MediaTrackKind
    Dim Ret As MediaTrackKind
    If m_pIMediaTrack <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaTrack, _
                  IMediaTrack_GetTrackKind, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TrackKind = Ret
End Property

Public Property Let Label(ByVal value As String)
    Dim Ret As String
    If m_pIMediaTrack <> 0& Then
        Dim hValue As Long
        hValue = CreateWindowsString(value)
        If Invoke(m_pIMediaTrack, _
                  IMediaTrack_PutLabel, _
                  hValue) = S_OK Then
        End If
        Call DeleteWindowsString(hValue)
    End If
End Property

Public Property Get Label() As String
    Dim Ret As String
    If m_pIMediaTrack <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaTrack, _
                  IMediaTrack_GetLabel, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Label = Ret
End Property

' IAudioTrack
Public Property Get PlaybackItem() As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIAudioTrack <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIAudioTrack, _
                  IAudioTrack_GetPlaybackItem, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set PlaybackItem = Ret
End Property

Public Property Get Name() As String
    Dim Ret As String
    If m_pIAudioTrack <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioTrack, _
                  IAudioTrack_GetName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Name = Ret
End Property

Public Property Get SupportInfo() As AudioTrackSupportInfo
    Dim Ret As AudioTrackSupportInfo
    If m_pIAudioTrack <> 0& Then
        Dim pIAudioTrackSupportInfo As Long
        If Invoke(m_pIAudioTrack, _
                  IAudioTrack_GetSupportInfo, _
                  VarPtr(pIAudioTrackSupportInfo)) = S_OK Then
            If pIAudioTrackSupportInfo <> 0& Then
                Set Ret = New AudioTrackSupportInfo
                Ret.Ifc = pIAudioTrackSupportInfo
            End If
        End If
    End If
    Set SupportInfo = Ret
End Property

' ----==== Functions ====----
' IAudioTrack
Public Function AddOpenFailed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAudioTrack <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAudioTrack, _
                  IAudioTrack_AddOpenFailed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddOpenFailed = Ret
End Function

Public Function RemoveOpenFailed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAudioTrack <> 0& And token <> 0@ Then
        If Invoke(m_pIAudioTrack, _
                  IAudioTrack_RemoveOpenFailed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveOpenFailed = Ret
End Function

Public Function GetEncodingProperties() As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioTrack <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioTrack, _
                  IAudioTrack_GetEncodingProperties, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set GetEncodingProperties = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

