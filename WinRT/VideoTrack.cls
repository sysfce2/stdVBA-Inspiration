VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Core_VideoTrack As String = "Windows.Media.Core.VideoTrack"
Private Const IID_IMediaTrack As String = "{03e1fafc-c931-491a-b46b-c10ee8c256b7}"
Private Const IID_IVideoTrack As String = "{99f3b7f3-e298-4396-bb6a-a51be6a2a20a}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaTrack
    IMediaTrack_GetId = 6
    IMediaTrack_GetLanguage = 7
    IMediaTrack_GetTrackKind = 8
    IMediaTrack_PutLabel = 9
    IMediaTrack_GetLabel = 10
                        
    ' IVideoTrack
    IVideoTrack_AddOpenFailed = 6
    IVideoTrack_RemoveOpenFailed = 7
    IVideoTrack_GetEncodingProperties = 8
    IVideoTrack_GetPlaybackItem = 9
    IVideoTrack_GetName = 10
    IVideoTrack_GetSupportInfo = 11
                        
End Enum

' ----==== Vars ====----
Private m_pIVideoTrack As Long
Private m_pIMediaTrack As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVideoTrack)
    Call ReleaseIfc(m_pIMediaTrack)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaTrack
End Property

Public Property Let Ifc(ByVal pIMediaTrack As Long)
    Call ReleaseIfc(m_pIVideoTrack)
    Call ReleaseIfc(m_pIMediaTrack)
    m_pIMediaTrack = pIMediaTrack
    Call QueryIfc(m_pIMediaTrack, _
                  IID_IVideoTrack, _
                  m_pIVideoTrack)
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

' IVideoTrack
Public Property Get PlaybackItem() As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIVideoTrack <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIVideoTrack, _
                  IVideoTrack_GetPlaybackItem, _
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
    If m_pIVideoTrack <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoTrack, _
                  IVideoTrack_GetName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Name = Ret
End Property

Public Property Get SupportInfo() As VideoTrackSupportInfo
    Dim Ret As VideoTrackSupportInfo
    If m_pIVideoTrack <> 0& Then
        Dim pIVideoTrackSupportInfo As Long
        If Invoke(m_pIVideoTrack, _
                  IVideoTrack_GetSupportInfo, _
                  VarPtr(pIVideoTrackSupportInfo)) = S_OK Then
            If pIVideoTrackSupportInfo <> 0& Then
                Set Ret = New VideoTrackSupportInfo
                Ret.Ifc = pIVideoTrackSupportInfo
            End If
        End If
    End If
    Set SupportInfo = Ret
End Property

' ----==== Functions ====----
' IVideoTrack
Public Function AddOpenFailed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIVideoTrack <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIVideoTrack, _
                  IVideoTrack_AddOpenFailed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddOpenFailed = Ret
End Function

Public Function RemoveOpenFailed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIVideoTrack <> 0& And token <> 0@ Then
        If Invoke(m_pIVideoTrack, _
                  IVideoTrack_RemoveOpenFailed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveOpenFailed = Ret
End Function

Public Function GetEncodingProperties() As VideoEncodingProperties
    Dim Ret As VideoEncodingProperties
    If m_pIVideoTrack <> 0& Then
        Dim pIVideoEncodingProperties As Long
        If Invoke(m_pIVideoTrack, _
                  IVideoTrack_GetEncodingProperties, _
                  VarPtr(pIVideoEncodingProperties)) = S_OK Then
            If pIVideoEncodingProperties <> 0& Then
                Set Ret = New VideoEncodingProperties
                Ret.Ifc = pIVideoEncodingProperties
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


