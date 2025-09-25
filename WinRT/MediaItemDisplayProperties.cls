VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaItemDisplayProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaItemDisplayProperties As String = "Windows.Media.Playback.MediaItemDisplayProperties"
Private Const IID_IMediaItemDisplayProperties As String = "{1e3c1b48-7097-4384-a217-c1291dfa8c16}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaItemDisplayProperties
    IMediaItemDisplayProperties_GetType = 6
    IMediaItemDisplayProperties_PutType = 7
    IMediaItemDisplayProperties_GetMusicProperties = 8
    IMediaItemDisplayProperties_GetVideoProperties = 9
    IMediaItemDisplayProperties_GetThumbnail = 10
    IMediaItemDisplayProperties_PutThumbnail = 11
    IMediaItemDisplayProperties_ClearAll = 12

End Enum

' ----==== Variablen ====----
Private m_pIMediaItemDisplayProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaItemDisplayProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaItemDisplayProperties
End Property

Public Property Let Ifc(ByVal pIMediaItemDisplayProperties As Long)
    Call ReleaseIfc(m_pIMediaItemDisplayProperties)
    m_pIMediaItemDisplayProperties = pIMediaItemDisplayProperties
End Property

' IMediaItemDisplayProperties
Public Property Get PlaybackType() As MediaPlaybackType
    Dim Ret As MediaPlaybackType
    If m_pIMediaItemDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaItemDisplayProperties, _
                  IMediaItemDisplayProperties_GetType, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackType = Ret
End Property

Public Property Let PlaybackType(ByVal value As MediaPlaybackType)
    If m_pIMediaItemDisplayProperties <> 0& Then
        If Invoke(m_pIMediaItemDisplayProperties, _
                  IMediaItemDisplayProperties_PutType, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get MusicProperties() As MusicDisplayProperties
    Dim Ret As MusicDisplayProperties
    If m_pIMediaItemDisplayProperties <> 0& Then
        Dim pIMusicDisplayProperties As Long
        If Invoke(m_pIMediaItemDisplayProperties, _
                  IMediaItemDisplayProperties_GetMusicProperties, _
                  VarPtr(pIMusicDisplayProperties)) = S_OK Then
            If pIMusicDisplayProperties <> 0& Then
                Set Ret = New MusicDisplayProperties
                Ret.Ifc = pIMusicDisplayProperties
            End If
        End If
    End If
    Set MusicProperties = Ret
End Property

Public Property Get VideoProperties() As VideoDisplayProperties
    Dim Ret As VideoDisplayProperties
    If m_pIMediaItemDisplayProperties <> 0& Then
        Dim pIVideoDisplayProperties As Long
        If Invoke(m_pIMediaItemDisplayProperties, _
                  IMediaItemDisplayProperties_GetVideoProperties, _
                  VarPtr(pIVideoDisplayProperties)) = S_OK Then
            If pIVideoDisplayProperties <> 0& Then
                Set Ret = New VideoDisplayProperties
                Ret.Ifc = pIVideoDisplayProperties
            End If
        End If
    End If
    Set VideoProperties = Ret
End Property

Public Property Get Thumbnail() As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIMediaItemDisplayProperties <> 0& Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIMediaItemDisplayProperties, _
                  IMediaItemDisplayProperties_GetThumbnail, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                Set Ret = New RandomAccessStreamReference
                Ret.Ifc = pIRandomAccessStreamReference
            End If
        End If
    End If
    Set Thumbnail = Ret
End Property

Public Property Let Thumbnail(ByVal value As RandomAccessStreamReference)
    If m_pIMediaItemDisplayProperties <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaItemDisplayProperties, _
                  IMediaItemDisplayProperties_PutThumbnail, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IMediaItemDisplayProperties
Public Function ClearAll() As Boolean
    Dim Ret As Boolean
    If m_pIMediaItemDisplayProperties <> 0& Then
        If Invoke(m_pIMediaItemDisplayProperties, _
                  IMediaItemDisplayProperties_ClearAll) = S_OK Then
            Ret = True
        End If
    End If
    ClearAll = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


