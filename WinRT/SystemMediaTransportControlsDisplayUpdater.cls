VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemMediaTransportControlsDisplayUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SystemMediaTransportControlsDisplayUpdater As String = "Windows.Media.SystemMediaTransportControlsDisplayUpdater"
Private Const IID_ISystemMediaTransportControlsDisplayUpdater As String = "{8abbc53e-fa55-4ecf-ad8e-c984e5dd1550}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    
    ' ISystemMediaTransportControlsDisplayUpdater
    ISystemMediaTransportControlsDisplayUpdater_GetType = 6
    ISystemMediaTransportControlsDisplayUpdater_PutType = 7
    ISystemMediaTransportControlsDisplayUpdater_GetAppMediaId = 8
    ISystemMediaTransportControlsDisplayUpdater_PutAppMediaId = 9
    ISystemMediaTransportControlsDisplayUpdater_GetThumbnail = 10
    ISystemMediaTransportControlsDisplayUpdater_PutThumbnail = 11
    ISystemMediaTransportControlsDisplayUpdater_GetMusicProperties = 12
    ISystemMediaTransportControlsDisplayUpdater_GetVideoProperties = 13
    ISystemMediaTransportControlsDisplayUpdater_GetImageProperties = 14
    ISystemMediaTransportControlsDisplayUpdater_CopyFromFileAsync = 15
    ISystemMediaTransportControlsDisplayUpdater_ClearAll = 16
    ISystemMediaTransportControlsDisplayUpdater_Update = 17

End Enum

' ----==== Variablen ====----
Private m_pISystemMediaTransportControlsDisplayUpdater As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISystemMediaTransportControlsDisplayUpdater)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pISystemMediaTransportControlsDisplayUpdater As Long)
    Call ReleaseIfc(m_pISystemMediaTransportControlsDisplayUpdater)
    m_pISystemMediaTransportControlsDisplayUpdater = pISystemMediaTransportControlsDisplayUpdater
End Property

Public Property Get Ifc() As Long
    Ifc = m_pISystemMediaTransportControlsDisplayUpdater
End Property

' ISystemMediaTransportControlsDisplayUpdater
Public Property Get PlaybackType() As MediaPlaybackType
    Dim Ret As MediaPlaybackType
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_GetType, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackType = Ret
End Property

Public Property Let PlaybackType(ByVal value As MediaPlaybackType)
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_PutType, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get AppMediaId() As String
    Dim Ret As String
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_GetAppMediaId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AppMediaId = Ret
End Property

Public Property Let AppMediaId(ByVal value As String)
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_PutAppMediaId, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Thumbnail() As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_GetThumbnail, _
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
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_PutThumbnail, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get MusicProperties() As MusicDisplayProperties
    Dim Ret As MusicDisplayProperties
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        Dim pIMusicDisplayProperties As Long
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_GetMusicProperties, _
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
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        Dim pIVideoDisplayProperties As Long
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_GetVideoProperties, _
                  VarPtr(pIVideoDisplayProperties)) = S_OK Then
            If pIVideoDisplayProperties <> 0& Then
                Set Ret = New VideoDisplayProperties
                Ret.Ifc = pIVideoDisplayProperties
            End If
        End If
    End If
    Set VideoProperties = Ret
End Property

Public Property Get ImageProperties() As ImageDisplayProperties
    Dim Ret As ImageDisplayProperties
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        Dim pIImageDisplayProperties As Long
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_GetImageProperties, _
                  VarPtr(pIImageDisplayProperties)) = S_OK Then
            If pIImageDisplayProperties <> 0& Then
                Set Ret = New ImageDisplayProperties
                Ret.Ifc = pIImageDisplayProperties
            End If
        End If
    End If
    Set ImageProperties = Ret
End Property

' ----==== Functions ====----
' ISystemMediaTransportControlsDisplayUpdater
Public Function CopyFromFileAsync(ByVal PlaybackType As MediaPlaybackType, _
                                  ByVal Source As StorageFile) As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& And _
       IsNotNothing(Source) Then
        Dim operation As Long
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_CopyFromFileAsync, _
                  PlaybackType, _
                  Source.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    CopyFromFileAsync = Ret
End Function

Public Function ClearAll() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_ClearAll) = S_OK Then
            Ret = True
        End If
    End If
    ClearAll = Ret
End Function

Public Function Update() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControlsDisplayUpdater <> 0& Then
        If Invoke(m_pISystemMediaTransportControlsDisplayUpdater, _
                  ISystemMediaTransportControlsDisplayUpdater_Update) = S_OK Then
            Ret = True
        End If
    End If
    Update = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


