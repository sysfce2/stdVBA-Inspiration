VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaPlayer As String = "Windows.Media.Playback.MediaPlayer"
Private Const IID_IMediaPlayer As String = "{381a83cb-6fff-499b-8d64-2885dfc1249e}"
Private Const IID_IMediaPlayer2 As String = "{3c841218-2123-4fc5-9082-2f883f77bdf5}"
Private Const IID_IMediaPlayer3 As String = "{ee0660da-031b-4feb-bd9b-92e0a0a8d299}"
Private Const IID_IMediaPlayer4 As String = "{80035db0-7448-4770-afcf-2a57450914c5}"
Private Const IID_IMediaPlayer5 As String = "{cfe537fd-f86a-4446-bf4d-c8e792b7b4b3}"
Private Const IID_IMediaPlayer6 As String = "{e0caa086-ae65-414c-b010-8bc55f00e692}"
Private Const IID_IMediaPlayer7 As String = "{5d1dc478-4500-4531-b3f4-777a71491f7f}"
Private Const IID_IMediaPlayerSource As String = "{bd4f8897-1423-4c3e-82c5-0fb1af94f715}"
Private Const IID_IMediaPlayerSource2 As String = "{82449b9f-7322-4c0b-b03b-3e69a48260c5}"
Private Const IID_IMediaPlayerEffects As String = "{85a1deda-cab6-4cc0-8be3-6035f4de2591}"
Private Const IID_IMediaPlayerEffects2 As String = "{fa419a79-1bbe-46c5-ae1f-8ee69fb3c2c7}"

' * Class implements the following interface:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaPlayer
    IMediaPlayer_GetAutoPlay = 6
    IMediaPlayer_PutAutoPlay = 7
'    DEPRECATED ("Use PlaybackSession.NaturalDuration instead of NaturalDuration.  For more info, see MSDN.")
'    IMediaPlayer_GetNaturalDuration = 8
'    DEPRECATED ("Use PlaybackSession.Position instead of Position.  For more info, see MSDN.")
'    IMediaPlayer_GetPosition = 9
'    DEPRECATED ("Use PlaybackSession.Position instead of Position.  For more info, see MSDN.")
'    IMediaPlayer_PutPosition = 10
'    DEPRECATED ("Use PlaybackSession.BufferingProgress instead of BufferingProgress.  For more info, see MSDN.")
'    IMediaPlayer_GetBufferingProgress = 11
'    DEPRECATED ("Use PlaybackSession.State instead of CurrentState.  For more info, see MSDN.")
'    IMediaPlayer_GetCurrentState = 12
'    DEPRECATED ("Use PlaybackSession.CanSeek instead of CanSeek.  For more info, see MSDN.")
'    IMediaPlayer_GetCanSeek = 13
'    DEPRECATED ("Use PlaybackSession.CanPause instead of CanPause.  For more info, see MSDN.")
'    IMediaPlayer_GetCanPause = 14
    IMediaPlayer_GetIsLoopingEnabled = 15
    IMediaPlayer_PutIsLoopingEnabled = 16
'    DEPRECATED ("Use PlaybackSession.IsProtected instead of IsProtected.  For more info, see MSDN.")
'    IMediaPlayer_GetIsProtected = 17
    IMediaPlayer_GetIsMuted = 18
    IMediaPlayer_PutIsMuted = 19
'    DEPRECATED ("Use PlaybackSession.PlaybackRate instead of PlaybackRate.  For more info, see MSDN.")
'    IMediaPlayer_GetPlaybackRate = 20
'    DEPRECATED ("Use PlaybackSession.PlaybackRate instead of PlaybackRate.  For more info, see MSDN.")
'    IMediaPlayer_PutPlaybackRate = 21
    IMediaPlayer_GetVolume = 22
    IMediaPlayer_PutVolume = 23
'    DEPRECATED ("Use media tracks on MediaPlaybackItem instead of PlaybackMediaMarkers.  For more info, see MSDN.")
'    IMediaPlayer_GetPlaybackMediaMarkers = 24
    IMediaPlayer_AddMediaOpened = 25
    IMediaPlayer_RemoveMediaOpened = 26
    IMediaPlayer_AddMediaEnded = 27
    IMediaPlayer_RemoveMediaEnded = 28
    IMediaPlayer_AddMediaFailed = 29
    IMediaPlayer_RemoveMediaFailed = 30
'    DEPRECATED ("Use PlaybackSession.PlaybackStateChanged instead of CurrentStateChanged.  For more info, see MSDN.")
'    IMediaPlayer_AddCurrentStateChanged = 31
'    DEPRECATED ("Use PlaybackSession.PlaybackStateChanged instead of CurrentStateChanged.  For more info, see MSDN.")
'    IMediaPlayer_RemoveCurrentStateChanged = 32
'    DEPRECATED ("Use media tracks on MediaPlaybackItem instead of PlaybackMediaMarkers.  For more info, see MSDN.")
'    IMediaPlayer_AddPlaybackMediaMarkerReached = 33
'    DEPRECATED ("Use media tracks on MediaPlaybackItem instead of PlaybackMediaMarkers.  For more info, see MSDN.")
'    IMediaPlayer_RemovePlaybackMediaMarkerReached = 34
'    DEPRECATED ("Use PlaybackSession.PlaybackRateChanged instead of MediaPlayerRateChanged.  For more info, see MSDN.")
'    IMediaPlayer_AddMediaPlayerRateChanged = 35
'    DEPRECATED ("Use PlaybackSession.PlaybackRateChanged instead of MediaPlayerRateChanged.  For more info, see MSDN.")
'    IMediaPlayer_RemoveMediaPlayerRateChanged = 36
    IMediaPlayer_AddVolumeChanged = 37
    IMediaPlayer_RemoveVolumeChanged = 38
'    DEPRECATED ("Use PlaybackSession.SeekCompleted instead of SeekCompleted.  For more info, see MSDN.")
'    IMediaPlayer_AddSeekCompleted = 39
'    DEPRECATED ("Use PlaybackSession.SeekCompleted instead of SeekCompleted.  For more info, see MSDN.")
'    IMediaPlayer_RemoveSeekCompleted = 40
'    DEPRECATED ("Use PlaybackSession.BufferingStarted instead of BufferingStarted.  For more info, see MSDN.")
'    IMediaPlayer_AddBufferingStarted = 41
'    DEPRECATED ("Use PlaybackSession.BufferingStarted instead of BufferingStarted.  For more info, see MSDN.")
'    IMediaPlayer_RemoveBufferingStarted = 42
'    DEPRECATED ("Use PlaybackSession.BufferingEnded instead of BufferingEnded.  For more info, see MSDN.")
'    IMediaPlayer_AddBufferingEnded = 43
'    DEPRECATED ("Use PlaybackSession.BufferingEnded instead of BufferingEnded.  For more info, see MSDN.")
'    IMediaPlayer_RemoveBufferingEnded = 44
    IMediaPlayer_Play = 45
    IMediaPlayer_Pause = 46
'    DEPRECATED ("Use Source instead of SetUriSource.  For more info, see MSDN.")
'    IMediaPlayer_SetUriSource = 47
                        
    ' IMediaPlayer2
    IMediaPlayer2_GetSystemMediaTransportControls = 6
    IMediaPlayer2_GetAudioCategory = 7
    IMediaPlayer2_PutAudioCategory = 8
    IMediaPlayer2_GetAudioDeviceType = 9
    IMediaPlayer2_PutAudioDeviceType = 10
                        
    ' IMediaPlayer3
    IMediaPlayer3_AddIsMutedChanged = 6
    IMediaPlayer3_RemoveIsMutedChanged = 7
    IMediaPlayer3_AddSourceChanged = 8
    IMediaPlayer3_RemoveSourceChanged = 9
    IMediaPlayer3_GetAudioBalance = 10
    IMediaPlayer3_PutAudioBalance = 11
    IMediaPlayer3_GetRealTimePlayback = 12
    IMediaPlayer3_PutRealTimePlayback = 13
    IMediaPlayer3_GetStereoscopicVideoRenderMode = 14
    IMediaPlayer3_PutStereoscopicVideoRenderMode = 15
    IMediaPlayer3_GetBreakManager = 16
    IMediaPlayer3_GetCommandManager = 17
    IMediaPlayer3_GetAudioDevice = 18
    IMediaPlayer3_PutAudioDevice = 19
    IMediaPlayer3_GetTimelineController = 20
    IMediaPlayer3_PutTimelineController = 21
    IMediaPlayer3_GetTimelineControllerPositionOffset = 22
    IMediaPlayer3_PutTimelineControllerPositionOffset = 23
    IMediaPlayer3_GetPlaybackSession = 24
    IMediaPlayer3_StepForwardOneFrame = 25
    IMediaPlayer3_StepBackwardOneFrame = 26
'    IMediaPlayer3_GetAsCastingSource = 27
                        
    ' IMediaPlayer4
    IMediaPlayer4_SetSurfaceSize = 6
'    IMediaPlayer4_GetSurface = 7
                        
    ' IMediaPlayer5
    IMediaPlayer5_AddVideoFrameAvailable = 6
    IMediaPlayer5_RemoveVideoFrameAvailable = 7
    IMediaPlayer5_GetIsVideoFrameServerEnabled = 8
    IMediaPlayer5_PutIsVideoFrameServerEnabled = 9
'    IMediaPlayer5_CopyFrameToVideoSurface = 10
'    IMediaPlayer5_CopyFrameToVideoSurfaceWithTargetRectangle = 11
'    IMediaPlayer5_CopyFrameToStereoscopicVideoSurfaces = 12
                        
    ' IMediaPlayer6
    IMediaPlayer6_AddSubtitleFrameChanged = 6
    IMediaPlayer6_RemoveSubtitleFrameChanged = 7
'    IMediaPlayer6_RenderSubtitlesToSurface = 8
'    IMediaPlayer6_RenderSubtitlesToSurfaceWithTargetRectangle = 9
                        
    ' IMediaPlayer7
    IMediaPlayer7_GetAudioStateMonitor = 6
    
    ' IMediaPlayerSource
'   IMediaPlayerSource_GetProtectionManager = 6
'   IMediaPlayerSource_PutProtectionManager = 7
'   DEPRECATED ("Use Source instead of SetFileSource.  For more info, see MSDN.")
'   IMediaPlayerSource_SetFileSource = 8
'   DEPRECATED ("Use Source instead of SetStreamSource.  For more info, see MSDN.")
'   IMediaPlayerSource_SetStreamSource = 9
'   DEPRECATED ("Use Source instead of SetMediaSource.  For more info, see MSDN.")
'   IMediaPlayerSource_SetMediaSource = 10
                        
    ' IMediaPlayerSource2
   IMediaPlayerSource2_GetSource = 6
   IMediaPlayerSource2_PutSource = 7
                        
    ' IMediaPlayerEffects
'   IMediaPlayerEffects_AddAudioEffect = 6
'   IMediaPlayerEffects_RemoveAllEffects = 7
                    
    ' IMediaPlayerEffects2
'   IMediaPlayerEffects2_AddVideoEffect = 6

End Enum

' ----==== Variablen ====----
Private m_pIMediaPlayer As Long
Private m_pIMediaPlayer2 As Long
Private m_pIMediaPlayer3 As Long
Private m_pIMediaPlayer4 As Long
Private m_pIMediaPlayer5 As Long
Private m_pIMediaPlayer6 As Long
Private m_pIMediaPlayer7 As Long
Private m_pIMediaPlayerSource As Long
Private m_pIMediaPlayerSource2 As Long
Private m_pIMediaPlayerEffects As Long
Private m_pIMediaPlayerEffects2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIMediaPlayer As Long
    If GetActivateInstance(Windows_Media_Playback_MediaPlayer, _
                           IID_IMediaPlayer, _
                           pIMediaPlayer) Then
        Me.Ifc = pIMediaPlayer
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaPlayerEffects2)
    Call ReleaseIfc(m_pIMediaPlayerEffects)
    Call ReleaseIfc(m_pIMediaPlayerSource2)
    Call ReleaseIfc(m_pIMediaPlayerSource)
    Call ReleaseIfc(m_pIMediaPlayer7)
    Call ReleaseIfc(m_pIMediaPlayer6)
    Call ReleaseIfc(m_pIMediaPlayer5)
    Call ReleaseIfc(m_pIMediaPlayer4)
    Call ReleaseIfc(m_pIMediaPlayer3)
    Call ReleaseIfc(m_pIMediaPlayer2)
    Call DisposeIfc(m_pIMediaPlayer) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaPlayer
End Property

Public Property Let Ifc(ByVal pIMediaPlayer As Long)
    Call ReleaseIfc(m_pIMediaPlayerEffects2)
    Call ReleaseIfc(m_pIMediaPlayerEffects)
    Call ReleaseIfc(m_pIMediaPlayerSource2)
    Call ReleaseIfc(m_pIMediaPlayerSource)
    Call ReleaseIfc(m_pIMediaPlayer7)
    Call ReleaseIfc(m_pIMediaPlayer6)
    Call ReleaseIfc(m_pIMediaPlayer5)
    Call ReleaseIfc(m_pIMediaPlayer4)
    Call ReleaseIfc(m_pIMediaPlayer3)
    Call ReleaseIfc(m_pIMediaPlayer2)
    Call DisposeIfc(m_pIMediaPlayer) ' <- IClosable
    m_pIMediaPlayer = pIMediaPlayer
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayer2, m_pIMediaPlayer2)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayer3, m_pIMediaPlayer3)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayer4, m_pIMediaPlayer4)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayer5, m_pIMediaPlayer5)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayer6, m_pIMediaPlayer6)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayer7, m_pIMediaPlayer7)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayerSource, m_pIMediaPlayerSource)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayerSource2, m_pIMediaPlayerSource2)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayerEffects, m_pIMediaPlayerEffects)
    Call QueryIfc(m_pIMediaPlayer, IID_IMediaPlayerEffects2, m_pIMediaPlayerEffects2)
End Property

' IMediaPlayer
Public Property Get AutoPlay() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_GetAutoPlay, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    AutoPlay = Ret
End Property

Public Property Let AutoPlay(ByVal value As Boolean)
    If m_pIMediaPlayer <> 0& Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_PutAutoPlay, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsLoopingEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_GetIsLoopingEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsLoopingEnabled = Ret
End Property

Public Property Let IsLoopingEnabled(ByVal value As Boolean)
    If m_pIMediaPlayer <> 0& Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_PutIsLoopingEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsMuted() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_GetIsMuted, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsMuted = Ret
End Property

Public Property Let IsMuted(ByVal value As Boolean)
    If m_pIMediaPlayer <> 0& Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_PutIsMuted, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Volume() As Double
    Dim Ret As Double
    If m_pIMediaPlayer <> 0& Then
        Dim value As Double
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_GetVolume, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Volume = Ret
End Property

Public Property Let Volume(ByVal value As Double)
    If m_pIMediaPlayer <> 0& Then
        If value < CDbl(0) Then value = CDbl(0)
        If value > CDbl(1) Then value = CDbl(1)
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_PutVolume, _
                  value) = S_OK Then
        End If
    End If
End Property

' IMediaPlayer2
Public Property Get SystemMediaTransportControls() As SystemMediaTransportControls
    Dim Ret As SystemMediaTransportControls
    If m_pIMediaPlayer2 <> 0& Then
        Dim pISystemMediaTransportControls As Long
        If Invoke(m_pIMediaPlayer2, _
                  IMediaPlayer2_GetSystemMediaTransportControls, _
                  VarPtr(pISystemMediaTransportControls)) = S_OK Then
            If pISystemMediaTransportControls <> 0& Then
                Set Ret = New SystemMediaTransportControls
                Ret.Ifc = pISystemMediaTransportControls
            End If
        End If
    End If
    Set SystemMediaTransportControls = Ret
End Property

Public Property Get AudioCategory() As MediaPlayerAudioCategory
    Dim Ret As MediaPlayerAudioCategory
    If m_pIMediaPlayer2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer2, _
                  IMediaPlayer2_GetAudioCategory, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AudioCategory = Ret
End Property

Public Property Let AudioCategory(ByVal value As MediaPlayerAudioCategory)
    If m_pIMediaPlayer <> 0& Then
        If Invoke(m_pIMediaPlayer2, _
                  IMediaPlayer2_PutAudioCategory, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get AudioDeviceType() As MediaPlayerAudioDeviceType
    Dim Ret As MediaPlayerAudioDeviceType
    If m_pIMediaPlayer2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer2, _
                  IMediaPlayer2_GetAudioDeviceType, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AudioDeviceType = Ret
End Property

Public Property Let AudioDeviceType(ByVal value As MediaPlayerAudioDeviceType)
    If m_pIMediaPlayer <> 0& Then
        If Invoke(m_pIMediaPlayer2, _
                  IMediaPlayer2_PutAudioDeviceType, _
                  value) = S_OK Then
        End If
    End If
End Property

' IMediaPlayer3
Public Property Get AudioBalance() As Double
    Dim Ret As Double
    If m_pIMediaPlayer3 <> 0& Then
        Dim value As Double
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetAudioBalance, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AudioBalance = Ret
End Property

Public Property Let AudioBalance(ByVal value As Double)
    If m_pIMediaPlayer3 <> 0& Then
        If value < CDbl(-1) Then value = CDbl(-1)
        If value > CDbl(1) Then value = CDbl(1)
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_PutAudioBalance, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get RealTimePlayback() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetRealTimePlayback, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    RealTimePlayback = Ret
End Property

Public Property Let RealTimePlayback(ByVal value As Boolean)
    If m_pIMediaPlayer3 <> 0& Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_PutRealTimePlayback, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get StereoscopicVideoRenderMode() As StereoscopicVideoRenderMode
    Dim Ret As StereoscopicVideoRenderMode
    If m_pIMediaPlayer3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetStereoscopicVideoRenderMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    StereoscopicVideoRenderMode = Ret
End Property

Public Property Let StereoscopicVideoRenderMode(ByVal value As StereoscopicVideoRenderMode)
    If m_pIMediaPlayer3 <> 0& Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_PutStereoscopicVideoRenderMode, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get BreakManager() As MediaBreakManager
    Dim Ret As MediaBreakManager
    If m_pIMediaPlayer3 <> 0& Then
        Dim pIMediaBreakManager As Long
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetBreakManager, _
                  VarPtr(pIMediaBreakManager)) = S_OK Then
            If pIMediaBreakManager <> 0& Then
                Set Ret = New MediaBreakManager
                Ret.Ifc = pIMediaBreakManager
            End If
        End If
    End If
    Set BreakManager = Ret
End Property

Public Property Get CommandManager() As MediaPlaybackCommandManager
    Dim Ret As MediaPlaybackCommandManager
    If m_pIMediaPlayer3 <> 0& Then
        Dim pIMediaPlaybackCommandManager As Long
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetCommandManager, _
                  VarPtr(pIMediaPlaybackCommandManager)) = S_OK Then
            If pIMediaPlaybackCommandManager <> 0& Then
                Set Ret = New MediaPlaybackCommandManager
                Ret.Ifc = pIMediaPlaybackCommandManager
            End If
        End If
    End If
    Set CommandManager = Ret
End Property

Public Property Get AudioDevice() As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIMediaPlayer3 <> 0& Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetAudioDevice, _
                  VarPtr(pIDeviceInformation)) = S_OK Then
            If pIDeviceInformation <> 0& Then
                Set Ret = New DeviceInformation
                Ret.Ifc = pIDeviceInformation
            End If
        End If
    End If
    Set AudioDevice = Ret
End Property

Public Property Let AudioDevice(ByVal value As DeviceInformation)
    If m_pIMediaPlayer3 <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_PutAudioDevice, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get TimelineController() As MediaTimelineController
    Dim Ret As MediaTimelineController
    If m_pIMediaPlayer3 <> 0& Then
        Dim pIMediaTimelineController As Long
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetTimelineController, _
                  VarPtr(pIMediaTimelineController)) = S_OK Then
            If pIMediaTimelineController <> 0& Then
                Set Ret = New MediaTimelineController
                Ret.Ifc = pIMediaTimelineController
            End If
        End If
    End If
    Set TimelineController = Ret
End Property

Public Property Let TimelineController(ByVal value As MediaTimelineController)
    If m_pIMediaPlayer3 <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_PutTimelineController, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get TimelineControllerPositionOffset() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaPlayer3 <> 0& Then
        Dim value As Currency
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetTimelineControllerPositionOffset, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set TimelineControllerPositionOffset = Ret
End Property

Public Property Let TimelineControllerPositionOffset(ByVal value As TimeSpan)
    If m_pIMediaPlayer3 <> 0& Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_PutTimelineControllerPositionOffset, _
                  value.TimeSpanValue) = S_OK Then
        End If
    End If
End Property

Public Property Get PlaybackSession() As MediaPlaybackSession
    Dim Ret As MediaPlaybackSession
    If m_pIMediaPlayer3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_GetPlaybackSession, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New MediaPlaybackSession
                Ret.Ifc = value
            End If
        End If
    End If
    Set PlaybackSession = Ret
End Property

' IMediaPlayer5
Public Property Get IsVideoFrameServerEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer5 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayer5, _
                  IMediaPlayer5_GetIsVideoFrameServerEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsVideoFrameServerEnabled = Ret
End Property

Public Property Let IsVideoFrameServerEnabled(ByVal value As Boolean)
    If m_pIMediaPlayer5 <> 0& Then
        If Invoke(m_pIMediaPlayer5, _
                  IMediaPlayer5_PutIsVideoFrameServerEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

' IMediaPlayer7
Public Property Get AudioStateMonitor() As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIMediaPlayer7 <> 0& Then
        Dim pIAudioStateMonitor As Long
        If Invoke(m_pIMediaPlayer7, _
                  IMediaPlayer7_GetAudioStateMonitor, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
    End If
    Set AudioStateMonitor = Ret
End Property

' IMediaPlayerSource
'   IMediaPlayerSource_GetProtectionManager = 6
'   IMediaPlayerSource_PutProtectionManager = 7

' IMediaPlayerSource2
Public Property Get Source() As MediaSource
    Dim Ret As MediaSource
    If m_pIMediaPlayerSource2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayerSource2, _
                  IMediaPlayerSource2_GetSource, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New MediaSource
                Ret.Ifc = value
            End If
        End If
    End If
    Set Source = Ret
End Property

Public Property Let Source(ByVal value As MediaSource)
    If m_pIMediaPlayerSource2 <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaPlayerSource2, _
                  IMediaPlayerSource2_PutSource, _
                  value.Ifc) = S_OK Then
        End If
    Else
        If Invoke(m_pIMediaPlayerSource2, _
                  IMediaPlayerSource2_PutSource, _
                  Nothing) = S_OK Then
        End If
    End If
End Property
                        
' ----==== Functions ====----
' IMediaPlayer
Public Function AddMediaOpened(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_AddMediaOpened, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddMediaOpened = Ret
End Function

Public Function RemoveMediaOpened(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_RemoveMediaOpened, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveMediaOpened = Ret
End Function

Public Function AddMediaEnded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_AddMediaEnded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddMediaEnded = Ret
End Function

Public Function RemoveMediaEnded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_RemoveMediaEnded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveMediaEnded = Ret
End Function

Public Function AddMediaFailed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_AddMediaFailed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddMediaFailed = Ret
End Function

Public Function RemoveMediaFailed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_RemoveMediaFailed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveMediaFailed = Ret
End Function

Public Function AddVolumeChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_AddVolumeChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddVolumeChanged = Ret
End Function

Public Function RemoveVolumeChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_RemoveVolumeChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveVolumeChanged = Ret
End Function

Public Function Play() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_Play) = S_OK Then
            Ret = True
        End If
    End If
    Play = Ret
End Function

Public Function Pause() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer <> 0& Then
        If Invoke(m_pIMediaPlayer, _
                  IMediaPlayer_Pause) = S_OK Then
            Ret = True
        End If
    End If
    Pause = Ret
End Function

' IMediaPlayer3
Public Function AddIsMutedChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer3 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_AddIsMutedChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddIsMutedChanged = Ret
End Function

Public Function RemoveIsMutedChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer3 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_RemoveIsMutedChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveIsMutedChanged = Ret
End Function

Public Function AddSourceChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer3 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_AddSourceChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddSourceChanged = Ret
End Function

Public Function RemoveSourceChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer3 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_RemoveSourceChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveSourceChanged = Ret
End Function

Public Function StepForwardOneFrame() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer3 <> 0& Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_StepForwardOneFrame) = S_OK Then
            Ret = True
        End If
    End If
    StepForwardOneFrame = Ret
End Function

Public Function StepBackwardOneFrame() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer3 <> 0& Then
        If Invoke(m_pIMediaPlayer3, _
                  IMediaPlayer3_StepBackwardOneFrame) = S_OK Then
            Ret = True
        End If
    End If
    StepBackwardOneFrame = Ret
End Function

'    IMediaPlayer3_GetAsCastingSource = 27

' IMediaPlayer4
Public Function SetSurfaceSize(ByVal value As Size) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer4 <> 0& Then
        If Invoke(m_pIMediaPlayer4, _
                  IMediaPlayer4_SetSurfaceSize, _
                  value.Height, _
                  value.Width) = S_OK Then
            Ret = True
        End If
    End If
    SetSurfaceSize = Ret
End Function

'    IMediaPlayer4_GetSurface = 7

' IMediaPlayer5
Public Function AddVideoFrameAvailable(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer5 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer5, _
                  IMediaPlayer5_AddVideoFrameAvailable, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddVideoFrameAvailable = Ret
End Function

Public Function RemoveVideoFrameAvailable(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer5 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer5, _
                  IMediaPlayer5_RemoveVideoFrameAvailable, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveVideoFrameAvailable = Ret
End Function

'    IMediaPlayer5_CopyFrameToVideoSurface = 10
'    IMediaPlayer5_CopyFrameToVideoSurfaceWithTargetRectangle = 11
'    IMediaPlayer5_CopyFrameToStereoscopicVideoSurfaces = 12

' IMediaPlayer6
Public Function AddSubtitleFrameChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlayer6 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlayer6, _
                  IMediaPlayer6_AddSubtitleFrameChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddSubtitleFrameChanged = Ret
End Function

Public Function RemoveSubtitleFrameChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlayer6 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlayer6, _
                  IMediaPlayer6_RemoveSubtitleFrameChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveSubtitleFrameChanged = Ret
End Function

'    IMediaPlayer6_RenderSubtitlesToSurface = 8
'    IMediaPlayer6_RenderSubtitlesToSurfaceWithTargetRectangle = 9

' IMediaPlayerEffects
'   IMediaPlayerEffects_AddAudioEffect = 6
'   IMediaPlayerEffects_RemoveAllEffects = 7
                    
' IMediaPlayerEffects2
'   IMediaPlayerEffects2_AddVideoEffect = 6

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
