VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Playback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

' Namespace Playback

Option Explicit

' ----==== Enums Windows.Media.Playback.xxx ====----
Public Enum MediaPlayerAudioCategory
    MediaPlayerAudioCategory_Other = 0
    MediaPlayerAudioCategory_Communications = 3
    MediaPlayerAudioCategory_Alerts = 4
    MediaPlayerAudioCategory_SoundEffects = 5
    MediaPlayerAudioCategory_GameEffects = 6
    MediaPlayerAudioCategory_GameMedia = 7
    MediaPlayerAudioCategory_GameChat = 8
    MediaPlayerAudioCategory_Speech = 9
    MediaPlayerAudioCategory_Movie = 10
    MediaPlayerAudioCategory_Media = 11
End Enum

Public Enum MediaPlayerAudioDeviceType
    MediaPlayerAudioDeviceType_Console = 0
    MediaPlayerAudioDeviceType_Multimedia = 1
    MediaPlayerAudioDeviceType_Communications = 2
End Enum

Public Enum MediaPlaybackState
    MediaPlaybackState_None = 0
    MediaPlaybackState_Opening = 1
    MediaPlaybackState_Buffering = 2
    MediaPlaybackState_Playing = 3
    MediaPlaybackState_Paused = 4
End Enum

Public Enum StereoscopicVideoRenderMode
    StereoscopicVideoRenderMode_Mono = 0
    StereoscopicVideoRenderMode_Stereo = 1
End Enum

Public Enum MediaPlayerError
    MediaPlayerError_Unknown = 0
    MediaPlayerError_Aborted = 1
    MediaPlayerError_NetworkError = 2
    MediaPlayerError_DecodingError = 3
    MediaPlayerError_SourceNotSupported = 4
End Enum

Public Enum AutoLoadedDisplayPropertyKind
    AutoLoadedDisplayPropertyKind_None = 0
    AutoLoadedDisplayPropertyKind_MusicOrVideo = 1
    AutoLoadedDisplayPropertyKind_Music = 2
    AutoLoadedDisplayPropertyKind_Video = 3
End Enum

Public Enum MediaBreakInsertionMethod
    MediaBreakInsertionMethod_Interrupt = 0
    MediaBreakInsertionMethod_Replace = 1
End Enum

Public Enum MediaCommandEnablingRule
    MediaCommandEnablingRule_Auto = 0
    MediaCommandEnablingRule_Always = 1
    MediaCommandEnablingRule_Never = 2
End Enum

' ----==== Namespace Playback.MediaPlayer ====----
Public Property Get MediaPlayer() As MediaPlayer
    Set MediaPlayer = New MediaPlayer
End Property

' ----==== Namespace Playback.MediaPlaybackSession ====----
Public Property Get MediaPlaybackSession() As MediaPlaybackSession
    Set MediaPlaybackSession = New MediaPlaybackSession
End Property

' ----==== Namespace Playback.MediaPlaybackItem ====----
Public Property Get MediaPlaybackItem() As MediaPlaybackItem
    Set MediaPlaybackItem = New MediaPlaybackItem
End Property

' ----==== Namespace Playback.MediaBreak ====----
Public Property Get MediaBreak() As MediaBreak
    Set MediaBreak = New MediaBreak
End Property

' ----==== Namespace Playback.MediaPlaybackList ====----
Public Property Get MediaPlaybackList() As MediaPlaybackList
    Set MediaPlaybackList = New MediaPlaybackList
End Property

