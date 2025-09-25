VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Media"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Media

Option Explicit

' ----==== Enums Windows.Media.xxx ====----
Public Enum AudioProcessing
    AudioProcessing_Default = 0
    AudioProcessing_Raw = 1
End Enum

Public Enum MediaPlaybackAutoRepeatMode
    MediaPlaybackAutoRepeatMode_None = 0
    MediaPlaybackAutoRepeatMode_Track = 1
    MediaPlaybackAutoRepeatMode_List = 2
End Enum

Public Enum MediaPlaybackType
    MediaPlaybackType_Unknown = 0
    MediaPlaybackType_Music = 1
    MediaPlaybackType_Video = 2
    MediaPlaybackType_Image = 3
End Enum
            
Public Enum MediaPlaybackStatus
    MediaPlaybackStatus_Closed = 0
    MediaPlaybackStatus_Changing = 1
    MediaPlaybackStatus_Stopped = 2
    MediaPlaybackStatus_Playing = 3
    MediaPlaybackStatus_Paused = 4
End Enum
            
Public Enum SoundLevel
    SoundLevel_Muted = 0
    SoundLevel_Low = 1
    SoundLevel_Full = 2
End Enum

Public Enum MediaTimelineControllerState
    MediaTimelineControllerState_Paused = 0
    MediaTimelineControllerState_Running = 1
    MediaTimelineControllerState_Stalled = 2
    MediaTimelineControllerState_Error = 3
End Enum

' ----==== Namespace Media.Ocr ====----
Public Property Get Ocr() As Ocr
    Set Ocr = New Ocr
End Property

' ----==== Namespace Media.Capture ====----
Public Property Get Capture() As MCapture
    Set Capture = New MCapture
End Property

' ----==== Namespace Media.Casting ====----
Public Property Get Casting() As Casting
    Set Casting = New Casting
End Property

' ----==== Namespace Media.Transcoding ====----
Public Property Get Transcoding() As Transcoding
    Set Transcoding = New Transcoding
End Property

' ----==== Namespace Media.FaceAnalysis ====----
Public Property Get FaceAnalysis() As FaceAnalysis
    Set FaceAnalysis = New FaceAnalysis
End Property

' ----==== Namespace Media.MediaProperties ====----
Public Property Get MediaProperties() As MediaProperties
    Set MediaProperties = New MediaProperties
End Property

' ----==== Namespace Media.Control ====----
Public Property Get Control() As Control
    Set Control = New Control
End Property

' ----==== Namespace Media.Playback ====----
Public Property Get Playback() As Playback
    Set Playback = New Playback
End Property

' ----==== Namespace Media.Core ====----
Public Property Get Core() As MCore
    Set Core = New MCore
End Property

' ----==== Namespace Media.SystemMediaTransportControls ====----
Public Property Get SystemMediaTransportControls() As SystemMediaTransportControls
    Set SystemMediaTransportControls = New SystemMediaTransportControls
End Property

' ----==== Namespace Media.SystemMediaTransportControlsDisplayUpdater ====----
Public Property Get SystemMediaTransportControlsDisplayUpdater() As SystemMediaTransportControlsDisplayUpdater
    Set SystemMediaTransportControlsDisplayUpdater = New SystemMediaTransportControlsDisplayUpdater
End Property

' ----==== Namespace Media.MusicDisplayProperties ====----
Public Property Get MusicDisplayProperties() As MusicDisplayProperties
    Set MusicDisplayProperties = New MusicDisplayProperties
End Property

' ----==== Namespace Media.VideoDisplayProperties ====----
Public Property Get VideoDisplayProperties() As VideoDisplayProperties
    Set VideoDisplayProperties = New VideoDisplayProperties
End Property

' ----==== Namespace Media.ImageDisplayProperties ====----
Public Property Get ImageDisplayProperties() As ImageDisplayProperties
    Set ImageDisplayProperties = New ImageDisplayProperties
End Property

' ----==== Namespace Media.SpeechSynthesis ====----
Public Property Get SpeechSynthesis() As SpeechSynthesis
    Set SpeechSynthesis = New SpeechSynthesis
End Property

' ----==== Namespace Media.SpeechRecognition ====----
Public Property Get SpeechRecognition() As SpeechRecognition
    Set SpeechRecognition = New SpeechRecognition
End Property

' ----==== Namespace Media.MediaTimelineController ====----
Public Property Get MediaTimelineController() As MediaTimelineController
    Set MediaTimelineController = New MediaTimelineController
End Property

' ----==== Namespace Media.Audio ====----
Public Property Get Audio() As Audio
    Set Audio = New Audio
End Property

' ----==== Namespace Media.Render ====----
Public Property Get Render() As Render
    Set Render = New Render
End Property

' ----==== Namespace Media.Devices ====----
Public Property Get Devices() As MDevices
    Set Devices = New Devices
End Property

