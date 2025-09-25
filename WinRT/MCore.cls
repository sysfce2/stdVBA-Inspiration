VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

' Namespace Core

Option Explicit

' ----==== Enums Windows.Media.Core.xxx ====----
Public Enum MediaTrackKind
    MediaTrackKind_Audio = 0
    MediaTrackKind_Video = 1
    MediaTrackKind_TimedMetadata = 2
End Enum

Public Enum MediaSourceStatus
    MediaSourceStatus_FullySupported = 0
    MediaSourceStatus_Unknown = 1
End Enum

Public Enum MediaDecoderStatus
    MediaDecoderStatus_FullySupported = 0
    MediaDecoderStatus_UnsupportedSubtype = 1
    MediaDecoderStatus_UnsupportedEncoderProperties = 2
    MediaDecoderStatus_Degraded = 3
End Enum

Public Enum AudioDecoderDegradation
    AudioDecoderDegradation_None = 0
    AudioDecoderDegradation_DownmixTo2Channels = 1
    AudioDecoderDegradation_DownmixTo6Channels = 2
    AudioDecoderDegradation_DownmixTo8Channels = 3
End Enum

Public Enum AudioDecoderDegradationReason
    AudioDecoderDegradationReason_None = 0
    AudioDecoderDegradationReason_LicensingRequirement = 1
    AudioDecoderDegradationReason_SpatialAudioNotSupported = 2
End Enum

' ----==== Namespace Core.MediaSource ====----
Public Property Get MediaSource() As MediaSource
    Set MediaSource = New MediaSource
End Property
