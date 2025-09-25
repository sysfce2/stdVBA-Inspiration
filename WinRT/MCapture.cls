VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Capture

Option Explicit

' ----==== Enums Windows.Media.Capture.xxx ====----
Public Enum KnownVideoProfile
    KnownVideoProfile_VideoRecording = 0
    KnownVideoProfile_HighQualityPhoto = 1
    KnownVideoProfile_BalancedVideoAndPhoto = 2
    KnownVideoProfile_VideoConferencing = 3
    KnownVideoProfile_PhotoSequence = 4
    KnownVideoProfile_HighFrameRate = 5
    KnownVideoProfile_VariablePhotoSequence = 6
    KnownVideoProfile_HdrWithWcgVideo = 7
    KnownVideoProfile_HdrWithWcgPhoto = 8
    KnownVideoProfile_VideoHdr8 = 9
    KnownVideoProfile_CompressedCamera = 10
End Enum

Public Enum VideoRotation
    VideoRotation_None = 0
    VideoRotation_Clockwise90Degrees = 1
    VideoRotation_Clockwise180Degrees = 2
    VideoRotation_Clockwise270Degrees = 3
End Enum

Public Enum PhotoCaptureSource
    PhotoCaptureSource_Auto = 0
    PhotoCaptureSource_VideoPreview = 1
    PhotoCaptureSource_Photo = 2
End Enum

Public Enum StreamingCaptureMode
    StreamingCaptureMode_AudioAndVideo = 0
    StreamingCaptureMode_Audio = 1
    StreamingCaptureMode_Video = 2
End Enum

Public Enum MediaCategory
    MediaCategory_Other = 0
    MediaCategory_Communications = 1
    MediaCategory_Media = 2
    MediaCategory_GameChat = 3
    MediaCategory_Speech = 4
    MediaCategory_FarFieldSpeech = 5
    MediaCategory_UniformSpeech = 6
    MediaCategory_VoiceTyping = 7
End Enum

Public Enum MediaStreamType
    MediaStreamType_VideoPreview = 0
    MediaStreamType_VideoRecord = 1
    MediaStreamType_Audio = 2
    MediaStreamType_Photo = 3
    MediaStreamType_Metadata = 4
End Enum

Public Enum VideoDeviceCharacteristic
    VideoDeviceCharacteristic_AllStreamsIndependent = 0
    VideoDeviceCharacteristic_PreviewRecordStreamsIdentical = 1
    VideoDeviceCharacteristic_PreviewPhotoStreamsIdentical = 2
    VideoDeviceCharacteristic_RecordPhotoStreamsIdentical = 3
    VideoDeviceCharacteristic_AllStreamsIdentical = 4
End Enum

' ----==== Namespace Capture.MediaCapture ====----
Public Property Get MediaCapture() As MediaCapture
    Set MediaCapture = New MediaCapture
End Property

' ----==== Namespace Capture.MediaCaptureVideoProfile ====----
Public Property Get MediaCaptureVideoProfile() As MediaCaptureVideoProfile
    Set MediaCaptureVideoProfile = New MediaCaptureVideoProfile
End Property

' ----==== Namespace Capture.MediaCaptureInitializationSettings ====----
Public Property Get MediaCaptureInitializationSettings() As MediaCaptureInitializationSettings
    Set MediaCaptureInitializationSettings = New MediaCaptureInitializationSettings
End Property

