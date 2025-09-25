VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace MediaProperties

Option Explicit

' ----==== Enums Windows.Media.MediaProperties.xxx ====----
Public Enum SphericalVideoFrameFormat
    SphericalVideoFrameFormat_None = 0
    SphericalVideoFrameFormat_Unsupported = 1
    SphericalVideoFrameFormat_Equirectangular = 2
End Enum

Public Enum StereoscopicVideoPackingMode
    StereoscopicVideoPackingMode_None = 0
    StereoscopicVideoPackingMode_SideBySide = 1
    StereoscopicVideoPackingMode_TopBottom = 2
End Enum

Public Enum AudioEncodingQuality
    AudioEncodingQuality_Auto = 0
    AudioEncodingQuality_High = 1
    AudioEncodingQuality_Medium = 2
    AudioEncodingQuality_Low = 3
End Enum

Public Enum VideoEncodingQuality
    VideoEncodingQuality_Auto = 0
    VideoEncodingQuality_HD1080p = 1
    VideoEncodingQuality_HD720p = 2
    VideoEncodingQuality_Wvga = 3
    VideoEncodingQuality_Ntsc = 4
    VideoEncodingQuality_Pal = 5
    VideoEncodingQuality_Vga = 6
    VideoEncodingQuality_Qvga = 7
    VideoEncodingQuality_Uhd2160p = 8
    VideoEncodingQuality_Uhd4320p = 9
End Enum

Public Enum MediaPixelFormat
    MediaPixelFormat_Nv12 = 0
    MediaPixelFormat_Bgra8 = 1
    MediaPixelFormat_P010 = 2
End Enum

Public Enum MediaRotation
    MediaRotation_None = 0
    MediaRotation_Clockwise90Degrees = 1
    MediaRotation_Clockwise180Degrees = 2
    MediaRotation_Clockwise270Degrees = 3
End Enum


' ----==== Namespace MediaProperties.MediaEncodingProfile ====----
Public Property Get MediaEncodingProfile() As MediaEncodingProfile
    Set MediaEncodingProfile = New MediaEncodingProfile
End Property

' ----==== Namespace MediaProperties.AudioEncodingProperties ====----
Public Property Get AudioEncodingProperties() As AudioEncodingProperties
    Set AudioEncodingProperties = New AudioEncodingProperties
End Property

' ----==== Namespace MediaProperties.VideoEncodingProperties ====----
Public Property Get VideoEncodingProperties() As VideoEncodingProperties
    Set VideoEncodingProperties = New VideoEncodingProperties
End Property

' ----==== Namespace MediaProperties.ImageEncodingProperties ====----
Public Property Get ImageEncodingProperties() As ImageEncodingProperties
    Set ImageEncodingProperties = New ImageEncodingProperties
End Property


