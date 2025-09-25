VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Transcoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Transcoding

Option Explicit

' ----==== Enums Windows.Media.Transcoding.xxx ===---
Public Enum MediaVideoProcessingAlgorithm
    MediaVideoProcessingAlgorithm_Default = 0
    MediaVideoProcessingAlgorithm_MrfCrf444 = 1
End Enum

Public Enum TranscodeFailureReason
    TranscodeFailureReason_None = 0
    TranscodeFailureReason_Unknown = 1
    TranscodeFailureReason_InvalidProfile = 2
    TranscodeFailureReason_CodecNotFound = 3
End Enum

' ----==== Namespace Transcoding.MediaTranscoder ====----
Public Property Get MediaTranscoder() As MediaTranscoder
    Set MediaTranscoder = New MediaTranscoder
End Property

' ----==== Namespace Transcoding.PrepareTranscodeResult ====----
Public Property Get PrepareTranscodeResult() As PrepareTranscodeResult
    Set PrepareTranscodeResult = New PrepareTranscodeResult
End Property

