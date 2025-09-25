VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlaybackSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaPlaybackSession As String = "Windows.Media.Playback.MediaPlaybackSession"
Private Const IID_IMediaPlaybackSession As String = "{c32b683d-0407-41ba-8946-8b345a5a5435}"
Private Const IID_IMediaPlaybackSession2 As String = "{f8ba7c79-1fc8-4097-ad70-c0fa18cc0050}"
Private Const IID_IMediaPlaybackSession3 As String = "{7ba2b41a-a3e2-405f-b77b-a4812c238b66}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaPlaybackSession
    IMediaPlaybackSession_AddPlaybackStateChanged = 6
    IMediaPlaybackSession_RemovePlaybackStateChanged = 7
    IMediaPlaybackSession_AddPlaybackRateChanged = 8
    IMediaPlaybackSession_RemovePlaybackRateChanged = 9
    IMediaPlaybackSession_AddSeekCompleted = 10
    IMediaPlaybackSession_RemoveSeekCompleted = 11
    IMediaPlaybackSession_AddBufferingStarted = 12
    IMediaPlaybackSession_RemoveBufferingStarted = 13
    IMediaPlaybackSession_AddBufferingEnded = 14
    IMediaPlaybackSession_RemoveBufferingEnded = 15
    IMediaPlaybackSession_AddBufferingProgressChanged = 16
    IMediaPlaybackSession_RemoveBufferingProgressChanged = 17
    IMediaPlaybackSession_AddDownloadProgressChanged = 18
    IMediaPlaybackSession_RemoveDownloadProgressChanged = 19
    IMediaPlaybackSession_AddNaturalDurationChanged = 20
    IMediaPlaybackSession_RemoveNaturalDurationChanged = 21
    IMediaPlaybackSession_AddPositionChanged = 22
    IMediaPlaybackSession_RemovePositionChanged = 23
    IMediaPlaybackSession_AddNaturalVideoSizeChanged = 24
    IMediaPlaybackSession_RemoveNaturalVideoSizeChanged = 25
    IMediaPlaybackSession_GetMediaPlayer = 26
    IMediaPlaybackSession_GetNaturalDuration = 27
    IMediaPlaybackSession_GetPosition = 28
    IMediaPlaybackSession_PutPosition = 29
    IMediaPlaybackSession_GetPlaybackState = 30
    IMediaPlaybackSession_GetCanSeek = 31
    IMediaPlaybackSession_GetCanPause = 32
    IMediaPlaybackSession_GetIsProtected = 33
    IMediaPlaybackSession_GetPlaybackRate = 34
    IMediaPlaybackSession_PutPlaybackRate = 35
    IMediaPlaybackSession_GetBufferingProgress = 36
    IMediaPlaybackSession_GetDownloadProgress = 37
    IMediaPlaybackSession_GetNaturalVideoHeight = 38
    IMediaPlaybackSession_GetNaturalVideoWidth = 39
    IMediaPlaybackSession_GetNormalizedSourceRect = 40
    IMediaPlaybackSession_PutNormalizedSourceRect = 41
    IMediaPlaybackSession_GetStereoscopicVideoPackingMode = 42
    IMediaPlaybackSession_PutStereoscopicVideoPackingMode = 43
                        
    ' IMediaPlaybackSession2
    IMediaPlaybackSession2_AddBufferedRangesChanged = 6
    IMediaPlaybackSession2_RemoveBufferedRangesChanged = 7
    IMediaPlaybackSession2_AddPlayedRangesChanged = 8
    IMediaPlaybackSession2_RemovePlayedRangesChanged = 9
    IMediaPlaybackSession2_AddSeekableRangesChanged = 10
    IMediaPlaybackSession2_RemoveSeekableRangesChanged = 11
    IMediaPlaybackSession2_AddSupportedPlaybackRatesChanged = 12
    IMediaPlaybackSession2_RemoveSupportedPlaybackRatesChanged = 13
'    IMediaPlaybackSession2_GetSphericalVideoProjection = 14
    IMediaPlaybackSession2_GetIsMirroring = 15
    IMediaPlaybackSession2_PutIsMirroring = 16
    IMediaPlaybackSession2_GetBufferedRanges = 17
    IMediaPlaybackSession2_GetPlayedRanges = 18
    IMediaPlaybackSession2_GetSeekableRanges = 19
    IMediaPlaybackSession2_IsSupportedPlaybackRateRange = 20
                        
    ' IMediaPlaybackSession3
    IMediaPlaybackSession3_GetPlaybackRotation = 6
    IMediaPlaybackSession3_PutPlaybackRotation = 7
'    IMediaPlaybackSession3_GetOutputDegradationPolicyState = 8
    
End Enum

' ----==== Types ====----
Private Type RECTF
    X As Single
    Y As Single
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pIMediaPlaybackSession As Long
Private m_pIMediaPlaybackSession2 As Long
Private m_pIMediaPlaybackSession3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaPlaybackSession3)
    Call ReleaseIfc(m_pIMediaPlaybackSession2)
    Call ReleaseIfc(m_pIMediaPlaybackSession)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaPlaybackSession
End Property

Public Property Let Ifc(ByVal pIMediaPlaybackSession As Long)
    Call ReleaseIfc(m_pIMediaPlaybackSession3)
    Call ReleaseIfc(m_pIMediaPlaybackSession2)
    Call ReleaseIfc(m_pIMediaPlaybackSession)
    m_pIMediaPlaybackSession = pIMediaPlaybackSession
    Call QueryIfc(m_pIMediaPlaybackSession, IID_IMediaPlaybackSession2, m_pIMediaPlaybackSession2)
    Call QueryIfc(m_pIMediaPlaybackSession, IID_IMediaPlaybackSession3, m_pIMediaPlaybackSession3)
End Property

' IMediaPlaybackSession
Public Property Get MediaPlayer() As MediaPlayer
    Dim Ret As MediaPlayer
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetMediaPlayer, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New MediaPlayer
                Ret.Ifc = value
            End If
        End If
    End If
    Set MediaPlayer = Ret
End Property

Public Property Get NaturalDuration() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetNaturalDuration, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set NaturalDuration = Ret
End Property

Public Property Get Position() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetPosition, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set Position = Ret
End Property

Public Property Let Position(ByVal value As TimeSpan)
    If m_pIMediaPlaybackSession <> 0& Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetPosition, _
                  value.TimeSpanValue) = S_OK Then
        End If
    End If
End Property

Public Property Get PlaybackState() As MediaPlaybackState
    Dim Ret As MediaPlaybackState
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetPlaybackState, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackState = Ret
End Property

Public Property Get CanSeek() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetCanSeek, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CanSeek = Ret
End Property

Public Property Get CanPause() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetCanPause, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CanPause = Ret
End Property

Public Property Get IsProtected() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetIsProtected, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsProtected = Ret
End Property

Public Property Get PlaybackRate() As Double
    Dim Ret As Double
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Double
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetPlaybackRate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackRate = Ret
End Property

Public Property Let PlaybackRate(ByVal value As Double)
    If m_pIMediaPlaybackSession <> 0& Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_PutPlaybackRate, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get BufferingProgress() As Double
    Dim Ret As Double
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Double
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetBufferingProgress, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BufferingProgress = Ret
End Property

Public Property Get DownloadProgress() As Double
    Dim Ret As Double
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Double
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetDownloadProgress, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DownloadProgress = Ret
End Property

Public Property Get NaturalVideoHeight() As Long
    Dim Ret As Long
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetNaturalVideoHeight, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NaturalVideoHeight = Ret
End Property

Public Property Get NaturalVideoWidth() As Long
    Dim Ret As Long
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetNaturalVideoWidth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NaturalVideoWidth = Ret
End Property

Public Property Get NormalizedSourceRect() As Rect
    Dim Ret As Rect
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As RECTF
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetNormalizedSourceRect, _
                  VarPtr(value)) = S_OK Then
            Ret.X = value.X
            Ret.Y = value.Y
            Ret.Width = value.Width
            Ret.Height = value.Height
        End If
    End If
    Set NormalizedSourceRect = Ret
End Property

Public Property Let NormalizedSourceRect(ByVal value As Rect)
    If m_pIMediaPlaybackSession <> 0& Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_PutNormalizedSourceRect, _
                  value.X, _
                  value.Y, _
                  value.Width, _
                  value.Height) = S_OK Then
        End If
    End If
End Property

Public Property Get StereoscopicVideoPackingMode() As StereoscopicVideoPackingMode
    Dim Ret As StereoscopicVideoPackingMode
    If m_pIMediaPlaybackSession <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_GetStereoscopicVideoPackingMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    StereoscopicVideoPackingMode = Ret
End Property

Public Property Let StereoscopicVideoPackingMode(ByVal value As StereoscopicVideoPackingMode)
    If m_pIMediaPlaybackSession <> 0& Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_PutStereoscopicVideoPackingMode, _
                  value) = S_OK Then
        End If
    End If
End Property
                        
' IMediaPlaybackSession2
'    IMediaPlaybackSession2_GetSphericalVideoProjection = 14

Public Property Get IsMirroring() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_GetIsMirroring, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsMirroring = Ret
End Property

Public Property Let IsMirroring(ByVal value As Boolean)
    If m_pIMediaPlaybackSession2 <> 0& Then
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_PutIsMirroring, _
                  value) = S_OK Then
        End If
    End If
End Property
                        
' IMediaPlaybackSession3
Public Property Get PlaybackRotation() As MediaRotation
    Dim Ret As MediaRotation
    If m_pIMediaPlaybackSession3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession3, _
                  IMediaPlaybackSession3_GetPlaybackRotation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackRotation = Ret
End Property

Public Property Let PlaybackRotation(ByVal value As MediaRotation)
    If m_pIMediaPlaybackSession3 <> 0& Then
        If Invoke(m_pIMediaPlaybackSession3, _
                  IMediaPlaybackSession3_PutPlaybackRotation, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IMediaPlaybackSession
Public Function AddPlaybackStateChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddPlaybackStateChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPlaybackStateChanged = Ret
End Function

Public Function RemovePlaybackStateChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemovePlaybackStateChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePlaybackStateChanged = Ret
End Function

Public Function AddPlaybackRateChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddPlaybackRateChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPlaybackRateChanged = Ret
End Function

Public Function RemovePlaybackRateChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemovePlaybackRateChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePlaybackRateChanged = Ret
End Function

Public Function AddSeekCompleted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddSeekCompleted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddSeekCompleted = Ret
End Function

Public Function RemoveSeekCompleted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemoveSeekCompleted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveSeekCompleted = Ret
End Function

Public Function AddBufferingStarted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddBufferingStarted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBufferingStarted = Ret
End Function

Public Function RemoveBufferingStarted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemoveBufferingStarted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBufferingStarted = Ret
End Function

Public Function AddBufferingEnded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddBufferingEnded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBufferingEnded = Ret
End Function

Public Function RemoveBufferingEnded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemoveBufferingEnded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBufferingEnded = Ret
End Function

Public Function AddBufferingProgressChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddBufferingProgressChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBufferingProgressChanged = Ret
End Function

Public Function RemoveBufferingProgressChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemoveBufferingProgressChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBufferingProgressChanged = Ret
End Function

Public Function AddDownloadProgressChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddDownloadProgressChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddDownloadProgressChanged = Ret
End Function

Public Function RemoveDownloadProgressChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemoveDownloadProgressChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveDownloadProgressChanged = Ret
End Function

Public Function AddNaturalDurationChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddNaturalDurationChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddNaturalDurationChanged = Ret
End Function

Public Function RemoveNaturalDurationChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemoveNaturalDurationChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveNaturalDurationChanged = Ret
End Function

Public Function AddPositionChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddPositionChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPositionChanged = Ret
End Function

Public Function RemovePositionChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemovePositionChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePositionChanged = Ret
End Function

Public Function AddNaturalVideoSizeChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_AddNaturalVideoSizeChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddNaturalVideoSizeChanged = Ret
End Function

Public Function RemoveNaturalVideoSizeChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession, _
                  IMediaPlaybackSession_RemoveNaturalVideoSizeChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveNaturalVideoSizeChanged = Ret
End Function
                        
' IMediaPlaybackSession2
Public Function AddBufferedRangesChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_AddBufferedRangesChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBufferedRangesChanged = Ret
End Function

Public Function RemoveBufferedRangesChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession2 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_RemoveBufferedRangesChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBufferedRangesChanged = Ret
End Function

Public Function AddPlayedRangesChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_AddPlayedRangesChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPlayedRangesChanged = Ret
End Function

Public Function RemovePlayedRangesChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession2 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_RemovePlayedRangesChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePlayedRangesChanged = Ret
End Function

Public Function AddSeekableRangesChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_AddSeekableRangesChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddSeekableRangesChanged = Ret
End Function

Public Function RemoveSeekableRangesChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession2 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_RemoveSeekableRangesChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveSeekableRangesChanged = Ret
End Function

Public Function AddSupportedPlaybackRatesChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackSession2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_AddSupportedPlaybackRatesChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddSupportedPlaybackRatesChanged = Ret
End Function

Public Function RemoveSupportedPlaybackRatesChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession2 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_RemoveSupportedPlaybackRatesChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveSupportedPlaybackRatesChanged = Ret
End Function

'    IMediaPlaybackSession2_GetSphericalVideoProjection = 14

Public Function GetBufferedRanges() As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
    If m_pIMediaPlaybackSession2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_GetBufferedRanges, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaTimeRange, pIVectorView) 'ReadOnlyList_MediaTimeRange
            End If
        End If
    End If
    Set GetBufferedRanges = Ret
End Function

Public Function GetPlayedRanges() As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
    If m_pIMediaPlaybackSession2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_GetPlayedRanges, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaTimeRange, pIVectorView) 'ReadOnlyList_MediaTimeRange
            End If
        End If
    End If
    Set GetPlayedRanges = Ret
End Function

Public Function GetSeekableRanges() As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
    If m_pIMediaPlaybackSession2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_GetSeekableRanges, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaTimeRange, pIVectorView) 'ReadOnlyList_MediaTimeRange
            End If
        End If
    End If
    Set GetSeekableRanges = Ret
End Function

Public Function IsSupportedPlaybackRateRange(ByVal rate1 As Double, _
                                             ByVal rate2 As Double) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackSession2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackSession2, _
                  IMediaPlaybackSession2_IsSupportedPlaybackRateRange, _
                  rate1, _
                  rate2, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsSupportedPlaybackRateRange = Ret
End Function

' IMediaPlaybackSession3
'    IMediaPlaybackSession3_GetOutputDegradationPolicyState = 8

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


