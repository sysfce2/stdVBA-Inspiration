VERSION 5.00
Begin VB.Form frmGlobalSystemMediaTransport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPlayPauseToggle 
      Caption         =   "PlayPauseToggle"
      Height          =   615
      Left            =   30
      TabIndex        =   3
      Top             =   1380
      Width           =   2865
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   615
      Left            =   30
      TabIndex        =   2
      Top             =   720
      Width           =   2865
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   615
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   2865
   End
   Begin VB.PictureBox Picture1 
      Height          =   4125
      Left            =   2940
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   0
      Top             =   30
      Width           =   4125
   End
End
Attribute VB_Name = "frmGlobalSystemMediaTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' Namespace Windows
Private Windows As New Windows

' Namespace Windows.Media.Control
Private SessionManager As GlobalSystemMediaTransportControlsSessionManager
Attribute SessionManager.VB_VarHelpID = -1
Private CurMediaSession As GlobalSystemMediaTransportControlsSession

Private m_pISessionsChangedEvent As Long
Private m_SessionsChangedEvent_Token As Currency
Private m_pIPlaybackInfoChangedEvent As Long
Private m_PlaybackInfoChangedEvent_Token As Currency
Private m_pICurrentSessionChangedEvent As Long
Private m_CurrentSessionChangedEvent_Token As Currency
Private m_pIMediaPropertiesChangedEvent As Long
Private m_MediaPropertiesChangedEvent_Token As Currency
Private m_pITimelinePropertiesChangedEvent As Long
Private m_TimelinePropertiesChangedEvent_Token As Currency

Private Sub Form_Load()
    cmdPlay.Enabled = False
    cmdPause.Enabled = False
    cmdPlayPauseToggle = False
    m_pISessionsChangedEvent = ITEH_SessionsChangedEvent.Create(Me)
    m_pIPlaybackInfoChangedEvent = ITEH_PlaybackInfoChanged.Create(Me)
    m_pICurrentSessionChangedEvent = ITEH_CurrentSessionChangedEvent.Create(Me)
    m_pIMediaPropertiesChangedEvent = ITEH_MediaPropertiesChanged.Create(Me)
    m_pITimelinePropertiesChangedEvent = ITEH_TimelinePropertiesChanged.Create(Me)
    Set SessionManager = Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager.RequestAsync
    If IsNotNothing(SessionManager) Then
        m_SessionsChangedEvent_Token = SessionManager.AddSessionsChanged(m_pISessionsChangedEvent)
        m_CurrentSessionChangedEvent_Token = SessionManager.AddCurrentSessionChanged(m_pICurrentSessionChangedEvent)
        Set CurMediaSession = SessionManager.GetCurrentSession
        If IsNotNothing(CurMediaSession) Then
            Me.Caption = "Session: " & CurMediaSession.SourceAppUserModelId
            Call GetPlaybackInfo
            Call GetMediaProperties
            Call GetTimeLineProperties
            m_PlaybackInfoChangedEvent_Token = CurMediaSession.AddPlaybackInfoChanged(m_pIPlaybackInfoChangedEvent)
            m_MediaPropertiesChangedEvent_Token = CurMediaSession.AddMediaPropertiesChanged(m_pIMediaPropertiesChangedEvent)
            m_TimelinePropertiesChangedEvent_Token = CurMediaSession.AddTimelinePropertiesChanged(m_pITimelinePropertiesChangedEvent)
        Else
            Me.Caption = "No running GlobalSystemMediaSessions found!"
            Picture1.Picture = LoadPicture
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsNotNothing(CurMediaSession) Then
        If m_pIPlaybackInfoChangedEvent <> 0& And m_PlaybackInfoChangedEvent_Token <> 0& Then
            If CurMediaSession.RemovePlaybackInfoChanged(m_PlaybackInfoChangedEvent_Token) Then
                m_pIPlaybackInfoChangedEvent = 0&
                m_PlaybackInfoChangedEvent_Token = 0&
            End If
        End If
        If m_pIMediaPropertiesChangedEvent <> 0& And m_MediaPropertiesChangedEvent_Token <> 0& Then
            If CurMediaSession.RemoveTimelinePropertiesChanged(m_MediaPropertiesChangedEvent_Token) Then
                m_pIMediaPropertiesChangedEvent = 0&
                m_MediaPropertiesChangedEvent_Token = 0&
            End If
        End If
        If m_pITimelinePropertiesChangedEvent <> 0& And m_TimelinePropertiesChangedEvent_Token <> 0& Then
            If CurMediaSession.RemoveTimelinePropertiesChanged(m_TimelinePropertiesChangedEvent_Token) Then
                m_pITimelinePropertiesChangedEvent = 0&
                m_TimelinePropertiesChangedEvent_Token = 0&
            End If
        End If
    End If
    If IsNotNothing(SessionManager) Then
        If m_pISessionsChangedEvent <> 0& And m_SessionsChangedEvent_Token <> 0& Then
            If SessionManager.RemoveSessionsChanged(m_SessionsChangedEvent_Token) Then
                m_pISessionsChangedEvent = 0&
                m_SessionsChangedEvent_Token = 0&
            End If
        End If
        If m_pICurrentSessionChangedEvent <> 0& And m_CurrentSessionChangedEvent_Token <> 0& Then
            If SessionManager.RemoveCurrentSessionChanged(m_CurrentSessionChangedEvent_Token) Then
                m_pICurrentSessionChangedEvent = 0&
                m_CurrentSessionChangedEvent_Token = 0&
            End If
        End If
    End If
    Set CurMediaSession = Nothing
    Set SessionManager = Nothing
End Sub

Private Sub cmdPause_Click()
    If IsNotNothing(CurMediaSession) Then
        Call CurMediaSession.TryPauseAsync
    End If
End Sub

Private Sub cmdPlay_Click()
    If IsNotNothing(CurMediaSession) Then
        Call CurMediaSession.TryPlayAsync
    End If
End Sub

Private Sub cmdPlayPauseToggle_Click()
    If IsNotNothing(CurMediaSession) Then
        Call CurMediaSession.TryTogglePlayPauseAsync
    End If
End Sub

Public Sub SessionsChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = SessionManager.Ifc Then
        If SessionManager.GetSessions.Size = 0 Then
            Me.Caption = "No running GlobalSystemMediaSessions found!"
            Picture1.Picture = LoadPicture
        End If
    End If
End Sub

Public Sub CurrentSessionChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = SessionManager.Ifc Then
        If IsNothing(CurMediaSession) Then
            Set CurMediaSession = SessionManager.GetCurrentSession
            If IsNotNothing(CurMediaSession) Then
                Me.Caption = "Session: " & CurMediaSession.SourceAppUserModelId
                Call GetPlaybackInfo
                Call GetMediaProperties
                Call GetTimeLineProperties
                m_PlaybackInfoChangedEvent_Token = CurMediaSession.AddPlaybackInfoChanged(m_pIPlaybackInfoChangedEvent)
                m_MediaPropertiesChangedEvent_Token = CurMediaSession.AddMediaPropertiesChanged(m_pIMediaPropertiesChangedEvent)
                m_TimelinePropertiesChangedEvent_Token = CurMediaSession.AddTimelinePropertiesChanged(m_pITimelinePropertiesChangedEvent)
            Else
                Me.Caption = "No running GlobalSystemMediaSessions found!"
                Picture1.Picture = LoadPicture
            End If
        End If
    End If
End Sub

Public Sub PlaybackInfoChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = CurMediaSession.Ifc Then
        Call GetPlaybackInfo
    End If
End Sub

Public Sub TimelinePropertiesChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = CurMediaSession.Ifc Then
        Call GetTimeLineProperties
    End If
End Sub

Public Sub MediaPropertiesChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = CurMediaSession.Ifc Then
        Call GetMediaProperties
    End If
End Sub

Private Sub GetMediaProperties()
    If IsNotNothing(CurMediaSession) Then
        Dim MediaProperties As GlobalSystemMediaTransportControlsSessionMediaProperties
        Set MediaProperties = CurMediaSession.TryGetMediaProperties
        If IsNotNothing(MediaProperties) Then
            Debug.Print "----==== MediaProperties ====----"
            Debug.Print "Artist: " & MediaProperties.Artist
            Debug.Print "Title: " & MediaProperties.Title
            Debug.Print "Subtitle: " & MediaProperties.Subtitle
            Debug.Print "AlbumTitle: " & MediaProperties.AlbumTitle
            Debug.Print "AlbumArtist: " & MediaProperties.AlbumArtist
            Debug.Print "AlbumTrackCount: " & MediaProperties.AlbumTrackCount
            Debug.Print "TrackNumber: " & MediaProperties.TrackNumber
            Debug.Print "PlaybackType: " & MediaProperties.PlaybackType
            Dim Genres As ReadOnlyList_1 'ReadOnlyList_String
            Set Genres = MediaProperties.Genres
            If IsNotNothing(Genres) Then
                If Genres.Size > 0& Then
                    Dim strGenre As String
                    Dim GenreItem As Long
                    For GenreItem = 0 To Genres.Size
                        strGenre = strGenre & Genres.GetAt(GenreItem) & ", "
                    Next
                    Debug.Print "Genres: " & strGenre
                End If
            End If
            Debug.Print
            Dim MediaThumbnail As RandomAccessStreamWithContentType
            Set MediaThumbnail = MediaProperties.Thumbnail.OpenReadAsync
            If IsNotNothing(MediaThumbnail) Then
                Dim pIStream As Long
                pIStream = MediaThumbnail.ToIStream
                If pIStream <> 0& Then
                    Picture1.Picture = GetPictureFromIStream(pIStream)
                    Call ReleaseIfc(pIStream)
                End If
                Set MediaThumbnail = Nothing
            End If
            Set MediaProperties = Nothing
        End If
    End If
End Sub

Private Sub GetTimeLineProperties()
    If IsNotNothing(CurMediaSession) Then
        Dim TimelineProperties As GlobalSystemMediaTransportControlsSessionTimelineProperties
        Set TimelineProperties = CurMediaSession.GetTimeLineProperties
        If IsNotNothing(TimelineProperties) Then
            Debug.Print "----==== TimelineProperties ====----"
            Debug.Print "StartTime: " & TimelineProperties.StartTime.VbDate
            Debug.Print "EndTime: " & TimelineProperties.EndTime.VbDate
            Debug.Print "LastUpdatedTime: " & TimelineProperties.LastUpdatedTime.VbDate
            Debug.Print "MaxSeekTime: " & TimelineProperties.MaxSeekTime.VbDate
            Debug.Print "MinSeekTime: " & TimelineProperties.MinSeekTime.VbDate
            Debug.Print "Position: " & TimelineProperties.Position.VbDate
            Debug.Print
            Set TimelineProperties = Nothing
        End If
    End If
End Sub

Private Sub GetPlaybackInfo()
    If IsNotNothing(CurMediaSession) Then
        Dim PlaybackInfo As GlobalSystemMediaTransportControlsSessionPlaybackInfo
        Set PlaybackInfo = CurMediaSession.GetPlaybackInfo
        If IsNotNothing(PlaybackInfo) Then
            Debug.Print "----==== PlaybackInfo ====----"
            Debug.Print "AutoRepeatMode: " & PlaybackInfo.AutoRepeatMode
            Debug.Print "IsShuffleActive: " & PlaybackInfo.IsShuffleActive
            Debug.Print "PlaybackRate: " & PlaybackInfo.PlaybackRate
            Debug.Print "PlaybackStatus: " & PlaybackInfo.PlaybackStatus
            Debug.Print "PlaybackType: " & PlaybackInfo.PlaybackType
            Dim PlaybackControls As GlobalSystemMediaTransportControlsSessionPlaybackControls
            Set PlaybackControls = PlaybackInfo.GetControls
            If IsNotNothing(PlaybackControls) Then
                Debug.Print "IsChannelUpEnabled: " & PlaybackControls.IsChannelUpEnabled
                Debug.Print "IsChannelDownEnabled: " & PlaybackControls.IsChannelDownEnabled
                Debug.Print "IsPreviousEnabled: " & PlaybackControls.IsPreviousEnabled
                Debug.Print "IsNextEnabled: " & PlaybackControls.IsNextEnabled

                Debug.Print "IsPlayEnabled: " & PlaybackControls.IsPlayEnabled
                cmdPlay.Enabled = PlaybackControls.IsPlayEnabled

                Debug.Print "IsPlaybackRateEnabled: " & PlaybackControls.IsPlaybackRateEnabled

                Debug.Print "IsPlayPauseToggleEnabled: " & PlaybackControls.IsPlayPauseToggleEnabled
                cmdPlayPauseToggle.Enabled = PlaybackControls.IsPlayPauseToggleEnabled

                Debug.Print "IsPlaybackPositionEnabled: " & PlaybackControls.IsPlaybackPositionEnabled

                Debug.Print "IsPauseEnabled: " & PlaybackControls.IsPauseEnabled
                cmdPause.Enabled = PlaybackControls.IsPauseEnabled

                Debug.Print "IsStopEnabled: " & PlaybackControls.IsStopEnabled
                Debug.Print "IsRecordEnabled: " & PlaybackControls.IsRecordEnabled
                Debug.Print "IsFastForwardEnabled: " & PlaybackControls.IsFastForwardEnabled
                Debug.Print "IsRepeatEnabled: " & PlaybackControls.IsRepeatEnabled
                Debug.Print "IsRewindEnabled: " & PlaybackControls.IsRewindEnabled
                Debug.Print "IsShuffleEnabled: " & PlaybackControls.IsShuffleEnabled
                Set PlaybackControls = Nothing
            End If
            Debug.Print
            Set PlaybackInfo = Nothing
        End If
    End If
End Sub
