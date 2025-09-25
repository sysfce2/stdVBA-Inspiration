VERSION 5.00
Begin VB.Form frmMediaPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MediaPlayer (Audio)"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDevice 
      Caption         =   "Select PlayBack Device"
      Height          =   465
      Left            =   2790
      TabIndex        =   9
      Top             =   570
      Width           =   2700
   End
   Begin VB.CommandButton cmdPlayWebRadio 
      Caption         =   "Play WebRadio"
      Height          =   465
      Left            =   2820
      TabIndex        =   8
      Top             =   60
      Width           =   2700
   End
   Begin VB.Frame frControls 
      Caption         =   "Controls"
      Height          =   1455
      Left            =   2790
      TabIndex        =   5
      Top             =   1080
      Width           =   2685
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   495
         Left            =   90
         TabIndex        =   7
         Top             =   840
         Width           =   2475
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   495
         Left            =   90
         TabIndex        =   6
         Top             =   300
         Width           =   2475
      End
   End
   Begin VB.Frame frBalVolMute 
      Caption         =   "Balance && Volume"
      Height          =   3300
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   2700
      Begin VB.CheckBox ckMute 
         Caption         =   "Mute"
         Height          =   225
         Left            =   90
         TabIndex        =   4
         Top             =   2970
         Width           =   2505
      End
      Begin VB.VScrollBar sbVolume 
         Height          =   2295
         Left            =   1200
         Max             =   100
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
      Begin VB.HScrollBar sbBalance 
         Height          =   255
         Left            =   90
         Max             =   100
         Min             =   -100
         TabIndex        =   2
         Top             =   270
         Width           =   2505
      End
   End
   Begin VB.CommandButton cmdOpenMedia 
      Caption         =   "OpenMedia"
      Height          =   465
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2700
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3780
      Left            =   5550
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3780
   End
End
Attribute VB_Name = "frmMediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' Namespace Windows
Private Windows As New Windows

'Namespace Windows.Storage.StorageFile
Private StorageFile As StorageFile

' Namespace Windows.Media.Core.MediaSource
Private MediaSource As MediaSource

' Namespace Windows.Media
Private SystemMediaTransportControls As SystemMediaTransportControls
Private SystemMediaTransportControlsDisplayUpdater As SystemMediaTransportControlsDisplayUpdater

' Namespace Windows.Media.Playback.MediaPlayer
Private MediaPlayer As MediaPlayer
Private MediaPlaybackSession As MediaPlaybackSession

' Namespace Windows.Devices.Enumeration.DeviceInformation
Private DeviceInformation As DeviceInformation

' ----==== for MediaPlayer Events ====----
Private m_pIMediaEndedEvent As Long
Private m_MediaEndedEvent_Token As Currency
Private m_pIMediaFailedEvent As Long
Private m_MediaFailedEvent_Token As Currency
Private m_pIMediaOpenedEvent As Long
Private m_MediaOpenedEvent_Token As Currency
Private m_pISourceChangedEvent As Long
Private m_SourceChangedEvent_Token As Currency
Private m_pIVolumeChangedEvent As Long
Private m_VolumeChangedEvent_Token As Currency
Private m_pIIsMutedChangedEvent As Long
Private m_IsMutedChangedEvent_Token As Currency

' ----==== for MediaPlaybackSession Events ====----
Private m_pIPlaybackStateChangedEvent As Long
Private m_PlaybackStateChangedEvent_Token As Currency
Private m_pIPlaybackRateChangedEvent As Long
Private m_PlaybackRateChangedEvent_Token As Currency
Private m_pISeekCompletedEvent As Long
Private m_SeekCompletedEvent_Token As Currency
Private m_pIBufferingStartedEvent As Long
Private m_BufferingStartedEvent_Token As Currency
Private m_pIBufferingEndedEvent As Long
Private m_BufferingEndedEvent_Token As Currency
Private m_pIBufferingProgressChangedEvent As Long
Private m_BufferingProgressChangedEvent_Token As Currency
Private m_pIDownloadProgressChangedEvent As Long
Private m_DownloadProgressChangedEvent_Token As Currency
Private m_pINaturalDurationChangedEvent As Long
Private m_NaturalDurationChangedEvent_Token As Currency
Private m_pIPositionChangedEvent As Long
Private m_PositionChangedEvent_Token As Currency
Private m_pINaturalVideoSizeChangedEvent As Long
Private m_NaturalVideoSizeChangedEvent_Token As Currency
Private m_pIBufferedRangesChangedEvent As Long
Private m_BufferedRangesChangedEvent_Token As Currency
Private m_pIPlayedRangesChangedEvent As Long
Private m_PlayedRangesChangedEvent_Token As Currency
Private m_pISeekableRangesChangedEvent As Long
Private m_SeekableRangesChangedEvent_Token As Currency
Private m_pISupportedPlaybackRatesChangedEvent As Long
Private m_SupportedPlaybackRatesChangedEvent_Token As Currency

Private Sub Form_Load()
    cmdPlay.Enabled = False
    cmdPause.Enabled = False
    frBalVolMute.Enabled = False

    ' ----==== for MediaPlayer Events ====----
    m_pIMediaEndedEvent = ITEH_MediaEnded.Create(Me)
    m_pIMediaFailedEvent = ITEH_MediaFailed.Create(Me)
    m_pIMediaOpenedEvent = ITEH_MediaOpened.Create(Me)
    m_pISourceChangedEvent = ITEH_SourceChanged.Create(Me)
    m_pIVolumeChangedEvent = ITEH_VolumeChanged.Create(Me)
    m_pIIsMutedChangedEvent = ITEH_IsMutedChanged.Create(Me)
    
    ' ----==== for MediaPlaybackSession Events ====----
    m_pIPlaybackStateChangedEvent = ITEH_PlaybackStateChanged.Create(Me)
    m_pIPlaybackRateChangedEvent = ITEH_PlaybackRateChanged.Create(Me)
    m_pISeekCompletedEvent = ITEH_SeekCompleted.Create(Me)
    m_pIBufferingStartedEvent = ITEH_BufferingStarted.Create(Me)
    m_pIBufferingEndedEvent = ITEH_BufferingEnded.Create(Me)
    m_pIBufferingProgressChangedEvent = ITEH_BufferingProgressChanged.Create(Me)
    m_pIDownloadProgressChangedEvent = ITEH_DownloadProgressChanged.Create(Me)
    m_pINaturalDurationChangedEvent = ITEH_NaturalDurationChanged.Create(Me)
    m_pIPositionChangedEvent = ITEH_PositionChanged.Create(Me)
    m_pINaturalVideoSizeChangedEvent = ITEH_NaturalVideoSizeChanged.Create(Me)
    m_pIBufferedRangesChangedEvent = ITEH_BufferedRangesChanged.Create(Me)
    m_pIPlayedRangesChangedEvent = ITEH_PlayedRangesChanged.Create(Me)
    m_pISeekableRangesChangedEvent = ITEH_SeekableRangesChanged.Create(Me)
    m_pISupportedPlaybackRatesChangedEvent = ITEH_SupportedPlaybackRatesChanged.Create(Me)

    Set MediaPlayer = Windows.Media.Playback.MediaPlayer
    If IsNotNothing(MediaPlayer) Then
    
        m_MediaEndedEvent_Token = MediaPlayer.AddMediaEnded(m_pIMediaEndedEvent)
        m_MediaFailedEvent_Token = MediaPlayer.AddMediaFailed(m_pIMediaFailedEvent)
        m_MediaOpenedEvent_Token = MediaPlayer.AddMediaOpened(m_pIMediaOpenedEvent)
        m_SourceChangedEvent_Token = MediaPlayer.AddSourceChanged(m_pISourceChangedEvent)
        m_VolumeChangedEvent_Token = MediaPlayer.AddVolumeChanged(m_pIVolumeChangedEvent)
        m_IsMutedChangedEvent_Token = MediaPlayer.AddIsMutedChanged(m_pIIsMutedChangedEvent)
        
        Debug.Print "----==== MediaPlayer ====----"
        Debug.Print "AudioCategory: " & MediaPlayer.AudioCategory
        Debug.Print "AudioDeviceType: " & MediaPlayer.AudioDeviceType
        Debug.Print "RealTimePlayback: " & MediaPlayer.RealTimePlayback
        Debug.Print "AutoPlay: " & MediaPlayer.AutoPlay
        Debug.Print "IsLoopingEnabled: " & MediaPlayer.IsLoopingEnabled
        Debug.Print "IsMuted: " & MediaPlayer.IsMuted
        Debug.Print "IsVideoFrameServerEnabled: " & MediaPlayer.IsVideoFrameServerEnabled
        Debug.Print "Volume: " & MediaPlayer.Volume
        Debug.Print "AudioBalance: " & MediaPlayer.AudioBalance
        Debug.Print
                
        MediaPlayer.AutoPlay = True
        sbVolume.Value = 100 - (MediaPlayer.Volume * 100)
        sbBalance.Value = MediaPlayer.AudioBalance * 100
        If MediaPlayer.IsMuted Then
            ckMute.Value = vbChecked
        Else
            ckMute.Value = vbUnchecked
        End If
        Set SystemMediaTransportControls = MediaPlayer.SystemMediaTransportControls
        If IsNotNothing(SystemMediaTransportControls) Then
            Set SystemMediaTransportControlsDisplayUpdater = SystemMediaTransportControls.DisplayUpdater
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsNotNothing(StorageFile) Then Set StorageFile = Nothing
    
    If IsNotNothing(MediaPlaybackSession) Then
        If m_pIPlaybackStateChangedEvent <> 0& And m_PlaybackStateChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemovePlaybackStateChanged(m_PlaybackStateChangedEvent_Token) Then
                ITEH_PlaybackStateChanged.Destroy
            End If
        End If
        If m_pIPlaybackRateChangedEvent <> 0& And m_PlaybackRateChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemovePlaybackRateChanged(m_PlaybackRateChangedEvent_Token) Then
                ITEH_PlaybackRateChanged.Destroy
            End If
        End If
        If m_pISeekCompletedEvent <> 0& And m_SeekCompletedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveSeekCompleted(m_SeekCompletedEvent_Token) Then
                ITEH_SeekCompleted.Destroy
            End If
        End If
        If m_pIBufferingStartedEvent <> 0& And m_BufferingStartedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveBufferingStarted(m_BufferingStartedEvent_Token) Then
                ITEH_BufferingStarted.Destroy
            End If
        End If
        If m_pIBufferingEndedEvent <> 0& And m_BufferingEndedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveBufferingEnded(m_BufferingEndedEvent_Token) Then
                ITEH_BufferingEnded.Destroy
            End If
        End If
        If m_pIBufferingProgressChangedEvent <> 0& And m_BufferingProgressChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveBufferingProgressChanged(m_BufferingProgressChangedEvent_Token) Then
                ITEH_BufferingProgressChanged.Destroy
            End If
        End If
        If m_pIDownloadProgressChangedEvent <> 0& And m_DownloadProgressChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveDownloadProgressChanged(m_DownloadProgressChangedEvent_Token) Then
                ITEH_DownloadProgressChanged.Destroy
            End If
        End If
        If m_pINaturalDurationChangedEvent <> 0& And m_NaturalDurationChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveNaturalDurationChanged(m_NaturalDurationChangedEvent_Token) Then
                ITEH_NaturalDurationChanged.Destroy
            End If
        End If
        If m_pIPositionChangedEvent <> 0& And m_PositionChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemovePositionChanged(m_PositionChangedEvent_Token) Then
                ITEH_PositionChanged.Destroy
            End If
        End If
        If m_pINaturalVideoSizeChangedEvent <> 0& And m_NaturalVideoSizeChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveNaturalVideoSizeChanged(m_NaturalVideoSizeChangedEvent_Token) Then
                ITEH_NaturalVideoSizeChanged.Destroy
            End If
        End If
        If m_pIBufferedRangesChangedEvent <> 0& And m_BufferedRangesChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveBufferedRangesChanged(m_BufferedRangesChangedEvent_Token) Then
                ITEH_BufferedRangesChanged.Destroy
            End If
        End If
        If m_pIPlayedRangesChangedEvent <> 0& And m_PlayedRangesChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemovePlayedRangesChanged(m_PlayedRangesChangedEvent_Token) Then
                ITEH_PlayedRangesChanged.Destroy
            End If
        End If
        If m_pISeekableRangesChangedEvent <> 0& And m_SeekableRangesChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveSeekableRangesChanged(m_SeekableRangesChangedEvent_Token) Then
                ITEH_SeekableRangesChanged.Destroy
            End If
        End If
        If m_pISupportedPlaybackRatesChangedEvent <> 0& And m_SupportedPlaybackRatesChangedEvent_Token <> 0& Then
            If MediaPlaybackSession.RemoveSupportedPlaybackRatesChanged(m_SupportedPlaybackRatesChangedEvent_Token) Then
                ITEH_SupportedPlaybackRatesChanged.Destroy
            End If
        End If
        Set MediaPlaybackSession = Nothing
    End If
    
    If IsNotNothing(MediaPlayer) Then
        MediaPlayer.Source = Nothing
        If m_pIIsMutedChangedEvent <> 0& And m_IsMutedChangedEvent_Token <> 0& Then
            If MediaPlayer.RemoveIsMutedChanged(m_IsMutedChangedEvent_Token) Then
                ITEH_IsMutedChanged.Destroy
            End If
        End If
        If m_pIVolumeChangedEvent <> 0& And m_VolumeChangedEvent_Token <> 0& Then
            If MediaPlayer.RemoveVolumeChanged(m_VolumeChangedEvent_Token) Then
                ITEH_VolumeChanged.Destroy
            End If
        End If
        If m_pISourceChangedEvent <> 0& And m_SourceChangedEvent_Token <> 0& Then
            If MediaPlayer.RemoveSourceChanged(m_SourceChangedEvent_Token) Then
                ITEH_SourceChanged.Destroy
            End If
        End If
        If m_pIMediaOpenedEvent <> 0& And m_MediaOpenedEvent_Token <> 0& Then
            If MediaPlayer.RemoveMediaOpened(m_MediaOpenedEvent_Token) Then
                ITEH_MediaOpened.Destroy
            End If
        End If
        If m_pIMediaFailedEvent <> 0& And m_MediaFailedEvent_Token <> 0& Then
            If MediaPlayer.RemoveMediaFailed(m_MediaFailedEvent_Token) Then
                ITEH_MediaFailed.Destroy
            End If
        End If
        If m_pIMediaEndedEvent <> 0& And m_MediaEndedEvent_Token <> 0& Then
            If MediaPlayer.RemoveMediaEnded(m_MediaEndedEvent_Token) Then
                ITEH_MediaEnded.Destroy
            End If
        End If
        Set MediaPlayer = Nothing
    End If

    If IsNotNothing(MediaSource) Then Set MediaSource = Nothing

    If IsNotNothing(SystemMediaTransportControlsDisplayUpdater) Then
        Set SystemMediaTransportControlsDisplayUpdater = Nothing
    End If
    
    If IsNotNothing(SystemMediaTransportControls) Then
        Set SystemMediaTransportControls = Nothing
    End If

End Sub

Private Sub cmdOpenMedia_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = FileOpenPicker.FileTypeFilter
        Call FileTypeFilter.Append(".mp3")
        If IsNotNothing(StorageFile) Then Set StorageFile = Nothing
        Set StorageFile = FileOpenPicker.PickSingleFileAsync
        If IsNotNothing(StorageFile) Then
            If IsNotNothing(MediaSource) Then Set MediaSource = Nothing
            If IsNotNothing(MediaPlayer) Then MediaPlayer.Source = Nothing
            If IsNotNothing(MediaPlaybackSession) Then Set MediaPlaybackSession = Nothing
            If IsNotNothing(SystemMediaTransportControlsDisplayUpdater) Then SystemMediaTransportControlsDisplayUpdater.ClearAll
            Set MediaSource = Windows.Media.Core.MediaSource.CreateFromStorageFile(StorageFile)
            If IsNotNothing(MediaSource) Then
                MediaPlayer.Source = MediaSource
                Set MediaPlaybackSession = MediaPlayer.PlaybackSession
                If IsNotNothing(MediaPlaybackSession) Then
                    If m_PlaybackStateChangedEvent_Token = 0& Then
                        m_PlaybackStateChangedEvent_Token = MediaPlaybackSession.AddPlaybackStateChanged(m_pIPlaybackStateChangedEvent)
                    End If
                    If m_PlaybackRateChangedEvent_Token = 0& Then
                        m_PlaybackRateChangedEvent_Token = MediaPlaybackSession.AddPlaybackRateChanged(m_pIPlaybackRateChangedEvent)
                    End If
                    If m_SeekCompletedEvent_Token = 0& Then
                        m_SeekCompletedEvent_Token = MediaPlaybackSession.AddSeekCompleted(m_pISeekCompletedEvent)
                    End If
                    If m_BufferingStartedEvent_Token = 0& Then
                        m_BufferingStartedEvent_Token = MediaPlaybackSession.AddBufferingStarted(m_pIBufferingStartedEvent)
                    End If
                    If m_BufferingEndedEvent_Token = 0& Then
                        m_BufferingEndedEvent_Token = MediaPlaybackSession.AddBufferingEnded(m_pIBufferingEndedEvent)
                    End If
                    If m_BufferingProgressChangedEvent_Token = 0& Then
                        m_BufferingProgressChangedEvent_Token = MediaPlaybackSession.AddBufferingProgressChanged(m_pIBufferingProgressChangedEvent)
                    End If
                    If m_DownloadProgressChangedEvent_Token = 0& Then
                        m_DownloadProgressChangedEvent_Token = MediaPlaybackSession.AddDownloadProgressChanged(m_pIDownloadProgressChangedEvent)
                    End If
                    If m_NaturalDurationChangedEvent_Token = 0& Then
                        m_NaturalDurationChangedEvent_Token = MediaPlaybackSession.AddNaturalDurationChanged(m_pINaturalDurationChangedEvent)
                    End If
                    If m_PositionChangedEvent_Token = 0& Then
                        m_PositionChangedEvent_Token = MediaPlaybackSession.AddPositionChanged(m_pIPositionChangedEvent)
                    End If
                    If m_NaturalVideoSizeChangedEvent_Token = 0& Then
                        m_NaturalVideoSizeChangedEvent_Token = MediaPlaybackSession.AddNaturalVideoSizeChanged(m_pINaturalVideoSizeChangedEvent)
                    End If
                    If m_BufferedRangesChangedEvent_Token = 0& Then
                        m_BufferedRangesChangedEvent_Token = MediaPlaybackSession.AddBufferedRangesChanged(m_pIBufferedRangesChangedEvent)
                    End If
                    If m_PlayedRangesChangedEvent_Token = 0& Then
                        m_PlayedRangesChangedEvent_Token = MediaPlaybackSession.AddPlayedRangesChanged(m_pIPlayedRangesChangedEvent)
                    End If
                    If m_SeekableRangesChangedEvent_Token = 0& Then
                        m_SeekableRangesChangedEvent_Token = MediaPlaybackSession.AddSeekableRangesChanged(m_pISeekableRangesChangedEvent)
                    End If
                    If m_SupportedPlaybackRatesChangedEvent_Token = 0& Then
                        m_SupportedPlaybackRatesChangedEvent_Token = MediaPlaybackSession.AddSupportedPlaybackRatesChanged(m_pISupportedPlaybackRatesChangedEvent)
                    End If
                    
                    Debug.Print "----==== MediaPlaybackSession ====----"
                    Debug.Print "CanPause: " & MediaPlaybackSession.CanPause
                    Debug.Print "CanSeek: " & MediaPlaybackSession.CanSeek
                    Debug.Print "IsProtected: " & MediaPlaybackSession.IsProtected
                    Debug.Print "PlaybackRate: " & MediaPlaybackSession.PlaybackRate
                    Debug.Print
                End If
            End If
        Else
            Debug.Print "FileOpenPicker = Cancel"
        End If
    End If
End Sub

Private Sub cmdPlayWebRadio_Click()

    If IsNotNothing(MediaSource) Then Set MediaSource = Nothing
    If IsNotNothing(MediaPlayer) Then MediaPlayer.Source = Nothing
    If IsNotNothing(MediaPlaybackSession) Then Set MediaPlaybackSession = Nothing
    If IsNotNothing(SystemMediaTransportControlsDisplayUpdater) Then SystemMediaTransportControlsDisplayUpdater.ClearAll
    
    ' "http://tunein.t4e.dj/hard/dsl/mp3"
    ' "http://listen.technobase.fm/tunein-mp3-pls"
    ' "https://intenseradio.live-streams.nl:18000/low"
    ' "http://www.schwarze-welle.de:7500"
    Set MediaSource = Windows.Media.Core.MediaSource.CreateFromUri(Windows.Foundation.Uri.CreateUri("https://intenseradio.live-streams.nl:18000/low"))
    If IsNotNothing(MediaSource) Then
        MediaPlayer.Source = MediaSource
        Set MediaPlaybackSession = MediaPlayer.PlaybackSession
        If IsNotNothing(MediaPlaybackSession) Then
            If m_PlaybackStateChangedEvent_Token = 0& Then
                m_PlaybackStateChangedEvent_Token = MediaPlaybackSession.AddPlaybackStateChanged(m_pIPlaybackStateChangedEvent)
            End If
            If m_PlaybackRateChangedEvent_Token = 0& Then
                m_PlaybackRateChangedEvent_Token = MediaPlaybackSession.AddPlaybackRateChanged(m_pIPlaybackRateChangedEvent)
            End If
            If m_SeekCompletedEvent_Token = 0& Then
                m_SeekCompletedEvent_Token = MediaPlaybackSession.AddSeekCompleted(m_pISeekCompletedEvent)
            End If
            If m_BufferingStartedEvent_Token = 0& Then
                m_BufferingStartedEvent_Token = MediaPlaybackSession.AddBufferingStarted(m_pIBufferingStartedEvent)
            End If
            If m_BufferingEndedEvent_Token = 0& Then
                m_BufferingEndedEvent_Token = MediaPlaybackSession.AddBufferingEnded(m_pIBufferingEndedEvent)
            End If
            If m_BufferingProgressChangedEvent_Token = 0& Then
                m_BufferingProgressChangedEvent_Token = MediaPlaybackSession.AddBufferingProgressChanged(m_pIBufferingProgressChangedEvent)
            End If
            If m_DownloadProgressChangedEvent_Token = 0& Then
                m_DownloadProgressChangedEvent_Token = MediaPlaybackSession.AddDownloadProgressChanged(m_pIDownloadProgressChangedEvent)
            End If
            If m_NaturalDurationChangedEvent_Token = 0& Then
                m_NaturalDurationChangedEvent_Token = MediaPlaybackSession.AddNaturalDurationChanged(m_pINaturalDurationChangedEvent)
            End If
            If m_PositionChangedEvent_Token = 0& Then
                m_PositionChangedEvent_Token = MediaPlaybackSession.AddPositionChanged(m_pIPositionChangedEvent)
            End If
            If m_NaturalVideoSizeChangedEvent_Token = 0& Then
                m_NaturalVideoSizeChangedEvent_Token = MediaPlaybackSession.AddNaturalVideoSizeChanged(m_pINaturalVideoSizeChangedEvent)
            End If
            If m_BufferedRangesChangedEvent_Token = 0& Then
                m_BufferedRangesChangedEvent_Token = MediaPlaybackSession.AddBufferedRangesChanged(m_pIBufferedRangesChangedEvent)
            End If
            If m_PlayedRangesChangedEvent_Token = 0& Then
                m_PlayedRangesChangedEvent_Token = MediaPlaybackSession.AddPlayedRangesChanged(m_pIPlayedRangesChangedEvent)
            End If
            If m_SeekableRangesChangedEvent_Token = 0& Then
                m_SeekableRangesChangedEvent_Token = MediaPlaybackSession.AddSeekableRangesChanged(m_pISeekableRangesChangedEvent)
            End If
            If m_SupportedPlaybackRatesChangedEvent_Token = 0& Then
                m_SupportedPlaybackRatesChangedEvent_Token = MediaPlaybackSession.AddSupportedPlaybackRatesChanged(m_pISupportedPlaybackRatesChangedEvent)
            End If
            
            Debug.Print "----==== MediaPlaybackSession ====----"
            Debug.Print "CanPause: " & MediaPlaybackSession.CanPause
            Debug.Print "CanSeek: " & MediaPlaybackSession.CanSeek
            Debug.Print "IsProtected: " & MediaPlaybackSession.IsProtected
            Debug.Print "PlaybackRate: " & MediaPlaybackSession.PlaybackRate
            Debug.Print
            
        End If
    End If
End Sub

Private Sub cmdDevice_Click()
    Dim DevicePicker As DevicePicker
    Set DevicePicker = Windows.Devices.Enumeration.DevicePicker
    If IsNotNothing(DevicePicker) Then
        Dim DevicePickerAppearance As DevicePickerAppearance
        Set DevicePickerAppearance = DevicePicker.Appearance
        If IsNotNothing(DevicePickerAppearance) Then DevicePickerAppearance.Title = "Select PlayBack Device"
        Dim DeviceClassList As List_DeviceClass
        Set DeviceClassList = DevicePicker.Filter.SupportedDeviceClasses
        If IsNotNothing(DeviceClassList) Then Call DeviceClassList.Append(DeviceClass_AudioRender)
        If IsNotNothing(DeviceInformation) Then Set DeviceInformation = Nothing
        Set DeviceInformation = DevicePicker.PickSingleDeviceAsync(Windows.Foundation.Rect(0, 0, 300, 300))
        If IsNotNothing(DeviceInformation) Then
            If IsNotNothing(MediaPlayer) Then
                MediaPlayer.AudioDevice = DeviceInformation
            End If
        End If
    End If
End Sub

Private Sub sbBalance_Change()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.AudioBalance = CDbl(sbBalance.Value / 100)
End Sub

Private Sub sbBalance_Scroll()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.AudioBalance = CDbl(sbBalance.Value / 100)
End Sub

Private Sub sbVolume_Change()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.Volume = CDbl((100 - sbVolume.Value) / 100)
End Sub

Private Sub sbVolume_Scroll()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.Volume = CDbl((100 - sbVolume.Value) / 100)
End Sub

Private Sub ckMute_Click()
    If IsNotNothing(MediaPlayer) Then
        If ckMute.Value = vbChecked Then
            MediaPlayer.IsMuted = True
        Else
            MediaPlayer.IsMuted = False
        End If
    End If
End Sub

Private Sub cmdPause_Click()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.Pause
End Sub

Private Sub cmdPlay_Click()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.Play
End Sub

' ----==== MediaPlayer Events ====----
Public Sub MediaOpenedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlayer.Ifc Then
        'Debug.Print "MediaOpened", sender, args
        frBalVolMute.Enabled = True
        If IsNotNothing(SystemMediaTransportControlsDisplayUpdater) And IsNotNothing(StorageFile) Then
            SystemMediaTransportControlsDisplayUpdater.ClearAll
            If SystemMediaTransportControlsDisplayUpdater.CopyFromFileAsync(MediaPlaybackType_Music, StorageFile) Then
                Call SystemMediaTransportControlsDisplayUpdater.Update
                Image1.Picture = GetPictureFromRandomAccessStreamReference(SystemMediaTransportControlsDisplayUpdater.Thumbnail)
                Dim MusicDisplayProperties As MusicDisplayProperties
                Set MusicDisplayProperties = SystemMediaTransportControlsDisplayUpdater.MusicProperties
                If IsNotNothing(MusicDisplayProperties) Then
                    Debug.Print "AlbumArtist: " & MusicDisplayProperties.AlbumArtist
                    Debug.Print "AlbumTitle: " & MusicDisplayProperties.AlbumTitle
                    Debug.Print "AlbumTrackCount: " & MusicDisplayProperties.AlbumTrackCount
                    Debug.Print "TrackNumber: " & MusicDisplayProperties.TrackNumber
                    Debug.Print "Artist: " & MusicDisplayProperties.Artist
                    Debug.Print "Title: " & MusicDisplayProperties.Title
                    Dim Genres As ReadOnlyList_1 'ReadOnlyList_String
                    Set Genres = MusicDisplayProperties.Genres
                    If IsNotNothing(Genres) Then
                        Dim GenresCount As Long
                        GenresCount = Genres.Size
                        If GenresCount > 0& Then
                            Dim strGenres As String
                            Dim GenresItem As Long
                            For GenresItem = 0 To GenresCount - 1
                                strGenres = strGenres & Genres.GetAt(GenresItem) & ", "
                            Next
                            Debug.Print "Genres: " & strGenres
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub MediaEndedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlayer.Ifc Then
        'Debug.Print "MediaEnded", sender, args
    End If
End Sub

Public Sub IsMutedChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlayer.Ifc Then
        'Debug.Print "IsMutedChanged", sender, args
    End If
End Sub

Public Sub SourceChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlayer.Ifc Then
        'Debug.Print "SourceChanged", sender, args
        frBalVolMute.Enabled = False
    End If
End Sub

Public Sub VolumeChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlayer.Ifc Then
        'Debug.Print "VolumeChanged", sender, args
    End If
End Sub

Public Sub MediaFailedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlayer.Ifc Then
        'Debug.Print "MediaFailed", sender, args
        If args <> 0& Then
            Dim MediaPlayerFailedEventArgs As New MediaPlayerFailedEventArgs
            MediaPlayerFailedEventArgs.Ifc = args
            Debug.Print "MediaFailed Error: " & MediaPlayerFailedEventArgs.Error
            Debug.Print "MediaFailed ExtendedErrorCode: " & "0x" & Hex$(MediaPlayerFailedEventArgs.ExtendedErrorCode)
            Debug.Print "MediaFailed ErrorMessage: " & MediaPlayerFailedEventArgs.ErrorMessage
        End If
    End If
End Sub

' ----==== MediaPlaybackSession Events ====----
Public Sub PlaybackStateChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "PlaybackStateChanged", sender, args
        Dim PlaybackState As MediaPlaybackState
        PlaybackState = MediaPlaybackSession.PlaybackState
        Debug.Print "PlaybackState: " & PlaybackState
        Select Case PlaybackState
            Case MediaPlaybackState.MediaPlaybackState_None
            Case MediaPlaybackState.MediaPlaybackState_Opening
                cmdPlay.Enabled = False
                cmdPause.Enabled = False
            Case MediaPlaybackState.MediaPlaybackState_Buffering
            Case MediaPlaybackState.MediaPlaybackState_Playing
                cmdPlay.Enabled = False
                cmdPause.Enabled = True
            Case MediaPlaybackState.MediaPlaybackState_Paused
                cmdPlay.Enabled = True
                cmdPause.Enabled = False
        End Select
    End If
End Sub

Public Sub PlaybackRateChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "PlaybackRateChanged", sender, args
        Debug.Print "PlaybackRate: " & MediaPlaybackSession.PlaybackRate
    End If
End Sub

Public Sub SeekCompletedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "SeekCompleted", sender, args
    End If
End Sub

Public Sub BufferingStartedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "BufferingStarted", sender, args
    End If
End Sub

Public Sub BufferingEndedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "BufferingEnded", sender, args
    End If
End Sub

Public Sub BufferingProgressChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "BufferingProgressChanged", sender, args
    End If
End Sub

Public Sub DownloadProgressChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "DownloadProgressChanged", sender, args
    End If
End Sub

Public Sub NaturalDurationChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "NaturalDurationChanged", sender, args
        Debug.Print "NaturalDuration: " & MediaPlaybackSession.NaturalDuration.VbDate
    End If
End Sub

Public Sub PositionChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "PositionChanged", sender, args
        Me.Caption = "MediaPlayer (Audio) Playing: " & MediaPlaybackSession.Position.VbDate & " to " & MediaPlaybackSession.NaturalDuration.VbDate
    End If
End Sub

Public Sub NaturalVideoSizeChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "NaturalVideoSizeChanged", sender, args
    End If
End Sub

Public Sub BufferedRangesChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "BufferedRangesChanged", sender, args
        Dim TimeRangeList As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
        Set TimeRangeList = MediaPlaybackSession.GetBufferedRanges
        If IsNotNothing(TimeRangeList) Then
            Dim TimeRangeCount As Long
            TimeRangeCount = TimeRangeList.Size
            If TimeRangeCount > 0& Then
                Dim TimeRangeItem As Long
                For TimeRangeItem = 0 To TimeRangeCount - 1
                    Dim TimeRange As MediaTimeRange
                    Set TimeRange = TimeRangeList.GetAt(TimeRangeItem)
                    If IsNotNothing(TimeRange) Then
                        'Debug.Print "BufferedRanges: " & TimeRange.ToString
                        Set TimeRange = Nothing
                    End If
                Next
            End If
            Set TimeRangeList = Nothing
        End If
    End If
End Sub

Public Sub PlayedRangesChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "PlayedRangesChanged", sender, args
        Dim TimeRangeList As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
        Set TimeRangeList = MediaPlaybackSession.GetPlayedRanges
        If IsNotNothing(TimeRangeList) Then
            Dim TimeRangeCount As Long
            TimeRangeCount = TimeRangeList.Size
            If TimeRangeCount > 0& Then
                Dim TimeRangeItem As Long
                For TimeRangeItem = 0 To TimeRangeCount - 1
                    Dim TimeRange As MediaTimeRange
                    Set TimeRange = TimeRangeList.GetAt(TimeRangeItem)
                    If IsNotNothing(TimeRange) Then
                        'Debug.Print "PlayedRanges: " & TimeRange.ToString
                        Set TimeRange = Nothing
                    End If
                Next
            End If
            Set TimeRangeList = Nothing
        End If
    End If
End Sub

Public Sub SeekableRangesChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "SeekableRangesChanged", sender, args
        Dim TimeRangeList As ReadOnlyList_1 'ReadOnlyList_MediaTimeRange
        Set TimeRangeList = MediaPlaybackSession.GetSeekableRanges
        If IsNotNothing(TimeRangeList) Then
            Dim TimeRangeCount As Long
            TimeRangeCount = TimeRangeList.Size
            If TimeRangeCount > 0& Then
                Dim TimeRangeItem As Long
                For TimeRangeItem = 0 To TimeRangeCount - 1
                    Dim TimeRange As MediaTimeRange
                    Set TimeRange = TimeRangeList.GetAt(TimeRangeItem)
                    If IsNotNothing(TimeRange) Then
                        'Debug.Print "SeekableRanges: " & TimeRange.ToString
                        Set TimeRange = Nothing
                    End If
                Next
            End If
            Set TimeRangeList = Nothing
        End If
    End If
End Sub

Public Sub SupportedPlaybackRatesChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = MediaPlaybackSession.Ifc Then
        'Debug.Print "SupportedPlaybackRatesChanged", sender, args
    End If
End Sub

