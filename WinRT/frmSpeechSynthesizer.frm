VERSION 5.00
Begin VB.Form frmSpeechSynthesizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpeechSynthesizer"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   574
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frVoices 
      Caption         =   "Voices"
      Height          =   645
      Left            =   30
      TabIndex        =   15
      Top             =   60
      Width           =   3855
      Begin VB.ComboBox cbVoices 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.Frame frOptions 
      Caption         =   "Volume"
      Height          =   645
      Index           =   0
      Left            =   30
      TabIndex        =   13
      Top             =   750
      Width           =   3855
      Begin VB.HScrollBar sbVolume 
         Height          =   285
         Left            =   90
         Max             =   100
         TabIndex        =   14
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.Frame frOptions 
      Caption         =   "Pitch"
      Height          =   645
      Index           =   1
      Left            =   30
      TabIndex        =   11
      Top             =   1440
      Width           =   3855
      Begin VB.HScrollBar sbPitch 
         Height          =   285
         Left            =   90
         Max             =   200
         TabIndex        =   12
         Top             =   240
         Width           =   3675
      End
   End
   Begin VB.Frame frOptions 
      Caption         =   "SpeakingRate"
      Height          =   645
      Index           =   2
      Left            =   30
      TabIndex        =   9
      Top             =   2130
      Width           =   3855
      Begin VB.HScrollBar sbSpeekingRate 
         Height          =   285
         Left            =   90
         Max             =   600
         Min             =   50
         TabIndex        =   10
         Top             =   240
         Value           =   50
         Width           =   3675
      End
   End
   Begin VB.Frame frText 
      Caption         =   "Text"
      Height          =   2715
      Left            =   3930
      TabIndex        =   6
      Top             =   60
      Width           =   4635
      Begin VB.TextBox tbText 
         Height          =   1755
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton cmdSythesizeText 
         Caption         =   "Text to Speech"
         Height          =   555
         Left            =   90
         TabIndex        =   7
         Top             =   2070
         Width           =   4455
      End
   End
   Begin VB.Frame frSsml 
      Caption         =   "Ssml"
      Height          =   2715
      Left            =   30
      TabIndex        =   3
      Top             =   2820
      Width           =   8535
      Begin VB.CommandButton cmdSynthesizeSsml 
         Caption         =   "Ssml to Speech: Infos auf https://www.w3.org/TR/speech-synthesis/"
         Height          =   555
         Left            =   90
         TabIndex        =   5
         Top             =   2070
         Width           =   8355
      End
      Begin VB.TextBox tbSsml 
         Height          =   1755
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   8355
      End
   End
   Begin VB.Frame frDivers 
      Caption         =   "Divers"
      Height          =   885
      Left            =   30
      TabIndex        =   0
      Top             =   5580
      Width           =   8535
      Begin VB.CommandButton cmdPlaySpeech 
         Caption         =   "Play Speech"
         Height          =   555
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   2760
      End
      Begin VB.CommandButton cmdPauseSpeech 
         Caption         =   "Pause Speech"
         Height          =   555
         Left            =   2910
         TabIndex        =   1
         Top             =   240
         Width           =   2745
      End
   End
End
Attribute VB_Name = "frmSpeechSynthesizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' Namespace Windows
Private Windows As New Windows

' Namespace Windows.Media.SpeechSynthesis
Private SpeechSynthesizer As SpeechSynthesizer
Attribute SpeechSynthesizer.VB_VarHelpID = -1
Private SpeechSynthesisStream As SpeechSynthesisStream
Private SpeechSynthesizerOptions As SpeechSynthesizerOptions
Private AllVoices As ReadOnlyList_1 'ReadOnlyList_VoiceInformation

' Namespace Windows.Media.Core.MediaSource
Private MediaSource As MediaSource

' Namespace Windows.Media.Playback.MediaPlayer
Private MediaPlayer As MediaPlayer

' Namespace Windows.Storage.Streams
Private RandomAccessStream As RandomAccessStream

Private Sub Form_Load()
    Set MediaPlayer = Windows.Media.Playback.MediaPlayer
    If IsNotNothing(MediaPlayer) Then
        MediaPlayer.AutoPlay = True
    End If
    Set SpeechSynthesizer = Windows.Media.SpeechSynthesis.SpeechSynthesizer
    If IsNotNothing(SpeechSynthesizer) Then
        Dim DefaultVoiceId As String
        DefaultVoiceId = SpeechSynthesizer.DefaultVoice.Id
        Set SpeechSynthesizerOptions = SpeechSynthesizer.Options
        If IsNotNothing(SpeechSynthesizerOptions) Then
            sbVolume.Value = SpeechSynthesizerOptions.AudioVolume * 100
            sbPitch.Value = SpeechSynthesizerOptions.AudioPitch * 100
            sbSpeekingRate.Value = SpeechSynthesizerOptions.SpeakingRate * 100
            tbText.Text = "Dieser Text, wird über die WinRT SpeechSynthesizer " & _
                          "COM-Interfaces ausgegeben. Wer Fehler " & _
                          "im Code findet, darf diese gern verbessern. ;-) Die Stimme hört sich " & _
                          "etwas natürlicher an, wenn man den Pitch so zwischen 0,7 und 0,8 einstellt."
            tbSsml.Text = "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='de-DE'>" & vbNewLine & _
                          "Hallo <prosody contour='(0%,+80Hz) (10%,+80%) (40%,+80Hz)'>Benutzer.</prosody> " & vbNewLine & _
                           "<break time='500ms'/>" & vbNewLine & _
                          "Auf Wiedersehen. <prosody rate='slow' contour='(0%,+20Hz) (10%,+30%) (40%,+10Hz)'>Bis zum nächsten mal.</prosody>" & vbNewLine & _
                          "</speak>"
            SpeechSynthesizerOptions.AppendedSilence = SpeechAppendedSilence_Min
            SpeechSynthesizerOptions.PunctuationSilence = SpeechPunctuationSilence_Min
        End If
        
        Set AllVoices = SpeechSynthesizer.AllVoices
        If IsNotNothing(AllVoices) Then
            Dim AllVoicesCount As Long
            AllVoicesCount = AllVoices.Size
            If AllVoicesCount > 0 Then
                Dim AllVoicesItem As Long
                Dim DefaultVoiceItem As Long
                For AllVoicesItem = 0 To AllVoicesCount - 1
                    Dim VoiceInformation As VoiceInformation
                    Set VoiceInformation = AllVoices.GetAt(AllVoicesItem)
                    If IsNotNothing(VoiceInformation) Then
                        If VoiceInformation.Id = DefaultVoiceId Then DefaultVoiceItem = AllVoicesItem
                        Dim strGender As String
                        Select Case VoiceInformation.Gender
                            Case VoiceGender.VoiceGender_Male
                                strGender = ", Male"
                            Case VoiceGender.VoiceGender_Female
                                strGender = ", Female"
                        End Select
                        Call cbVoices.AddItem(Replace$(VoiceInformation.Description, ")", strGender & ")"))
                        Set VoiceInformation = Nothing
                    End If
                Next
                cbVoices.ListIndex = DefaultVoiceItem
            End If
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsNotNothing(MediaPlayer) Then
        MediaPlayer.Source = Nothing
        Set MediaPlayer = Nothing
    End If
    If IsNotNothing(MediaSource) Then Set MediaSource = Nothing
    If IsNotNothing(RandomAccessStream) Then Set RandomAccessStream = Nothing
    If IsNotNothing(SpeechSynthesisStream) Then Set SpeechSynthesisStream = Nothing
    If IsNotNothing(AllVoices) Then Set AllVoices = Nothing
    If IsNotNothing(SpeechSynthesizerOptions) Then Set SpeechSynthesizerOptions = Nothing
    If IsNotNothing(SpeechSynthesizer) Then Set SpeechSynthesizer = Nothing
End Sub

Private Sub cmdSythesizeText_Click()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.Source = Nothing
    If IsNotNothing(MediaSource) Then Set MediaSource = Nothing
    If IsNotNothing(RandomAccessStream) Then Set RandomAccessStream = Nothing
    If IsNotNothing(SpeechSynthesisStream) Then Set SpeechSynthesisStream = Nothing
    Set SpeechSynthesisStream = SpeechSynthesizer.SynthesizeTextToStreamAsync(tbText.Text)
    If IsNotNothing(SpeechSynthesisStream) Then
        Set RandomAccessStream = SpeechSynthesisStream.RandomAccessStream
        If IsNotNothing(RandomAccessStream) Then
            Set MediaSource = Windows.Media.Core.MediaSource.CreateFromStream(RandomAccessStream, vbNullString)
            If IsNotNothing(MediaSource) Then
                MediaPlayer.Source = MediaSource
            End If
        End If
    End If
End Sub

Private Sub cmdSynthesizeSsml_Click()
    If IsNotNothing(MediaPlayer) Then MediaPlayer.Source = Nothing
    If IsNotNothing(MediaSource) Then Set MediaSource = Nothing
    If IsNotNothing(RandomAccessStream) Then Set RandomAccessStream = Nothing
    If IsNotNothing(SpeechSynthesisStream) Then Set SpeechSynthesisStream = Nothing
    Set SpeechSynthesisStream = SpeechSynthesizer.SynthesizeSsmlToStreamAsync(tbSsml.Text)
    If IsNotNothing(SpeechSynthesisStream) Then
        Set RandomAccessStream = SpeechSynthesisStream.RandomAccessStream
        If IsNotNothing(RandomAccessStream) Then
            Set MediaSource = Windows.Media.Core.MediaSource.CreateFromStream(RandomAccessStream, vbNullString)
            If IsNotNothing(MediaSource) Then
                MediaPlayer.Source = MediaSource
            End If
        End If
    End If
End Sub

Private Sub cmdPauseSpeech_Click()
    Call MediaPlayer.Pause
End Sub

Private Sub cmdPlaySpeech_Click()
    Call MediaPlayer.Play
End Sub

Private Sub cbVoices_Click()
    If IsNotNothing(SpeechSynthesizer) Then
        SpeechSynthesizer.voice = AllVoices.GetAt(cbVoices.ListIndex)
    End If
End Sub

Private Sub sbVolume_Change()
    If IsNotNothing(SpeechSynthesizerOptions) Then
        SpeechSynthesizerOptions.AudioVolume = CDbl(sbVolume.Value / 100)
        frOptions(0).Caption = "Volume: " & Format$(CDbl(sbVolume.Value / 100), "0.00")
    End If
End Sub

Private Sub sbVolume_Scroll()
    If IsNotNothing(SpeechSynthesizerOptions) Then
        SpeechSynthesizerOptions.AudioVolume = CDbl(sbVolume.Value / 100)
        frOptions(0).Caption = "Volume: " & Format$(CDbl(sbVolume.Value / 100), "0.00")
    End If
End Sub

Private Sub sbPitch_Change()
    If IsNotNothing(SpeechSynthesizerOptions) Then
        SpeechSynthesizerOptions.AudioPitch = CDbl(sbPitch.Value / 100)
        frOptions(1).Caption = "Pitch: " & Format$(CDbl(sbPitch.Value / 100), "0.00")
    End If
End Sub

Private Sub sbPitch_Scroll()
    If IsNotNothing(SpeechSynthesizerOptions) Then
        SpeechSynthesizerOptions.AudioPitch = CDbl(sbPitch.Value / 100)
        frOptions(1).Caption = "Pitch: " & Format$(CDbl(sbPitch.Value / 100), "0.00")
    End If
End Sub

Private Sub sbSpeekingRate_Change()
    If IsNotNothing(SpeechSynthesizerOptions) Then
        SpeechSynthesizerOptions.SpeakingRate = CDbl(sbSpeekingRate.Value / 100)
        frOptions(2).Caption = "SpeekingRate: " & Format$(CDbl(sbSpeekingRate.Value / 100), "0.00")
    End If
End Sub

Private Sub sbSpeekingRate_Scroll()
    If IsNotNothing(SpeechSynthesizerOptions) Then
        SpeechSynthesizerOptions.SpeakingRate = CDbl(sbSpeekingRate.Value / 100)
        frOptions(2).Caption = "SpeekingRate: " & Format$(CDbl(sbSpeekingRate.Value / 100), "0.00")
    End If
End Sub

