VERSION 5.00
Begin VB.Form frmSpeechRecognizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpeechRecognizer"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRecognizerUI 
      BackColor       =   &H0000FF00&
      Caption         =   "Start SpeechRecognizerUI"
      Height          =   555
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   3000
   End
   Begin VB.TextBox txtResult 
      Height          =   4245
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   660
      Width           =   6015
   End
   Begin VB.CommandButton cmdContinuousRecognizer 
      BackColor       =   &H0000FF00&
      Caption         =   "Start ContinuousSpeechRecognizer"
      Height          =   555
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   3000
   End
End
Attribute VB_Name = "frmSpeechRecognizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' Namespace Windows
Private Windows As New Windows

' Namespace Windows.Media.SpeechRecognition
Private SpeechRecognizer As SpeechRecognizer
Private SpeechContinuousRecognitionSession As SpeechContinuousRecognitionSession

Private m_IsRecognize As Boolean

' ----==== for SpeechRecognizer Events ====---
Private m_pISpeechRecognizerStateChangedEvent As Long
Private m_SpeechRecognizerStateChangedEvent_Token As Currency

' ----==== for SpeechContinuousRecognitionSession Events ====---
Private m_pISpeechContinuousRecognitionCompletedEvent As Long
Private m_SpeechContinuousRecognitionCompletedEventToken As Currency
Private m_pISpeechContinuousRecognitionResultGeneratedEvent As Long
Private m_SpeechContinuousRecognitionResultGeneratedEventToken As Currency

Private Sub Form_Load()
    m_pISpeechRecognizerStateChangedEvent = ITEH_SpeechRecognizerStateChanged.Create(Me)
    m_pISpeechContinuousRecognitionCompletedEvent = ITEH_SpeechContinuousRecognitionCompleted.Create(Me)
    m_pISpeechContinuousRecognitionResultGeneratedEvent = ITEH_SpeechContinuousRecognitionResultGenerated.Create(Me)

    Set SpeechRecognizer = Windows.Media.SpeechRecognition.SpeechRecognizer
    If IsNotNothing(SpeechRecognizer) Then

        Dim SpeechRecognizerTimeouts As SpeechRecognizerTimeouts
        Set SpeechRecognizerTimeouts = SpeechRecognizer.Timeouts
        If IsNotNothing(SpeechRecognizerTimeouts) Then
            Debug.Print "Default InitialSilenceTimeout: " & SpeechRecognizerTimeouts.InitialSilenceTimeout.VbDate
            Debug.Print "Default EndSilenceTimeout: " & SpeechRecognizerTimeouts.EndSilenceTimeout.VbDate
            Debug.Print "Default BabbleTimeout: " & SpeechRecognizerTimeouts.BabbleTimeout.VbDate
        End If

        Dim SystemSpeechLanguage As Language
        Set SystemSpeechLanguage = SpeechRecognizer.SystemSpeechLanguage
        If IsNotNothing(SystemSpeechLanguage) Then
            Debug.Print "SystemSpeechLanguage: " & SystemSpeechLanguage.DisplayName
        End If

        Dim CurrentLanguage As Language
        Set CurrentLanguage = SpeechRecognizer.CurrentLanguage
        If IsNotNothing(CurrentLanguage) Then
            Debug.Print "CurrentSpeechRecognizerLanguage: " & CurrentLanguage.DisplayName
        End If

        Dim SupportedTopicLanguages As ReadOnlyList_1 'ReadOnlyList_Language
        Set SupportedTopicLanguages = SpeechRecognizer.SupportedTopicLanguages
        If IsNotNothing(SupportedTopicLanguages) Then
            Dim SupportedTopicLanguagesCount As Long
            SupportedTopicLanguagesCount = SupportedTopicLanguages.Size
            Debug.Print "SupportedTopicLanguages Count: " & SupportedTopicLanguages.Size
            If SupportedTopicLanguagesCount > 0& Then
                Dim SupportedTopicLanguagesItem As Long
                For SupportedTopicLanguagesItem = 0 To SupportedTopicLanguagesCount - 1
                    Debug.Print vbTab & SupportedTopicLanguages.GetAt(SupportedTopicLanguagesItem).DisplayName
                Next
            End If
        End If

        Dim SupportedGrammarLanguages As ReadOnlyList_1 'ReadOnlyList_Language
        Set SupportedGrammarLanguages = SpeechRecognizer.SupportedTopicLanguages
        If IsNotNothing(SupportedGrammarLanguages) Then
            Dim SupportedGrammarLanguagesCount As Long
            SupportedGrammarLanguagesCount = SupportedGrammarLanguages.Size
            Debug.Print "SupportedGrammarLanguages Count: " & SupportedGrammarLanguages.Size
            If SupportedGrammarLanguagesCount > 0& Then
                Dim SupportedGrammarLanguagesItem As Long
                For SupportedGrammarLanguagesItem = 0 To SupportedGrammarLanguagesCount - 1
                    Debug.Print vbTab & SupportedGrammarLanguages.GetAt(SupportedGrammarLanguagesItem).DisplayName
                Next
            End If
        End If
        Set SpeechRecognizer = Nothing
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsNotNothing(SpeechContinuousRecognitionSession) Then
        Call SpeechContinuousRecognitionSession.StopAsync
        If m_SpeechContinuousRecognitionCompletedEventToken <> 0& Then
            If SpeechContinuousRecognitionSession.RemoveCompleted(m_SpeechContinuousRecognitionCompletedEventToken) Then
                Call ITEH_SpeechContinuousRecognitionCompleted.Destroy
            End If
        End If
        If m_SpeechContinuousRecognitionResultGeneratedEventToken <> 0& Then
            If SpeechContinuousRecognitionSession.RemoveResultGenerated(m_SpeechContinuousRecognitionResultGeneratedEventToken) Then
                Call ITEH_SpeechContinuousRecognitionResultGenerated.Destroy
            End If
        End If
        Set SpeechContinuousRecognitionSession = Nothing
    End If
    If IsNotNothing(SpeechRecognizer) Then
        Call SpeechRecognizer.StopRecognitionAsync
        If m_SpeechRecognizerStateChangedEvent_Token <> 0& Then
            If SpeechRecognizer.RemoveStateChanged(m_SpeechRecognizerStateChangedEvent_Token) Then
                Call ITEH_SpeechRecognizerStateChanged.Destroy
            End If
        End If
        Set SpeechRecognizer = Nothing
    End If
End Sub

Private Sub cmdRecognizerUI_Click()
    Set SpeechRecognizer = Nothing
    Set SpeechRecognizer = Windows.Media.SpeechRecognition.SpeechRecognizer
    If IsNotNothing(SpeechRecognizer) Then
        m_SpeechRecognizerStateChangedEvent_Token = SpeechRecognizer.AddStateChanged(m_pISpeechRecognizerStateChangedEvent)
        If Not m_IsRecognize Then
            m_IsRecognize = True
            With cmdRecognizerUI
                .BackColor = vbRed
                .Caption = "Please say something."
            End With
            With cmdContinuousRecognizer
                .BackColor = vbYellow
                .Enabled = False
            End With
        
        
            Dim SpeechRecognizerUIOptions As SpeechRecognizerUIOptions
            Set SpeechRecognizerUIOptions = SpeechRecognizer.UIOptions
            If IsNotNothing(SpeechRecognizerUIOptions) Then
                SpeechRecognizerUIOptions.IsReadBackEnabled = True
                SpeechRecognizerUIOptions.ShowConfirmation = True
            End If
        
'            Dim SpeechRecognizerTimeouts As SpeechRecognizerTimeouts
'            Set SpeechRecognizerTimeouts = SpeechRecognizer.Timeouts
'            If IsNotNothing(SpeechRecognizerTimeouts) Then
'                SpeechRecognizerTimeouts.InitialSilenceTimeout = Windows.Foundation.TimeSpan.FromSeconds(10)
'                SpeechRecognizerTimeouts.EndSilenceTimeout = Windows.Foundation.TimeSpan.FromSeconds(10)
'                SpeechRecognizerTimeouts.BabbleTimeout = Windows.Foundation.TimeSpan.FromSeconds(10)
'            End If
            
            Dim SpeechRecognitionCompilationResult As SpeechRecognitionCompilationResult
            Set SpeechRecognitionCompilationResult = SpeechRecognizer.CompileConstraintsAsync
            If IsNotNothing(SpeechRecognitionCompilationResult) Then
                Dim Status As SpeechRecognitionResultStatus
                Status = SpeechRecognitionCompilationResult.Status
                Debug.Print GetSpeechRecognitionResultStatusText(Status)
                If Status = SpeechRecognitionResultStatus_Success Then
                    Dim SpeechRecognitionResult As SpeechRecognitionResult
                    Set SpeechRecognitionResult = SpeechRecognizer.RecognizeWithUIAsync
                    If IsNotNothing(SpeechRecognitionResult) Then
                        Call GetSpeechRecognitionResult(SpeechRecognitionResult)
                    Else
                        Debug.Print "Please activate online speech recognition: Windows Settings -> Speech recognition -> Online speech recognition -> On"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdContinuousRecognizer_Click()
    Set SpeechRecognizer = Nothing
    Set SpeechRecognizer = Windows.Media.SpeechRecognition.SpeechRecognizer
    If IsNotNothing(SpeechRecognizer) Then
    
        If Not m_IsRecognize Then
            m_IsRecognize = True
            With cmdContinuousRecognizer
                .BackColor = vbRed
                .Caption = "Stop ContinuousSpeechRecognizer"
            End With
            With cmdRecognizerUI
                .BackColor = vbYellow
                .Enabled = False
            End With
            
           Dim SpeechRecognizerTimeouts As SpeechRecognizerTimeouts
           Set SpeechRecognizerTimeouts = SpeechRecognizer.Timeouts
           If IsNotNothing(SpeechRecognizerTimeouts) Then
               SpeechRecognizerTimeouts.InitialSilenceTimeout = Windows.Foundation.TimeSpan.FromSeconds(10)
               SpeechRecognizerTimeouts.EndSilenceTimeout = Windows.Foundation.TimeSpan.FromSeconds(10)
               SpeechRecognizerTimeouts.BabbleTimeout = Windows.Foundation.TimeSpan.FromSeconds(10)
           End If
    
            Dim SpeechRecognitionCompilationResult As SpeechRecognitionCompilationResult
            Set SpeechRecognitionCompilationResult = SpeechRecognizer.CompileConstraintsAsync
            If IsNotNothing(SpeechRecognitionCompilationResult) Then
                Dim Status As SpeechRecognitionResultStatus
                Status = SpeechRecognitionCompilationResult.Status
                Debug.Print GetSpeechRecognitionResultStatusText(Status)
                If Status = SpeechRecognitionResultStatus_Success Then
                    Set SpeechContinuousRecognitionSession = SpeechRecognizer.ContinuousRecognitionSession
                    If IsNotNothing(SpeechContinuousRecognitionSession) Then
                        m_SpeechContinuousRecognitionCompletedEventToken = SpeechContinuousRecognitionSession.AddCompleted(m_pISpeechContinuousRecognitionCompletedEvent)
                        m_SpeechContinuousRecognitionResultGeneratedEventToken = SpeechContinuousRecognitionSession.AddResultGenerated(m_pISpeechContinuousRecognitionResultGeneratedEvent)
                        If Not SpeechContinuousRecognitionSession.StartAsync Then
                            Debug.Print "Please activate online speech recognition: Windows Settings -> Speech recognition -> Online speech recognition -> On"
                            Call StopContinuousSpeechRecognizer
                        End If
                    End If
                End If
            End If
        Else
            Call StopContinuousSpeechRecognizer
        End If
    End If
End Sub
    
Private Sub StopContinuousSpeechRecognizer()
    m_IsRecognize = False
    With cmdContinuousRecognizer
        .BackColor = vbGreen
        .Caption = "Start ContinuousSpeechRecognizer"
    End With
    With cmdRecognizerUI
        .BackColor = vbGreen
        .Enabled = True
    End With
    If IsNotNothing(SpeechContinuousRecognitionSession) Then
        If m_SpeechContinuousRecognitionCompletedEventToken <> 0& Then
            If SpeechContinuousRecognitionSession.RemoveCompleted(m_SpeechContinuousRecognitionCompletedEventToken) Then
                m_SpeechContinuousRecognitionCompletedEventToken = 0&
            End If
            If m_SpeechContinuousRecognitionResultGeneratedEventToken <> 0& Then
                If SpeechContinuousRecognitionSession.RemoveResultGenerated(m_SpeechContinuousRecognitionResultGeneratedEventToken) Then
                    m_SpeechContinuousRecognitionResultGeneratedEventToken = 0&
                End If
                Set SpeechContinuousRecognitionSession = Nothing
            End If
            If IsNotNothing(SpeechRecognizer) Then
                Set SpeechRecognizer = Nothing
            End If
        End If
    End If
End Sub

Private Sub GetSpeechRecognitionResult(ByVal Result As SpeechRecognitionResult)
    If IsNotNothing(Result) Then
        Debug.Print "Status: " & GetSpeechRecognitionResultStatusText(Result.Status)
        Debug.Print "Confidence: " & Result.Confidence
        Debug.Print "PhraseStartTime: " & Result.PhraseStartTime.VbDate
        Debug.Print "PhraseDuration: " & Result.PhraseDuration.VbDate
        Debug.Print "RawConfidence: " & Result.RawConfidence
        
        Dim RulePaths As ReadOnlyList_1 'ReadOnlyList_String
        Set RulePaths = Result.RulePath
        If IsNotNothing(RulePaths) Then
            Debug.Print "RulePaths: " & RulePaths.Size
            Debug.Print "RulePath(0): " & RulePaths.GetAt(0)
        End If
        
        Dim Alternates As ReadOnlyList_1 'ReadOnlyList_SpeechRecognitionResult
        Set Alternates = Result.GetAlternates(5)
        If IsNotNothing(Alternates) Then
            Dim AlternatesCount As Long
            AlternatesCount = Alternates.Size
            Debug.Print "Alternates: " & AlternatesCount
            If AlternatesCount > 0& Then
                Dim AlternatesItem As Long
                For AlternatesItem = 0 To AlternatesCount - 1
                    Debug.Print "Alternates(" & CStr(AlternatesItem) & ").Text: " & Alternates.GetAt(AlternatesItem).Text
                Next
            End If
        End If
        
        Dim RecognitionResult As String
        RecognitionResult = Result.Text
        If RecognitionResult <> vbNullString Then
            txtResult.Text = txtResult.Text & RecognitionResult & vbNewLine
            txtResult.SelStart = Len(txtResult.Text)
        End If
    End If
End Sub

Private Function GetSpeechRecognitionResultStatusText(ByVal Status As SpeechRecognitionResultStatus) As String
    Dim Ret As String
    Select Case Status
    Case SpeechRecognitionResultStatus_Success
        Ret = "The recognition session or compilation succeeded."
    Case SpeechRecognitionResultStatus_TopicLanguageNotSupported
        Ret = "A topic constraint was set for an unsupported language."
    Case SpeechRecognitionResultStatus_GrammarLanguageMismatch
        Ret = "The language of the speech recognizer does not match the language of a grammar."
    Case SpeechRecognitionResultStatus_GrammarCompilationFailure
        Ret = "A grammar failed to compile."
    Case SpeechRecognitionResultStatus_AudioQualityFailure
        Ret = "Audio problems caused recognition to fail."
    Case SpeechRecognitionResultStatus_UserCanceled
        Ret = "User canceled recognition session."
    Case SpeechRecognitionResultStatus_Unknown
        Ret = "An unknown problem caused recognition or compilation to fail."
    Case SpeechRecognitionResultStatus_TimeoutExceeded
        Ret = "A timeout due to extended silence or poor audio caused recognition to fail."
    Case SpeechRecognitionResultStatus_PauseLimitExceeded
        Ret = "An extended pause, or excessive processing time, caused recognition to fail."
    Case SpeechRecognitionResultStatus_NetworkFailure
        Ret = "Network problems caused recognition to fail."
    Case SpeechRecognitionResultStatus_MicrophoneUnavailable
        Ret = "Lack of a microphone caused recognition to fail."
    End Select
    GetSpeechRecognitionResultStatusText = Ret
End Function

Private Function GetSpeechRecognizerStateText(ByVal State As SpeechRecognizerState) As String
    Dim Ret As String
    Select Case State
        Case SpeechRecognizerState_Idle
            Ret = "Indicates that speech recognition is not active and the speech recognizer is not capturing (listening for) audio input."
        Case SpeechRecognizerState_Capturing
            Ret = "Indicates that the speech recognizer is capturing (listening for) audio input from the user."
        Case SpeechRecognizerState_Processing
            Ret = "Indicates that the speech recognizer is processing (attempting to recognize) audio input from the user. The recognizer is no longer capturing (listening for) audio input from the user."
        Case SpeechRecognizerState_SoundStarted
            Ret = "Indicates that the speech recognizer has detected sound on the audio stream."
        Case SpeechRecognizerState_SoundEnded
            Ret = "Indicates that the speech recognizer no longer detects sound on the audio stream."
        Case SpeechRecognizerState_SpeechDetected
            Ret = "Indicates that the speech recognizer has detected speech input on the audio stream."
        Case SpeechRecognizerState_Paused
            Ret = "Indicates that the speech recognition session is still active, but the speech recognizer is no longer processing (attempting to recognize) audio input. Ongoing audio input is buffered."
    End Select
    GetSpeechRecognizerStateText = Ret
End Function

' ----==== SpeechRecognizer Events ====---
Public Sub SpeechRecognizerStateChangedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = SpeechRecognizer.Ifc And args <> 0& Then
        Dim SpeechRecognizerStateChangedEventArgs As New SpeechRecognizerStateChangedEventArgs
        SpeechRecognizerStateChangedEventArgs.Ifc = args
        Dim State As SpeechRecognizerState
        State = SpeechRecognizerStateChangedEventArgs.State
        Debug.Print GetSpeechRecognizerStateText(State)
        If State = SpeechRecognizerState_Idle Then
            m_IsRecognize = False
            With cmdRecognizerUI
                .BackColor = vbGreen
                .Caption = "Start SpeechRecognizerUI"
            End With
            With cmdContinuousRecognizer
                .BackColor = vbGreen
                .Enabled = True
            End With
            If IsNotNothing(SpeechRecognizer) Then
                Call SpeechRecognizer.StopRecognitionAsync
                If m_SpeechRecognizerStateChangedEvent_Token <> 0& Then
                    If SpeechRecognizer.RemoveStateChanged(m_SpeechRecognizerStateChangedEvent_Token) Then
                        m_SpeechRecognizerStateChangedEvent_Token = 0&
                    End If
                End If
                Set SpeechRecognizer = Nothing
            End If
        End If
    End If
End Sub

' ----==== SpeechContinuousRecognitionSession Events ====---
Public Sub SpeechContinuousRecognitionCompletedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = SpeechContinuousRecognitionSession.Ifc Then
        If args <> 0& Then
            Dim SpeechContinuousRecognitionCompletedEventArgs As New SpeechContinuousRecognitionCompletedEventArgs
            SpeechContinuousRecognitionCompletedEventArgs.Ifc = args
            Dim Status As SpeechRecognitionResultStatus
            Status = SpeechContinuousRecognitionCompletedEventArgs.Status
            Debug.Print GetSpeechRecognitionResultStatusText(Status)
            If Status = SpeechRecognitionResultStatus_TimeoutExceeded Then
                Call StopContinuousSpeechRecognizer
            End If
        End If
    End If
End Sub

Public Sub SpeechContinuousRecognitionResultGeneratedEvent(ByVal sender As Long, ByVal args As Long)
    If sender = SpeechContinuousRecognitionSession.Ifc Then
        If args <> 0& Then
            Dim SpeechContinuousRecognitionResultGeneratedEventArgs As New SpeechContinuousRecognitionResultGeneratedEventArgs
            Dim SpeechRecognitionResult As SpeechRecognitionResult
            SpeechContinuousRecognitionResultGeneratedEventArgs.Ifc = args
            Set SpeechRecognitionResult = SpeechContinuousRecognitionResultGeneratedEventArgs.Result
            If IsNotNothing(SpeechRecognitionResult) Then
                Call GetSpeechRecognitionResult(SpeechRecognitionResult)
            End If
        End If
    End If
End Sub
