VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

' Namespace SpeechRecognition

Option Explicit

' ----==== Enums Windows.Media.SpeechRecognition.xxx ====----
Public Enum SpeechRecognizerState
    SpeechRecognizerState_Idle = 0
    SpeechRecognizerState_Capturing = 1
    SpeechRecognizerState_Processing = 2
    SpeechRecognizerState_SoundStarted = 3
    SpeechRecognizerState_SoundEnded = 4
    SpeechRecognizerState_SpeechDetected = 5
    SpeechRecognizerState_Paused = 6
End Enum

Public Enum SpeechRecognitionResultStatus
    SpeechRecognitionResultStatus_Success = 0
    SpeechRecognitionResultStatus_TopicLanguageNotSupported = 1
    SpeechRecognitionResultStatus_GrammarLanguageMismatch = 2
    SpeechRecognitionResultStatus_GrammarCompilationFailure = 3
    SpeechRecognitionResultStatus_AudioQualityFailure = 4
    SpeechRecognitionResultStatus_UserCanceled = 5
    SpeechRecognitionResultStatus_Unknown = 6
    SpeechRecognitionResultStatus_TimeoutExceeded = 7
    SpeechRecognitionResultStatus_PauseLimitExceeded = 8
    SpeechRecognitionResultStatus_NetworkFailure = 9
    SpeechRecognitionResultStatus_MicrophoneUnavailable = 10
End Enum

Public Enum SpeechRecognitionConfidence
    SpeechRecognitionConfidence_High = 0
    SpeechRecognitionConfidence_Medium = 1
    SpeechRecognitionConfidence_Low = 2
    SpeechRecognitionConfidence_Rejected = 3
End Enum

Public Enum SpeechRecognitionConstraintType
    SpeechRecognitionConstraintType_Topic = 0
    SpeechRecognitionConstraintType_List = 1
    SpeechRecognitionConstraintType_Grammar = 2
    SpeechRecognitionConstraintType_VoiceCommandDefinition = 3
End Enum

Public Enum SpeechRecognitionConstraintProbability
    SpeechRecognitionConstraintProbability_Default = 0
    SpeechRecognitionConstraintProbability_Min = 1
    SpeechRecognitionConstraintProbability_Max = 2
End Enum

Public Enum SpeechContinuousRecognitionMode
    SpeechContinuousRecognitionMode_Default = 0
    SpeechContinuousRecognitionMode_PauseOnRecognition = 1
End Enum

Public Enum SpeechRecognitionScenario
    SpeechRecognitionScenario_WebSearch = 0
    SpeechRecognitionScenario_Dictation = 1
    SpeechRecognitionScenario_FormFilling = 2
End Enum

' ----==== Namespace SpeechRecognition.SpeechRecognizer ====----
Public Property Get SpeechRecognizer() As SpeechRecognizer
    Set SpeechRecognizer = New SpeechRecognizer
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognizerTimeouts ====----
Public Property Get SpeechRecognizerTimeouts() As SpeechRecognizerTimeouts
    Set SpeechRecognizerTimeouts = New SpeechRecognizerTimeouts
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognizerUIOptions ====----
Public Property Get SpeechRecognizerUIOptions() As SpeechRecognizerUIOptions
    Set SpeechRecognizerUIOptions = New SpeechRecognizerUIOptions
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognitionCompilationResult ====----
Public Property Get SpeechRecognitionCompilationResult() As SpeechRecognitionCompilationResult
    Set SpeechRecognitionCompilationResult = New SpeechRecognitionCompilationResult
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognitionResult ====----
Public Property Get SpeechRecognitionResult() As SpeechRecognitionResult
    Set SpeechRecognitionResult = New SpeechRecognitionResult
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognitionConstraint ====----
Public Property Get SpeechRecognitionConstraint() As SpeechRecognitionConstraint
    Set SpeechRecognitionConstraint = New SpeechRecognitionConstraint
End Property

' ----==== Namespace SpeechRecognition.SpeechContinuousRecognitionSession ====----
Public Property Get SpeechContinuousRecognitionSession() As SpeechContinuousRecognitionSession
    Set SpeechContinuousRecognitionSession = New SpeechContinuousRecognitionSession
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognitionSemanticInterpretation ====----
Public Property Get SpeechRecognitionSemanticInterpretation() As SpeechRecognitionSemanticInterpretation
    Set SpeechRecognitionSemanticInterpretation = New SpeechRecognitionSemanticInterpretation
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognitionTopicConstraint ====----
Public Property Get SpeechRecognitionTopicConstraint() As SpeechRecognitionTopicConstraint
    Set SpeechRecognitionTopicConstraint = New SpeechRecognitionTopicConstraint
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognitionListConstraint ====----
Public Property Get SpeechRecognitionListConstraint() As SpeechRecognitionListConstraint
    Set SpeechRecognitionListConstraint = New SpeechRecognitionListConstraint
End Property

' ----==== Namespace SpeechRecognition.SpeechRecognitionGrammarFileConstraint ====----
Public Property Get SpeechRecognitionGrammarFileConstraint() As SpeechRecognitionGrammarFileConstraint
    Set SpeechRecognitionGrammarFileConstraint = New SpeechRecognitionGrammarFileConstraint
End Property

