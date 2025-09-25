VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechSynthesis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

' Namespace SpeechSynthesis

Option Explicit

' ----==== Enums Windows.Media.SpeechSynthesis.xxx ====----
Public Enum VoiceGender
    VoiceGender_Male = 0
    VoiceGender_Female = 1
End Enum

Public Enum SpeechAppendedSilence
    SpeechAppendedSilence_Default = 0
    SpeechAppendedSilence_Min = 1
End Enum

Public Enum SpeechPunctuationSilence
    SpeechPunctuationSilence_Default = 0
    SpeechPunctuationSilence_Min = 1
End Enum

' ----==== Namespace SpeechSynthesis.SpeechSynthesizer ====----
Public Property Get SpeechSynthesizer() As SpeechSynthesizer
    Set SpeechSynthesizer = New SpeechSynthesizer
End Property

' ----==== Namespace SpeechSynthesis.VoiceInformation ====----
Public Property Get VoiceInformation() As VoiceInformation
    Set VoiceInformation = New VoiceInformation
End Property

' ----==== Namespace SpeechSynthesis.SpeechSynthesizerOptions ====----
Public Property Get SpeechSynthesizerOptions() As SpeechSynthesizerOptions
    Set SpeechSynthesizerOptions = New SpeechSynthesizerOptions
End Property

' ----==== Namespace SpeechSynthesis.SpeechSynthesisStream ====----
Public Property Get SpeechSynthesisStream() As SpeechSynthesisStream
    Set SpeechSynthesisStream = New SpeechSynthesisStream
End Property
