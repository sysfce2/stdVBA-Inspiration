VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechSynthesizerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechSynthesis_SpeechSynthesizerOptions As String = "Windows.Media.SpeechSynthesis.SpeechSynthesizerOptions"
Private Const IID_ISpeechSynthesizerOptions As String = "{a0e23871-cc3d-43c9-91b1-ee185324d83d}"
Private Const IID_ISpeechSynthesizerOptions2 As String = "{1cbef60e-119c-4bed-b118-d250c3a25793}"
Private Const IID_ISpeechSynthesizerOptions3 As String = "{401ed877-902c-4814-a582-a5d0c0769fa8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechSynthesizerOptions
    ISpeechSynthesizerOptions_GetIncludeWordBoundaryMetadata = 6
    ISpeechSynthesizerOptions_PutIncludeWordBoundaryMetadata = 7
    ISpeechSynthesizerOptions_GetIncludeSentenceBoundaryMetadata = 8
    ISpeechSynthesizerOptions_PutIncludeSentenceBoundaryMetadata = 9
                        
    ' ISpeechSynthesizerOptions2
    ISpeechSynthesizerOptions2_GetAudioVolume = 6
    ISpeechSynthesizerOptions2_PutAudioVolume = 7
    ISpeechSynthesizerOptions2_GetSpeakingRate = 8
    ISpeechSynthesizerOptions2_PutSpeakingRate = 9
    ISpeechSynthesizerOptions2_GetAudioPitch = 10
    ISpeechSynthesizerOptions2_PutAudioPitch = 11
                        
    ' ISpeechSynthesizerOptions3
    ISpeechSynthesizerOptions3_GetAppendedSilence = 6
    ISpeechSynthesizerOptions3_PutAppendedSilence = 7
    ISpeechSynthesizerOptions3_GetPunctuationSilence = 8
    ISpeechSynthesizerOptions3_PutPunctuationSilence = 9
    
End Enum

' ----==== Variablen ====----
Private m_pISpeechSynthesizerOptions As Long
Private m_pISpeechSynthesizerOptions2 As Long
Private m_pISpeechSynthesizerOptions3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechSynthesizerOptions3)
    Call ReleaseIfc(m_pISpeechSynthesizerOptions2)
    Call ReleaseIfc(m_pISpeechSynthesizerOptions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechSynthesizerOptions
End Property

Public Property Let Ifc(ByVal pISpeechSynthesizerOptions As Long)
    Call ReleaseIfc(m_pISpeechSynthesizerOptions3)
    Call ReleaseIfc(m_pISpeechSynthesizerOptions2)
    Call ReleaseIfc(m_pISpeechSynthesizerOptions)
    m_pISpeechSynthesizerOptions = pISpeechSynthesizerOptions
    Call QueryIfc(m_pISpeechSynthesizerOptions, IID_ISpeechSynthesizerOptions2, m_pISpeechSynthesizerOptions2)
    Call QueryIfc(m_pISpeechSynthesizerOptions, IID_ISpeechSynthesizerOptions3, m_pISpeechSynthesizerOptions3)
End Property

' ISpeechSynthesizerOptions
Public Property Get IncludeWordBoundaryMetadata() As Boolean
    Dim Ret As Boolean
    If m_pISpeechSynthesizerOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechSynthesizerOptions, _
                  ISpeechSynthesizerOptions_GetIncludeWordBoundaryMetadata, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IncludeWordBoundaryMetadata = Ret
End Property

Public Property Let IncludeWordBoundaryMetadata(ByVal value As Boolean)
    If m_pISpeechSynthesizerOptions <> 0& Then
        If Invoke(m_pISpeechSynthesizerOptions, _
                  ISpeechSynthesizerOptions_PutIncludeWordBoundaryMetadata, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IncludeSentenceBoundaryMetadata() As Boolean
    Dim Ret As Boolean
    If m_pISpeechSynthesizerOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechSynthesizerOptions, _
                  ISpeechSynthesizerOptions_GetIncludeSentenceBoundaryMetadata, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IncludeSentenceBoundaryMetadata = Ret
End Property

Public Property Let IncludeSentenceBoundaryMetadata(ByVal value As Boolean)
    If m_pISpeechSynthesizerOptions <> 0& Then
        If Invoke(m_pISpeechSynthesizerOptions, _
                  ISpeechSynthesizerOptions_PutIncludeSentenceBoundaryMetadata, _
                  value) = S_OK Then
        End If
    End If
End Property
                        
' ISpeechSynthesizerOptions2
Public Property Get AudioVolume() As Double
    Dim Ret As Double
    If m_pISpeechSynthesizerOptions2 <> 0& Then
        Dim value As Double
        If Invoke(m_pISpeechSynthesizerOptions2, _
                  ISpeechSynthesizerOptions2_GetAudioVolume, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AudioVolume = Ret
End Property

Public Property Let AudioVolume(ByVal value As Double)
    If m_pISpeechSynthesizerOptions2 <> 0& Then
        If Invoke(m_pISpeechSynthesizerOptions2, _
                  ISpeechSynthesizerOptions2_PutAudioVolume, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get SpeakingRate() As Double
    Dim Ret As Double
    If m_pISpeechSynthesizerOptions2 <> 0& Then
        Dim value As Double
        If Invoke(m_pISpeechSynthesizerOptions2, _
                  ISpeechSynthesizerOptions2_GetSpeakingRate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SpeakingRate = Ret
End Property

Public Property Let SpeakingRate(ByVal value As Double)
    If m_pISpeechSynthesizerOptions2 <> 0& Then
        If Invoke(m_pISpeechSynthesizerOptions2, _
                  ISpeechSynthesizerOptions2_PutSpeakingRate, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get AudioPitch() As Double
    Dim Ret As Double
    If m_pISpeechSynthesizerOptions2 <> 0& Then
        Dim value As Double
        If Invoke(m_pISpeechSynthesizerOptions2, _
                  ISpeechSynthesizerOptions2_GetAudioPitch, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AudioPitch = Ret
End Property

Public Property Let AudioPitch(ByVal value As Double)
    If m_pISpeechSynthesizerOptions2 <> 0& Then
        If Invoke(m_pISpeechSynthesizerOptions2, _
                  ISpeechSynthesizerOptions2_PutAudioPitch, _
                  value) = S_OK Then
        End If
    End If
End Property
                        
' ISpeechSynthesizerOptions3
Public Property Get AppendedSilence() As SpeechAppendedSilence
    Dim Ret As SpeechAppendedSilence
    If m_pISpeechSynthesizerOptions3 <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechSynthesizerOptions3, _
                  ISpeechSynthesizerOptions3_GetAppendedSilence, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AppendedSilence = Ret
End Property

Public Property Let AppendedSilence(ByVal value As SpeechAppendedSilence)
    If m_pISpeechSynthesizerOptions3 <> 0& Then
        If Invoke(m_pISpeechSynthesizerOptions3, _
                  ISpeechSynthesizerOptions3_PutAppendedSilence, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get PunctuationSilence() As SpeechPunctuationSilence
    Dim Ret As SpeechPunctuationSilence
    If m_pISpeechSynthesizerOptions3 <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechSynthesizerOptions3, _
                  ISpeechSynthesizerOptions3_GetPunctuationSilence, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PunctuationSilence = Ret
End Property

Public Property Let PunctuationSilence(ByVal value As SpeechPunctuationSilence)
    If m_pISpeechSynthesizerOptions3 <> 0& Then
        If Invoke(m_pISpeechSynthesizerOptions3, _
                  ISpeechSynthesizerOptions3_PutPunctuationSilence, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


