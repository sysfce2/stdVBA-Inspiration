VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognizer As String = "Windows.Media.SpeechRecognition.SpeechRecognizer"
Private Const IID_ISpeechRecognizer As String = "{0bc3c9cb-c26a-40f2-aeb5-8096b2e48073}"
Private Const IID_ISpeechRecognizer2 As String = "{63c9baf1-91e3-4ea4-86a1-7c3867d084a6}"
Private Const IID_ISpeechRecognizerStatic As String = "{87a35eac-a7dc-4b0b-bcc9-24f47c0b7ebf}"
Private Const IID_ISpeechRecognizerStatic2 As String = "{1d1b0d95-7565-4ef9-a2f3-ba15162a96cf}"
Private Const IID_ISpeechRecognizerFactory As String = "{60c488dd-7fb8-4033-ac70-d046f64818e1}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognizerFactory
    ISpeechRecognizerFactory_Create = 6
    
    ' ISpeechRecognizerStatic
    ISpeechRecognizerStatic_GetSystemSpeechLanguage = 6
    ISpeechRecognizerStatic_GetSupportedTopicLanguages = 7
    ISpeechRecognizerStatic_GetSupportedGrammarLanguages = 8
    
    ' ISpeechRecognizerStatic2
    ISpeechRecognizerStatic2_TrySetSystemSpeechLanguageAsync = 6
    
    ' ISpeechRecognizer
    ISpeechRecognizer_GetCurrentLanguage = 6
    ISpeechRecognizer_GetConstraints = 7
    ISpeechRecognizer_GetTimeouts = 8
    ISpeechRecognizer_GetUIOptions = 9
    ISpeechRecognizer_CompileConstraintsAsync = 10
    ISpeechRecognizer_RecognizeAsync = 11
    ISpeechRecognizer_RecognizeWithUIAsync = 12
'    ISpeechRecognizer_AddRecognitionQualityDegrading = 13
'    ISpeechRecognizer_RemoveRecognitionQualityDegrading = 14
    ISpeechRecognizer_AddStateChanged = 15
    ISpeechRecognizer_RemoveStateChanged = 16
                        
    ' ISpeechRecognizer2
    ISpeechRecognizer2_GetContinuousRecognitionSession = 6
    ISpeechRecognizer2_GetState = 7
    ISpeechRecognizer2_StopRecognitionAsync = 8
'    ISpeechRecognizer2_AddHypothesisGenerated = 9
'    ISpeechRecognizer2_RemoveHypothesisGenerated = 10
    
End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognizer As Long
Private m_pISpeechRecognizer2 As Long
Private m_pISpeechRecognizerStatic As Long
Private m_pISpeechRecognizerStatic2 As Long
Private m_pISpeechRecognizerFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_SpeechRecognition_SpeechRecognizer, _
                            IID_ISpeechRecognizerFactory, _
                            m_pISpeechRecognizerFactory) Then
    End If
    If GetActivationFactory(Windows_Media_SpeechRecognition_SpeechRecognizer, _
                            IID_ISpeechRecognizerStatic, _
                            m_pISpeechRecognizerStatic) Then
        Call QueryIfc(m_pISpeechRecognizerStatic, _
                      IID_ISpeechRecognizerStatic2, _
                      m_pISpeechRecognizerStatic2)
    End If
    Dim pISpeechRecognizer As Long
    If GetActivateInstance(Windows_Media_SpeechRecognition_SpeechRecognizer, _
                           IID_ISpeechRecognizer, _
                           pISpeechRecognizer) Then
        Me.Ifc = pISpeechRecognizer
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognizerFactory)
    Call ReleaseIfc(m_pISpeechRecognizerStatic2)
    Call ReleaseIfc(m_pISpeechRecognizerStatic)
    Call ReleaseIfc(m_pISpeechRecognizer2)
    Call DisposeIfc(m_pISpeechRecognizer) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognizer
End Property

Public Property Let Ifc(ByVal pISpeechRecognizer As Long)
    Call ReleaseIfc(m_pISpeechRecognizer2)
    Call DisposeIfc(m_pISpeechRecognizer) ' <- IClosable
    m_pISpeechRecognizer = pISpeechRecognizer
    Call QueryIfc(m_pISpeechRecognizer, IID_ISpeechRecognizer2, m_pISpeechRecognizer2)
End Property

' ISpeechRecognizerStatic
Public Property Get SystemSpeechLanguage() As Language
    Dim Ret As Language
    If m_pISpeechRecognizerStatic <> 0& Then
        Dim pILanguage As Long
        If Invoke(m_pISpeechRecognizerStatic, _
                  ISpeechRecognizerStatic_GetSystemSpeechLanguage, _
                  VarPtr(pILanguage)) = S_OK Then
            If pILanguage <> 0& Then
                Set Ret = New Language
                Ret.Ifc = pILanguage
            End If
        End If
    End If
    Set SystemSpeechLanguage = Ret
End Property

Public Property Get SupportedTopicLanguages() As ReadOnlyList_1 'ReadOnlyList_Language
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_Language
    If m_pISpeechRecognizerStatic <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pISpeechRecognizerStatic, _
                  ISpeechRecognizerStatic_GetSupportedTopicLanguages, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_Language, pIVectorView) 'ReadOnlyList_Language
            End If
        End If
    End If
    Set SupportedTopicLanguages = Ret
End Property

Public Property Get SupportedGrammarLanguages() As ReadOnlyList_1 'ReadOnlyList_Language
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_Language
    If m_pISpeechRecognizerStatic <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pISpeechRecognizerStatic, _
                  ISpeechRecognizerStatic_GetSupportedGrammarLanguages, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_Language, pIVectorView) 'ReadOnlyList_Language
            End If
        End If
    End If
    Set SupportedGrammarLanguages = Ret
End Property
    
' ISpeechRecognizer
Public Property Get CurrentLanguage() As Language
    Dim Ret As Language
    If m_pISpeechRecognizer <> 0& Then
        Dim pILanguage As Long
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_GetCurrentLanguage, _
                  VarPtr(pILanguage)) = S_OK Then
            If pILanguage <> 0& Then
                Set Ret = New Language
                Ret.Ifc = pILanguage
            End If
        End If
    End If
    Set CurrentLanguage = Ret
End Property

Public Property Get Constraints() As List_SpeechRecognitionConstraint
    Dim Ret As List_SpeechRecognitionConstraint
    If m_pISpeechRecognizer <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_GetConstraints, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_SpeechRecognitionConstraint
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set Constraints = Ret
End Property

Public Property Get Timeouts() As SpeechRecognizerTimeouts
    Dim Ret As SpeechRecognizerTimeouts
    If m_pISpeechRecognizer <> 0& Then
        Dim pISpeechRecognizerTimeouts As Long
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_GetTimeouts, _
                  VarPtr(pISpeechRecognizerTimeouts)) = S_OK Then
            If pISpeechRecognizerTimeouts <> 0& Then
                Set Ret = New SpeechRecognizerTimeouts
                Ret.Ifc = pISpeechRecognizerTimeouts
            End If
        End If
    End If
    Set Timeouts = Ret
End Property

Public Property Get UIOptions() As SpeechRecognizerUIOptions
    Dim Ret As SpeechRecognizerUIOptions
    If m_pISpeechRecognizer <> 0& Then
        Dim pISpeechRecognizerUIOptions As Long
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_GetUIOptions, _
                  VarPtr(pISpeechRecognizerUIOptions)) = S_OK Then
            If pISpeechRecognizerUIOptions <> 0& Then
                Set Ret = New SpeechRecognizerUIOptions
                Ret.Ifc = pISpeechRecognizerUIOptions
            End If
        End If
    End If
    Set UIOptions = Ret
End Property
                        
' ISpeechRecognizer2
Public Property Get ContinuousRecognitionSession() As SpeechContinuousRecognitionSession
    Dim Ret As SpeechContinuousRecognitionSession
    If m_pISpeechRecognizer2 <> 0& Then
        Dim pISpeechContinuousRecognitionSession As Long
        If Invoke(m_pISpeechRecognizer2, _
                  ISpeechRecognizer2_GetContinuousRecognitionSession, _
                  VarPtr(pISpeechContinuousRecognitionSession)) = S_OK Then
            If pISpeechContinuousRecognitionSession <> 0& Then
                Set Ret = New SpeechContinuousRecognitionSession
                Ret.Ifc = pISpeechContinuousRecognitionSession
            End If
        End If
    End If
    Set ContinuousRecognitionSession = Ret
End Property

Public Property Get State() As SpeechRecognizerState
    Dim Ret As SpeechRecognizerState
    If m_pISpeechRecognizer2 <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognizer2, _
                  ISpeechRecognizer2_GetState, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    State = Ret
End Property

' ----==== Functions ====----
' ISpeechRecognizerFactory
Public Function Create(ByVal lang As Language) As SpeechRecognizer
    Dim Ret As SpeechRecognizer
    If m_pISpeechRecognizerFactory <> 0& And _
       IsNotNothing(lang) Then
        Dim pISpeechRecognizer As Long
        If Invoke(m_pISpeechRecognizerFactory, _
                  ISpeechRecognizerFactory_Create, _
                  lang.Ifc, _
                  VarPtr(pISpeechRecognizer)) = S_OK Then
            If pISpeechRecognizer <> 0& Then
                Set Ret = New SpeechRecognizer
                Ret.Ifc = pISpeechRecognizer
            End If
        End If
    End If
    Set Create = Ret
End Function
    
' ISpeechRecognizerStatic2
Public Function TrySetSystemSpeechLanguageAsync(ByVal speechLanguage As Language) As Boolean
    Dim Ret As Boolean
    If m_pISpeechRecognizerStatic2 <> 0& And _
       IsNotNothing(speechLanguage) Then
        Dim Result As Long
        If Invoke(m_pISpeechRecognizerStatic2, _
                  ISpeechRecognizerStatic2_TrySetSystemSpeechLanguageAsync, _
                  speechLanguage.Ifc, _
                  VarPtr(Result)) = S_OK Then
            If Result <> 0& Then
                If Await(Result) Then
                    Ret = Result
                End If
            End If
        End If
    End If
    TrySetSystemSpeechLanguageAsync = Ret
End Function
    
' ISpeechRecognizer
Public Function CompileConstraintsAsync() As SpeechRecognitionCompilationResult
    Dim Ret As SpeechRecognitionCompilationResult
    If m_pISpeechRecognizer <> 0& Then
        Dim operation As Long
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_CompileConstraintsAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New SpeechRecognitionCompilationResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set CompileConstraintsAsync = Ret
End Function

Public Function RecognizeAsync() As SpeechRecognitionResult
    Dim Ret As SpeechRecognitionResult
    If m_pISpeechRecognizer <> 0& Then
        Dim operation As Long
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_RecognizeAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New SpeechRecognitionResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set RecognizeAsync = Ret
End Function

Public Function RecognizeWithUIAsync() As SpeechRecognitionResult
    Dim Ret As SpeechRecognitionResult
    If m_pISpeechRecognizer <> 0& Then
        Dim operation As Long
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_RecognizeWithUIAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New SpeechRecognitionResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set RecognizeWithUIAsync = Ret
End Function

'    ISpeechRecognizer_AddRecognitionQualityDegrading = 13
'    ISpeechRecognizer_RemoveRecognitionQualityDegrading = 14

Public Function AddStateChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISpeechRecognizer <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_AddStateChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStateChanged = Ret
End Function

Public Function RemoveStateChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISpeechRecognizer <> 0& And token <> 0@ Then
        If Invoke(m_pISpeechRecognizer, _
                  ISpeechRecognizer_RemoveStateChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStateChanged = Ret
End Function
                        
' ISpeechRecognizer2
Public Function StopRecognitionAsync() As Boolean
    Dim Ret As Boolean
    If m_pISpeechRecognizer2 <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognizer2, _
                  ISpeechRecognizer2_StopRecognitionAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    StopRecognitionAsync = Ret
End Function

'    ISpeechRecognizer2_AddHypothesisGenerated = 9
'    ISpeechRecognizer2_RemoveHypothesisGenerated = 10

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

