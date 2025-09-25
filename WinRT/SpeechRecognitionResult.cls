VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognitionResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognitionResult As String = "Windows.Media.SpeechRecognition.SpeechRecognitionResult"
Private Const IID_ISpeechRecognitionResult As String = "{4e303157-034e-4652-857e-d0454cc4beec}"
Private Const IID_ISpeechRecognitionResult2 As String = "{af7ed1ba-451b-4166-a0c1-1ffe84032d03}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognitionResult
    ISpeechRecognitionResult_GetStatus = 6
    ISpeechRecognitionResult_GetText = 7
    ISpeechRecognitionResult_GetConfidence = 8
    ISpeechRecognitionResult_GetSemanticInterpretation = 9
    ISpeechRecognitionResult_GetAlternates = 10
    ISpeechRecognitionResult_GetConstraint = 11
    ISpeechRecognitionResult_GetRulePath = 12
    ISpeechRecognitionResult_GetRawConfidence = 13
                        
    ' ISpeechRecognitionResult2
    ISpeechRecognitionResult2_GetPhraseStartTime = 6
    ISpeechRecognitionResult2_GetPhraseDuration = 7

End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognitionResult As Long
Private m_pISpeechRecognitionResult2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognitionResult2)
    Call ReleaseIfc(m_pISpeechRecognitionResult)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognitionResult
End Property

Public Property Let Ifc(ByVal pISpeechRecognitionResult As Long)
    Call ReleaseIfc(m_pISpeechRecognitionResult2)
    Call ReleaseIfc(m_pISpeechRecognitionResult)
    m_pISpeechRecognitionResult = pISpeechRecognitionResult
    Call QueryIfc(m_pISpeechRecognitionResult, IID_ISpeechRecognitionResult2, m_pISpeechRecognitionResult2)
End Property

' ISpeechRecognitionResult
Public Property Get Status() As SpeechRecognitionResultStatus
    Dim Ret As SpeechRecognitionResultStatus
    If m_pISpeechRecognitionResult <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
End Property

Public Property Get Text() As String
    Dim Ret As String
    If m_pISpeechRecognitionResult <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Text = Ret
End Property

Public Property Get Confidence() As SpeechRecognitionConfidence
    Dim Ret As SpeechRecognitionConfidence
    If m_pISpeechRecognitionResult <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetConfidence, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Confidence = Ret
End Property

Public Property Get SemanticInterpretation() As SpeechRecognitionSemanticInterpretation
    Dim Ret As SpeechRecognitionSemanticInterpretation
    If m_pISpeechRecognitionResult <> 0& Then
        Dim pISpeechRecognitionSemanticInterpretation As Long
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetSemanticInterpretation, _
                  VarPtr(pISpeechRecognitionSemanticInterpretation)) = S_OK Then
            If pISpeechRecognitionSemanticInterpretation <> 0& Then
                Set Ret = New SpeechRecognitionSemanticInterpretation
                Ret.Ifc = pISpeechRecognitionSemanticInterpretation
            End If
        End If
    End If
    Set SemanticInterpretation = Ret
End Property

Public Property Get Constraint() As SpeechRecognitionConstraint
    Dim Ret As SpeechRecognitionConstraint
    If m_pISpeechRecognitionResult <> 0& Then
        Dim pISpeechRecognitionConstraint As Long
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetConstraint, _
                  VarPtr(pISpeechRecognitionConstraint)) = S_OK Then
            If pISpeechRecognitionConstraint <> 0& Then
                Set Ret = New SpeechRecognitionConstraint
                Ret.Ifc = pISpeechRecognitionConstraint
            End If
        End If
    End If
    Set Constraint = Ret
End Property

Public Property Get RulePath() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pISpeechRecognitionResult <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetRulePath, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set RulePath = Ret
End Property

Public Property Get RawConfidence() As Double
    Dim Ret As Double
    If m_pISpeechRecognitionResult <> 0& Then
        Dim value As Double
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetRawConfidence, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    RawConfidence = Ret
End Property

' ISpeechRecognitionResult2
Public Property Get PhraseStartTime() As DateTime
    Dim Ret As DateTime
    If m_pISpeechRecognitionResult2 <> 0& Then
        Dim value As Currency
        If Invoke(m_pISpeechRecognitionResult2, _
                  ISpeechRecognitionResult2_GetPhraseStartTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set PhraseStartTime = Ret
End Property

Public Property Get PhraseDuration() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISpeechRecognitionResult2 <> 0& Then
        Dim value As Currency
        If Invoke(m_pISpeechRecognitionResult2, _
                  ISpeechRecognitionResult2_GetPhraseDuration, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set PhraseDuration = Ret
End Property

' ----==== Functions ====----
' ISpeechRecognitionResult
Public Function GetAlternates(ByVal maxAlternates As Long) As ReadOnlyList_1 'ReadOnlyList_SpeechRecognitionResult
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_SpeechRecognitionResult
    If m_pISpeechRecognitionResult <> 0& And _
       maxAlternates > 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pISpeechRecognitionResult, _
                  ISpeechRecognitionResult_GetAlternates, _
                  maxAlternates, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_SpeechRecognitionResult, pIVectorView) 'ReadOnlyList_SpeechRecognitionResult
            End If
        End If
    End If
    Set GetAlternates = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
