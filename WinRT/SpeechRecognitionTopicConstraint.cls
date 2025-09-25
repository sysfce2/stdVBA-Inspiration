VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognitionTopicConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognitionTopicConstraint As String = "Windows.Media.SpeechRecognition.SpeechRecognitionTopicConstraint"
Private Const IID_ISpeechRecognitionConstraint As String = "{79ac1628-4d68-43c4-8911-40dc4101b55b}"
Private Const IID_ISpeechRecognitionTopicConstraint As String = "{bf6fdf19-825d-4e69-a681-36e48cf1c93e}"
Private Const IID_ISpeechRecognitionTopicConstraintFactory As String = "{6e6863df-ec05-47d7-a5df-56a3431e58d2}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognitionTopicConstraintFactory
    ISpeechRecognitionTopicConstraintFactory_Create = 6
    ISpeechRecognitionTopicConstraintFactory_CreateWithTag = 7
    
    ' ISpeechRecognitionTopicConstraint
    ISpeechRecognitionTopicConstraint_GetScenario = 6
    ISpeechRecognitionTopicConstraint_GetTopicHint = 7
    
    ' ISpeechRecognitionConstraint
    ISpeechRecognitionConstraint_GetIsEnabled = 6
    ISpeechRecognitionConstraint_PutIsEnabled = 7
    ISpeechRecognitionConstraint_GetTag = 8
    ISpeechRecognitionConstraint_PutTag = 9
    ISpeechRecognitionConstraint_GetType = 10
    ISpeechRecognitionConstraint_GetProbability = 11
    ISpeechRecognitionConstraint_PutProbability = 12
    
End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognitionConstraint As Long
Private m_pISpeechRecognitionTopicConstraint As Long
Private m_pISpeechRecognitionTopicConstraintFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_SpeechRecognition_SpeechRecognitionTopicConstraint, _
                            IID_ISpeechRecognitionTopicConstraintFactory, _
                            m_pISpeechRecognitionTopicConstraintFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognitionTopicConstraintFactory)
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
    Call ReleaseIfc(m_pISpeechRecognitionTopicConstraint)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognitionTopicConstraint
End Property

Public Property Let Ifc(ByVal pISpeechRecognitionTopicConstraint As Long)
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
    Call ReleaseIfc(m_pISpeechRecognitionTopicConstraint)
    m_pISpeechRecognitionTopicConstraint = pISpeechRecognitionTopicConstraint
    Call QueryIfc(m_pISpeechRecognitionTopicConstraint, IID_ISpeechRecognitionConstraint, m_pISpeechRecognitionConstraint)
End Property

' ISpeechRecognitionTopicConstraint
Public Property Get Scenario() As SpeechRecognitionScenario
    Dim Ret As SpeechRecognitionScenario
    If m_pISpeechRecognitionTopicConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionTopicConstraint, _
                  ISpeechRecognitionTopicConstraint_GetScenario, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Scenario = Ret
End Property

Public Property Get TopicHint() As String
    Dim Ret As String
    If m_pISpeechRecognitionTopicConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionTopicConstraint, _
                  ISpeechRecognitionTopicConstraint_GetTopicHint, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    TopicHint = Ret
End Property

' ISpeechRecognitionConstraint
Public Property Get IsEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISpeechRecognitionConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_GetIsEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsEnabled = Ret
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    If m_pISpeechRecognitionConstraint <> 0& Then
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_PutIsEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get tag() As String
    Dim Ret As String
    If m_pISpeechRecognitionConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_GetTag, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    tag = Ret
End Property

Public Property Let tag(ByVal value As String)
    If m_pISpeechRecognitionConstraint <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_PutTag, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get ConstraintType() As SpeechRecognitionConstraintType
    Dim Ret As SpeechRecognitionConstraintType
    If m_pISpeechRecognitionConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_GetType, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ConstraintType = Ret
End Property

Public Property Get Probability() As SpeechRecognitionConstraintProbability
    Dim Ret As SpeechRecognitionConstraintProbability
    If m_pISpeechRecognitionConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_GetProbability, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Probability = Ret
End Property

Public Property Let Probability(ByVal value As SpeechRecognitionConstraintProbability)
    If m_pISpeechRecognitionConstraint <> 0& Then
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_PutProbability, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' ISpeechRecognitionTopicConstraintFactory
Public Function Create(ByVal Scenario As SpeechRecognitionScenario, _
                       ByVal TopicHint As String) As SpeechRecognitionTopicConstraint
    Dim Ret As SpeechRecognitionTopicConstraint
    If m_pISpeechRecognitionTopicConstraintFactory <> 0& Then
        Dim pISpeechRecognitionTopicConstraint As Long
        Dim hTopicHint As Long
        hTopicHint = CreateWindowsString(TopicHint)
        If Invoke(m_pISpeechRecognitionTopicConstraintFactory, _
                  ISpeechRecognitionTopicConstraintFactory_Create, _
                  Scenario, _
                  hTopicHint, _
                  VarPtr(pISpeechRecognitionTopicConstraint)) = S_OK Then
            If pISpeechRecognitionTopicConstraint <> 0& Then
                Set Ret = New SpeechRecognitionTopicConstraint
                Ret.Ifc = pISpeechRecognitionTopicConstraint
            End If
        End If
        Call DeleteWindowsString(hTopicHint)
    End If
    Set Create = Ret
End Function

Public Function CreateWithTag(ByVal Scenario As SpeechRecognitionScenario, _
                              ByVal TopicHint As String, _
                              ByVal tag As String) As SpeechRecognitionTopicConstraint
    Dim Ret As SpeechRecognitionTopicConstraint
    If m_pISpeechRecognitionTopicConstraintFactory <> 0& Then
        Dim pISpeechRecognitionTopicConstraint As Long
        Dim hTopicHint As Long
        Dim hTag As Long
        hTopicHint = CreateWindowsString(TopicHint)
        hTag = CreateWindowsString(tag)
        If Invoke(m_pISpeechRecognitionTopicConstraintFactory, _
                  ISpeechRecognitionTopicConstraintFactory_CreateWithTag, _
                  Scenario, _
                  hTopicHint, _
                  hTag, _
                  VarPtr(pISpeechRecognitionTopicConstraint)) = S_OK Then
            If pISpeechRecognitionTopicConstraint <> 0& Then
                Set Ret = New SpeechRecognitionTopicConstraint
                Ret.Ifc = pISpeechRecognitionTopicConstraint
            End If
        End If
        Call DeleteWindowsString(hTag)
        Call DeleteWindowsString(hTopicHint)
    End If
    Set CreateWithTag = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


