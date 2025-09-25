VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognitionConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognitionConstraint As String = "Windows.Media.SpeechRecognition.SpeechRecognitionConstraint"
Private Const IID_ISpeechRecognitionConstraint As String = "{79ac1628-4d68-43c4-8911-40dc4101b55b}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
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

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognitionConstraint
End Property

Public Property Let Ifc(ByVal pISpeechRecognitionConstraint As Long)
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
    m_pISpeechRecognitionConstraint = pISpeechRecognitionConstraint
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

Public Property Get Tag() As String
    Dim Ret As String
    If m_pISpeechRecognitionConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognitionConstraint, _
                  ISpeechRecognitionConstraint_GetTag, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Tag = Ret
End Property

Public Property Let Tag(ByVal value As String)
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

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


