VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognitionGrammarFileConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognitionGrammarFileConstraint As String = "Windows.Media.SpeechRecognition.SpeechRecognitionGrammarFileConstraint"
Private Const IID_ISpeechRecognitionConstraint As String = "{79ac1628-4d68-43c4-8911-40dc4101b55b}"
Private Const IID_ISpeechRecognitionGrammarFileConstraint As String = "{b5031a8f-85ca-4fa4-b11a-474fc41b3835}"
Private Const IID_ISpeechRecognitionGrammarFileConstraintFactory As String = "{3da770eb-c479-4c27-9f19-89974ef392d1}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognitionGrammarFileConstraintFactory
    ISpeechRecognitionGrammarFileConstraintFactory_Create = 6
    ISpeechRecognitionGrammarFileConstraintFactory_CreateWithTag = 7
    
    ' ISpeechRecognitionGrammarFileConstraint
    ISpeechRecognitionGrammarFileConstraint_GetGrammarFile = 6
    
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
Private m_pISpeechRecognitionGrammarFileConstraint As Long
Private m_pISpeechRecognitionGrammarFileConstraintFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_SpeechRecognition_SpeechRecognitionGrammarFileConstraint, _
                            IID_ISpeechRecognitionGrammarFileConstraintFactory, _
                            m_pISpeechRecognitionGrammarFileConstraintFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognitionGrammarFileConstraintFactory)
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
    Call ReleaseIfc(m_pISpeechRecognitionGrammarFileConstraint)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognitionGrammarFileConstraint
End Property

Public Property Let Ifc(ByVal pISpeechRecognitionGrammarFileConstraint As Long)
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
    Call ReleaseIfc(m_pISpeechRecognitionGrammarFileConstraint)
    m_pISpeechRecognitionGrammarFileConstraint = pISpeechRecognitionGrammarFileConstraint
    Call QueryIfc(m_pISpeechRecognitionGrammarFileConstraint, IID_ISpeechRecognitionConstraint, m_pISpeechRecognitionConstraint)
End Property

' ISpeechRecognitionGrammarFileConstraint
Public Property Get GrammarFile() As StorageFile
    Dim Ret As StorageFile
    If m_pISpeechRecognitionGrammarFileConstraint <> 0& Then
        Dim pIStorageFile As Long
        If Invoke(m_pISpeechRecognitionGrammarFileConstraint, _
                  ISpeechRecognitionGrammarFileConstraint_GetGrammarFile, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                Set Ret = New StorageFile
                Ret.Ifc = pIStorageFile
            End If
        End If
    End If
    Set GrammarFile = Ret
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
' ISpeechRecognitionGrammarFileConstraintFactory
Public Function Create(ByVal file As StorageFile) As SpeechRecognitionGrammarFileConstraint
    Dim Ret As SpeechRecognitionGrammarFileConstraint
    If m_pISpeechRecognitionGrammarFileConstraintFactory <> 0& And _
       IsNotNothing(file) Then
        Dim pISpeechRecognitionGrammarFileConstraint As Long
        If Invoke(m_pISpeechRecognitionGrammarFileConstraintFactory, _
                  ISpeechRecognitionGrammarFileConstraintFactory_Create, _
                  file.Ifc, _
                  VarPtr(pISpeechRecognitionGrammarFileConstraint)) = S_OK Then
            If pISpeechRecognitionGrammarFileConstraint <> 0& Then
                Set Ret = New SpeechRecognitionGrammarFileConstraint
                Ret.Ifc = pISpeechRecognitionGrammarFileConstraint
            End If
        End If
    End If
    Set Create = Ret
End Function

Public Function CreateWithTag(ByVal file As StorageFile, _
                              ByVal Tag As String) As SpeechRecognitionGrammarFileConstraint
    Dim Ret As SpeechRecognitionGrammarFileConstraint
    If m_pISpeechRecognitionGrammarFileConstraintFactory <> 0& And _
       IsNotNothing(file) Then
        Dim pISpeechRecognitionGrammarFileConstraint As Long
        Dim hTag As Long
        hTag = CreateWindowsString(Tag)
        If Invoke(m_pISpeechRecognitionGrammarFileConstraintFactory, _
                  ISpeechRecognitionGrammarFileConstraintFactory_CreateWithTag, _
                  file.Ifc, _
                  hTag, _
                  VarPtr(pISpeechRecognitionGrammarFileConstraint)) = S_OK Then
            If pISpeechRecognitionGrammarFileConstraint <> 0& Then
                Set Ret = New SpeechRecognitionGrammarFileConstraint
                Ret.Ifc = pISpeechRecognitionGrammarFileConstraint
            End If
        End If
        Call DeleteWindowsString(hTag)
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


