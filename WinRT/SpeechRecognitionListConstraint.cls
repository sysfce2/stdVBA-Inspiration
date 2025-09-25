VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognitionListConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognitionListConstraint As String = "Windows.Media.SpeechRecognition.SpeechRecognitionListConstraint"
Private Const IID_ISpeechRecognitionConstraint As String = "{79ac1628-4d68-43c4-8911-40dc4101b55b}"
Private Const IID_ISpeechRecognitionListConstraint As String = "{09c487e9-e4ad-4526-81f2-4946fb481d98}"
Private Const IID_ISpeechRecognitionListConstraintFactory As String = "{40f3cdc7-562a-426a-9f3b-3b4e282be1d5}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognitionListConstraintFactory
    ISpeechRecognitionListConstraintFactory_Create = 6
    ISpeechRecognitionListConstraintFactory_CreateWithTag = 7
    
    ' ISpeechRecognitionListConstraint
    ISpeechRecognitionListConstraint_GetCommands = 6
    
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
Private m_pISpeechRecognitionListConstraint As Long
Private m_pISpeechRecognitionListConstraintFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_SpeechRecognition_SpeechRecognitionListConstraint, _
                            IID_ISpeechRecognitionListConstraintFactory, _
                            m_pISpeechRecognitionListConstraintFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognitionListConstraintFactory)
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
    Call ReleaseIfc(m_pISpeechRecognitionListConstraint)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognitionListConstraint
End Property

Public Property Let Ifc(ByVal pISpeechRecognitionListConstraint As Long)
    Call ReleaseIfc(m_pISpeechRecognitionConstraint)
    Call ReleaseIfc(m_pISpeechRecognitionListConstraint)
    m_pISpeechRecognitionListConstraint = pISpeechRecognitionListConstraint
    Call QueryIfc(m_pISpeechRecognitionListConstraint, IID_ISpeechRecognitionConstraint, m_pISpeechRecognitionConstraint)
End Property

' ISpeechRecognitionListConstraint
Public Property Get Commands() As List_String
    Dim Ret As List_String
    If m_pISpeechRecognitionListConstraint <> 0& Then
        Dim pIList_String As Long
        If Invoke(m_pISpeechRecognitionListConstraint, _
                  ISpeechRecognitionListConstraint_GetCommands, _
                  VarPtr(pIList_String)) = S_OK Then
            If pIList_String <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIList_String
            End If
        End If
    End If
    Set Commands = Ret
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
' ISpeechRecognitionListConstraintFactory
Public Function Create(ByVal Commands As List_String) As SpeechRecognitionListConstraint
    Dim Ret As SpeechRecognitionListConstraint
    If m_pISpeechRecognitionListConstraintFactory <> 0& And _
       IsNotNothing(Commands) Then
        Dim pISpeechRecognitionListConstraint As Long
        If Invoke(m_pISpeechRecognitionListConstraintFactory, _
                  ISpeechRecognitionListConstraintFactory_Create, _
                  Commands.Ifc, _
                  VarPtr(pISpeechRecognitionListConstraint)) = S_OK Then
            If pISpeechRecognitionListConstraint <> 0& Then
                Set Ret = New SpeechRecognitionListConstraint
                Ret.Ifc = pISpeechRecognitionListConstraint
            End If
        End If
    End If
    Set Create = Ret
End Function

Public Function CreateWithTag(ByVal Commands As List_String, _
                              ByVal Tag As String) As SpeechRecognitionListConstraint
    Dim Ret As SpeechRecognitionListConstraint
    If m_pISpeechRecognitionListConstraintFactory <> 0& And _
       IsNotNothing(Commands) Then
        Dim pISpeechRecognitionListConstraint As Long
        Dim hTag As Long
        hTag = CreateWindowsString(Tag)
        If Invoke(m_pISpeechRecognitionListConstraintFactory, _
                  ISpeechRecognitionListConstraintFactory_CreateWithTag, _
                  Commands.Ifc, _
                  hTag, _
                  VarPtr(pISpeechRecognitionListConstraint)) = S_OK Then
            If pISpeechRecognitionListConstraint <> 0& Then
                Set Ret = New SpeechRecognitionListConstraint
                Ret.Ifc = pISpeechRecognitionListConstraint
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


