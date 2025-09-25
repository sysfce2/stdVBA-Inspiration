VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognizerUIOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognizerUIOptions As String = "Windows.Media.SpeechRecognition.SpeechRecognizerUIOptions"
Private Const IID_ISpeechRecognizerUIOptions As String = "{7888d641-b92b-44ba-a25f-d1864630641f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognizerUIOptions
    ISpeechRecognizerUIOptions_GetExampleText = 6
    ISpeechRecognizerUIOptions_PutExampleText = 7
    ISpeechRecognizerUIOptions_GetAudiblePrompt = 8
    ISpeechRecognizerUIOptions_PutAudiblePrompt = 9
    ISpeechRecognizerUIOptions_GetIsReadBackEnabled = 10
    ISpeechRecognizerUIOptions_PutIsReadBackEnabled = 11
    ISpeechRecognizerUIOptions_GetShowConfirmation = 12
    ISpeechRecognizerUIOptions_PutShowConfirmation = 13
                        
End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognizerUIOptions As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognizerUIOptions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognizerUIOptions
End Property

Public Property Let Ifc(ByVal pISpeechRecognizerUIOptions As Long)
    Call ReleaseIfc(m_pISpeechRecognizerUIOptions)
    m_pISpeechRecognizerUIOptions = pISpeechRecognizerUIOptions
End Property

' ISpeechRecognizerUIOptions
Public Property Get ExampleText() As String
    Dim Ret As String
    If m_pISpeechRecognizerUIOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_GetExampleText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ExampleText = Ret
End Property

Public Property Let ExampleText(ByVal value As String)
    If m_pISpeechRecognizerUIOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_PutExampleText, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get AudiblePrompt() As String
    Dim Ret As String
    If m_pISpeechRecognizerUIOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_GetAudiblePrompt, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AudiblePrompt = Ret
End Property

Public Property Let AudiblePrompt(ByVal value As String)
    If m_pISpeechRecognizerUIOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_PutAudiblePrompt, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get IsReadBackEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISpeechRecognizerUIOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_GetIsReadBackEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsReadBackEnabled = Ret
End Property

Public Property Let IsReadBackEnabled(ByVal value As Boolean)
    If m_pISpeechRecognizerUIOptions <> 0& Then
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_PutIsReadBackEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get ShowConfirmation() As Boolean
    Dim Ret As Boolean
    If m_pISpeechRecognizerUIOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_GetShowConfirmation, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    ShowConfirmation = Ret
End Property

Public Property Let ShowConfirmation(ByVal value As Boolean)
    If m_pISpeechRecognizerUIOptions <> 0& Then
        If Invoke(m_pISpeechRecognizerUIOptions, _
                  ISpeechRecognizerUIOptions_PutShowConfirmation, _
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


