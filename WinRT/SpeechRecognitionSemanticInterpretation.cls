VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechRecognitionSemanticInterpretation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechRecognition_SpeechRecognitionSemanticInterpretation As String = "Windows.Media.SpeechRecognition.SpeechRecognitionSemanticInterpretation"
Private Const IID_ISpeechRecognitionSemanticInterpretation As String = "{aae1da9b-7e32-4c1f-89fe-0c65f486f52e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechRecognitionSemanticInterpretation
    ISpeechRecognitionSemanticInterpretation_GetProperties = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pISpeechRecognitionSemanticInterpretation As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISpeechRecognitionSemanticInterpretation)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechRecognitionSemanticInterpretation
End Property

Public Property Let Ifc(ByVal pISpeechRecognitionSemanticInterpretation As Long)
    Call ReleaseIfc(m_pISpeechRecognitionSemanticInterpretation)
    m_pISpeechRecognitionSemanticInterpretation = pISpeechRecognitionSemanticInterpretation
End Property

' ISpeechRecognitionSemanticInterpretation
'    ISpeechRecognitionSemanticInterpretation_GetProperties = 6

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


