VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechSynthesizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechSynthesis_SpeechSynthesizer As String = "Windows.Media.SpeechSynthesis.SpeechSynthesizer"
Private Const IID_ISpeechSynthesizer As String = "{ce9f7c76-97f4-4ced-ad68-d51c458e45c6}"
Private Const IID_ISpeechSynthesizer2 As String = "{a7c5ecb2-4339-4d6a-bbf8-c7a4f1544c2e}"
Private Const IID_IInstalledVoicesStatic As String = "{7d526ecc-7533-4c3f-85be-888c2baeebdc}"
Private Const IID_IInstalledVoicesStatic2 As String = "{64255f2e-358d-4058-be9a-fd3fcb423530}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IInstalledVoicesStatic
    IInstalledVoicesStatic_GetAllVoices = 6
    IInstalledVoicesStatic_GetDefaultVoice = 7
    
    ' IInstalledVoicesStatic2
    IInstalledVoicesStatic2_TrySetDefaultVoiceAsync = 6
    
    ' ISpeechSynthesizer
    ISpeechSynthesizer_SynthesizeTextToStreamAsync = 6
    ISpeechSynthesizer_SynthesizeSsmlToStreamAsync = 7
    ISpeechSynthesizer_PutVoice = 8
    ISpeechSynthesizer_GetVoice = 9
    
    ' ISpeechSynthesizer2
    ISpeechSynthesizer2_GetOptions = 6

End Enum

' ----==== Variablen ====----
Private m_pISpeechSynthesizer As Long
Private m_pISpeechSynthesizer2 As Long
Private m_pIInstalledVoicesStatic As Long
Private m_pIInstalledVoicesStatic2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_SpeechSynthesis_SpeechSynthesizer, _
                            IID_IInstalledVoicesStatic, _
                            m_pIInstalledVoicesStatic) Then
        Call QueryIfc(m_pIInstalledVoicesStatic, _
                      IID_IInstalledVoicesStatic2, _
                      m_pIInstalledVoicesStatic2)
    End If
    Dim pISpeechSynthesizer As Long
    If GetActivateInstance(Windows_Media_SpeechSynthesis_SpeechSynthesizer, _
                           IID_ISpeechSynthesizer, _
                           pISpeechSynthesizer) Then
        Me.Ifc = pISpeechSynthesizer
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIInstalledVoicesStatic2)
    Call ReleaseIfc(m_pIInstalledVoicesStatic)
    Call ReleaseIfc(m_pISpeechSynthesizer2)
    Call DisposeIfc(m_pISpeechSynthesizer) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechSynthesizer
End Property

Public Property Let Ifc(ByVal pISpeechSynthesizer As Long)
    Call ReleaseIfc(m_pISpeechSynthesizer2)
    Call DisposeIfc(m_pISpeechSynthesizer) ' <- IClosable
    m_pISpeechSynthesizer = pISpeechSynthesizer
    Call QueryIfc(m_pISpeechSynthesizer, IID_ISpeechSynthesizer2, m_pISpeechSynthesizer2)
End Property

' IInstalledVoicesStatic
Public Property Get AllVoices() As ReadOnlyList_1 'ReadOnlyList_VoiceInformation
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_VoiceInformation
    If m_pIInstalledVoicesStatic <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIInstalledVoicesStatic, _
                  IInstalledVoicesStatic_GetAllVoices, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_VoiceInformation, pIVectorView) 'ReadOnlyList_VoiceInformation
            End If
        End If
    End If
    Set AllVoices = Ret
End Property

Public Property Get DefaultVoice() As VoiceInformation
    Dim Ret As VoiceInformation
    If m_pIInstalledVoicesStatic <> 0& Then
        Dim pIVoiceInformation As Long
        If Invoke(m_pIInstalledVoicesStatic, _
                  IInstalledVoicesStatic_GetDefaultVoice, _
                  VarPtr(pIVoiceInformation)) = S_OK Then
            If pIVoiceInformation <> 0& Then
                Set Ret = New VoiceInformation
                Ret.Ifc = pIVoiceInformation
            End If
        End If
    End If
    Set DefaultVoice = Ret
End Property

' ISpeechSynthesizer
Public Property Let voice(ByVal value As VoiceInformation)
    If m_pISpeechSynthesizer <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pISpeechSynthesizer, _
                  ISpeechSynthesizer_PutVoice, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get voice() As VoiceInformation
    Dim Ret As VoiceInformation
    If m_pISpeechSynthesizer <> 0& Then
        Dim pIVoiceInformation As Long
        If Invoke(m_pISpeechSynthesizer, _
                  ISpeechSynthesizer_GetVoice, _
                  VarPtr(pIVoiceInformation)) = S_OK Then
            If pIVoiceInformation <> 0& Then
                Set Ret = New VoiceInformation
                Ret.Ifc = pIVoiceInformation
            End If
        End If
    End If
    Set voice = Ret
End Property

' ISpeechSynthesizer2
Public Property Get Options() As SpeechSynthesizerOptions
    Dim Ret As SpeechSynthesizerOptions
    If m_pISpeechSynthesizer2 <> 0& Then
        Dim pISpeechSynthesizerOptions As Long
        If Invoke(m_pISpeechSynthesizer2, _
                  ISpeechSynthesizer2_GetOptions, _
                  VarPtr(pISpeechSynthesizerOptions)) = S_OK Then
            If pISpeechSynthesizerOptions <> 0& Then
                Set Ret = New SpeechSynthesizerOptions
                Ret.Ifc = pISpeechSynthesizerOptions
            End If
        End If
    End If
    Set Options = Ret
End Property

' ----==== Functions ====----
' IInstalledVoicesStatic2
Public Function TrySetDefaultVoiceAsync(ByVal voice As VoiceInformation) As Boolean
    Dim Ret As Boolean
    If m_pIInstalledVoicesStatic2 <> 0& And _
       IsNotNothing(voice) Then
        Dim Result As Long
        If Invoke(m_pIInstalledVoicesStatic2, _
                  IInstalledVoicesStatic2_TrySetDefaultVoiceAsync, _
                  voice.Ifc, _
                  VarPtr(Result)) = S_OK Then
            If Result <> 0& Then
                If Await(Result) Then
                    Ret = CBool(Result)
                End If
            End If
        End If
    End If
End Function

' ISpeechSynthesizer
Public Function SynthesizeTextToStreamAsync(ByVal Text As String) As SpeechSynthesisStream
    Dim Ret As SpeechSynthesisStream
    If m_pISpeechSynthesizer <> 0& Then
        Dim pISpeechSynthesisStream As Long
        Dim hString As Long
        hString = CreateWindowsString(Text)
        If Invoke(m_pISpeechSynthesizer, _
                  ISpeechSynthesizer_SynthesizeTextToStreamAsync, _
                  hString, _
                  VarPtr(pISpeechSynthesisStream)) = S_OK Then
            If pISpeechSynthesisStream <> 0& Then
                If Await(pISpeechSynthesisStream) Then
                    Set Ret = New SpeechSynthesisStream
                    Ret.Ifc = pISpeechSynthesisStream
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set SynthesizeTextToStreamAsync = Ret
End Function

Public Function SynthesizeSsmlToStreamAsync(ByVal Ssml As String) As SpeechSynthesisStream
    Dim Ret As SpeechSynthesisStream
    If m_pISpeechSynthesizer <> 0& Then
        Dim pISpeechSynthesisStream As Long
        Dim hString As Long
        hString = CreateWindowsString(Ssml)
        If Invoke(m_pISpeechSynthesizer, _
                  ISpeechSynthesizer_SynthesizeSsmlToStreamAsync, _
                  hString, _
                  VarPtr(pISpeechSynthesisStream)) = S_OK Then
            If pISpeechSynthesisStream <> 0& Then
                If Await(pISpeechSynthesisStream) Then
                    Set Ret = New SpeechSynthesisStream
                    Ret.Ifc = pISpeechSynthesisStream
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set SynthesizeSsmlToStreamAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
