VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoiceInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechSynthesis_VoiceInformation As String = "Windows.Media.SpeechSynthesis.VoiceInformation"
Private Const IID_IVoiceInformation As String = "{b127d6a4-1291-4604-aa9c-83134083352c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IVoiceInformation
    IVoiceInformation_GetDisplayName = 6
    IVoiceInformation_GetId = 7
    IVoiceInformation_GetLanguage = 8
    IVoiceInformation_GetDescription = 9
    IVoiceInformation_GetGender = 10

End Enum

' ----==== Variablen ====----
Private m_pIVoiceInformation As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVoiceInformation)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVoiceInformation
End Property

Public Property Let Ifc(ByVal pIVoiceInformation As Long)
    Call ReleaseIfc(m_pIVoiceInformation)
    m_pIVoiceInformation = pIVoiceInformation
End Property

' IVoiceInformation
Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIVoiceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIVoiceInformation, _
                  IVoiceInformation_GetDisplayName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get Id() As String
    Dim Ret As String
    If m_pIVoiceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIVoiceInformation, _
                  IVoiceInformation_GetId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Id = Ret
End Property

Public Property Get Language() As String
    Dim Ret As String
    If m_pIVoiceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIVoiceInformation, _
                  IVoiceInformation_GetLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Language = Ret
End Property

Public Property Get Description() As String
    Dim Ret As String
    If m_pIVoiceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIVoiceInformation, _
                  IVoiceInformation_GetDescription, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Description = Ret
End Property

Public Property Get Gender() As VoiceGender
    Dim Ret As VoiceGender
    If m_pIVoiceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIVoiceInformation, _
                  IVoiceInformation_GetGender, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Gender = Ret
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

