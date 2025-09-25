VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlaybackCommandManagerCommandBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaPlaybackCommandManagerCommandBehavior As String = "Windows.Media.Playback.MediaPlaybackCommandManagerCommandBehavior"
Private Const IID_IMediaPlaybackCommandManagerCommandBehavior As String = "{786c1e78-ce78-4a10-afd6-843fcbb90c2e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaPlaybackCommandManagerCommandBehavior
    IMediaPlaybackCommandManagerCommandBehavior_GetCommandManager = 6
    IMediaPlaybackCommandManagerCommandBehavior_GetIsEnabled = 7
    IMediaPlaybackCommandManagerCommandBehavior_GetEnablingRule = 8
    IMediaPlaybackCommandManagerCommandBehavior_PutEnablingRule = 9
    IMediaPlaybackCommandManagerCommandBehavior_AddIsEnabledChanged = 10
    IMediaPlaybackCommandManagerCommandBehavior_RemoveIsEnabledChanged = 11
                        
End Enum

' ----==== Variablen ====----
Private m_pIMediaPlaybackCommandManagerCommandBehavior As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaPlaybackCommandManagerCommandBehavior)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaPlaybackCommandManagerCommandBehavior
End Property

Public Property Let Ifc(ByVal pIMediaPlaybackCommandManagerCommandBehavior As Long)
    Call ReleaseIfc(m_pIMediaPlaybackCommandManagerCommandBehavior)
    m_pIMediaPlaybackCommandManagerCommandBehavior = pIMediaPlaybackCommandManagerCommandBehavior
End Property

' IMediaPlaybackCommandManagerCommandBehavior
Public Property Get CommandManager() As MediaPlaybackCommandManager
    Dim Ret As MediaPlaybackCommandManager
    If m_pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
        Dim pIMediaPlaybackCommandManager As Long
        If Invoke(m_pIMediaPlaybackCommandManagerCommandBehavior, _
                  IMediaPlaybackCommandManagerCommandBehavior_GetCommandManager, _
                  VarPtr(pIMediaPlaybackCommandManager)) = S_OK Then
            If pIMediaPlaybackCommandManager <> 0& Then
                Set Ret = New MediaPlaybackCommandManager
                Ret.Ifc = pIMediaPlaybackCommandManager
            End If
        End If
    End If
    Set CommandManager = Ret
End Property

Public Property Get IsEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackCommandManagerCommandBehavior, _
                  IMediaPlaybackCommandManagerCommandBehavior_GetIsEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsEnabled = Ret
End Property

Public Property Get EnablingRule() As MediaCommandEnablingRule
    Dim Ret As MediaCommandEnablingRule
    If m_pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackCommandManagerCommandBehavior, _
                  IMediaPlaybackCommandManagerCommandBehavior_GetEnablingRule, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    EnablingRule = Ret
End Property

Public Property Let EnablingRule(ByVal value As MediaCommandEnablingRule)
    If m_pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
        If Invoke(m_pIMediaPlaybackCommandManagerCommandBehavior, _
                  IMediaPlaybackCommandManagerCommandBehavior_PutEnablingRule, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IMediaPlaybackCommandManagerCommandBehavior
Public Function AddIsEnabledChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManagerCommandBehavior <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManagerCommandBehavior, _
                  IMediaPlaybackCommandManagerCommandBehavior_AddIsEnabledChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddIsEnabledChanged = Ret
End Function

Public Function RemoveIsEnabledChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManagerCommandBehavior <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManagerCommandBehavior, _
                  IMediaPlaybackCommandManagerCommandBehavior_RemoveIsEnabledChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveIsEnabledChanged = Ret
End Function
    
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

