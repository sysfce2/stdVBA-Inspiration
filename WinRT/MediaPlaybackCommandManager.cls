VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlaybackCommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaPlaybackCommandManager As String = "Windows.Media.Playback.MediaPlaybackCommandManager"
Private Const IID_IMediaPlaybackCommandManager As String = "{5acee5a6-5cb6-4a5a-8521-cc86b1c1ed37}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaPlaybackCommandManager
    IMediaPlaybackCommandManager_GetIsEnabled = 6
    IMediaPlaybackCommandManager_PutIsEnabled = 7
    IMediaPlaybackCommandManager_GetMediaPlayer = 8
    IMediaPlaybackCommandManager_GetPlayBehavior = 9
    IMediaPlaybackCommandManager_GetPauseBehavior = 10
    IMediaPlaybackCommandManager_GetNextBehavior = 11
    IMediaPlaybackCommandManager_GetPreviousBehavior = 12
    IMediaPlaybackCommandManager_GetFastForwardBehavior = 13
    IMediaPlaybackCommandManager_GetRewindBehavior = 14
    IMediaPlaybackCommandManager_GetShuffleBehavior = 15
    IMediaPlaybackCommandManager_GetAutoRepeatModeBehavior = 16
    IMediaPlaybackCommandManager_GetPositionBehavior = 17
    IMediaPlaybackCommandManager_GetRateBehavior = 18
    IMediaPlaybackCommandManager_AddPlayReceived = 19
    IMediaPlaybackCommandManager_RemovePlayReceived = 20
    IMediaPlaybackCommandManager_AddPauseReceived = 21
    IMediaPlaybackCommandManager_RemovePauseReceived = 22
    IMediaPlaybackCommandManager_AddNextReceived = 23
    IMediaPlaybackCommandManager_RemoveNextReceived = 24
    IMediaPlaybackCommandManager_AddPreviousReceived = 25
    IMediaPlaybackCommandManager_RemovePreviousReceived = 26
    IMediaPlaybackCommandManager_AddFastForwardReceived = 27
    IMediaPlaybackCommandManager_RemoveFastForwardReceived = 28
    IMediaPlaybackCommandManager_AddRewindReceived = 29
    IMediaPlaybackCommandManager_RemoveRewindReceived = 30
    IMediaPlaybackCommandManager_AddShuffleReceived = 31
    IMediaPlaybackCommandManager_RemoveShuffleReceived = 32
    IMediaPlaybackCommandManager_AddAutoRepeatModeReceived = 33
    IMediaPlaybackCommandManager_RemoveAutoRepeatModeReceived = 34
    IMediaPlaybackCommandManager_AddPositionReceived = 35
    IMediaPlaybackCommandManager_RemovePositionReceived = 36
    IMediaPlaybackCommandManager_AddRateReceived = 37
    IMediaPlaybackCommandManager_RemoveRateReceived = 38
                        
End Enum

' ----==== Variablen ====----
Private m_pIMediaPlaybackCommandManager As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaPlaybackCommandManager)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaPlaybackCommandManager
End Property

Public Property Let Ifc(ByVal pIMediaPlaybackCommandManager As Long)
    Call ReleaseIfc(m_pIMediaPlaybackCommandManager)
    m_pIMediaPlaybackCommandManager = pIMediaPlaybackCommandManager
End Property

' IMediaPlaybackCommandManager
Public Property Get IsEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetIsEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsEnabled = Ret
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    If m_pIMediaPlaybackCommandManager <> 0& Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_PutIsEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get MediaPlayer() As MediaPlayer
    Dim Ret As MediaPlayer
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlayer As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetMediaPlayer, _
                  VarPtr(pIMediaPlayer)) = S_OK Then
            If pIMediaPlayer <> 0& Then
                Set Ret = New MediaPlayer
                Ret.Ifc = pIMediaPlayer
            End If
        End If
    End If
    Set MediaPlayer = Ret
End Property

Public Property Get PlayBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetPlayBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set PlayBehavior = Ret
End Property

Public Property Get PauseBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetPauseBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set PauseBehavior = Ret
End Property

Public Property Get NextBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetNextBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set NextBehavior = Ret
End Property

Public Property Get PreviousBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetPreviousBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set PreviousBehavior = Ret
End Property

Public Property Get FastForwardBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetFastForwardBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set FastForwardBehavior = Ret
End Property

Public Property Get RewindBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetRewindBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set RewindBehavior = Ret
End Property

Public Property Get ShuffleBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetShuffleBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set ShuffleBehavior = Ret
End Property

Public Property Get AutoRepeatModeBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetAutoRepeatModeBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set AutoRepeatModeBehavior = Ret
End Property

Public Property Get PositionBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetPositionBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set PositionBehavior = Ret
End Property

Public Property Get RateBehavior() As MediaPlaybackCommandManagerCommandBehavior
    Dim Ret As MediaPlaybackCommandManagerCommandBehavior
    If m_pIMediaPlaybackCommandManager <> 0& Then
        Dim pIMediaPlaybackCommandManagerCommandBehavior As Long
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_GetRateBehavior, _
                  VarPtr(pIMediaPlaybackCommandManagerCommandBehavior)) = S_OK Then
            If pIMediaPlaybackCommandManagerCommandBehavior <> 0& Then
                Set Ret = New MediaPlaybackCommandManagerCommandBehavior
                Ret.Ifc = pIMediaPlaybackCommandManagerCommandBehavior
            End If
        End If
    End If
    Set RateBehavior = Ret
End Property

' ----==== Functions ====----
' IMediaPlaybackCommandManager
Public Function AddPlayReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddPlayReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPlayReceived = Ret
End Function

Public Function RemovePlayReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemovePlayReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePlayReceived = Ret
End Function

Public Function AddPauseReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddPauseReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPauseReceived = Ret
End Function

Public Function RemovePauseReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemovePauseReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePauseReceived = Ret
End Function

Public Function AddNextReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddNextReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddNextReceived = Ret
End Function

Public Function RemoveNextReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemoveNextReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveNextReceived = Ret
End Function

Public Function AddPreviousReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddPreviousReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPreviousReceived = Ret
End Function

Public Function RemovePreviousReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemovePreviousReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePreviousReceived = Ret
End Function

Public Function AddFastForwardReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddFastForwardReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddFastForwardReceived = Ret
End Function

Public Function RemoveFastForwardReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemoveFastForwardReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveFastForwardReceived = Ret
End Function

Public Function AddRewindReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddRewindReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddRewindReceived = Ret
End Function

Public Function RemoveRewindReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemoveRewindReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveRewindReceived = Ret
End Function

Public Function AddShuffleReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddShuffleReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddShuffleReceived = Ret
End Function

Public Function RemoveShuffleReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemoveShuffleReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveShuffleReceived = Ret
End Function

Public Function AddAutoRepeatModeReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddAutoRepeatModeReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAutoRepeatModeReceived = Ret
End Function

Public Function RemoveAutoRepeatModeReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemoveAutoRepeatModeReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAutoRepeatModeReceived = Ret
End Function

Public Function AddPositionReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddPositionReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPositionReceived = Ret
End Function

Public Function RemovePositionReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemovePositionReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePositionReceived = Ret
End Function

Public Function AddRateReceived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackCommandManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_AddRateReceived, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddRateReceived = Ret
End Function

Public Function RemoveRateReceived(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackCommandManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackCommandManager, _
                  IMediaPlaybackCommandManager_RemoveRateReceived, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveRateReceived = Ret
End Function
    
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

