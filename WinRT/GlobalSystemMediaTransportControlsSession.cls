VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSystemMediaTransportControlsSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Control_GlobalSystemMediaTransportControlsSession As String = "Windows.Media.Control.GlobalSystemMediaTransportControlsSession"
Private Const IID_IGlobalSystemMediaTransportControlsSession As String = "{7148c835-9b14-5ae2-ab85-dc9b1c14e1a8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGlobalSystemMediaTransportControlsSession
    IGlobalSystemMediaTransportControlsSession_GetSourceAppUserModelId = 6
    IGlobalSystemMediaTransportControlsSession_TryGetMediaPropertiesAsync = 7
    IGlobalSystemMediaTransportControlsSession_GetTimelineProperties = 8
    IGlobalSystemMediaTransportControlsSession_GetPlaybackInfo = 9
    IGlobalSystemMediaTransportControlsSession_TryPlayAsync = 10
    IGlobalSystemMediaTransportControlsSession_TryPauseAsync = 11
    IGlobalSystemMediaTransportControlsSession_TryStopAsync = 12
    IGlobalSystemMediaTransportControlsSession_TryRecordAsync = 13
    IGlobalSystemMediaTransportControlsSession_TryFastForwardAsync = 14
    IGlobalSystemMediaTransportControlsSession_TryRewindAsync = 15
    IGlobalSystemMediaTransportControlsSession_TrySkipNextAsync = 16
    IGlobalSystemMediaTransportControlsSession_TrySkipPreviousAsync = 17
    IGlobalSystemMediaTransportControlsSession_TryChangeChannelUpAsync = 18
    IGlobalSystemMediaTransportControlsSession_TryChangeChannelDownAsync = 19
    IGlobalSystemMediaTransportControlsSession_TryTogglePlayPauseAsync = 20
    IGlobalSystemMediaTransportControlsSession_TryChangeAutoRepeatModeAsync = 21
    IGlobalSystemMediaTransportControlsSession_TryChangePlaybackRateAsync = 22
    IGlobalSystemMediaTransportControlsSession_TryChangeShuffleActiveAsync = 23
    IGlobalSystemMediaTransportControlsSession_TryChangePlaybackPositionAsync = 24
    IGlobalSystemMediaTransportControlsSession_AddTimelinePropertiesChanged = 25
    IGlobalSystemMediaTransportControlsSession_RemoveTimelinePropertiesChanged = 26
    IGlobalSystemMediaTransportControlsSession_AddPlaybackInfoChanged = 27
    IGlobalSystemMediaTransportControlsSession_RemovePlaybackInfoChanged = 28
    IGlobalSystemMediaTransportControlsSession_AddMediaPropertiesChanged = 29
    IGlobalSystemMediaTransportControlsSession_RemoveMediaPropertiesChanged = 30
                        
End Enum

' ----==== Variablen ====----
Private m_pIGlobalSystemMediaTransportControlsSession As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSession)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGlobalSystemMediaTransportControlsSession
End Property

Public Property Let Ifc(ByVal pIGlobalSystemMediaTransportControlsSession As Long)
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSession)
    m_pIGlobalSystemMediaTransportControlsSession = pIGlobalSystemMediaTransportControlsSession
End Property

' IGlobalSystemMediaTransportControlsSession
Public Property Get SourceAppUserModelId() As String
    Dim Ret As String
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim hString As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_GetSourceAppUserModelId, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    SourceAppUserModelId = Ret
End Property

' ----==== Functions ====----
' IGlobalSystemMediaTransportControlsSession
Public Function TryGetMediaProperties() As GlobalSystemMediaTransportControlsSessionMediaProperties
    Dim Ret As GlobalSystemMediaTransportControlsSessionMediaProperties
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim pIGlobalSystemMediaTransportControlsSessionMediaProperties As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryGetMediaPropertiesAsync, _
                  VarPtr(pIGlobalSystemMediaTransportControlsSessionMediaProperties)) = S_OK Then
            If pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
                If Await(pIGlobalSystemMediaTransportControlsSessionMediaProperties) Then
                    Set Ret = New GlobalSystemMediaTransportControlsSessionMediaProperties
                    Ret.Ifc = pIGlobalSystemMediaTransportControlsSessionMediaProperties
                End If
            End If
        End If
    End If
    Set TryGetMediaProperties = Ret
End Function

Public Function GetTimeLineProperties() As GlobalSystemMediaTransportControlsSessionTimelineProperties
    Dim Ret As GlobalSystemMediaTransportControlsSessionTimelineProperties
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim pIGlobalSystemMediaTransportControlsSessionTimelineProperties As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_GetTimelineProperties, _
                  VarPtr(pIGlobalSystemMediaTransportControlsSessionTimelineProperties)) = S_OK Then
            If pIGlobalSystemMediaTransportControlsSessionTimelineProperties <> 0& Then
                Set Ret = New GlobalSystemMediaTransportControlsSessionTimelineProperties
                Ret.Ifc = pIGlobalSystemMediaTransportControlsSessionTimelineProperties
            End If
        End If
    End If
    Set GetTimeLineProperties = Ret
End Function

Public Function GetPlaybackInfo() As GlobalSystemMediaTransportControlsSessionPlaybackInfo
    Dim Ret As GlobalSystemMediaTransportControlsSessionPlaybackInfo
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim pIGlobalSystemMediaTransportControlsSessionPlaybackInfo As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_GetPlaybackInfo, _
                  VarPtr(pIGlobalSystemMediaTransportControlsSessionPlaybackInfo)) = S_OK Then
            If pIGlobalSystemMediaTransportControlsSessionPlaybackInfo <> 0& Then
                Set Ret = New GlobalSystemMediaTransportControlsSessionPlaybackInfo
                Ret.Ifc = pIGlobalSystemMediaTransportControlsSessionPlaybackInfo
            End If
        End If
    End If
    Set GetPlaybackInfo = Ret
End Function

Public Function TryPlayAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryPlayAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryPlayAsync = Ret
End Function

Public Function TryPauseAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryPauseAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryPauseAsync = Ret
End Function

Public Function TryStopAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryStopAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryStopAsync = Ret
End Function

Public Function TryRecordAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryRecordAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryRecordAsync = Ret
End Function

Public Function TryFastForwardAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryFastForwardAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryFastForwardAsync = Ret
End Function

Public Function TryRewindAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryRewindAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryRewindAsync = Ret
End Function

Public Function TrySkipNextAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TrySkipNextAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TrySkipNextAsync = Ret
End Function

Public Function TrySkipPreviousAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TrySkipPreviousAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TrySkipPreviousAsync = Ret
End Function

Public Function TryChangeChannelUpAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryChangeChannelUpAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryChangeChannelUpAsync = Ret
End Function

Public Function TryChangeChannelDownAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryChangeChannelDownAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryChangeChannelDownAsync = Ret
End Function

Public Function TryTogglePlayPauseAsync() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryTogglePlayPauseAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryTogglePlayPauseAsync = Ret
End Function

Public Function TryChangeAutoRepeatModeAsync(ByVal requestedAutoRepeatMode As MediaPlaybackAutoRepeatMode) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryChangeAutoRepeatModeAsync, _
                  requestedAutoRepeatMode, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryChangeAutoRepeatModeAsync = Ret
End Function

Public Function TryChangePlaybackRateAsync(ByVal requestedPlaybackRate As Double) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryChangePlaybackRateAsync, _
                  requestedPlaybackRate, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryChangePlaybackRateAsync = Ret
End Function

Public Function TryChangeShuffleActiveAsync(ByVal requestedShuffleState As Boolean) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryChangeShuffleActiveAsync, _
                  requestedShuffleState, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryChangeShuffleActiveAsync = Ret
End Function

Public Function TryChangePlaybackPositionAsync(ByVal requestedPlaybackPosition As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_TryChangePlaybackPositionAsync, _
                  requestedPlaybackPosition, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    TryChangePlaybackPositionAsync = Ret
End Function

Public Function AddTimelinePropertiesChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_AddTimelinePropertiesChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddTimelinePropertiesChanged = Ret
End Function

Public Function RemoveTimelinePropertiesChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& And token <> 0@ Then
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_RemoveTimelinePropertiesChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveTimelinePropertiesChanged = Ret
End Function

Public Function AddPlaybackInfoChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_AddPlaybackInfoChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPlaybackInfoChanged = Ret
End Function

Public Function RemovePlaybackInfoChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& And token <> 0@ Then
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_RemovePlaybackInfoChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePlaybackInfoChanged = Ret
End Function

Public Function AddMediaPropertiesChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_AddMediaPropertiesChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddMediaPropertiesChanged = Ret
End Function

Public Function RemoveMediaPropertiesChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSession <> 0& And token <> 0@ Then
        If Invoke(m_pIGlobalSystemMediaTransportControlsSession, _
                  IGlobalSystemMediaTransportControlsSession_RemoveMediaPropertiesChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveMediaPropertiesChanged = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

