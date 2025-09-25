VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemMediaTransportControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SystemMediaTransportControls As String = "Windows.Media.SystemMediaTransportControls"
Private Const IID_ISystemMediaTransportControls As String = "{99fa3ff4-1742-42a6-902e-087d41f965ec}"
Private Const IID_ISystemMediaTransportControls2 As String = "{ea98d2f6-7f3c-4af2-a586-72889808efb1}"
Private Const IID_ISystemMediaTransportControlsStatics As String = "{43ba380a-eca4-4832-91ab-d415fae484c6}"
Private Const IID_ISystemMediaTransportControlsInterop As String = "{ddb0472d-c911-4a1f-86d9-dc3d71a95f5a}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISystemMediaTransportControlsStatics
    ISystemMediaTransportControlsStatics_GetForCurrentView = 6
    
    ' ISystemMediaTransportControlsInterop
    ISystemMediaTransportControlsInterop_GetForWindow = 6
    
    ' ISystemMediaTransportControls
    ISystemMediaTransportControls_GetPlaybackStatus = 6
    ISystemMediaTransportControls_PutPlaybackStatus = 7
    ISystemMediaTransportControls_GetDisplayUpdater = 8
    ISystemMediaTransportControls_GetSoundLevel = 9
    ISystemMediaTransportControls_GetIsEnabled = 10
    ISystemMediaTransportControls_PutIsEnabled = 11
    ISystemMediaTransportControls_GetIsPlayEnabled = 12
    ISystemMediaTransportControls_PutIsPlayEnabled = 13
    ISystemMediaTransportControls_GetIsStopEnabled = 14
    ISystemMediaTransportControls_PutIsStopEnabled = 15
    ISystemMediaTransportControls_GetIsPauseEnabled = 16
    ISystemMediaTransportControls_PutIsPauseEnabled = 17
    ISystemMediaTransportControls_GetIsRecordEnabled = 18
    ISystemMediaTransportControls_PutIsRecordEnabled = 19
    ISystemMediaTransportControls_GetIsFastForwardEnabled = 20
    ISystemMediaTransportControls_PutIsFastForwardEnabled = 21
    ISystemMediaTransportControls_GetIsRewindEnabled = 22
    ISystemMediaTransportControls_PutIsRewindEnabled = 23
    ISystemMediaTransportControls_GetIsPreviousEnabled = 24
    ISystemMediaTransportControls_PutIsPreviousEnabled = 25
    ISystemMediaTransportControls_GetIsNextEnabled = 26
    ISystemMediaTransportControls_PutIsNextEnabled = 27
    ISystemMediaTransportControls_GetIsChannelUpEnabled = 28
    ISystemMediaTransportControls_PutIsChannelUpEnabled = 29
    ISystemMediaTransportControls_GetIsChannelDownEnabled = 30
    ISystemMediaTransportControls_PutIsChannelDownEnabled = 31
    ISystemMediaTransportControls_AddButtonPressed = 32
    ISystemMediaTransportControls_RemoveButtonPressed = 33
    ISystemMediaTransportControls_AddPropertyChanged = 34
    ISystemMediaTransportControls_RemovePropertyChanged = 35
                    
    ' ISystemMediaTransportControls2
    ISystemMediaTransportControls2_GetAutoRepeatMode = 6
    ISystemMediaTransportControls2_PutAutoRepeatMode = 7
    ISystemMediaTransportControls2_GetShuffleEnabled = 8
    ISystemMediaTransportControls2_PutShuffleEnabled = 9
    ISystemMediaTransportControls2_GetPlaybackRate = 10
    ISystemMediaTransportControls2_PutPlaybackRate = 11
'    ISystemMediaTransportControls2_UpdateTimelineProperties = 12
    ISystemMediaTransportControls2_AddPlaybackPositionChangeRequested = 13
    ISystemMediaTransportControls2_RemovePlaybackPositionChangeRequested = 14
    ISystemMediaTransportControls2_AddPlaybackRateChangeRequested = 15
    ISystemMediaTransportControls2_RemovePlaybackRateChangeRequested = 16
    ISystemMediaTransportControls2_AddShuffleEnabledChangeRequested = 17
    ISystemMediaTransportControls2_RemoveShuffleEnabledChangeRequested = 18
    ISystemMediaTransportControls2_AddAutoRepeatModeChangeRequested = 19
    ISystemMediaTransportControls2_RemoveAutoRepeatModeChangeRequested = 20
                        
End Enum

' ----==== Variablen ====----
Private m_pISystemMediaTransportControlsStatics As Long
Private m_pISystemMediaTransportControls As Long
Private m_pISystemMediaTransportControls2 As Long
Private m_pISystemMediaTransportControlsInterop As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_SystemMediaTransportControls, _
                            IID_ISystemMediaTransportControlsStatics, _
                            m_pISystemMediaTransportControlsStatics) Then
        Call QueryIfc(m_pISystemMediaTransportControlsStatics, _
                      IID_ISystemMediaTransportControlsInterop, _
                      m_pISystemMediaTransportControlsInterop)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISystemMediaTransportControls2)
    Call ReleaseIfc(m_pISystemMediaTransportControls)
    Call ReleaseIfc(m_pISystemMediaTransportControlsStatics)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pISystemMediaTransportControls As Long)
    Call ReleaseIfc(m_pISystemMediaTransportControls2)
    Call ReleaseIfc(m_pISystemMediaTransportControls)
    m_pISystemMediaTransportControls = pISystemMediaTransportControls
    Call QueryIfc(m_pISystemMediaTransportControls, IID_ISystemMediaTransportControls2, m_pISystemMediaTransportControls2)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pISystemMediaTransportControls
End Property

' ISystemMediaTransportControls
Public Property Get PlaybackStatus() As MediaPlaybackStatus
    Dim Ret As MediaPlaybackStatus
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetPlaybackStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackStatus = Ret
End Property

Public Property Let PlaybackStatus(ByVal value As MediaPlaybackStatus)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutPlaybackStatus, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get DisplayUpdater() As SystemMediaTransportControlsDisplayUpdater
    Dim Ret As SystemMediaTransportControlsDisplayUpdater
    If m_pISystemMediaTransportControls <> 0& Then
        Dim pISystemMediaTransportControlsDisplayUpdater As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetDisplayUpdater, _
                  VarPtr(pISystemMediaTransportControlsDisplayUpdater)) = S_OK Then
            If pISystemMediaTransportControlsDisplayUpdater <> 0& Then
                Set Ret = New SystemMediaTransportControlsDisplayUpdater
                Ret.Ifc = pISystemMediaTransportControlsDisplayUpdater
            End If
        End If
    End If
    Set DisplayUpdater = Ret
End Property

Public Property Get SoundLevel() As SoundLevel
    Dim Ret As SoundLevel
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetSoundLevel, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SoundLevel = Ret
End Property

Public Property Get IsEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsEnabled = Ret
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsPlayEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsPlayEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPlayEnabled = Ret
End Property

Public Property Let IsPlayEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsPlayEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsStopEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsStopEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsStopEnabled = Ret
End Property

Public Property Let IsStopEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsStopEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsPauseEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsPauseEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPauseEnabled = Ret
End Property

Public Property Let IsPauseEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsPauseEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsRecordEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsRecordEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsRecordEnabled = Ret
End Property

Public Property Let IsRecordEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsRecordEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsFastForwardEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsFastForwardEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsFastForwardEnabled = Ret
End Property

Public Property Let IsFastForwardEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsFastForwardEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsRewindEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsRewindEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsRewindEnabled = Ret
End Property

Public Property Let IsRewindEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsRewindEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsPreviousEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsPreviousEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPreviousEnabled = Ret
End Property

Public Property Let IsPreviousEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsPreviousEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsNextEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsNextEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsNextEnabled = Ret
End Property

Public Property Let IsNextEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsNextEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsChannelUpEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsChannelUpEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsChannelUpEnabled = Ret
End Property

Public Property Let IsChannelUpEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsChannelUpEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsChannelDownEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_GetIsChannelDownEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsChannelDownEnabled = Ret
End Property

Public Property Let IsChannelDownEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls <> 0& Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_PutIsChannelDownEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

' ISystemMediaTransportControls2
Public Property Get AutoRepeatMode() As MediaPlaybackAutoRepeatMode
    Dim Ret As MediaPlaybackAutoRepeatMode
    If m_pISystemMediaTransportControls2 <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_GetAutoRepeatMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AutoRepeatMode = Ret
End Property

Public Property Let AutoRepeatMode(ByVal value As MediaPlaybackAutoRepeatMode)
    If m_pISystemMediaTransportControls2 <> 0& Then
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_PutAutoRepeatMode, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get ShuffleEnabled() As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls2 <> 0& Then
        Dim value As Long
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_GetShuffleEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    ShuffleEnabled = Ret
End Property

Public Property Let ShuffleEnabled(ByVal value As Boolean)
    If m_pISystemMediaTransportControls2 <> 0& Then
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_PutShuffleEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get PlaybackRate() As Double
    Dim Ret As Double
    If m_pISystemMediaTransportControls2 <> 0& Then
        Dim value As Double
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_GetPlaybackRate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackRate = Ret
End Property

Public Property Let PlaybackRate(ByVal value As Double)
    If m_pISystemMediaTransportControls2 <> 0& Then
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_PutPlaybackRate, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' ISystemMediaTransportControlsStatics
Public Function GetForCurrentView() As SystemMediaTransportControls
    Dim Ret As SystemMediaTransportControls
    If m_pISystemMediaTransportControlsStatics <> 0& Then
        Dim pISystemMediaTransportControls As Long
        If Invoke(m_pISystemMediaTransportControlsStatics, _
                  ISystemMediaTransportControlsStatics_GetForCurrentView, _
                  VarPtr(pISystemMediaTransportControls)) = S_OK Then
            If pISystemMediaTransportControls <> 0& Then
                Set Ret = New SystemMediaTransportControls
                Ret.Ifc = pISystemMediaTransportControls
            End If
        End If
    End If
    Set GetForCurrentView = Ret
End Function

' ISystemMediaTransportControlsInterop
Public Function GetForWindow(ByVal hWndAppWindow As Long) As SystemMediaTransportControls
    Dim Ret As SystemMediaTransportControls
    If m_pISystemMediaTransportControlsInterop <> 0& Then
        Dim pISystemMediaTransportControls As Long
        If Invoke(m_pISystemMediaTransportControlsInterop, _
                  ISystemMediaTransportControlsInterop_GetForWindow, _
                  hWndAppWindow, _
                  VarPtr(Str2Guid(IID_ISystemMediaTransportControls)), _
                  VarPtr(pISystemMediaTransportControls)) = S_OK Then
            If pISystemMediaTransportControls <> 0& Then
                Set Ret = New SystemMediaTransportControls
                Ret.Ifc = pISystemMediaTransportControls
            End If
        End If
    End If
    Set GetForWindow = Ret
End Function

' ISystemMediaTransportControls
Public Function AddButtonPressed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISystemMediaTransportControls <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_AddButtonPressed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddButtonPressed = Ret
End Function

Public Function RemoveButtonPressed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& And token <> 0@ Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_RemoveButtonPressed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveButtonPressed = Ret
End Function

Public Function AddPropertyChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISystemMediaTransportControls <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_AddPropertyChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPropertyChanged = Ret
End Function

Public Function RemovePropertyChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls <> 0& And token <> 0@ Then
        If Invoke(m_pISystemMediaTransportControls, _
                  ISystemMediaTransportControls_RemovePropertyChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePropertyChanged = Ret
End Function

' ISystemMediaTransportControls2
'    ISystemMediaTransportControls2_UpdateTimelineProperties = 12

Public Function AddPlaybackPositionChangeRequested(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISystemMediaTransportControls2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_AddPlaybackPositionChangeRequested, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPlaybackPositionChangeRequested = Ret
End Function

Public Function RemovePlaybackPositionChangeRequested(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls2 <> 0& And token <> 0@ Then
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_RemovePlaybackPositionChangeRequested, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePlaybackPositionChangeRequested = Ret
End Function

Public Function AddPlaybackRateChangeRequested(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISystemMediaTransportControls2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_AddPlaybackRateChangeRequested, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPlaybackRateChangeRequested = Ret
End Function

Public Function RemovePlaybackRateChangeRequested(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls2 <> 0& And token <> 0@ Then
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_RemovePlaybackRateChangeRequested, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePlaybackRateChangeRequested = Ret
End Function

Public Function AddShuffleEnabledChangeRequested(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISystemMediaTransportControls2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_AddShuffleEnabledChangeRequested, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddShuffleEnabledChangeRequested = Ret
End Function

Public Function RemoveShuffleEnabledChangeRequested(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls2 <> 0& And token <> 0@ Then
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_RemoveShuffleEnabledChangeRequested, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveShuffleEnabledChangeRequested = Ret
End Function

Public Function AddAutoRepeatModeChangeRequested(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pISystemMediaTransportControls2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_AddAutoRepeatModeChangeRequested, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAutoRepeatModeChangeRequested = Ret
End Function

Public Function RemoveAutoRepeatModeChangeRequested(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pISystemMediaTransportControls2 <> 0& And token <> 0@ Then
        If Invoke(m_pISystemMediaTransportControls2, _
                  ISystemMediaTransportControls2_RemoveAutoRepeatModeChangeRequested, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAutoRepeatModeChangeRequested = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


