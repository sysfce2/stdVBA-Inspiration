VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSystemMediaTransportControlsSessionPlaybackControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Control_GlobalSystemMediaTransportControlsSessionPlaybackControls As String = "Windows.Media.Control.GlobalSystemMediaTransportControlsSessionPlaybackControls"
Private Const IID_IGlobalSystemMediaTransportControlsSessionPlaybackControls As String = "{6501a3e6-bc7a-503a-bb1b-68f158f3fb03}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGlobalSystemMediaTransportControlsSessionPlaybackControls
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlayEnabled = 6
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPauseEnabled = 7
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsStopEnabled = 8
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsRecordEnabled = 9
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsFastForwardEnabled = 10
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsRewindEnabled = 11
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsNextEnabled = 12
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPreviousEnabled = 13
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsChannelUpEnabled = 14
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsChannelDownEnabled = 15
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlayPauseToggleEnabled = 16
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsShuffleEnabled = 17
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsRepeatEnabled = 18
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlaybackRateEnabled = 19
    IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlaybackPositionEnabled = 20
                        
End Enum

' ----==== Variablen ====----
Private m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls
End Property

Public Property Let Ifc(ByVal pIGlobalSystemMediaTransportControlsSessionPlaybackControls As Long)
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls)
    m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls = pIGlobalSystemMediaTransportControlsSessionPlaybackControls
End Property

' IGlobalSystemMediaTransportControlsSessionPlaybackControls
Public Property Get IsPlayEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlayEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPlayEnabled = Ret
End Property

Public Property Get IsPauseEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPauseEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPauseEnabled = Ret
End Property

Public Property Get IsStopEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsStopEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsStopEnabled = Ret
End Property

Public Property Get IsRecordEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsRecordEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsRecordEnabled = Ret
End Property

Public Property Get IsFastForwardEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsFastForwardEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsFastForwardEnabled = Ret
End Property

Public Property Get IsRewindEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsRewindEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsRewindEnabled = Ret
End Property

Public Property Get IsNextEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsNextEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsNextEnabled = Ret
End Property

Public Property Get IsPreviousEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPreviousEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPreviousEnabled = Ret
End Property

Public Property Get IsChannelUpEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsChannelUpEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsChannelUpEnabled = Ret
End Property

Public Property Get IsChannelDownEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsChannelDownEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsChannelDownEnabled = Ret
End Property

Public Property Get IsPlayPauseToggleEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlayPauseToggleEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPlayPauseToggleEnabled = Ret
End Property

Public Property Get IsShuffleEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsShuffleEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsShuffleEnabled = Ret
End Property

Public Property Get IsRepeatEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsRepeatEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsRepeatEnabled = Ret
End Property

Public Property Get IsPlaybackRateEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlaybackRateEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPlaybackRateEnabled = Ret
End Property

Public Property Get IsPlaybackPositionEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackControls, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackControls_GetIsPlaybackPositionEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPlaybackPositionEnabled = Ret
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
