VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSystemMediaTransportControlsSessionPlaybackInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Control_GlobalSystemMediaTransportControlsSessionPlaybackInfo As String = "Windows.Media.Control.GlobalSystemMediaTransportControlsSessionPlaybackInfo"
Private Const IID_IGlobalSystemMediaTransportControlsSessionPlaybackInfo As String = "{94b4b6cf-e8ba-51ad-87a7-c10ade106127}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
                        
    ' IGlobalSystemMediaTransportControlsSessionPlaybackInfo
    IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetControls = 6
    IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetPlaybackStatus = 7
    IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetPlaybackType = 8
    IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetAutoRepeatMode = 9
    IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetPlaybackRate = 10
    IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetIsShuffleActive = 11
                        
    ' IReference_xxx
    IReference_GetValue = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo
End Property

Public Property Let Ifc(ByVal pIGlobalSystemMediaTransportControlsSessionPlaybackInfo As Long)
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo)
    m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo = pIGlobalSystemMediaTransportControlsSessionPlaybackInfo
End Property

' IGlobalSystemMediaTransportControlsSessionPlaybackInfo
Public Property Get GetControls() As GlobalSystemMediaTransportControlsSessionPlaybackControls
    Dim Ret As GlobalSystemMediaTransportControlsSessionPlaybackControls
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo <> 0& Then
        Dim pIGlobalSystemMediaTransportControlsSessionPlaybackControls As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetControls, _
                  VarPtr(pIGlobalSystemMediaTransportControlsSessionPlaybackControls)) = S_OK Then
            If pIGlobalSystemMediaTransportControlsSessionPlaybackControls <> 0& Then
                Set Ret = New GlobalSystemMediaTransportControlsSessionPlaybackControls
                Ret.Ifc = pIGlobalSystemMediaTransportControlsSessionPlaybackControls
            End If
        End If
    End If
    Set GetControls = Ret
End Property

Public Property Get PlaybackStatus() As GlobalSystemMediaTransportControlsSessionPlaybackStatus
    Dim Ret As GlobalSystemMediaTransportControlsSessionPlaybackStatus
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetPlaybackStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PlaybackStatus = Ret
End Property

Public Property Get PlaybackType() As MediaPlaybackType
    Dim Ret As MediaPlaybackType
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetPlaybackType, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    PlaybackType = Ret
End Property

Public Property Get AutoRepeatMode() As MediaPlaybackAutoRepeatMode
    Dim Ret As MediaPlaybackAutoRepeatMode
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetAutoRepeatMode, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    AutoRepeatMode = Ret
End Property

Public Property Get PlaybackRate() As Double
    Dim Ret As Double
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetPlaybackRate, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Double
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Ret = value
                End If
            Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    PlaybackRate = Ret
End Property

Public Property Get IsShuffleActive() As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionPlaybackInfo, _
                  IGlobalSystemMediaTransportControlsSessionPlaybackInfo_GetIsShuffleActive, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Ret = CBool(value)
                End If
            Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    IsShuffleActive = Ret
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
