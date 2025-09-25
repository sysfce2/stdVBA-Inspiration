VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaTimelineController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_MediaTimelineController As String = "Windows.Media.MediaTimelineController"
Private Const IID_IMediaTimelineController As String = "{8ed361f3-0b78-4360-bf71-0c841999ea1b}"
Private Const IID_IMediaTimelineController2 As String = "{ef74ea38-9e72-4df9-8355-6e90c81bbadd}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaTimelineController
    IMediaTimelineController_Start = 6
    IMediaTimelineController_Resume = 7
    IMediaTimelineController_Pause = 8
    IMediaTimelineController_GetPosition = 9
    IMediaTimelineController_PutPosition = 10
    IMediaTimelineController_GetClockRate = 11
    IMediaTimelineController_PutClockRate = 12
    IMediaTimelineController_GetState = 13
    IMediaTimelineController_AddPositionChanged = 14
    IMediaTimelineController_RemovePositionChanged = 15
    IMediaTimelineController_AddStateChanged = 16
    IMediaTimelineController_RemoveStateChanged = 17
    
    ' IMediaTimelineController2
    IMediaTimelineController2_GetDuration = 6
'    IMediaTimelineController2_PutDuration = 7
    IMediaTimelineController2_GetIsLoopingEnabled = 8
    IMediaTimelineController2_PutIsLoopingEnabled = 9
    IMediaTimelineController2_AddFailed = 10
    IMediaTimelineController2_RemoveFailed = 11
    IMediaTimelineController2_AddEnded = 12
    IMediaTimelineController2_RemoveEnded = 13
                    
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaTimelineController As Long
Private m_pIMediaTimelineController2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIMediaTimelineController As Long
    If GetActivateInstance(Windows_Media_MediaTimelineController, _
                           IID_IMediaTimelineController, _
                           pIMediaTimelineController) Then
        Me.Ifc = pIMediaTimelineController
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaTimelineController2)
    Call ReleaseIfc(m_pIMediaTimelineController)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaTimelineController
End Property

Public Property Let Ifc(ByVal pIMediaTimelineController As Long)
    Call ReleaseIfc(m_pIMediaTimelineController2)
    Call ReleaseIfc(m_pIMediaTimelineController)
    m_pIMediaTimelineController = pIMediaTimelineController
    Call QueryIfc(m_pIMediaTimelineController, IID_IMediaTimelineController2, m_pIMediaTimelineController2)
End Property

' IMediaTimelineController
Public Property Get Position() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaTimelineController <> 0& Then
        Dim value As Currency
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_GetPosition, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set Position = Ret
End Property

Public Property Let Position(ByVal value As TimeSpan)
    If m_pIMediaTimelineController <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_PutPosition, _
                  value.TimeSpanValue) = S_OK Then
        End If
    End If
End Property

Public Property Get ClockRate() As Double
    Dim Ret As Double
    If m_pIMediaTimelineController <> 0& Then
        Dim value As Double
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_GetClockRate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ClockRate = Ret
End Property

Public Property Let ClockRate(ByVal value As Double)
    If m_pIMediaTimelineController <> 0& Then
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_PutClockRate, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get State() As MediaTimelineControllerState
    Dim Ret As MediaTimelineControllerState
    If m_pIMediaTimelineController <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_GetState, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    State = Ret
End Property
    
' IMediaTimelineController2
Public Property Get Duration() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaTimelineController2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaTimelineController2, _
                  IMediaTimelineController2_GetDuration, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Currency
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Set Ret = New TimeSpan
                    Ret.TimeSpanValue = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Set Duration = Ret
End Property

'    IMediaTimelineController2_PutDuration = 7

Public Property Get IsLoopingEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaTimelineController2, _
                  IMediaTimelineController2_GetIsLoopingEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsLoopingEnabled = Ret
End Property

Public Property Let IsLoopingEnabled(ByVal value As Boolean)
    If m_pIMediaTimelineController2 <> 0& Then
        If Invoke(m_pIMediaTimelineController2, _
                  IMediaTimelineController2_PutIsLoopingEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IMediaTimelineController
Public Function StartMediaTimelineController() As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController <> 0& Then
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_Start) = S_OK Then
            Ret = True
        End If
    End If
    StartMediaTimelineController = Ret
End Function

Public Function ResumeMediaTimelineController() As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController <> 0& Then
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_Resume) = S_OK Then
            Ret = True
        End If
    End If
    ResumeMediaTimelineController = Ret
End Function

Public Function PauseMediaTimelineController() As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController <> 0& Then
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_Pause) = S_OK Then
            Ret = True
        End If
    End If
    PauseMediaTimelineController = Ret
End Function

Public Function AddPositionChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaTimelineController <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_AddPositionChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPositionChanged = Ret
End Function

Public Function RemovePositionChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_RemovePositionChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePositionChanged = Ret
End Function

Public Function AddStateChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaTimelineController <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_AddStateChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStateChanged = Ret
End Function

Public Function RemoveStateChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaTimelineController, _
                  IMediaTimelineController_RemoveStateChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStateChanged = Ret
End Function
    
' IMediaTimelineController2
Public Function AddFailed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaTimelineController2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaTimelineController2, _
                  IMediaTimelineController2_AddFailed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddFailed = Ret
End Function

Public Function RemoveFailed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController2 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaTimelineController2, _
                  IMediaTimelineController2_RemoveFailed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveFailed = Ret
End Function

Public Function AddEnded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaTimelineController2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaTimelineController2, _
                  IMediaTimelineController2_AddEnded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddEnded = Ret
End Function

Public Function RemoveEnded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaTimelineController2 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaTimelineController2, _
                  IMediaTimelineController2_RemoveEnded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveEnded = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


