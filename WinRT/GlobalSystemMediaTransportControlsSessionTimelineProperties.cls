VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSystemMediaTransportControlsSessionTimelineProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Control_GlobalSystemMediaTransportControlsSessionTimelineProperties As String = "Windows.Media.Control.GlobalSystemMediaTransportControlsSessionTimelineProperties"
Private Const IID_IGlobalSystemMediaTransportControlsSessionTimelineProperties As String = "{ede34136-6f25-588d-8ecf-ea5b6735aaa5}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGlobalSystemMediaTransportControlsSessionTimelineProperties
    IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetStartTime = 6
    IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetEndTime = 7
    IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetMinSeekTime = 8
    IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetMaxSeekTime = 9
    IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetPosition = 10
    IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetLastUpdatedTime = 11

End Enum

' ----==== Variablen ====----
Private m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties
End Property

Public Property Let Ifc(ByVal pIGlobalSystemMediaTransportControlsSessionTimelineProperties As Long)
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties)
    m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties = pIGlobalSystemMediaTransportControlsSessionTimelineProperties
End Property

' IGlobalSystemMediaTransportControlsSessionTimelineProperties
Public Property Get StartTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties, _
                  IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetStartTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set StartTime = Ret
End Property

Public Property Get EndTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties, _
                  IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetEndTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set EndTime = Ret
End Property

Public Property Get MinSeekTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties, _
                  IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetMinSeekTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set MinSeekTime = Ret
End Property

Public Property Get MaxSeekTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties, _
                  IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetMaxSeekTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set MaxSeekTime = Ret
End Property

Public Property Get Position() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties, _
                  IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetPosition, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set Position = Ret
End Property

Public Property Get LastUpdatedTime() As DateTime
    Dim Ret As DateTime
    If m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionTimelineProperties, _
                  IGlobalSystemMediaTransportControlsSessionTimelineProperties_GetLastUpdatedTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set LastUpdatedTime = Ret
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

