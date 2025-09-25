VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceWatcher As String = "Windows.Devices.Enumeration.DeviceWatcher"
Private Const IID_IDeviceWatcher As String = "{c9eab97d-8f6b-4f96-a9f4-abc814e22271}"
Private Const IID_IDeviceWatcher2 As String = "{ff08456e-ed14-49e9-9a69-8117c54ae971}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDeviceWatcher
    IDeviceWatcher_AddAdded = 6
    IDeviceWatcher_RemoveAdded = 7
    IDeviceWatcher_AddUpdated = 8
    IDeviceWatcher_RemoveUpdated = 9
    IDeviceWatcher_AddRemoved = 10
    IDeviceWatcher_RemoveRemoved = 11
    IDeviceWatcher_AddEnumerationCompleted = 12
    IDeviceWatcher_RemoveEnumerationCompleted = 13
    IDeviceWatcher_AddStopped = 14
    IDeviceWatcher_RemoveStopped = 15
    IDeviceWatcher_GetStatus = 16
    IDeviceWatcher_Start = 17
    IDeviceWatcher_Stop = 18
                        
    ' IDeviceWatcher2
'    IDeviceWatcher2_GetBackgroundTrigger = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceWatcher As Long
Private m_pIDeviceWatcher2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceWatcher2)
    Call ReleaseIfc(m_pIDeviceWatcher)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceWatcher
End Property

Public Property Let Ifc(ByVal pIDeviceWatcher As Long)
    Call ReleaseIfc(m_pIDeviceWatcher2)
    Call ReleaseIfc(m_pIDeviceWatcher)
    m_pIDeviceWatcher = pIDeviceWatcher
    Call QueryIfc(m_pIDeviceWatcher, IID_IDeviceWatcher2, m_pIDeviceWatcher2)
End Property

' IDeviceWatcher
Public Property Get Status() As DeviceWatcherStatus
    Dim Ret As DeviceWatcherStatus
    If m_pIDeviceWatcher <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
End Property

' ----==== Functions ====----
' IDeviceWatcher
Public Function AddAdded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDeviceWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_AddAdded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAdded = Ret
End Function

Public Function RemoveAdded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_RemoveAdded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAdded = Ret
End Function

Public Function AddUpdated(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDeviceWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_AddUpdated, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddUpdated = Ret
End Function

Public Function RemoveUpdated(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_RemoveUpdated, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveUpdated = Ret
End Function

Public Function AddRemoved(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDeviceWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_AddRemoved, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddRemoved = Ret
End Function

Public Function RemoveRemoved(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_RemoveRemoved, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveRemoved = Ret
End Function

Public Function AddEnumerationCompleted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDeviceWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_AddEnumerationCompleted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddEnumerationCompleted = Ret
End Function

Public Function RemoveEnumerationCompleted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_RemoveEnumerationCompleted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveEnumerationCompleted = Ret
End Function

Public Function AddStopped(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDeviceWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_AddStopped, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStopped = Ret
End Function

Public Function RemoveStopped(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_RemoveStopped, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStopped = Ret
End Function

Public Function StartWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIDeviceWatcher <> 0& Then
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_Start) = S_OK Then
            Ret = True
        End If
    End If
    StartWatcher = Ret
End Function

Public Function StopWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIDeviceWatcher <> 0& Then
        If Invoke(m_pIDeviceWatcher, _
                  IDeviceWatcher_Stop) = S_OK Then
            Ret = True
        End If
    End If
    StopWatcher = Ret
End Function
                        
    ' IDeviceWatcher2
'    IDeviceWatcher2_GetBackgroundTrigger = 6

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


