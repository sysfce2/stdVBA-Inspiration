VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppResourceGroupInfoWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_AppResourceGroupInfoWatcher As String = "Windows.System.AppResourceGroupInfoWatcher"
Private Const IID_IAppResourceGroupInfoWatcher As String = "{d9b0a0fd-6e5a-4c72-8b17-09fec4a212bd}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAppResourceGroupInfoWatcher
    IAppResourceGroupInfoWatcher_AddAdded = 6
    IAppResourceGroupInfoWatcher_RemoveAdded = 7
    IAppResourceGroupInfoWatcher_AddRemoved = 8
    IAppResourceGroupInfoWatcher_RemoveRemoved = 9
    IAppResourceGroupInfoWatcher_AddEnumerationCompleted = 10
    IAppResourceGroupInfoWatcher_RemoveEnumerationCompleted = 11
    IAppResourceGroupInfoWatcher_AddStopped = 12
    IAppResourceGroupInfoWatcher_RemoveStopped = 13
    IAppResourceGroupInfoWatcher_AddExecutionStateChanged = 14
    IAppResourceGroupInfoWatcher_RemoveExecutionStateChanged = 15
    IAppResourceGroupInfoWatcher_GetStatus = 16
    IAppResourceGroupInfoWatcher_Start = 17
    IAppResourceGroupInfoWatcher_Stop = 18

End Enum

' ----==== Variablen ====----
Private m_pIAppResourceGroupInfoWatcher As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAppResourceGroupInfoWatcher)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIAppResourceGroupInfoWatcher As Long)
    Call ReleaseIfc(m_pIAppResourceGroupInfoWatcher)
    m_pIAppResourceGroupInfoWatcher = pIAppResourceGroupInfoWatcher
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIAppResourceGroupInfoWatcher
End Property

' IAppResourceGroupInfoWatcher
Public Property Get Status() As AppResourceGroupInfoWatcherStatus
    Dim Ret As AppResourceGroupInfoWatcherStatus
    If m_pIAppResourceGroupInfoWatcher <> 0& Then
        Dim value As Long
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
End Property

' ----==== Functions ====----
' IAppResourceGroupInfoWatcher
Public Function AddAdded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppResourceGroupInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_AddAdded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAdded = Ret
End Function

Public Function RemoveAdded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_RemoveAdded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAdded = Ret
End Function

Public Function AddRemoved(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppResourceGroupInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_AddRemoved, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddRemoved = Ret
End Function

Public Function RemoveRemoved(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_RemoveRemoved, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveRemoved = Ret
End Function

Public Function AddEnumerationCompleted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppResourceGroupInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_AddEnumerationCompleted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddEnumerationCompleted = Ret
End Function

Public Function RemoveEnumerationCompleted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_RemoveEnumerationCompleted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveEnumerationCompleted = Ret
End Function

Public Function AddStopped(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppResourceGroupInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_AddStopped, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStopped = Ret
End Function

Public Function RemoveStopped(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_RemoveStopped, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStopped = Ret
End Function

Public Function AddExecutionStateChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppResourceGroupInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_AddExecutionStateChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddExecutionStateChanged = Ret
End Function

Public Function RemoveExecutionStateChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_RemoveExecutionStateChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveExecutionStateChanged = Ret
End Function

Public Function StartWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfoWatcher <> 0& Then
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_Start) = S_OK Then
            Ret = True
        End If
    End If
    StartWatcher = Ret
End Function

Public Function StopWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfoWatcher <> 0& Then
        If Invoke(m_pIAppResourceGroupInfoWatcher, _
                  IAppResourceGroupInfoWatcher_Stop) = S_OK Then
            Ret = True
        End If
    End If
    StopWatcher = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

