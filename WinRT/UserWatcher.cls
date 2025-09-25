VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_UserWatcher As String = "Windows.System.UserWatcher"
Private Const IID_IUserWatcher As String = "{155eb23b-242a-45e0-a2e9-3171fc6a7fbb}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IUserWatcher
    IUserWatcher_GetStatus = 6
    IUserWatcher_Start = 7
    IUserWatcher_Stop = 8
    IUserWatcher_AddAdded = 9
    IUserWatcher_RemoveAdded = 10
    IUserWatcher_AddRemoved = 11
    IUserWatcher_RemoveRemoved = 12
    IUserWatcher_AddUpdated = 13
    IUserWatcher_RemoveUpdated = 14
    IUserWatcher_AddAuthenticationStatusChanged = 15
    IUserWatcher_RemoveAuthenticationStatusChanged = 16
    IUserWatcher_AddAuthenticationStatusChanging = 17
    IUserWatcher_RemoveAuthenticationStatusChanging = 18
    IUserWatcher_AddEnumerationCompleted = 19
    IUserWatcher_RemoveEnumerationCompleted = 20
    IUserWatcher_AddStopped = 21
    IUserWatcher_RemoveStopped = 22
                    
End Enum

' ----==== Vars ====----
Private m_pIUserWatcher As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUserWatcher)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUserWatcher
End Property

Public Property Let Ifc(ByVal pIUserWatcher As Long)
    Call ReleaseIfc(m_pIUserWatcher)
    m_pIUserWatcher = pIUserWatcher
End Property

' IUserWatcher
Public Property Get Status() As UserWatcherStatus
    Dim Ret As UserWatcherStatus
    If m_pIUserWatcher <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
End Property

' ----==== Functions ====----
' IUserWatcher
Public Function StartWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_Start) = S_OK Then
            Ret = True
        End If
    End If
    StartWatcher = Ret
End Function

Public Function StopWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_Stop) = S_OK Then
            Ret = True
        End If
    End If
    StopWatcher = Ret
End Function

Public Function AddAdded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_AddAdded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAdded = Ret
End Function

Public Function RemoveAdded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_RemoveAdded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAdded = Ret
End Function

Public Function AddRemoved(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_AddRemoved, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddRemoved = Ret
End Function

Public Function RemoveRemoved(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_RemoveRemoved, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveRemoved = Ret
End Function

Public Function AddUpdated(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_AddUpdated, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddUpdated = Ret
End Function

Public Function RemoveUpdated(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_RemoveUpdated, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveUpdated = Ret
End Function

Public Function AddAuthenticationStatusChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_AddAuthenticationStatusChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAuthenticationStatusChanged = Ret
End Function

Public Function RemoveAuthenticationStatusChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_RemoveAuthenticationStatusChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAuthenticationStatusChanged = Ret
End Function

Public Function AddAuthenticationStatusChanging(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_AddAuthenticationStatusChanging, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAuthenticationStatusChanging = Ret
End Function

Public Function RemoveAuthenticationStatusChanging(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_RemoveAuthenticationStatusChanging, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAuthenticationStatusChanging = Ret
End Function

Public Function AddEnumerationCompleted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_AddEnumerationCompleted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddEnumerationCompleted = Ret
End Function

Public Function RemoveEnumerationCompleted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_RemoveEnumerationCompleted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveEnumerationCompleted = Ret
End Function

Public Function AddStopped(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_AddStopped, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStopped = Ret
End Function

Public Function RemoveStopped(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIUserWatcher, _
                  IUserWatcher_RemoveStopped, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStopped = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

