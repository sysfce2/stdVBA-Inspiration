VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppDiagnosticInfoWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_AppDiagnosticInfoWatcher As String = "Windows.System.AppDiagnosticInfoWatcher"
Private Const IID_IAppDiagnosticInfoWatcher As String = "{75575070-01d3-489a-9325-52f9cc6ede0a}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAppDiagnosticInfoWatcher
    IAppDiagnosticInfoWatcher_AddAdded = 6
    IAppDiagnosticInfoWatcher_RemoveAdded = 7
    IAppDiagnosticInfoWatcher_AddRemoved = 8
    IAppDiagnosticInfoWatcher_RemoveRemoved = 9
    IAppDiagnosticInfoWatcher_AddEnumerationCompleted = 10
    IAppDiagnosticInfoWatcher_RemoveEnumerationCompleted = 11
    IAppDiagnosticInfoWatcher_AddStopped = 12
    IAppDiagnosticInfoWatcher_RemoveStopped = 13
    IAppDiagnosticInfoWatcher_GetStatus = 14
    IAppDiagnosticInfoWatcher_Start = 15
    IAppDiagnosticInfoWatcher_Stop = 16

End Enum

' ----==== Variablen ====----
Private m_pIAppDiagnosticInfoWatcher As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAppDiagnosticInfoWatcher)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIAppDiagnosticInfoWatcher As Long)
    Call ReleaseIfc(m_pIAppDiagnosticInfoWatcher)
    m_pIAppDiagnosticInfoWatcher = pIAppDiagnosticInfoWatcher
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIAppDiagnosticInfoWatcher
End Property

' IAppDiagnosticInfoWatcher
Public Property Get Status() As AppDiagnosticInfoWatcherStatus
    Dim Ret As AppDiagnosticInfoWatcherStatus
    If m_pIAppDiagnosticInfoWatcher <> 0& Then
        Dim value As Long
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
End Property

' ----==== Functions ====----
' IAppDiagnosticInfoWatcher
Public Function AddAdded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppDiagnosticInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_AddAdded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAdded = Ret
End Function

Public Function RemoveAdded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppDiagnosticInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_RemoveAdded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAdded = Ret
End Function

Public Function AddRemoved(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppDiagnosticInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_AddRemoved, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddRemoved = Ret
End Function

Public Function RemoveRemoved(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppDiagnosticInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_RemoveRemoved, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveRemoved = Ret
End Function

Public Function AddEnumerationCompleted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppDiagnosticInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_AddEnumerationCompleted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddEnumerationCompleted = Ret
End Function

Public Function RemoveEnumerationCompleted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppDiagnosticInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_RemoveEnumerationCompleted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveEnumerationCompleted = Ret
End Function

Public Function AddStopped(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAppDiagnosticInfoWatcher <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_AddStopped, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStopped = Ret
End Function

Public Function RemoveStopped(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAppDiagnosticInfoWatcher <> 0& And token <> 0@ Then
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_RemoveStopped, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStopped = Ret
End Function

Public Function StartWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIAppDiagnosticInfoWatcher <> 0& Then
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_Start) = S_OK Then
            Ret = True
        End If
    End If
    StartWatcher = Ret
End Function

Public Function StopWatcher() As Boolean
    Dim Ret As Boolean
    If m_pIAppDiagnosticInfoWatcher <> 0& Then
        If Invoke(m_pIAppDiagnosticInfoWatcher, _
                  IAppDiagnosticInfoWatcher_Stop) = S_OK Then
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


