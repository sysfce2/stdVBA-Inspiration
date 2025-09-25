VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSystemMediaTransportControlsSessionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Control_GlobalSystemMediaTransportControlsSessionManager As String = "Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager"
Private Const IID_IGlobalSystemMediaTransportControlsSessionManager As String = "{cace8eac-e86e-504a-ab31-5ff8ff1bce49}"
Private Const IID_IGlobalSystemMediaTransportControlsSessionManagerStatics As String = "{2050c4ee-11a0-57de-aed7-c97c70338245}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGlobalSystemMediaTransportControlsSessionManagerStatics
    IGlobalSystemMediaTransportControlsSessionManagerStatics_RequestAsync = 6
    
    ' IGlobalSystemMediaTransportControlsSessionManager
    IGlobalSystemMediaTransportControlsSessionManager_GetCurrentSession = 6
    IGlobalSystemMediaTransportControlsSessionManager_GetSessions = 7
    IGlobalSystemMediaTransportControlsSessionManager_AddCurrentSessionChanged = 8
    IGlobalSystemMediaTransportControlsSessionManager_RemoveCurrentSessionChanged = 9
    IGlobalSystemMediaTransportControlsSessionManager_AddSessionsChanged = 10
    IGlobalSystemMediaTransportControlsSessionManager_RemoveSessionsChanged = 11
                        
End Enum

' ----==== Variablen ====----
Private m_pIGlobalSystemMediaTransportControlsSessionManager As Long
Private m_pIGlobalSystemMediaTransportControlsSessionManagerStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Control_GlobalSystemMediaTransportControlsSessionManager, _
                            IID_IGlobalSystemMediaTransportControlsSessionManagerStatics, _
                            m_pIGlobalSystemMediaTransportControlsSessionManagerStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionManager)
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionManagerStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGlobalSystemMediaTransportControlsSessionManager
End Property

Public Property Let Ifc(ByVal pIGlobalSystemMediaTransportControlsSessionManager As Long)
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionManager)
    m_pIGlobalSystemMediaTransportControlsSessionManager = pIGlobalSystemMediaTransportControlsSessionManager
End Property

' ----==== Functions ====----
' IGlobalSystemMediaTransportControlsSessionManagerStatics
Public Function RequestAsync() As GlobalSystemMediaTransportControlsSessionManager
    Dim Ret As GlobalSystemMediaTransportControlsSessionManager
    If m_pIGlobalSystemMediaTransportControlsSessionManagerStatics <> 0& Then
        Dim pIGlobalSystemMediaTransportControlsSessionManager As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionManagerStatics, _
                  IGlobalSystemMediaTransportControlsSessionManagerStatics_RequestAsync, _
                  VarPtr(pIGlobalSystemMediaTransportControlsSessionManager)) = S_OK Then
            If pIGlobalSystemMediaTransportControlsSessionManager <> 0& Then
                If Await(pIGlobalSystemMediaTransportControlsSessionManager) Then
                    Set Ret = New GlobalSystemMediaTransportControlsSessionManager
                    Ret.Ifc = pIGlobalSystemMediaTransportControlsSessionManager
                End If
            End If
        End If
    End If
    Set RequestAsync = Ret
End Function

' IGlobalSystemMediaTransportControlsSessionManager
Public Function GetCurrentSession() As GlobalSystemMediaTransportControlsSession
    Dim Ret As GlobalSystemMediaTransportControlsSession
    If m_pIGlobalSystemMediaTransportControlsSessionManager <> 0& Then
        Dim pIGlobalSystemMediaTransportControlsSession As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionManager, _
                  IGlobalSystemMediaTransportControlsSessionManager_GetCurrentSession, _
                  VarPtr(pIGlobalSystemMediaTransportControlsSession)) = S_OK Then
            If pIGlobalSystemMediaTransportControlsSession <> 0& Then
                Set Ret = New GlobalSystemMediaTransportControlsSession
                Ret.Ifc = pIGlobalSystemMediaTransportControlsSession
            End If
        End If
    End If
    Set GetCurrentSession = Ret
End Function

Public Function GetSessions() As ReadOnlyList_1 'ReadOnlyList_GlobalSystemMediaTransportControlsSession
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_GlobalSystemMediaTransportControlsSession
    If m_pIGlobalSystemMediaTransportControlsSessionManager <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionManager, _
                  IGlobalSystemMediaTransportControlsSessionManager_GetSessions, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_GlobalSystemMediaTransportControlsSession, pIVectorView) 'ReadOnlyList_GlobalSystemMediaTransportControlsSession
            End If
        End If
    End If
    Set GetSessions = Ret
End Function

Public Function AddCurrentSessionChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGlobalSystemMediaTransportControlsSessionManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionManager, _
                  IGlobalSystemMediaTransportControlsSessionManager_AddCurrentSessionChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddCurrentSessionChanged = Ret
End Function

Public Function RemoveCurrentSessionChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionManager <> 0& And token <> 0@ Then
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionManager, _
                  IGlobalSystemMediaTransportControlsSessionManager_RemoveCurrentSessionChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveCurrentSessionChanged = Ret
End Function

Public Function AddSessionsChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGlobalSystemMediaTransportControlsSessionManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionManager, _
                  IGlobalSystemMediaTransportControlsSessionManager_AddSessionsChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddSessionsChanged = Ret
End Function

Public Function RemoveSessionsChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGlobalSystemMediaTransportControlsSessionManager <> 0& And token <> 0@ Then
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionManager, _
                  IGlobalSystemMediaTransportControlsSessionManager_RemoveSessionsChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveSessionsChanged = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
