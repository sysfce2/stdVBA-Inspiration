VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_UserNotification As String = "Windows.UI.Notifications.UserNotification"
Private Const IID_IUserNotification As String = "{adf7e52f-4e53-42d5-9c33-eb5ea515b23e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IUserNotification
    IUserNotification_GetNotification = 6
    IUserNotification_GetAppInfo = 7
    IUserNotification_GetId = 8
    IUserNotification_GetCreationTime = 9
                        
End Enum

' ----==== Variablen ====----
Private m_pIUserNotification As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUserNotification)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUserNotification
End Property

Public Property Let Ifc(ByVal pIUserNotification As Long)
    Call ReleaseIfc(m_pIUserNotification)
    m_pIUserNotification = pIUserNotification
End Property

' IUserNotification
Public Property Get Notification() As Notification
    Dim Ret As Notification
    If m_pIUserNotification <> 0& Then
        Dim pINotification As Long
        If Invoke(m_pIUserNotification, _
                  IUserNotification_GetNotification, _
                  VarPtr(pINotification)) = S_OK Then
            If pINotification <> 0& Then
                Set Ret = New Notification
                Ret.Ifc = pINotification
            End If
        End If
    End If
    Set Notification = Ret
End Property

Public Property Get AppInfo() As AppInfo
    Dim Ret As AppInfo
    If m_pIUserNotification <> 0& Then
        Dim pIAppInfo As Long
        If Invoke(m_pIUserNotification, _
                  IUserNotification_GetAppInfo, _
                  VarPtr(pIAppInfo)) = S_OK Then
            If pIAppInfo <> 0& Then
                Set Ret = New AppInfo
                Ret.Ifc = pIAppInfo
            End If
        End If
    End If
    Set Notification = Ret
End Property

Public Property Get Id() As Long
    Dim Ret As Long
    If m_pIUserNotification <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUserNotification, _
                  IUserNotification_GetId, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Id = Ret
End Property

Public Property Get CreationTime() As DateTime
    Dim Ret As DateTime
    If m_pIUserNotification <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIUserNotification, _
                  IUserNotification_GetCreationTime, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = Value
        End If
    End If
    Set CreationTime = Ret
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


