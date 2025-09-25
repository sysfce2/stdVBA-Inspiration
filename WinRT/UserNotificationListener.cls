VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserNotificationListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_Management_UserNotificationListener As String = "Windows.UI.Notifications.Management.UserNotificationListener"
Private Const IID_IUserNotificationListener As String = "{62553e41-8a06-4cef-8215-6033a5be4b03}"
Private Const IID_IUserNotificationListenerStatics As String = "{ff6123cf-4386-4aa3-b73d-b804e5b63b23}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IUserNotificationListenerStatics
    IUserNotificationListenerStatics_GetCurrent = 6
    
    ' IUserNotificationListener
    IUserNotificationListener_RequestAccessAsync = 6
    IUserNotificationListener_GetAccessStatus = 7
    IUserNotificationListener_AddNotificationChanged = 8
    IUserNotificationListener_RemoveNotificationChanged = 9
    IUserNotificationListener_GetNotificationsAsync = 10
    IUserNotificationListener_GetNotification = 11
    IUserNotificationListener_ClearNotifications = 12
    IUserNotificationListener_RemoveNotification = 13
                            
End Enum

' ----==== Variablen ====----
Private m_pIUserNotificationListener As Long
Private m_pIUserNotificationListenerStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Notifications_Management_UserNotificationListener, _
                            IID_IUserNotificationListenerStatics, _
                            m_pIUserNotificationListenerStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUserNotificationListener)
    Call ReleaseIfc(m_pIUserNotificationListenerStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUserNotificationListener
End Property

Public Property Let Ifc(ByVal pIUserNotificationListener As Long)
    Call ReleaseIfc(m_pIUserNotificationListener)
    m_pIUserNotificationListener = pIUserNotificationListener
End Property

' IUserNotificationListenerStatics
Public Property Get Current() As UserNotificationListener
    Dim Ret As UserNotificationListener
    If m_pIUserNotificationListenerStatics <> 0& Then
        Dim pIUserNotificationListener As Long
        If Invoke(m_pIUserNotificationListenerStatics, _
                  IUserNotificationListenerStatics_GetCurrent, _
                  VarPtr(pIUserNotificationListener)) = S_OK Then
            If pIUserNotificationListener <> 0& Then
                Set Ret = New UserNotificationListener
                Ret.Ifc = pIUserNotificationListener
            End If
        End If
    End If
    Set Current = Ret
End Property

' ----==== Functions ====----
' IUserNotificationListener
Public Function RequestAccessAsync() As UserNotificationListenerAccessStatus
    Dim Ret As UserNotificationListenerAccessStatus
    If m_pIUserNotificationListener <> 0& Then
        Dim operation As Long
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_RequestAccessAsync, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = operation
            End If
        End If
    End If
    RequestAccessAsync = Ret
End Function

Public Function GetAccessStatus() As UserNotificationListenerAccessStatus
    Dim Ret As UserNotificationListenerAccessStatus
    If m_pIUserNotificationListener <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_GetAccessStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetAccessStatus = Ret
End Function

Public Function AddNotificationChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIUserNotificationListener <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_AddNotificationChanged, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddNotificationChanged = Ret
End Function

Public Function RemoveNotificationChanged(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIUserNotificationListener <> 0& Then
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_RemoveNotificationChanged, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveNotificationChanged = Ret
End Function

Public Function GetNotificationsAsync(ByVal kinds As NotificationKinds) As ReadOnlyList_1 'ReadOnlyList_UserNotification
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_UserNotification
    If m_pIUserNotificationListener <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_GetNotificationsAsync, _
                  kinds, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_UserNotification, pIVectorView) 'ReadOnlyList_UserNotification
                End If
            End If
        End If
    End If
    Set GetNotificationsAsync = Ret
End Function

Public Function GetNotification(ByVal notificationId As Long) As UserNotification
    Dim Ret As New UserNotification
    If m_pIUserNotificationListener <> 0& Then
        Dim pIUserNotification As Long
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_GetNotification, _
                  notificationId, _
                  VarPtr(pIUserNotification)) = S_OK Then
            Ret.Ifc = pIUserNotification
        End If
    End If
    Set GetNotification = Ret
End Function

Public Function ClearNotifications() As Boolean
    Dim Ret As Boolean
    If m_pIUserNotificationListener <> 0& Then
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_ClearNotifications) = S_OK Then
            Ret = True
        End If
    End If
    ClearNotifications = Ret
End Function

Public Function RemoveNotification(ByVal notificationId As Long) As Boolean
    Dim Ret As Boolean
    If m_pIUserNotificationListener <> 0& Then
        If Invoke(m_pIUserNotificationListener, _
                  IUserNotificationListener_RemoveNotification, _
                  notificationId) = S_OK Then
            Ret = True
        End If
    End If
    RemoveNotification = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


