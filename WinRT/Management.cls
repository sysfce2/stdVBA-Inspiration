VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Management

Option Explicit

' ----==== Enums Windows.UI.Notifications.Management.xxx ====----
Public Enum NotificationKinds
    NotificationKinds_Unknown = 0
    NotificationKinds_Toast = 1
End Enum

Public Enum UserNotificationListenerAccessStatus
    UserNotificationListenerAccessStatus_Unspecified = 0
    UserNotificationListenerAccessStatus_Allowed = 1
    UserNotificationListenerAccessStatus_Denied = 2
End Enum

' ----==== Namespace Management.UserNotificationListener ====----
Public Property Get UserNotificationListener() As UserNotificationListener
    Set UserNotificationListener = New UserNotificationListener
End Property

