VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

' Namespace Notifications

Option Explicit

' ----==== Enums Windows.Notifications.xxx ====----
Public Enum UserNotificationChangedKind
    UserNotificationChangedKind_Added = 0
    UserNotificationChangedKind_Removed = 1
End Enum

Public Enum AdaptiveNotificationContentKind
    AdaptiveNotificationContentKind_Text = 0
End Enum

Public Enum ToastTemplateType
    ToastImageAndText01 = 0
    ToastImageAndText02 = 1
    ToastImageAndText03 = 2
    ToastImageAndText04 = 3
    ToastText01 = 4
    ToastText02 = 5
    ToastText03 = 6
    ToastText04 = 7
End Enum

Public Enum NotificationMirroring
    NotificationMirroring_Allowed = 0
    NotificationMirroring_Disabled = 1
End Enum

Public Enum ToastNotificationPriority
    ToastNotificationPriority_Default = 0
    ToastNotificationPriority_High = 1
End Enum

Public Enum NotificationSetting
    Enabled = 0
    DisabledForApplication = 1
    DisabledForUser = 2
    DisabledByGroupPolicy = 3
    DisabledByManifest = 4
End Enum

Public Enum NotificationUpdateResult
    NotificationUpdateResult_Succeeded = 0
    NotificationUpdateResult_Failed = 1
    NotificationUpdateResult_NotificationNotFound = 2
End Enum

Public Enum ToastDismissalReason
    ToastDismissalReason_UserCanceled = 0
    ToastDismissalReason_ApplicationHidden = 1
    ToastDismissalReason_TimedOut = 2
End Enum

Public Enum ToastNotificationMode
    ToastNotificationMode_Unrestricted = 0
    ToastNotificationMode_PriorityOnly = 1
    ToastNotificationMode_AlarmsOnly = 2
End Enum

' ----==== Namespace Notifications.Management ====----
Public Property Get Management() As Management
    Set Management = New Management
End Property

' ----==== Namespace Notifications.Notification ====----
Public Property Get notification() As notification
    Set notification = New notification
End Property

' ----==== Namespace Notifications.UserNotification ====----
Public Property Get UserNotification() As UserNotification
    Set UserNotification = New UserNotification
End Property

' ----==== Namespace Notifications.NotificationVisual ====----
Public Property Get NotificationVisual() As NotificationVisual
    Set NotificationVisual = New NotificationVisual
End Property

' ----==== Namespace Notifications.NotificationBinding ====----
Public Property Get NotificationBinding() As NotificationBinding
    Set NotificationBinding = New NotificationBinding
End Property

' ----==== Namespace Notifications.AdaptiveNotificationText ====----
Public Property Get AdaptiveNotificationText() As AdaptiveNotificationText
    Set AdaptiveNotificationText = New AdaptiveNotificationText
End Property

' ----==== Namespace Notifications.ToastNotificationManager ====----
Public Property Get ToastNotificationManager() As ToastNotificationManager
    Set ToastNotificationManager = New ToastNotificationManager
End Property

' ----==== Namespace Notifications.ToastNotificationHistory ====----
Public Property Get ToastNotificationHistory() As ToastNotificationHistory
    Set ToastNotificationHistory = New ToastNotificationHistory
End Property

' ----==== Namespace Notifications.ToastNotifier ====----
Public Property Get ToastNotifier() As ToastNotifier
    Set ToastNotifier = New ToastNotifier
End Property

' ----==== Namespace Notifications.ToastNotification ====----
Public Property Get ToastNotification() As ToastNotification
    Set ToastNotification = New ToastNotification
End Property

' ----==== Namespace Notifications.ScheduledToastNotification ====----
Public Property Get ScheduledToastNotification() As ScheduledToastNotification
    Set ScheduledToastNotification = New ScheduledToastNotification
End Property

' ----==== Namespace Notifications.NotificationData ====----
Public Property Get NotificationData() As NotificationData
    Set NotificationData = New NotificationData
End Property

' ----==== Namespace Notifications.ToastNotificationManagerForUser ====----
Public Property Get ToastNotificationManagerForUser() As ToastNotificationManagerForUser
    Set ToastNotificationManagerForUser = New ToastNotificationManagerForUser
End Property

