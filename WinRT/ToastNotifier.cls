VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 02/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastNotifier As String = "Windows.UI.Notifications.ToastNotifier"
Private Const IID_IToastNotifier As String = "{75927b93-03f3-41ec-91d3-6e5bac1b38e7}"
Private Const IID_IToastNotifier2 As String = "{354389c6-7c01-4bd5-9c20-604340cd2b74}"
Private Const IID_IToastNotifier3 As String = "{ae75a04a-3b0c-51ad-b7e8-b08ab6052549}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastNotifier
    IToastNotifier_Show = 6
    IToastNotifier_Hide = 7
    IToastNotifier_GetSetting = 8
    IToastNotifier_AddToSchedule = 9
    IToastNotifier_RemoveFromSchedule = 10
    IToastNotifier_GetScheduledToastNotifications = 11
                        
    ' IToastNotifier2
    IToastNotifier2_UpdateWithTagAndGroup = 6
    IToastNotifier2_UpdateWithTag = 7
                        
    ' IToastNotifier3
    IToastNotifier3_AddScheduledToastNotificationShowing = 6
'    IToastNotifier3_RemoveScheduledToastNotificationShowing = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIToastNotifier As Long
Private m_pIToastNotifier2 As Long
Private m_pIToastNotifier3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastNotifier)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastNotifier
End Property

Public Property Let Ifc(ByVal pIToastNotifier As Long)
    Call ReleaseIfc(m_pIToastNotifier3)
    Call ReleaseIfc(m_pIToastNotifier2)
    Call ReleaseIfc(m_pIToastNotifier)
    m_pIToastNotifier = pIToastNotifier
    Call QueryIfc(m_pIToastNotifier, IID_IToastNotifier2, m_pIToastNotifier2)
    Call QueryIfc(m_pIToastNotifier, IID_IToastNotifier3, m_pIToastNotifier3)
End Property

'' *************** EventHandler ***************
'Public Property Let AddScheduledToastShowingEventHandler(ByVal handler As Long)
'    Dim showingToken As Long
'    If m_pIToastNotifier3 <> 0& And handler <> 0& Then
'        If Invoke(m_pIToastNotifier3, _
'                  IToastNotifier3_AddScheduledToastNotificationShowing, _
'                  handler, _
'                  VarPtr(showingToken)) = S_OK Then
'        End If
'    End If
'End Property
'' *************** EventHandler ***************

' IToastNotifier
Public Property Get Setting() As NotificationSetting
    Dim Ret As NotificationSetting
    If m_pIToastNotifier <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotifier, _
                  IToastNotifier_GetSetting, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Setting = Ret
End Property

' ----==== Functions ====----
' IToastNotifier
Public Function Show(ByVal notification As ToastNotification) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotifier <> 0& And _
       IsNotNothing(notification) Then
        If Invoke(m_pIToastNotifier, _
                  IToastNotifier_Show, _
                  notification.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    Show = Ret
End Function

Public Function Hide(ByVal notification As ToastNotification) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotifier <> 0& And _
       IsNotNothing(notification) Then
        If Invoke(m_pIToastNotifier, _
                  IToastNotifier_Hide, _
                  notification.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    Hide = Ret
End Function

Public Function AddToSchedule(ByVal ScheduledToast As ScheduledToastNotification) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotifier <> 0& And _
       IsNotNothing(ScheduledToast) Then
        If Invoke(m_pIToastNotifier, _
                  IToastNotifier_AddToSchedule, _
                  ScheduledToast.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    AddToSchedule = Ret
End Function

Public Function RemoveFromSchedule(ByVal ScheduledToast As ScheduledToastNotification) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotifier <> 0& And _
       IsNotNothing(ScheduledToast) <> 0& Then
        If Invoke(m_pIToastNotifier, _
                  IToastNotifier_RemoveFromSchedule, _
                  ScheduledToast.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    RemoveFromSchedule = Ret
End Function

Public Function GetScheduledToastNotifications() As ReadOnlyList_1 'ReadOnlyList_ScheduledToastNotification
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_ScheduledToastNotification
    If m_pIToastNotifier <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIToastNotifier, _
                  IToastNotifier_GetScheduledToastNotifications, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_ScheduledToastNotification, pIVectorView) 'ReadOnlyList_ScheduledToastNotification
            End If
        End If
    End If
    Set GetScheduledToastNotifications = Ret
End Function

' IToastNotifier2
Public Function UpdateWithTagAndGroup(ByVal Data As NotificationData, _
                                      ByVal Tag As String, _
                                      ByVal Group As String) As NotificationUpdateResult
    Dim Ret As NotificationUpdateResult
    If m_pIToastNotifier2 <> 0& And _
       IsNotNothing(Data) Then
        Dim hTag As Long
        Dim hGroup As Long
        hTag = CreateWindowsString(Tag)
        hGroup = CreateWindowsString(Group)
        Dim value As Long
        If Invoke(m_pIToastNotifier2, _
                  IToastNotifier2_UpdateWithTagAndGroup, _
                  Data.Ifc, _
                  hTag, _
                  hGroup, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
        Call DeleteWindowsString(hGroup)
        Call DeleteWindowsString(hTag)
    End If
    UpdateWithTagAndGroup = Ret
End Function

Public Function UpdateWithTag(ByVal Data As NotificationData, _
                              ByVal Tag As String) As NotificationUpdateResult
    Dim Ret As NotificationUpdateResult
    If m_pIToastNotifier2 <> 0& And _
       IsNotNothing(Data) Then
        Dim hTag As Long
        Dim value As Long
        hTag = CreateWindowsString(Tag)
        If Invoke(m_pIToastNotifier2, _
                  IToastNotifier2_UpdateWithTag, _
                  Data.Ifc, _
                  hTag, _
                  VarPtr(value)) = S_OK Then
                Ret = value
        End If
        Call DeleteWindowsString(hTag)
    End If
    UpdateWithTag = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

