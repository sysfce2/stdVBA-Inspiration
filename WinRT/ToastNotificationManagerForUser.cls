VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastNotificationManagerForUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastNotificationManagerForUser As String = "Windows.UI.Notifications.ToastNotificationManagerForUser"
Private Const IID_IToastNotificationManagerForUser As String = "{79ab57f6-43fe-487b-8a7f-99567200ae94}"
Private Const IID_IToastNotificationManagerForUser2 As String = "{679c64b7-81ab-42c2-8819-c958767753f4}"
Private Const IID_IToastNotificationManagerForUser3 As String = "{3efcb176-6cc1-56dc-973b-251f7aacb1c5}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastNotificationManagerForUser
    IToastNotificationManagerForUser_CreateToastNotifier = 6
    IToastNotificationManagerForUser_CreateToastNotifierWithId = 7
    IToastNotificationManagerForUser_GetHistory = 8
    IToastNotificationManagerForUser_GetUser = 9
                        
    ' IToastNotificationManagerForUser2
    IToastNotificationManagerForUser2_GetToastNotifierForToastCollectionIdAsync = 6
    IToastNotificationManagerForUser2_GetHistoryForToastCollectionIdAsync = 7
'    IToastNotificationManagerForUser2_GetToastCollectionManager = 8
'    IToastNotificationManagerForUser2_GetToastCollectionManagerWithAppId = 9
                        
    ' IToastNotificationManagerForUser3
    IToastNotificationManagerForUser3_GetNotificationMode = 6
    IToastNotificationManagerForUser3_AddNotificationModeChanged = 7
    IToastNotificationManagerForUser3_RemoveNotificationModeChanged = 8
                        
End Enum

' ----==== Variablen ====----
Private m_pIToastNotificationManagerForUser As Long
Private m_pIToastNotificationManagerForUser2 As Long
Private m_pIToastNotificationManagerForUser3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastNotificationManagerForUser3)
    Call ReleaseIfc(m_pIToastNotificationManagerForUser2)
    Call ReleaseIfc(m_pIToastNotificationManagerForUser)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastNotificationManagerForUser
End Property

Public Property Let Ifc(ByVal pIToastNotificationManagerForUser As Long)
    Call ReleaseIfc(m_pIToastNotificationManagerForUser3)
    Call ReleaseIfc(m_pIToastNotificationManagerForUser2)
    Call ReleaseIfc(m_pIToastNotificationManagerForUser)
    m_pIToastNotificationManagerForUser = pIToastNotificationManagerForUser
    Call QueryIfc(m_pIToastNotificationManagerForUser, IID_IToastNotificationManagerForUser2, m_pIToastNotificationManagerForUser2)
    Call QueryIfc(m_pIToastNotificationManagerForUser, IID_IToastNotificationManagerForUser3, m_pIToastNotificationManagerForUser3)
End Property

' IToastNotificationManagerForUser
Public Property Get History() As ToastNotificationHistory
    Dim Ret As ToastNotificationHistory
    If m_pIToastNotificationManagerForUser <> 0& Then
        Dim pIToastNotificationHistory As Long
        If Invoke(m_pIToastNotificationManagerForUser, _
                  IToastNotificationManagerForUser_GetHistory, _
                  VarPtr(pIToastNotificationHistory)) = S_OK Then
            If pIToastNotificationHistory <> 0& Then
                Set Ret = New ToastNotificationHistory
                Ret.Ifc = pIToastNotificationHistory
            End If
        End If
    End If
    Set History = Ret
End Property

Public Property Get User() As User
    Dim Ret As User
    If m_pIToastNotificationManagerForUser <> 0& Then
        Dim pIUser As Long
        If Invoke(m_pIToastNotificationManagerForUser, _
                  IToastNotificationManagerForUser_GetUser, _
                  VarPtr(pIUser)) = S_OK Then
            If pIUser <> 0& Then
                Set Ret = New User
                Ret.Ifc = pIUser
            End If
        End If
    End If
    Set User = Ret
End Property
                        
' IToastNotificationManagerForUser3
Public Property Get NotificationMode() As ToastNotificationMode
    Dim Ret As ToastNotificationMode
    If m_pIToastNotificationManagerForUser3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotificationManagerForUser3, _
                  IToastNotificationManagerForUser3_GetNotificationMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NotificationMode = Ret
End Property

' ----==== Functions ====----
' IToastNotificationManagerForUser
Public Function CreateToastNotifier() As ToastNotifier
    Dim Ret As ToastNotifier
    If m_pIToastNotificationManagerForUser <> 0& Then
        Dim pIToastNotifier As Long
        If Invoke(m_pIToastNotificationManagerForUser, _
                  IToastNotificationManagerForUser_CreateToastNotifier, _
                  VarPtr(pIToastNotifier)) = S_OK Then
            If pIToastNotifier <> 0& Then
                Set Ret = New ToastNotifier
                Ret.Ifc = pIToastNotifier
            End If
        End If
    End If
    Set CreateToastNotifier = Ret
End Function

Public Function CreateToastNotifierWithId(ByVal applicationId As String) As ToastNotifier
    Dim Ret As ToastNotifier
    If m_pIToastNotificationManagerForUser <> 0& Then
        Dim pIToastNotifier As Long
        Dim hApplicationId As Long
        hApplicationId = CreateWindowsString(applicationId)
        If Invoke(m_pIToastNotificationManagerForUser, _
                  IToastNotificationManagerForUser_CreateToastNotifierWithId, _
                  hApplicationId, _
                  VarPtr(pIToastNotifier)) = S_OK Then
            If pIToastNotifier <> 0& Then
                Set Ret = New ToastNotifier
                Ret.Ifc = pIToastNotifier
            End If
        End If
        Call DeleteWindowsString(hApplicationId)
    End If
    Set CreateToastNotifierWithId = Ret
End Function

' IToastNotificationManagerForUser2
Public Function GetToastNotifierForToastCollectionIdAsync(ByVal collectionId As String) As ToastNotifier
    Dim Ret As ToastNotifier
    If m_pIToastNotificationManagerForUser2 <> 0& Then
        Dim pIToastNotifier As Long
        Dim hCollectionId As Long
        hCollectionId = CreateWindowsString(collectionId)
        If Invoke(m_pIToastNotificationManagerForUser2, _
                  IToastNotificationManagerForUser2_GetToastNotifierForToastCollectionIdAsync, _
                  hCollectionId, _
                  VarPtr(pIToastNotifier)) = S_OK Then
            If pIToastNotifier <> 0& Then
                If Await(pIToastNotifier) Then
                    Set Ret = New ToastNotifier
                    Ret.Ifc = pIToastNotifier
                End If
            End If
        End If
        Call DeleteWindowsString(hCollectionId)
    End If
    Set GetToastNotifierForToastCollectionIdAsync = Ret
End Function

Public Function GetHistoryForToastCollectionIdAsync(ByVal collectionId As String) As ToastNotificationHistory
    Dim Ret As ToastNotificationHistory
    If m_pIToastNotificationManagerForUser2 <> 0& Then
        Dim pIToastNotificationHistory As Long
        Dim hCollectionId As Long
        hCollectionId = CreateWindowsString(collectionId)
        If Invoke(m_pIToastNotificationManagerForUser2, _
                  IToastNotificationManagerForUser2_GetHistoryForToastCollectionIdAsync, _
                  hCollectionId, _
                  VarPtr(pIToastNotificationHistory)) = S_OK Then
            If pIToastNotificationHistory <> 0& Then
                If Await(pIToastNotificationHistory) Then
                    Set Ret = New ToastNotificationHistory
                    Ret.Ifc = pIToastNotificationHistory
                End If
            End If
        End If
        Call DeleteWindowsString(hCollectionId)
    End If
    Set GetHistoryForToastCollectionIdAsync = Ret
End Function

'    IToastNotificationManagerForUser2_GetToastCollectionManager = 8
'    IToastNotificationManagerForUser2_GetToastCollectionManagerWithAppId = 9

' IToastNotificationManagerForUser3
Public Function AddNotificationModeChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIToastNotificationManagerForUser3 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIToastNotificationManagerForUser3, _
                  IToastNotificationManagerForUser3_AddNotificationModeChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddNotificationModeChanged = Ret
End Function

Public Function RemoveNotificationModeChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationManagerForUser3 <> 0& And token <> 0@ Then
        If Invoke(m_pIToastNotificationManagerForUser3, _
                  IToastNotificationManagerForUser3_RemoveNotificationModeChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveNotificationModeChanged = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


