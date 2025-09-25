VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastNotificationHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastNotificationHistory As String = "Windows.UI.Notifications.ToastNotificationHistory"
Private Const IID_IToastNotificationHistory As String = "{5caddc63-01d3-4c97-986f-0533483fee14}"
Private Const IID_IToastNotificationHistory2 As String = "{3bc3d253-2f31-4092-9129-8ad5abf067da}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastNotificationHistory
    IToastNotificationHistory_RemoveGroup = 6
    IToastNotificationHistory_RemoveGroupWithId = 7
    IToastNotificationHistory_RemoveGroupedTagWithId = 8
    IToastNotificationHistory_RemoveGroupedTag = 9
    IToastNotificationHistory_Remove = 10
    IToastNotificationHistory_Clear = 11
    IToastNotificationHistory_ClearWithId = 12

    ' IToastNotificationHistory2
    IToastNotificationHistory2_GetHistory = 6
    IToastNotificationHistory2_GetHistoryWithId = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIToastNotificationHistory As Long
Private m_pIToastNotificationHistory2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastNotificationHistory2)
    Call ReleaseIfc(m_pIToastNotificationHistory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastNotificationHistory
End Property

Public Property Let Ifc(ByVal pIToastNotificationHistory As Long)
    Call ReleaseIfc(m_pIToastNotificationHistory2)
    Call ReleaseIfc(m_pIToastNotificationHistory)
    m_pIToastNotificationHistory = pIToastNotificationHistory
    Call QueryIfc(m_pIToastNotificationHistory, IID_IToastNotificationHistory2, m_pIToastNotificationHistory2)
End Property

' ----==== Functions ====----
' IToastNotificationHistory
Public Function RemoveGroup(ByVal Group As String) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationHistory <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Group)
        If Invoke(m_pIToastNotificationHistory, _
                  IToastNotificationHistory_RemoveGroup, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    RemoveGroup = Ret
End Function

Public Function RemoveGroupWithId(ByVal Group As String, _
                                  ByVal applicationId As String) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationHistory <> 0& Then
        Dim hGroup As Long
        Dim hApplicationId As Long
        hGroup = CreateWindowsString(Group)
        hApplicationId = CreateWindowsString(applicationId)
        If Invoke(m_pIToastNotificationHistory, _
                  IToastNotificationHistory_RemoveGroupWithId, _
                  hGroup, _
                  hApplicationId) = S_OK Then
                Ret = True
            End If
        Call DeleteWindowsString(hApplicationId)
        Call DeleteWindowsString(hGroup)
    End If
    RemoveGroupWithId = Ret
End Function

Public Function RemoveGroupedTagWithId(ByVal Tag As String, _
                                       ByVal Group As String, _
                                       ByVal applicationId As String) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationHistory <> 0& Then
        Dim hTag As Long
        Dim hGroup As Long
        Dim hApplicationId As Long
        hTag = CreateWindowsString(Tag)
        hGroup = CreateWindowsString(Group)
        hApplicationId = CreateWindowsString(applicationId)
        If Invoke(m_pIToastNotificationHistory, _
                  IToastNotificationHistory_RemoveGroupedTagWithId, _
                  hTag, _
                  hGroup, _
                  hApplicationId) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hApplicationId)
        Call DeleteWindowsString(hGroup)
        Call DeleteWindowsString(hTag)
    End If
    RemoveGroupedTagWithId = Ret
End Function

Public Function RemoveGroupedTag(ByVal Tag As String, _
                                 ByVal Group As String) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationHistory <> 0& Then
        Dim hTag As Long
        Dim hGroup As Long
        hTag = CreateWindowsString(Tag)
        hGroup = CreateWindowsString(Group)
        If Invoke(m_pIToastNotificationHistory, _
                  IToastNotificationHistory_RemoveGroupedTag, _
                  hTag, _
                  hGroup) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hGroup)
        Call DeleteWindowsString(hTag)
    End If
    RemoveGroupedTag = Ret
End Function

Public Function Remove(ByVal Tag As String) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationHistory <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Tag)
        If Invoke(m_pIToastNotificationHistory, _
                  IToastNotificationHistory_Remove, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    Remove = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationHistory <> 0& Then
        If Invoke(m_pIToastNotificationHistory, _
                  IToastNotificationHistory_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

Public Function ClearWithId(ByVal applicationId As String) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationHistory <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(applicationId)
        If Invoke(m_pIToastNotificationHistory, _
                  IToastNotificationHistory_ClearWithId, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    ClearWithId = Ret
End Function

' IToastNotificationHistory2
Public Function GetHistory() As ReadOnlyList_1 'ReadOnlyList_ToastNotification
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_ToastNotification
    If m_pIToastNotificationHistory2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIToastNotificationHistory2, _
                  IToastNotificationHistory2_GetHistory, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_ToastNotification, pIVectorView) 'ReadOnlyList_ToastNotification
            End If
        End If
    End If
    Set GetHistory = Ret
End Function

Public Function GetHistoryWithId(ByVal applicationId As String) As ReadOnlyList_1 'ReadOnlyList_ToastNotification
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_ToastNotification
    If m_pIToastNotificationHistory2 <> 0& Then
        Dim hString As Long
        Dim pIVectorView As Long
        hString = CreateWindowsString(applicationId)
        If Invoke(m_pIToastNotificationHistory2, _
                  IToastNotificationHistory2_GetHistoryWithId, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_ToastNotification, pIVectorView) 'ReadOnlyList_ToastNotification
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    GetHistoryWithId = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


