VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastNotificationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastNotificationManager As String = "Windows.UI.Notifications.ToastNotificationManager"
Private Const IID_IToastNotificationManagerStatics As String = "{50ac103f-d235-4598-bbef-98fe4d1a3ad4}"
Private Const IID_IToastNotificationManagerStatics2 As String = "{7ab93c52-0e48-4750-ba9d-1a4113981847}"
Private Const IID_IToastNotificationManagerStatics4 As String = "{8f993fd3-e516-45fb-8130-398e93fa52c3}"
Private Const IID_IToastNotificationManagerStatics5 As String = "{d6f5f569-d40d-407c-8989-88cab42cfd14}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastNotificationManagerStatics
    IToastNotificationManagerStatics_CreateToastNotifier = 6
    IToastNotificationManagerStatics_CreateToastNotifierWithId = 7
    IToastNotificationManagerStatics_GetTemplateContent = 8
    
    ' IToastNotificationManagerStatics2
    IToastNotificationManagerStatics2_GetHistory = 6
    
    ' IToastNotificationManagerStatics4
    IToastNotificationManagerStatics4_GetForUser = 6
    IToastNotificationManagerStatics4_ConfigureNotificationMirroring = 7
    
    ' IToastNotificationManagerStatics5
    IToastNotificationManagerStatics5_GetDefault = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIToastNotificationManagerStatics As Long
Private m_pIToastNotificationManagerStatics2 As Long
Private m_pIToastNotificationManagerStatics4 As Long
Private m_pIToastNotificationManagerStatics5 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Notifications_ToastNotificationManager, _
                            IID_IToastNotificationManagerStatics, _
                            m_pIToastNotificationManagerStatics) Then
        Call QueryIfc(m_pIToastNotificationManagerStatics, _
                      IID_IToastNotificationManagerStatics2, _
                      m_pIToastNotificationManagerStatics2)
        Call QueryIfc(m_pIToastNotificationManagerStatics, _
                      IID_IToastNotificationManagerStatics4, _
                      m_pIToastNotificationManagerStatics4)
        Call QueryIfc(m_pIToastNotificationManagerStatics, _
                      IID_IToastNotificationManagerStatics5, _
                      m_pIToastNotificationManagerStatics5)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastNotificationManagerStatics5)
    Call ReleaseIfc(m_pIToastNotificationManagerStatics4)
    Call ReleaseIfc(m_pIToastNotificationManagerStatics2)
    Call ReleaseIfc(m_pIToastNotificationManagerStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastNotificationManagerStatics
End Property

' IToastNotificationManagerStatics2
Public Property Get History() As ToastNotificationHistory
    Dim Ret As ToastNotificationHistory
    If m_pIToastNotificationManagerStatics2 <> 0& Then
        Dim pIToastNotificationHistory As Long
        If Invoke(m_pIToastNotificationManagerStatics2, _
                  IToastNotificationManagerStatics2_GetHistory, _
                  VarPtr(pIToastNotificationHistory)) = S_OK Then
            If pIToastNotificationHistory <> 0& Then
                Set Ret = New ToastNotificationHistory
                Ret.Ifc = pIToastNotificationHistory
            End If
        End If
    End If
    Set History = Ret
End Property

' ----==== Functions ====----
' IToastNotificationManagerStatics
Public Function CreateToastNotifier() As ToastNotifier
    Dim Ret As ToastNotifier
    If m_pIToastNotificationManagerStatics <> 0& Then
        Dim pIToastNotifier As Long
        If Invoke(m_pIToastNotificationManagerStatics, _
                  IToastNotificationManagerStatics_CreateToastNotifier, _
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
    If m_pIToastNotificationManagerStatics <> 0& Then
        Dim hString As Long
        Dim pIToastNotifier As Long
        hString = CreateWindowsString(applicationId)
        If Invoke(m_pIToastNotificationManagerStatics, _
                  IToastNotificationManagerStatics_CreateToastNotifierWithId, _
                  hString, _
                  VarPtr(pIToastNotifier)) = S_OK Then
            If pIToastNotifier <> 0& Then
                Set Ret = New ToastNotifier
                Ret.Ifc = pIToastNotifier
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateToastNotifierWithId = Ret
End Function

Public Function GetTemplateContent(ByVal templateType As ToastTemplateType) As XmlDocument
    Dim Ret As XmlDocument
    If m_pIToastNotificationManagerStatics <> 0& Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIToastNotificationManagerStatics, _
                  IToastNotificationManagerStatics_GetTemplateContent, _
                  templateType, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                Set Ret = New XmlDocument
                Ret.Ifc = pIXmlDocument
            End If
        End If
    End If
    Set GetTemplateContent = Ret
End Function

' IToastNotificationManagerStatics4
Public Function GetForUser(ByVal value As User) As ToastNotificationManagerForUser
    Dim Ret As ToastNotificationManagerForUser
    If m_pIToastNotificationManagerStatics4 <> 0& _
       And IsNotNothing(value) Then
        Dim pIToastNotificationManagerForUser As Long
        If Invoke(m_pIToastNotificationManagerStatics4, _
                  IToastNotificationManagerStatics4_GetForUser, _
                  value.Ifc, _
                  VarPtr(pIToastNotificationManagerForUser)) = S_OK Then
            If pIToastNotificationManagerForUser <> 0& Then
                Set Ret = New ToastNotificationManagerForUser
                Ret.Ifc = pIToastNotificationManagerForUser
            End If
        End If
    End If
    Set GetForUser = Ret
End Function

Public Function ConfigureNotificationMirroring(ByVal value As NotificationMirroring) As Boolean
    Dim Ret As Boolean
    If m_pIToastNotificationManagerStatics4 <> 0& Then
        If Invoke(m_pIToastNotificationManagerStatics4, _
                  IToastNotificationManagerStatics4_ConfigureNotificationMirroring, _
                  value) = S_OK Then
            Ret = True
        End If
    End If
    ConfigureNotificationMirroring = Ret
End Function
    
' IToastNotificationManagerStatics5
Public Function GetDefault() As ToastNotificationManagerForUser
    Dim Ret As ToastNotificationManagerForUser
    If m_pIToastNotificationManagerStatics5 <> 0& Then
        Dim pIToastNotificationManagerForUser As Long
        If Invoke(m_pIToastNotificationManagerStatics5, _
                  IToastNotificationManagerStatics5_GetDefault, _
                  VarPtr(pIToastNotificationManagerForUser)) = S_OK Then
            If pIToastNotificationManagerForUser <> 0& Then
                Set Ret = New ToastNotificationManagerForUser
                Ret.Ifc = pIToastNotificationManagerForUser
            End If
        End If
    End If
    Set GetDefault = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


