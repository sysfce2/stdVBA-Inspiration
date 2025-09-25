VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserNotificationChangedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_UserNotificationChangedEventArgs As String = "Windows.UI.Notifications.IUserNotificationChangedEventArgs"
Private Const IID_IUserNotificationChangedEventArgs As String = "{b6bd6839-79cf-4b25-82c0-0ce1eef81f8c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IUserNotificationChangedEventArgs
    IUserNotificationChangedEventArgs_GetChangeKind = 6
    IUserNotificationChangedEventArgs_GetUserNotificationId = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIUserNotificationChangedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUserNotificationChangedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUserNotificationChangedEventArgs
End Property

Public Property Let Ifc(ByVal pIUserNotificationChangedEventArgs As Long)
    Call ReleaseIfc(m_pIUserNotificationChangedEventArgs)
    m_pIUserNotificationChangedEventArgs = pIUserNotificationChangedEventArgs
End Property

' IUserNotificationChangedEventArgs
Public Property Get ChangeKind() As UserNotificationChangedKind
    Dim Ret As UserNotificationChangedKind
    If m_pIUserNotificationChangedEventArgs <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUserNotificationChangedEventArgs, _
                  IUserNotificationChangedEventArgs_GetChangeKind, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    ChangeKind = Ret
End Property

Public Property Get UserNotificationId() As Long
    Dim Ret As Long
    If m_pIUserNotificationChangedEventArgs <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUserNotificationChangedEventArgs, _
                  IUserNotificationChangedEventArgs_GetUserNotificationId, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    UserNotificationId = Ret
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


