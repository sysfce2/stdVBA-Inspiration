VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 03/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_Notification As String = "Windows.UI.Notifications.Notification"
Private Const IID_INotification As String = "{108037fe-eb76-4f82-97bc-da07530a2e20}"
Private Const IReference_Windows_Foundation_DateTime As String = "{5541d8a7-497c-5aa4-86fc-7713adbf2a2c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' INotification
    INotification_GetExpirationTime = 6
    INotification_PutExpirationTime = 7
    INotification_GetVisual = 8
    INotification_PutVisual = 9
                        
    ' IReference_xxx
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pINotification As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pINotification)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pINotification
End Property

Public Property Let Ifc(ByVal pINotification As Long)
    Call ReleaseIfc(m_pINotification)
    m_pINotification = pINotification
End Property

' INotification
Public Property Get ExpirationTime() As DateTime
    Dim Ret As DateTime
    If m_pINotification <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pINotification, _
                  INotification_GetExpirationTime, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Currency
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Set Ret = New DateTime
                    Ret.DateTimeValue = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Set ExpirationTime = Ret
End Property

Public Property Let ExpirationTime(ByVal value As DateTime)
    If m_pINotification <> 0& And _
       IsNotNothing(value) Then
        Dim PropertyValue As New PropertyValue
        Dim pIInspectable As Long
        pIInspectable = PropertyValue.CreateDateTime(value)
        Dim pIReference As Long
        If QueryIfc(pIInspectable, _
                    IReference_Windows_Foundation_DateTime, _
                    pIReference) Then
            If Invoke(m_pINotification, _
                      INotification_PutExpirationTime, _
                      pIReference) = S_OK Then
            End If
            Call ReleaseIfc(pIReference)
        End If
        Call ReleaseIfc(pIInspectable)
    End If
End Property

Public Property Get Visual() As NotificationVisual
    Dim Ret As NotificationVisual
    If m_pINotification <> 0& Then
        Dim pINotificationVisual As Long
        If Invoke(m_pINotification, _
                  INotification_GetVisual, _
                  VarPtr(pINotificationVisual)) = S_OK Then
            If pINotificationVisual <> 0& Then
                Set Ret = New NotificationVisual
                Ret.Ifc = pINotificationVisual
            End If
        End If
    End If
    Set Visual = Ret
End Property

Public Property Let Visual(ByVal value As NotificationVisual)
    If m_pINotification <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pINotification, _
                  INotification_PutVisual, _
                  value.Ifc) = S_OK Then
        End If
    End If
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


