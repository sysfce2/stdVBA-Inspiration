VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScheduledToastNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ScheduledToastNotification As String = "Windows.UI.Notifications.ScheduledToastNotification"
Private Const IID_IScheduledToastNotification As String = "{79f577f8-0de7-48cd-9740-9b370490c838}"
Private Const IID_IScheduledToastNotification2 As String = "{a66ea09c-31b4-43b0-b5dd-7a40e85363b1}"
Private Const IID_IScheduledToastNotification3 As String = "{98429e8b-bd32-4a3b-9d15-22aea49462a1}"
Private Const IID_IScheduledToastNotification4 As String = "{1d4761fd-bdef-4e4a-96be-0101369b58d2}"
Private Const IID_IScheduledToastNotificationFactory As String = "{e7bed191-0bb9-4189-8394-31761b476fd7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

     ' IScheduledToastNotificationFactory
    IScheduledToastNotificationFactory_CreateScheduledToastNotification = 6
    IScheduledToastNotificationFactory_CreateScheduledToastNotificationRecurring = 7 '<- Deprecated in Windows 10!
                        
    ' IScheduledToastNotification
    IScheduledToastNotification_GetContent = 6
    IScheduledToastNotification_GetDeliveryTime = 7
    IScheduledToastNotification_GetSnoozeInterval = 8
    IScheduledToastNotification_GetMaximumSnoozeCount = 9
    IScheduledToastNotification_PutId = 10
    IScheduledToastNotification_GetId = 11
                        
    ' IScheduledToastNotification2
    IScheduledToastNotification2_PutTag = 6
    IScheduledToastNotification2_GetTag = 7
    IScheduledToastNotification2_PutGroup = 8
    IScheduledToastNotification2_GetGroup = 9
    IScheduledToastNotification2_PutSuppressPopup = 10
    IScheduledToastNotification2_GetSuppressPopup = 11
                              
    ' IScheduledToastNotification3
    IScheduledToastNotification3_GetNotificationMirroring = 6
    IScheduledToastNotification3_PutNotificationMirroring = 7
    IScheduledToastNotification3_GetRemoteId = 8
    IScheduledToastNotification3_PutRemoteId = 9

    ' IScheduledToastNotification4
    IScheduledToastNotification4_GetExpirationTime = 6
    IScheduledToastNotification4_PutExpirationTime = 7
 
    ' IReference_xxx
    IReference_GetValue = 6
 
End Enum

' ----==== Variablen ====----
Private m_pIScheduledToastNotification As Long
Private m_pIScheduledToastNotification2 As Long
Private m_pIScheduledToastNotification3 As Long
Private m_pIScheduledToastNotification4 As Long
Private m_pIScheduledToastNotificationFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Notifications_ScheduledToastNotification, _
                            IID_IScheduledToastNotificationFactory, _
                            m_pIScheduledToastNotificationFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIScheduledToastNotification4)
    Call ReleaseIfc(m_pIScheduledToastNotification3)
    Call ReleaseIfc(m_pIScheduledToastNotification2)
    Call ReleaseIfc(m_pIScheduledToastNotification)
    Call ReleaseIfc(m_pIScheduledToastNotificationFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIScheduledToastNotification
End Property

Public Property Let Ifc(ByVal pIScheduledToastNotification As Long)
    Call ReleaseIfc(m_pIScheduledToastNotification4)
    Call ReleaseIfc(m_pIScheduledToastNotification3)
    Call ReleaseIfc(m_pIScheduledToastNotification2)
    Call ReleaseIfc(m_pIScheduledToastNotification)
    m_pIScheduledToastNotification = pIScheduledToastNotification
    Call QueryIfc(m_pIScheduledToastNotification, IID_IScheduledToastNotification2, m_pIScheduledToastNotification2)
    Call QueryIfc(m_pIScheduledToastNotification, IID_IScheduledToastNotification3, m_pIScheduledToastNotification3)
    Call QueryIfc(m_pIScheduledToastNotification, IID_IScheduledToastNotification4, m_pIScheduledToastNotification4)
End Property

' IScheduledToastNotification
Public Property Get Content() As XmlDocument
    Dim Ret As XmlDocument
    If m_pIScheduledToastNotification <> 0& Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIScheduledToastNotification, _
                  IScheduledToastNotification_GetContent, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                Set Ret = New XmlDocument
                Ret.Ifc = pIXmlDocument
            End If
        End If
    End If
    Set Content = Ret
End Property

Public Property Get DeliveryTime() As DateTime
    Dim Ret As DateTime
    If m_pIScheduledToastNotification <> 0& Then
        Dim value As Currency
        If Invoke(m_pIScheduledToastNotification, _
                  IScheduledToastNotification_GetDeliveryTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set DeliveryTime = Ret
End Property

Public Property Get snoozeInterval() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIScheduledToastNotification <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIScheduledToastNotification, _
                  IScheduledToastNotification_GetSnoozeInterval, _
                  VarPtr(pIReference)) = S_OK Then
            Dim value As Currency
            If Invoke(pIReference, _
                      IReference_GetValue, _
                      VarPtr(value)) = S_OK Then
                Set Ret = New TimeSpan
                Ret.TimeSpanValue = value
            End If
            Call ReleaseIfc(pIReference)
        End If
    End If
    Set snoozeInterval = Ret
End Property

Public Property Get maximumSnoozeCount() As Long
    Dim Ret As Long
    If m_pIScheduledToastNotification <> 0& Then
        Dim value As Long
        If Invoke(m_pIScheduledToastNotification, _
                  IScheduledToastNotification_GetMaximumSnoozeCount, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    maximumSnoozeCount = Ret
End Property

Public Property Let Id(ByVal value As String)
    If m_pIScheduledToastNotification <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIScheduledToastNotification, _
                  IScheduledToastNotification_PutId, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Id() As String
    Dim Ret As String
    If m_pIScheduledToastNotification <> 0& Then
        Dim value As Long
        If Invoke(m_pIScheduledToastNotification, _
                  IScheduledToastNotification_GetId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Id = Ret
End Property

' IScheduledToastNotification2
Public Property Let Tag(ByVal value As String)
    If m_pIScheduledToastNotification2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIScheduledToastNotification2, _
                  IScheduledToastNotification2_PutTag, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Tag() As String
    Dim Ret As String
    If m_pIScheduledToastNotification2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIScheduledToastNotification2, _
                  IScheduledToastNotification2_GetTag, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Tag = Ret
End Property

Public Property Let Group(ByVal value As String)
    If m_pIScheduledToastNotification2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIScheduledToastNotification2, _
                  IScheduledToastNotification2_PutGroup, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Group() As String
    Dim Ret As String
    If m_pIScheduledToastNotification2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIScheduledToastNotification2, _
                  IScheduledToastNotification2_GetGroup, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Group = Ret
End Property

Public Property Let SuppressPopup(ByVal value As Boolean)
    If m_pIScheduledToastNotification2 <> 0& Then
        If Invoke(m_pIScheduledToastNotification2, _
                  IScheduledToastNotification2_PutSuppressPopup, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get SuppressPopup() As Boolean
    Dim Ret As Boolean
    If m_pIScheduledToastNotification2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIScheduledToastNotification2, _
                  IScheduledToastNotification2_GetSuppressPopup, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    SuppressPopup = Ret
End Property

' IScheduledToastNotification3
Public Property Let NotificationMirroring(ByVal value As NotificationMirroring)
    If m_pIScheduledToastNotification3 <> 0& Then
        If Invoke(m_pIScheduledToastNotification3, _
                  IScheduledToastNotification3_PutNotificationMirroring, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get NotificationMirroring() As NotificationMirroring
    Dim Ret As NotificationMirroring
    If m_pIScheduledToastNotification3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIScheduledToastNotification3, _
                  IScheduledToastNotification3_GetNotificationMirroring, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NotificationMirroring = Ret
End Property

Public Property Let RemoteId(ByVal value As String)
    If m_pIScheduledToastNotification3 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIScheduledToastNotification3, _
                  IScheduledToastNotification3_PutRemoteId, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get RemoteId() As String
    Dim Ret As String
    If m_pIScheduledToastNotification3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIScheduledToastNotification3, _
                  IScheduledToastNotification3_GetRemoteId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    RemoteId = Ret
End Property

' IScheduledToastNotification4
Public Property Get ExpirationTime() As DateTime
    Dim Ret As DateTime
    If m_pIScheduledToastNotification4 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIScheduledToastNotification4, _
                  IScheduledToastNotification4_GetExpirationTime, _
                  VarPtr(pIReference)) = S_OK Then
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
    Set ExpirationTime = Ret
End Property

Public Property Let ExpirationTime(ByVal value As DateTime)
    If m_pIScheduledToastNotification4 <> 0& And _
       IsNotNothing(value) Then
        Dim PropertyValue As New PropertyValue
        Dim pIReference As Long
        pIReference = PropertyValue.CreateDateTime(value)
        If Invoke(m_pIScheduledToastNotification4, _
                  IScheduledToastNotification4_PutExpirationTime, _
                  pIReference) = S_OK Then
        End If
        Call ReleaseIfc(pIReference)
    End If
End Property

' ----==== Functions ====----
' IScheduledToastNotificationFactory
Public Function CreateScheduledToastNotification(ByVal Content As XmlDocument, _
                                                 ByVal DeliveryTime As DateTime) As ScheduledToastNotification
    Dim Ret As ScheduledToastNotification
    If m_pIScheduledToastNotificationFactory <> 0& And _
       IsNotNothing(Content) And _
       IsNotNothing(DeliveryTime) Then
        Dim pIScheduledToastNotification As Long
        If Invoke(m_pIScheduledToastNotificationFactory, _
                  IScheduledToastNotificationFactory_CreateScheduledToastNotification, _
                  Content.Ifc, _
                  DeliveryTime.DateTimeValue, _
                  VarPtr(pIScheduledToastNotification)) = S_OK Then
            If pIScheduledToastNotification <> 0& Then
                Set Ret = New ScheduledToastNotification
                Ret.Ifc = pIScheduledToastNotification
            End If
        End If
    End If
    Set CreateScheduledToastNotification = Ret
End Function

' Deprecated in Windows 10!
Public Function CreateScheduledToastNotificationRecurring(ByVal Content As XmlDocument, _
                                                          ByVal DeliveryTime As DateTime, _
                                                          ByVal snoozeInterval As TimeSpan, _
                                                          ByVal maximumSnoozeCount As Long) As ScheduledToastNotification
    Dim Ret As ScheduledToastNotification
    If m_pIScheduledToastNotificationFactory <> 0& And _
       IsNotNothing(Content) And _
       IsNotNothing(DeliveryTime) And _
       IsNotNothing(snoozeInterval) Then
        Dim pIScheduledToastNotification As Long
        If Invoke(m_pIScheduledToastNotificationFactory, _
                  IScheduledToastNotificationFactory_CreateScheduledToastNotificationRecurring, _
                  Content.Ifc, _
                  DeliveryTime.DateTimeValue, _
                  snoozeInterval.TimeSpanValue, _
                  maximumSnoozeCount, _
                  VarPtr(pIScheduledToastNotification)) = S_OK Then
            If pIScheduledToastNotification <> 0& Then
                Set Ret = New ScheduledToastNotification
                Ret.Ifc = pIScheduledToastNotification
            End If
        End If
    End If
    Set CreateScheduledToastNotificationRecurring = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


