VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastNotification As String = "Windows.UI.Notifications.ToastNotification"
Private Const IID_IToastNotification As String = "{997e2675-059e-4e60-8b06-1760917c8b80}"
Private Const IID_IToastNotification2 As String = "{9dfb9fd1-143a-490e-90bf-b9fba7132de7}"
Private Const IID_IToastNotification3 As String = "{31e8aed8-8141-4f99-bc0a-c4ed21297d77}"
Private Const IID_IToastNotification4 As String = "{15154935-28ea-4727-88e9-c58680e2d118}"
Private Const IID_IToastNotification6 As String = "{43ebfe53-89ae-5c1e-a279-3aecfe9b6f54}"
Private Const IID_IToastNotificationFactory As String = "{04124b20-82c6-4229-b109-fd9ed4662b53}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastNotificationFactory
    IToastNotificationFactory_CreateToastNotification = 6
    
    ' IToastNotification
    IToastNotification_GetContent = 6
    IToastNotification_PutExpirationTime = 7
    IToastNotification_GetExpirationTime = 8
    IToastNotification_AddDismissed = 9
'    IToastNotification_RemoveDismissed = 10
    IToastNotification_AddActivated = 11
'    IToastNotification_RemoveActivated = 12
    IToastNotification_AddFailed = 13
'    IToastNotification_RemoveFailed = 14
    
    ' IToastNotification2
    IToastNotification2_PutTag = 6
    IToastNotification2_GetTag = 7
    IToastNotification2_PutGroup = 8
    IToastNotification2_GetGroup = 9
    IToastNotification2_PutSuppressPopup = 10
    IToastNotification2_GetSuppressPopup = 11
    
    ' IToastNotification3
    IToastNotification3_GetNotificationMirroring = 6
    IToastNotification3_PutNotificationMirroring = 7
    IToastNotification3_GetRemoteId = 8
    IToastNotification3_PutRemoteId = 9
    
    ' IToastNotification4
    IToastNotification4_GetData = 6
    IToastNotification4_PutData = 7
    IToastNotification4_GetPriority = 8
    IToastNotification4_PutPriority = 9
    
    ' IToastNotification6
    IToastNotification6_GetExpiresOnReboot = 6
    IToastNotification6_PutExpiresOnReboot = 7
    
    ' IReference_xxx
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIToastNotification As Long
Private m_pIToastNotification2 As Long
Private m_pIToastNotification3 As Long
Private m_pIToastNotification4 As Long
Private m_pIToastNotification6 As Long
Private m_pIToastNotificationFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Notifications_ToastNotification, _
                            IID_IToastNotificationFactory, _
                            m_pIToastNotificationFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastNotification6)
    Call ReleaseIfc(m_pIToastNotification4)
    Call ReleaseIfc(m_pIToastNotification3)
    Call ReleaseIfc(m_pIToastNotification2)
    Call ReleaseIfc(m_pIToastNotification)
    Call ReleaseIfc(m_pIToastNotificationFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastNotification
End Property

Public Property Let Ifc(ByVal pIToastNotification As Long)
    Call ReleaseIfc(m_pIToastNotification6)
    Call ReleaseIfc(m_pIToastNotification4)
    Call ReleaseIfc(m_pIToastNotification3)
    Call ReleaseIfc(m_pIToastNotification2)
    Call ReleaseIfc(m_pIToastNotification)
    m_pIToastNotification = pIToastNotification
    Call QueryIfc(m_pIToastNotification, IID_IToastNotification2, m_pIToastNotification2)
    Call QueryIfc(m_pIToastNotification, IID_IToastNotification3, m_pIToastNotification3)
    Call QueryIfc(m_pIToastNotification, IID_IToastNotification4, m_pIToastNotification4)
    Call QueryIfc(m_pIToastNotification, IID_IToastNotification6, m_pIToastNotification6)
End Property

' *************** EventHandler ***************
Public Property Let AddToastActivatedEventHandler(ByVal handler As Long)
    Dim token As Currency
    If m_pIToastNotification <> 0& And handler <> 0& Then
        If Invoke(m_pIToastNotification, _
                  IToastNotification_AddActivated, _
                  handler, _
                  VarPtr(token)) = S_OK Then
        End If
    End If
End Property

Public Property Let AddToastDismissedEventHandler(ByVal handler As Long)
    Dim token As Currency
    If m_pIToastNotification <> 0& And handler <> 0& Then
        If Invoke(m_pIToastNotification, _
                  IToastNotification_AddDismissed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
        End If
    End If
End Property

Public Property Let AddToastFailedEventHandler(ByVal handler As Long)
    Dim token As Currency
    If m_pIToastNotification <> 0& And handler <> 0& Then
        If Invoke(m_pIToastNotification, _
                  IToastNotification_AddFailed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
        End If
    End If
End Property
' *************** EventHandler ***************

' IToastNotification
Public Property Get Content() As XmlDocument
    Dim Ret As XmlDocument
    If m_pIToastNotification <> 0& Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIToastNotification, _
                  IToastNotification_GetContent, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                Set Ret = New XmlDocument
                Ret.Ifc = pIXmlDocument
            End If
        End If
    End If
    Set Content = Ret
End Property

Public Property Let ExpirationTime(ByVal value As DateTime)
    If m_pIToastNotification <> 0& And _
       IsNotNothing(value) Then
        Dim PropertyValue As New PropertyValue
        Dim pIReference As Long
        pIReference = PropertyValue.CreateDateTime(value)
        If Invoke(m_pIToastNotification, _
                  IToastNotification_PutExpirationTime, _
                  pIReference) = S_OK Then
        End If
        Call ReleaseIfc(pIReference)
    End If
End Property

Public Property Get ExpirationTime() As DateTime
    Dim Ret As DateTime
    If m_pIToastNotification <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIToastNotification, _
                  IToastNotification_GetExpirationTime, _
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

' IToastNotification2
Public Property Let Tag(ByVal value As String)
    If m_pIToastNotification2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIToastNotification2, _
                  IToastNotification2_PutTag, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Tag() As String
    Dim Ret As String
    If m_pIToastNotification2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotification2, _
                  IToastNotification2_GetTag, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Tag = Ret
End Property

Public Property Let Group(ByVal value As String)
    If m_pIToastNotification2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIToastNotification2, _
                  IToastNotification2_PutGroup, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Group() As String
    Dim Ret As String
    If m_pIToastNotification2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotification2, _
                  IToastNotification2_GetGroup, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Group = Ret
End Property

Public Property Let SuppressPopup(ByVal value As Boolean)
    If m_pIToastNotification2 <> 0& Then
        If Invoke(m_pIToastNotification2, _
                  IToastNotification2_PutSuppressPopup, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get SuppressPopup() As Boolean
    Dim Ret As Boolean
    If m_pIToastNotification2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotification2, _
                  IToastNotification2_GetSuppressPopup, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    SuppressPopup = Ret
End Property

' IToastNotification3
Public Property Get NotificationMirroring() As NotificationMirroring
    Dim Ret As NotificationMirroring
    If m_pIToastNotification3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotification3, _
                  IToastNotification3_GetNotificationMirroring, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NotificationMirroring = Ret
End Property

Public Property Let NotificationMirroring(ByVal value As NotificationMirroring)
    If m_pIToastNotification3 <> 0& Then
        If Invoke(m_pIToastNotification3, _
                  IToastNotification3_PutNotificationMirroring, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get RemoteId() As String
    Dim Ret As String
    If m_pIToastNotification3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotification3, _
                  IToastNotification3_GetRemoteId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    RemoteId = Ret
End Property

Public Property Let RemoteId(ByVal value As String)
    If m_pIToastNotification3 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIToastNotification3, _
                  IToastNotification3_PutRemoteId, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

' IToastNotification4
Public Property Get Data() As NotificationData
    Dim Ret As NotificationData
    If m_pIToastNotification4 <> 0& Then
        Dim pINotificationData As Long
        If Invoke(m_pIToastNotification4, _
                  IToastNotification4_GetData, _
                  VarPtr(pINotificationData)) = S_OK Then
            If pINotificationData <> 0& Then
                Set Ret = New NotificationData
                Ret.Ifc = pINotificationData
            End If
        End If
    End If
    Set Data = Ret
End Property

Public Property Let Data(ByVal value As NotificationData)
    If m_pIToastNotification4 <> 0& And _
       IsNotNothing(value) <> 0& Then
        If Invoke(m_pIToastNotification4, _
                  IToastNotification4_PutData, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get Priority() As ToastNotificationPriority
    Dim Ret As ToastNotificationPriority
    If m_pIToastNotification4 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotification4, _
                  IToastNotification4_GetPriority, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Priority = Ret
End Property

Public Property Let Priority(ByVal value As ToastNotificationPriority)
    If m_pIToastNotification4 <> 0& Then
        If Invoke(m_pIToastNotification4, _
                  IToastNotification4_PutPriority, _
                  value) = S_OK Then
        End If
    End If
End Property

' IToastNotification6
Public Property Get ExpiresOnReboot() As Boolean
    Dim Ret As Boolean
    If m_pIToastNotification6 <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastNotification6, _
                  IToastNotification6_GetExpiresOnReboot, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    ExpiresOnReboot = Ret
End Property

Public Property Let ExpiresOnReboot(ByVal value As Boolean)
    If m_pIToastNotification6 <> 0& Then
        If Invoke(m_pIToastNotification6, _
                  IToastNotification6_PutExpiresOnReboot, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IToastNotificationFactory
Public Function CreateToastNotification(ByVal Content As XmlDocument) As ToastNotification
    Dim Ret As ToastNotification
    If m_pIToastNotificationFactory <> 0& And _
       IsNotNothing(Content) Then
        Dim pIToastNotification As Long
        If Invoke(m_pIToastNotificationFactory, _
                  IToastNotificationFactory_CreateToastNotification, _
                  Content.Ifc, _
                  VarPtr(pIToastNotification)) = S_OK Then
            If pIToastNotification <> 0& Then
                Set Ret = New ToastNotification
                Ret.Ifc = pIToastNotification
            End If
        End If
    End If
    Set CreateToastNotification = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


