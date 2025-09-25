VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevicePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DevicePicker As String = "Windows.Devices.Enumeration.DevicePicker"
Private Const IID_IDevicePicker As String = "{84997aa2-034a-4440-8813-7d0bd479bf5a}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDevicePicker
    IDevicePicker_GetFilter = 6
    IDevicePicker_GetAppearance = 7
    IDevicePicker_GetRequestedProperties = 8
    IDevicePicker_AddDeviceSelected = 9
    IDevicePicker_RemoveDeviceSelected = 10
    IDevicePicker_AddDisconnectButtonClicked = 11
    IDevicePicker_RemoveDisconnectButtonClicked = 12
    IDevicePicker_AddDevicePickerDismissed = 13
    IDevicePicker_RemoveDevicePickerDismissed = 14
    IDevicePicker_Show = 15
    IDevicePicker_ShowWithPlacement = 16
    IDevicePicker_PickSingleDeviceAsync = 17
    IDevicePicker_PickSingleDeviceAsyncWithPlacement = 18
    IDevicePicker_Hide = 19
    IDevicePicker_SetDisplayStatus = 20
                        
End Enum

' ----==== Variablen ====----
Private m_pIDevicePicker As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_Devices_Enumeration_DevicePicker, _
                           IID_IDevicePicker, _
                           m_pIDevicePicker) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDevicePicker)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDevicePicker
End Property

Public Property Let Ifc(ByVal pIDevicePicker As Long)
    Call ReleaseIfc(m_pIDevicePicker)
    m_pIDevicePicker = pIDevicePicker
End Property

' IDevicePicker
Public Property Get Filter() As DevicePickerFilter
    Dim Ret As DevicePickerFilter
    If m_pIDevicePicker <> 0& Then
        Dim pIDevicePickerFilter As Long
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_GetFilter, _
                  VarPtr(pIDevicePickerFilter)) = S_OK Then
            If pIDevicePickerFilter <> 0& Then
                Set Ret = New DevicePickerFilter
                Ret.Ifc = pIDevicePickerFilter
            End If
        End If
    End If
    Set Filter = Ret
End Property

Public Property Get Appearance() As DevicePickerAppearance
    Dim Ret As DevicePickerAppearance
    If m_pIDevicePicker <> 0& Then
        Dim pIDevicePickerAppearance As Long
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_GetAppearance, _
                  VarPtr(pIDevicePickerAppearance)) = S_OK Then
            If pIDevicePickerAppearance <> 0& Then
                Set Ret = New DevicePickerAppearance
                Ret.Ifc = pIDevicePickerAppearance
            End If
        End If
    End If
    Set Appearance = Ret
End Property

Public Property Get RequestedProperties() As List_String
    Dim Ret As List_String
    If m_pIDevicePicker <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_GetRequestedProperties, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set RequestedProperties = Ret
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pIDevicePicker <> 0& Then
        If InitWithWindow(m_pIDevicePicker, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' IDevicePicker
Public Function AddDeviceSelected(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDevicePicker <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_AddDeviceSelected, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddDeviceSelected = Ret
End Function

Public Function RemoveDeviceSelected(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDevicePicker <> 0& Then
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_RemoveDeviceSelected, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveDeviceSelected = Ret
End Function

Public Function AddDisconnectButtonClicked(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDevicePicker <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_AddDisconnectButtonClicked, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddDisconnectButtonClicked = Ret
End Function

Public Function RemoveDisconnectButtonClicked(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDevicePicker <> 0& Then
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_RemoveDisconnectButtonClicked, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveDisconnectButtonClicked = Ret
End Function

Public Function AddDevicePickerDismissed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDevicePicker <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_AddDevicePickerDismissed, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddDevicePickerDismissed = Ret
End Function

Public Function RemoveDevicePickerDismissed(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDevicePicker <> 0& Then
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_RemoveDevicePickerDismissed, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveDevicePickerDismissed = Ret
End Function

Public Function Show(ByVal selection As Rect) As Boolean
    Dim Ret As Boolean
    If m_pIDevicePicker <> 0& And _
       IsNotNothing(selection) Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_Show, _
                  selection.X, _
                  selection.Y, _
                  selection.Width, _
                  selection.Height) = S_OK Then
            Ret = True
        End If
    End If
    Show = Ret
End Function

Public Function ShowWithPlacement(ByVal selection As Rect, _
                                  ByVal preferredPlacement As Placement) As Boolean
    Dim Ret As Boolean
    If m_pIDevicePicker <> 0& And _
       IsNotNothing(selection) Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_ShowWithPlacement, _
                  selection.X, _
                  selection.Y, _
                  selection.Width, _
                  selection.Height, _
                  preferredPlacement) = S_OK Then
            Ret = True
        End If
    End If
    ShowWithPlacement = Ret
End Function

Public Function PickSingleDeviceAsync(ByVal selection As Rect) As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIDevicePicker <> 0& And _
       IsNotNothing(selection) Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_PickSingleDeviceAsync, _
                  selection.X, _
                  selection.Y, _
                  selection.Width, _
                  selection.Height, _
                  VarPtr(pIDeviceInformation)) = S_OK Then
            If pIDeviceInformation <> 0& Then
                If Await(pIDeviceInformation) Then
                    Set Ret = New DeviceInformation
                    Ret.Ifc = pIDeviceInformation
                End If
            End If
        End If
    End If
    Set PickSingleDeviceAsync = Ret
End Function

Public Function PickSingleDeviceAsyncWithPlacement(ByVal selection As Rect, _
                                                   ByVal preferredPlacement As Placement) As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIDevicePicker <> 0& And _
       IsNotNothing(selection) Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_PickSingleDeviceAsyncWithPlacement, _
                  selection.X, _
                  selection.Y, _
                  selection.Width, _
                  selection.Height, _
                  preferredPlacement, _
                  VarPtr(pIDeviceInformation)) = S_OK Then
            If pIDeviceInformation <> 0& Then
                If Await(pIDeviceInformation) Then
                    Set Ret = New DeviceInformation
                    Ret.Ifc = pIDeviceInformation
                End If
            End If
        End If
    End If
    Set PickSingleDeviceAsyncWithPlacement = Ret
End Function

Public Function Hide() As Boolean
    Dim Ret As Boolean
    If m_pIDevicePicker <> 0& Then
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_Hide) = S_OK Then
            Ret = True
        End If
    End If
    Hide = Ret
End Function

Public Function SetDisplayStatus(ByVal Device As DeviceInformation, _
                                 ByVal Status As String, _
                                 ByVal Options As DevicePickerDisplayStatusOptions) As Boolean
    Dim Ret As Boolean
    If m_pIDevicePicker <> 0& And _
       IsNotNothing(Device) Then
        Dim hString As Long
        hString = CreateWindowsString(Status)
        If Invoke(m_pIDevicePicker, _
                  IDevicePicker_SetDisplayStatus, _
                  Device.Ifc, _
                  hString, _
                  Options) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    SetDisplayStatus = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


