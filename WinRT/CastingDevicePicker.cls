VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CastingDevicePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Casting_CastingDevicePicker As String = "Windows.Media.Casting.CastingDevicePicker"
Private Const IID_ICastingDevicePicker As String = "{dcd39924-0591-49be-aacb-4b82ee756a95}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICastingDevicePicker
    ICastingDevicePicker_GetFilter = 6
    ICastingDevicePicker_GetAppearance = 7
    ICastingDevicePicker_AddCastingDeviceSelected = 8
    ICastingDevicePicker_RemoveCastingDeviceSelected = 9
    ICastingDevicePicker_AddCastingDevicePickerDismissed = 10
    ICastingDevicePicker_RemoveCastingDevicePickerDismissed = 11
    ICastingDevicePicker_Show = 12
    ICastingDevicePicker_ShowWithPlacement = 13
    ICastingDevicePicker_Hide = 14
                        
End Enum

' ----==== Variablen ====----
Private m_pICastingDevicePicker As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_Media_Casting_CastingDevicePicker, _
                           IID_ICastingDevicePicker, _
                           m_pICastingDevicePicker) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICastingDevicePicker)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICastingDevicePicker
End Property

Public Property Let Ifc(ByVal pICastingDevicePicker As Long)
    Call ReleaseIfc(m_pICastingDevicePicker)
    m_pICastingDevicePicker = pICastingDevicePicker
End Property

' ICastingDevicePicker
Public Property Get Filter() As CastingDevicePickerFilter
    Dim Ret As CastingDevicePickerFilter
    If m_pICastingDevicePicker <> 0& Then
        Dim pICastingDevicePickerFilter As Long
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_GetFilter, _
                  VarPtr(pICastingDevicePickerFilter)) = S_OK Then
            If pICastingDevicePickerFilter <> 0& Then
                Set Ret = New CastingDevicePickerFilter
                Ret.Ifc = pICastingDevicePickerFilter
            End If
        End If
    End If
    Set Filter = Ret
End Property

Public Property Get Appearance() As DevicePickerAppearance
    Dim Ret As DevicePickerAppearance
    If m_pICastingDevicePicker <> 0& Then
        Dim pICastingDevicePickerAppearance As Long
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_GetAppearance, _
                  VarPtr(pICastingDevicePickerAppearance)) = S_OK Then
            If pICastingDevicePickerAppearance <> 0& Then
                Set Ret = New DevicePickerAppearance
                Ret.Ifc = pICastingDevicePickerAppearance
            End If
        End If
    End If
    Set Appearance = Ret
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pICastingDevicePicker <> 0& Then
        If InitWithWindow(m_pICastingDevicePicker, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' ICastingDevicePicker
Public Function AddCastingDeviceSelected(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pICastingDevicePicker <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_AddCastingDeviceSelected, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddCastingDeviceSelected = Ret
End Function

Public Function RemoveCastingDeviceSelected(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pICastingDevicePicker <> 0& Then
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_RemoveCastingDeviceSelected, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveCastingDeviceSelected = Ret
End Function

Public Function AddCastingDevicePickerDismissed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pICastingDevicePicker <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_AddCastingDevicePickerDismissed, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddCastingDevicePickerDismissed = Ret
End Function

Public Function RemoveCastingDevicePickerDismissed(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pICastingDevicePicker <> 0& Then
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_RemoveCastingDevicePickerDismissed, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveCastingDevicePickerDismissed = Ret
End Function

Public Function Show(ByVal selection As Rect) As Boolean
    Dim Ret As Boolean
    If m_pICastingDevicePicker <> 0& And _
       IsNotNothing(selection) Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_Show, _
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
    If m_pICastingDevicePicker <> 0& And _
       IsNotNothing(selection) Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_ShowWithPlacement, _
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

Public Function Hide() As Boolean
    Dim Ret As Boolean
    If m_pICastingDevicePicker <> 0& Then
        If Invoke(m_pICastingDevicePicker, _
                  ICastingDevicePicker_Hide) = S_OK Then
            Ret = True
        End If
    End If
    Hide = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


