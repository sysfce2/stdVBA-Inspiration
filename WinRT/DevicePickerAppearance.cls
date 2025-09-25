VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevicePickerAppearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const WindowsDevicesEnumerationDevicePickerAppearance As String = "Windows.Devices.Enumeration.DevicePickerAppearance"
Private Const IID_IDevicePickerAppearance As String = "{e69a12c6-e627-4ed8-9b6c-460af445e56d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDevicePickerAppearance
    IDevicePickerAppearance_GetTitle = 6
    IDevicePickerAppearance_PutTitle = 7
    IDevicePickerAppearance_GetForegroundColor = 8
    IDevicePickerAppearance_PutForegroundColor = 9
    IDevicePickerAppearance_GetBackgroundColor = 10
    IDevicePickerAppearance_PutBackgroundColor = 11
    IDevicePickerAppearance_GetAccentColor = 12
    IDevicePickerAppearance_PutAccentColor = 13
    IDevicePickerAppearance_GetSelectedForegroundColor = 14
    IDevicePickerAppearance_PutSelectedForegroundColor = 15
    IDevicePickerAppearance_GetSelectedBackgroundColor = 16
    IDevicePickerAppearance_PutSelectedBackgroundColor = 17
    IDevicePickerAppearance_GetSelectedAccentColor = 18
    IDevicePickerAppearance_PutSelectedAccentColor = 19
                        
End Enum

' ----==== Variablen ====----
Private m_pIDevicePickerAppearance As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDevicePickerAppearance)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDevicePickerAppearance
End Property

Public Property Let Ifc(ByVal pIDevicePickerAppearanceAppearance As Long)
    Call ReleaseIfc(m_pIDevicePickerAppearance)
    m_pIDevicePickerAppearance = pIDevicePickerAppearanceAppearance
End Property

' IDevicePickerAppearance
Public Property Get Title() As String
    Dim Ret As String
    If m_pIDevicePickerAppearance <> 0& Then
        Dim Value As Long
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_GetTitle, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal Value As String)
    If m_pIDevicePickerAppearance <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get ForegroundColor() As Color
    Dim Ret As Color
    If m_pIDevicePickerAppearance <> 0& Then
        Dim Value As Long
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_GetForegroundColor, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set ForegroundColor = Ret
End Property

Public Property Let ForegroundColor(ByVal Value As Color)
    If m_pIDevicePickerAppearance <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_PutForegroundColor, _
                  Value.ColorValue) = S_OK Then
        End If
    End If
End Property

Public Property Get BackgroundColor() As Color
    Dim Ret As Color
    If m_pIDevicePickerAppearance <> 0& Then
        Dim Value As Long
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_GetBackgroundColor, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set BackgroundColor = Ret
End Property

Public Property Let BackgroundColor(ByVal Value As Color)
    If m_pIDevicePickerAppearance <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_PutBackgroundColor, _
                  Value.ColorValue) = S_OK Then
        End If
    End If
End Property

Public Property Get AccentColor() As Color
    Dim Ret As Color
    If m_pIDevicePickerAppearance <> 0& Then
        Dim Value As Long
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_GetAccentColor, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set AccentColor = Ret
End Property

Public Property Let AccentColor(ByVal Value As Color)
    If m_pIDevicePickerAppearance <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_PutAccentColor, _
                  Value.ColorValue) = S_OK Then
        End If
    End If
End Property

Public Property Get SelectedForegroundColor() As Color
    Dim Ret As Color
    If m_pIDevicePickerAppearance <> 0& Then
        Dim Value As Long
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_GetSelectedForegroundColor, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SelectedForegroundColor = Ret
End Property

Public Property Let SelectedForegroundColor(ByVal Value As Color)
    If m_pIDevicePickerAppearance <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_PutSelectedForegroundColor, _
                  Value.ColorValue) = S_OK Then
        End If
    End If
End Property

Public Property Get SelectedBackgroundColor() As Color
    Dim Ret As Color
    If m_pIDevicePickerAppearance <> 0& Then
        Dim Value As Long
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_GetSelectedBackgroundColor, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SelectedBackgroundColor = Ret
End Property

Public Property Let SelectedBackgroundColor(ByVal Value As Color)
    If m_pIDevicePickerAppearance <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_PutSelectedBackgroundColor, _
                  Value.ColorValue) = S_OK Then
        End If
    End If
End Property

Friend Property Get SelectedAccentColor() As Color
    Dim Ret As Color
    If m_pIDevicePickerAppearance <> 0& Then
        Dim Value As Long
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_GetSelectedAccentColor, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SelectedAccentColor = Ret
End Property

Public Property Let SelectedAccentColor(ByVal Value As Color)
    If m_pIDevicePickerAppearance <> 0& Then
        If Invoke(m_pIDevicePickerAppearance, _
                  IDevicePickerAppearance_PutSelectedAccentColor, _
                  Value.ColorValue) = S_OK Then
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


