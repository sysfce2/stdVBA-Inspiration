VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Casting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Casting

Option Explicit

' ----==== Enums Windows.Media.Casting.xxx ====----
Public Enum CastingPlaybackTypes
    CastingPlaybackTypes_None = &H0&
    CastingPlaybackTypes_Audio = &H1&
    CastingPlaybackTypes_Video = &H2&
    CastingPlaybackTypes_Picture = &H4&
End Enum

Public Enum CastingConnectionState
    CastingConnectionState_Disconnected = 0
    CastingConnectionState_Connected = 1
    CastingConnectionState_Rendering = 2
    CastingConnectionState_Disconnecting = 3
    CastingConnectionState_Connecting = 4
End Enum

Public Enum CastingConnectionErrorStatus
    CastingConnectionErrorStatus_Succeeded = 0
    CastingConnectionErrorStatus_DeviceDidNotRespond = 1
    CastingConnectionErrorStatus_DeviceError = 2
    CastingConnectionErrorStatus_DeviceLocked = 3
    CastingConnectionErrorStatus_ProtectedPlaybackFailed = 4
    CastingConnectionErrorStatus_InvalidCastingSource = 5
    CastingConnectionErrorStatus_Unknown = 6
End Enum

' ----==== Namespace Casting.CastingDevice ====----
Public Property Get CastingDevice() As CastingDevice
    Set CastingDevice = New CastingDevice
End Property

' ----==== Namespace Casting.CastingSource ====----
Public Property Get CastingSource() As CastingSource
    Set CastingSource = New CastingSource
End Property

' ----==== Namespace Casting.CastingDevicePicker ====----
Public Property Get CastingDevicePicker() As CastingDevicePicker
    Set CastingDevicePicker = New CastingDevicePicker
End Property

' ----==== Namespace Casting.CastingDevicePickerFilter ====----
Public Property Get CastingDevicePickerFilter() As CastingDevicePickerFilter
    Set CastingDevicePickerFilter = New CastingDevicePickerFilter
End Property

