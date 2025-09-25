VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Devices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Devices

Option Explicit

' ----==== Namespace Devices.Enumeration ====----
Public Property Get Enumeration() As Enumeration
    Set Enumeration = New Enumeration
End Property

' ----==== Namespace Devices.Geolocation ====----
Public Property Get Geolocation() As Geolocation
    Set Geolocation = New Geolocation
End Property

' ----==== Namespace Devices.Input ====----
Public Property Get InputDevice() As InputDevice
    Set InputDevice = New InputDevice
End Property

