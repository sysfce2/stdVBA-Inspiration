VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Windows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Windows

Option Explicit

' ----==== Namespace Windows.UI ====----
Public Property Get UI() As UI
    Set UI = New UI
End Property

' ----==== Namespace Windows.Data ====----
Public Property Get Data() As Data
    Set Data = New Data
End Property

' ----==== Namespace Windows.Media ====----
Public Property Get Media() As Media
    Set Media = New Media
End Property

' ----==== Namespace Windows.Sytem ====----
Public Property Get System() As System
    Set System = New System
End Property

' ----==== Namespace Windows.Sytem ====----
Public Property Get Storage() As Storage
    Set Storage = New Storage
End Property

' ----==== Namespace Windows.Devices ====----
Public Property Get Devices() As Devices
    Set Devices = New Devices
End Property

' ----==== Namespace Windows.Graphics ====----
Public Property Get Graphics() As Graphics
    Set Graphics = New Graphics
End Property

' ----==== Namespace Windows.Security ====----
Public Property Get Security() As Security
    Set Security = New Security
End Property

' ----==== Namespace Windows.Foundation ====----
Public Property Get Foundation() As Foundation
    Set Foundation = New Foundation
End Property

' ----==== Namespace Windows.Globalization ====----
Public Property Get Globalization() As Globalization
    Set Globalization = New Globalization
End Property

' ----==== Namespace Windows.ApplicationModel ====----
Public Property Get ApplicationModel() As ApplicationModel
    Set ApplicationModel = New ApplicationModel
End Property

