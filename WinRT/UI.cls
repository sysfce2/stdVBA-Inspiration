VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace UI

Option Explicit

' ----==== Namespace UI.Core ====----
Public Property Get Core() As UICore
    Set Core = New UICore
End Property

' ----==== Namespace UI.Color ====----
Public Property Get Color(Optional ByVal value As Long) As Color
    Set Color = New Color
    Color.ColorValue = value
End Property

' ----==== Namespace UI.Colors ====----
Public Property Get Colors() As Colors
    Set Colors = New Colors
End Property

' ----==== Namespace UI.Popups ====----
Public Property Get Popups() As Popups
    Set Popups = New Popups
End Property

' ----==== Namespace UI.ViewManagement ====----
Public Property Get ViewManagement() As ViewManagement
    Set ViewManagement = New ViewManagement
End Property

' ----==== Namespace UI.Notifications ====----
Public Property Get Notifications() As Notifications
    Set Notifications = New Notifications
End Property

