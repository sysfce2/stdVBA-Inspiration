VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

' Namespace Control

Option Explicit

' ----==== Enums Windows.Media.Control.xxx ====----
Public Enum GlobalSystemMediaTransportControlsSessionPlaybackStatus
    GlobalSystemMediaTransportControlsSessionPlaybackStatus_Closed = 0
    GlobalSystemMediaTransportControlsSessionPlaybackStatus_Opened = 1
    GlobalSystemMediaTransportControlsSessionPlaybackStatus_Changing = 2
    GlobalSystemMediaTransportControlsSessionPlaybackStatus_Stopped = 3
    GlobalSystemMediaTransportControlsSessionPlaybackStatus_Playing = 4
    GlobalSystemMediaTransportControlsSessionPlaybackStatus_Paused = 5
End Enum

' ----==== Namespace Control.GlobalSystemMediaTransportControlsSessionManager ====----
Public Property Get GlobalSystemMediaTransportControlsSessionManager() As GlobalSystemMediaTransportControlsSessionManager
    Set GlobalSystemMediaTransportControlsSessionManager = New GlobalSystemMediaTransportControlsSessionManager
End Property

' ----==== Namespace Control.GlobalSystemMediaTransportControlsSession ====----
Public Property Get GlobalSystemMediaTransportControlsSession() As GlobalSystemMediaTransportControlsSession
    Set GlobalSystemMediaTransportControlsSession = New GlobalSystemMediaTransportControlsSession
End Property

' ----==== Namespace Control.GlobalSystemMediaTransportControlsSessionTimelineProperties ====----
Public Property Get GlobalSystemMediaTransportControlsSessionTimelineProperties() As GlobalSystemMediaTransportControlsSessionTimelineProperties
    Set GlobalSystemMediaTransportControlsSessionTimelineProperties = New GlobalSystemMediaTransportControlsSessionTimelineProperties
End Property

' ----==== Namespace Control.GlobalSystemMediaTransportControlsSessionPlaybackInfo ====----
Public Property Get GlobalSystemMediaTransportControlsSessionPlaybackInfo() As GlobalSystemMediaTransportControlsSessionPlaybackInfo
    Set GlobalSystemMediaTransportControlsSessionPlaybackInfo = New GlobalSystemMediaTransportControlsSessionPlaybackInfo
End Property

' ----==== Namespace Control.GlobalSystemMediaTransportControlsSessionPlaybackControls ====----
Public Property Get GlobalSystemMediaTransportControlsSessionPlaybackControls() As GlobalSystemMediaTransportControlsSessionPlaybackControls
    Set GlobalSystemMediaTransportControlsSessionPlaybackControls = New GlobalSystemMediaTransportControlsSessionPlaybackControls
End Property

' ----==== Namespace Control.GlobalSystemMediaTransportControlsSessionMediaProperties ====----
Public Property Get GlobalSystemMediaTransportControlsSessionMediaProperties() As GlobalSystemMediaTransportControlsSessionMediaProperties
    Set GlobalSystemMediaTransportControlsSessionMediaProperties = New GlobalSystemMediaTransportControlsSessionMediaProperties
End Property

