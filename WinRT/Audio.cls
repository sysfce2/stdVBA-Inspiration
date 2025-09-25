VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Audio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

' Namespace Audio

Option Explicit

' ----==== Enums Windows.Media.Audio.xxx ====----

' ----==== Namespace Audio.AudioStateMonitor ====----
Public Property Get AudioStateMonitor() As AudioStateMonitor
    Set AudioStateMonitor = New AudioStateMonitor
End Property


