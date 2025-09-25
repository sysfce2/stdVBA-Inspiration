VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UICore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Core

Option Explicit

' ----==== Namespace Core.CoreWindowFlyout ====----
Public Property Get CoreWindowFlyout() As CoreWindowFlyout
    Set CoreWindowFlyout = New CoreWindowFlyout
End Property

' ----==== Namespace Core.CoreWindowDialog ====----
Public Property Get CoreWindowDialog() As CoreWindowDialog
    Set CoreWindowDialog = New CoreWindowDialog
End Property

' ----==== Namespace Core.CoreWindowPopupShowingEventArgs ====----
Public Property Get CoreWindowPopupShowingEventArgs() As CoreWindowPopupShowingEventArgs
    Set CoreWindowPopupShowingEventArgs = New CoreWindowPopupShowingEventArgs
End Property

