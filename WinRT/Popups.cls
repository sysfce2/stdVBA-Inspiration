VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Popups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Popups

Option Explicit

' ----==== Enums Windows.UI.Popups.xxx ====----
Public Enum Placement
    Placement_Default = 0
    Placement_Above = 1
    Placement_Below = 2
    Placement_Left = 3
    Placement_Right = 4
End Enum

Public Enum MessageDialogOptions
    MessageDialogOptions_None = 0
    MessageDialogOptions_AcceptUserInputAfterDelay = 1
End Enum

' ----==== Namespace Popups.PopupMenu ====----
Public Property Get PopupMenu() As PopupMenu
    Set PopupMenu = New PopupMenu
End Property

' ----==== Namespace Popups.UICommand ====----
Public Property Get UICommand() As UICommand
    Set UICommand = New UICommand
End Property

' ----==== Namespace Popups.MessageDialog ====----
Public Property Get MessageDialog() As MessageDialog
    Set MessageDialog = New MessageDialog
End Property

' ----==== Namespace Popups.UICommandSeparator ====----
Public Property Get UICommandSeparator() As UICommandSeparator
    Set UICommandSeparator = New UICommandSeparator
End Property

' ----==== Namespace Popups.UICommandInvokedHandler ====----
Public Property Get UICommandInvokedHandler() As UICommandInvokedHandler
    Set UICommandInvokedHandler = New UICommandInvokedHandler
End Property

