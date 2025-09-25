VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----==== Enums Windows_UI_ViewManagement.xxx ====----
Public Enum ViewSizePreference
    ViewSizePreference_Default = 0
    ViewSizePreference_UseLess = 1
    ViewSizePreference_UseHalf = 2
    ViewSizePreference_UseMore = 3
    ViewSizePreference_UseMinimum = 4
    ViewSizePreference_UseNone = 5
    ViewSizePreference_Custom = 6
End Enum

