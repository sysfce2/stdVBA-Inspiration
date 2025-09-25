VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Render"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

' Namespace Render

Option Explicit

' ----==== Enums Windows.Media.Render.xxx ====----
Public Enum AudioRenderCategory
    AudioRenderCategory_Other = 0
    ' DEPRECATEDENUMERATOR ("ForegroundOnlyMedia is deprecated and might not work on all platforms. For more info, see MSDN.")
    AudioRenderCategory_ForegroundOnlyMedia = 1
    ' DEPRECATEDENUMERATOR ("BackgroundCapableMedia is deprecated and might not work on all platforms. For more info, see MSDN.")
    AudioRenderCategory_BackgroundCapableMedia = 2
    AudioRenderCategory_Communications = 3
    AudioRenderCategory_Alerts = 4
    AudioRenderCategory_SoundEffects = 5
    AudioRenderCategory_GameEffects = 6
    AudioRenderCategory_GameMedia = 7
    AudioRenderCategory_GameChat = 8
    AudioRenderCategory_Speech = 9
    AudioRenderCategory_Movie = 10
    AudioRenderCategory_Media = 11
End Enum

