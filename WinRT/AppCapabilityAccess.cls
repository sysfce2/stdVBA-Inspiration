VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppCapabilityAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

' Namespace AppCapabilityAccess

Option Explicit

' ----==== Enums Windows.Security.Authorization.AppCapabilityAccess.xxx ====----
Public Enum AppCapabilityAccessStatus
    AppCapabilityAccessStatus_DeniedBySystem = 0
    AppCapabilityAccessStatus_NotDeclaredByApp = 1
    AppCapabilityAccessStatus_DeniedByUser = 2
    AppCapabilityAccessStatus_UserPromptRequired = 3
    AppCapabilityAccessStatus_Allowed = 4
End Enum

