VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

' Namespace ApplicationModel

Option Explicit

' ----==== Enums Windows.ApplicationModel.xxx ====----
Public Enum AppExecutionContext
    AppExecutionContext_Unknown = 0
    AppExecutionContext_Host = 1
    AppExecutionContext_Guest = 2
End Enum

Public Enum PackageSignatureKind
    PackageSignatureKind_None = 0
    PackageSignatureKind_Developer = 1
    PackageSignatureKind_Enterprise = 2
    PackageSignatureKind_Store = 3
    PackageSignatureKind_System = 4
End Enum

' ----==== Namespace ApplicationModel.AppInfo ====----
Public Property Get AppInfo() As AppInfo
    Set AppInfo = New AppInfo
End Property

' ----==== Namespace ApplicationModel.Package ====----
Public Property Get Package() As Package
    Set Package = New Package
End Property

' ----==== Namespace ApplicationModel.PackageId ====----
Public Property Get PackageId() As PackageId
    Set PackageId = New PackageId
End Property

' ----==== Namespace ApplicationModel.PackageStatus ====----
Public Property Get PackageStatus() As PackageStatus
    Set PackageStatus = New PackageStatus
End Property

' ----==== Namespace ApplicationModel.AppDisplayInfo ====----
Public Property Get AppDisplayInfo() As AppDisplayInfo
    Set AppInfo = New AppDisplayInfo
End Property

