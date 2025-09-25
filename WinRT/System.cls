VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace System

Option Explicit

' ----==== Enums Windows_System.xxx ====----
Public Enum LaunchUriStatus
    LaunchUriStatus_Success = 0
    LaunchUriStatus_AppUnavailable = 1
    LaunchUriStatus_ProtocolUnavailable = 2
    LaunchUriStatus_Unknown = 3
End Enum

Public Enum ProcessorArchitecture
    ProcessorArchitecture_X86 = 0
    ProcessorArchitecture_Arm = 5
    ProcessorArchitecture_X64 = 9
    ProcessorArchitecture_Neutral = 11
    ProcessorArchitecture_Arm64 = 12
    ProcessorArchitecture_X86OnArm64 = 14
    ProcessorArchitecture_Unknown = 65535
End Enum

Public Enum DiagnosticAccessStatus
    DiagnosticAccessStatus_Unspecified = 0
    DiagnosticAccessStatus_Denied = 1
    DiagnosticAccessStatus_Limited = 2
    DiagnosticAccessStatus_Allowed = 3
End Enum

Public Enum AppDiagnosticInfoWatcherStatus
    AppDiagnosticInfoWatcherStatus_Created = 0
    AppDiagnosticInfoWatcherStatus_Started = 1
    AppDiagnosticInfoWatcherStatus_EnumerationCompleted = 2
    AppDiagnosticInfoWatcherStatus_Stopping = 3
    AppDiagnosticInfoWatcherStatus_Stopped = 4
    AppDiagnosticInfoWatcherStatus_Aborted = 5
End Enum

Public Enum AppResourceGroupInfoWatcherStatus
    AppResourceGroupInfoWatcherStatus_Created = 0
    AppResourceGroupInfoWatcherStatus_Started = 1
    AppResourceGroupInfoWatcherStatus_EnumerationCompleted = 2
    AppResourceGroupInfoWatcherStatus_Stopping = 3
    AppResourceGroupInfoWatcherStatus_Stopped = 4
    AppResourceGroupInfoWatcherStatus_Aborted = 5
End Enum

Public Enum UserAuthenticationStatus
    UserAuthenticationStatus_Unauthenticated = 0
    UserAuthenticationStatus_LocallyAuthenticated = 1
    UserAuthenticationStatus_RemotelyAuthenticated = 2
End Enum

Public Enum UserType
    UserType_LocalUser = 0
    UserType_RemoteUser = 1
    UserType_LocalGuest = 2
    UserType_RemoteGuest = 3
    UserType_SystemManaged = 4
End Enum

Public Enum UserPictureSize
    UserPictureSize_Size64x64 = 0
    UserPictureSize_Size208x208 = 1
    UserPictureSize_Size424x424 = 2
    UserPictureSize_Size1080x1080 = 3
End Enum

Public Enum UserAgeConsentGroup
    UserAgeConsentGroup_Child = 0
    UserAgeConsentGroup_Minor = 1
    UserAgeConsentGroup_Adult = 2
End Enum

Public Enum UserAgeConsentResult
    UserAgeConsentResult_NotEnforced = 0
    UserAgeConsentResult_Included = 1
    UserAgeConsentResult_NotIncluded = 2
    UserAgeConsentResult_Unknown = 3
    UserAgeConsentResult_Ambiguous = 4
End Enum

Public Enum UserWatcherStatus
    UserWatcherStatus_Created = 0
    UserWatcherStatus_Started = 1
    UserWatcherStatus_EnumerationCompleted = 2
    UserWatcherStatus_Stopping = 3
    UserWatcherStatus_Stopped = 4
    UserWatcherStatus_Aborted = 5
End Enum

' ----==== Namespace System.Launcher ====----
Public Property Get Launcher() As Launcher
    Set Launcher = New Launcher
End Property

' ----==== Namespace System.Threading ====----
Public Property Get Threading() As Threading
    Set Threading = New Threading
End Property

' ----==== Namespace System.Diagnostics ====----
Public Property Get Diagnostics() As Diagnostics
    Set Diagnostics = New Diagnostics
End Property

' ----==== Namespace System.LauncherOptions ====----
Public Property Get LauncherOptions() As LauncherOptions
    Set LauncherOptions = New LauncherOptions
End Property

' ----==== Namespace System.LaunchUriResult ====----
Public Property Get LaunchUriResult() As LaunchUriResult
    Set LaunchUriResult = New LaunchUriResult
End Property

' ----==== Namespace System.FolderLauncherOptions ====----
Public Property Get FolderLauncherOptions() As FolderLauncherOptions
    Set FolderLauncherOptions = New FolderLauncherOptions
End Property

' ----==== Namespace System.AppDiagnosticInfo ====----
Public Property Get AppDiagnosticInfo() As AppDiagnosticInfo
    Set AppDiagnosticInfo = New AppDiagnosticInfo
End Property

' ----==== Namespace System.AppActivationResult ====----
Public Property Get AppActivationResult() As AppActivationResult
    Set AppActivationResult = New AppActivationResult
End Property

' ----==== Namespace System.AppResourceGroupInfo ====----
Public Property Get AppResourceGroupInfo() As AppResourceGroupInfo
    Set AppResourceGroupInfo = New AppResourceGroupInfo
End Property

' ----==== Namespace System.AppDiagnosticInfoWatcher ====----
Public Property Get AppDiagnosticInfoWatcher() As AppDiagnosticInfoWatcher
    Set AppDiagnosticInfoWatcher = New AppDiagnosticInfoWatcher
End Property

' ----==== Namespace System.AppResourceGroupInfoWatcher ====----
Public Property Get AppResourceGroupInfoWatcher() As AppResourceGroupInfoWatcher
    Set AppResourceGroupInfoWatcher = New AppResourceGroupInfoWatcher
End Property

' ----==== Namespace System.User ====----
Public Property Get User() As User
    Set User = New User
End Property

' ----==== Namespace System.KnownUserProperties ====----
Public Property Get KnownUserProperties() As KnownUserProperties
    Set KnownUserProperties = New KnownUserProperties
End Property

