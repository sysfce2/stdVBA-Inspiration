VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Enumeration

Option Explicit

' ----==== Enums Windows.Devices.Enumeration.xxx ====----
Public Enum Panel
    Panel_Unknown = 0
    Panel_Front = 1
    Panel_Back = 2
    Panel_Top = 3
    Panel_Bottom = 4
    Panel_Left = 5
    Panel_Right = 6
End Enum
                
Public Enum DeviceClass
    DeviceClass_All = 0
    DeviceClass_AudioCapture = 1
    DeviceClass_AudioRender = 2
    DeviceClass_PortableStorageDevice = 3
    DeviceClass_VideoCapture = 4
    DeviceClass_ImageScanner = 5
    DeviceClass_Location = 6
End Enum

Public Enum DeviceInformationKind
    DeviceInformationKind_Unknown = 0
    DeviceInformationKind_DeviceInterface = 1
    DeviceInformationKind_DeviceContainer = 2
    DeviceInformationKind_Device = 3
    DeviceInformationKind_DeviceInterfaceClass = 4
    DeviceInformationKind_AssociationEndpoint = 5
    DeviceInformationKind_AssociationEndpointContainer = 6
    DeviceInformationKind_AssociationEndpointService = 7
    DeviceInformationKind_DevicePanel = 8
End Enum

Public Enum DevicePickerDisplayStatusOptions
    DevicePickerDisplayStatusOptions_None = &H0&
    DevicePickerDisplayStatusOptions_ShowProgress = &H1&
    DevicePickerDisplayStatusOptions_ShowDisconnectButton = &H2&
    DevicePickerDisplayStatusOptions_ShowRetryButton = &H4&
End Enum

Public Enum DevicePairingProtectionLevel
    DevicePairingProtectionLevel_Default = 0
    DevicePairingProtectionLevel_None = 1
    DevicePairingProtectionLevel_Encryption = 2
    DevicePairingProtectionLevel_EncryptionAndAuthentication = 3
End Enum

Public Enum DevicePairingKinds
    DevicePairingKinds_None = &H0&
    DevicePairingKinds_ConfirmOnly = &H1&
    DevicePairingKinds_DisplayPin = &H2&
    DevicePairingKinds_ProvidePin = &H4&
    DevicePairingKinds_ConfirmPinMatch = &H8&
    DevicePairingKinds_ProvidePasswordCredential = &H10&
End Enum

Public Enum DevicePairingResultStatus
    DevicePairingResultStatus_Paired = 0
    DevicePairingResultStatus_NotReadyToPair = 1
    DevicePairingResultStatus_NotPaired = 2
    DevicePairingResultStatus_AlreadyPaired = 3
    DevicePairingResultStatus_ConnectionRejected = 4
    DevicePairingResultStatus_TooManyConnections = 5
    DevicePairingResultStatus_HardwareFailure = 6
    DevicePairingResultStatus_AuthenticationTimeout = 7
    DevicePairingResultStatus_AuthenticationNotAllowed = 8
    DevicePairingResultStatus_AuthenticationFailure = 9
    DevicePairingResultStatus_NoSupportedProfiles = 10
    DevicePairingResultStatus_ProtectionLevelCouldNotBeMet = 11
    DevicePairingResultStatus_AccessDenied = 12
    DevicePairingResultStatus_InvalidCeremonyData = 13
    DevicePairingResultStatus_PairingCanceled = 14
    DevicePairingResultStatus_OperationAlreadyInProgress = 15
    DevicePairingResultStatus_RequiredHandlerNotRegistered = 16
    DevicePairingResultStatus_RejectedByHandler = 17
    DevicePairingResultStatus_RemoteDeviceHasAssociation = 18
    DevicePairingResultStatus_Failed = 19
End Enum

Public Enum DeviceUnpairingResultStatus
    DeviceUnpairingResultStatus_Unpaired = 0
    DeviceUnpairingResultStatus_AlreadyUnpaired = 1
    DeviceUnpairingResultStatus_OperationAlreadyInProgress = 2
    DeviceUnpairingResultStatus_AccessDenied = 3
    DeviceUnpairingResultStatus_Failed = 4
End Enum

Public Enum DeviceWatcherStatus
    DeviceWatcherStatus_Created = 0
    DeviceWatcherStatus_Started = 1
    DeviceWatcherStatus_EnumerationCompleted = 2
    DeviceWatcherStatus_Stopping = 3
    DeviceWatcherStatus_Stopped = 4
    DeviceWatcherStatus_Aborted = 5
End Enum

Public Enum DeviceWatcherEventKind
    DeviceWatcherEventKind_Add = 0
    DeviceWatcherEventKind_Update = 1
    DeviceWatcherEventKind_Remove = 2
End Enum

' ----==== Namespace Enumeration.DevicePicker ====----
Public Property Get DevicePicker() As DevicePicker
    Set DevicePicker = New DevicePicker
End Property

' ----==== Namespace Enumeration.EnclosureLocation ====----
Public Property Get EnclosureLocation() As EnclosureLocation
    Set EnclosureLocation = New EnclosureLocation
End Property

' ----==== Namespace Enumeration.DevicePickerFilter ====----
Public Property Get DevicePickerFilter() As DevicePickerFilter
    Set DevicePickerFilter = New DevicePickerFilter
End Property

' ----==== Namespace Enumeration.DeviceInformation ====----
Public Property Get DeviceInformation() As DeviceInformation
    Set DeviceInformation = New DeviceInformation
End Property

' ----==== Namespace Enumeration.DevicePickerAppearance ====----
Public Property Get DevicePickerAppearance() As DevicePickerAppearance
    Set DevicePickerAppearance = New DevicePickerAppearance
End Property

' ----==== Namespace Enumeration.DeviceInformationPairing ====----
Public Property Get DeviceInformationPairing() As DeviceInformationPairing
    Set DeviceInformationPairing = New DeviceInformationPairing
End Property

