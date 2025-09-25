VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceInformationCustomPairing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceInformationCustomPairing As String = "Windows.Devices.Enumeration.DeviceInformationCustomPairing"
Private Const IID_IDeviceInformationCustomPairing As String = "{85138c02-4ee6-4914-8370-107a39144c0e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
                        
    ' IDeviceInformationCustomPairing
    IDeviceInformationCustomPairing_PairAsync = 6
    IDeviceInformationCustomPairing_PairWithProtectionLevelAsync = 7
    IDeviceInformationCustomPairing_PairWithProtectionLevelAndSettingsAsync = 8
    IDeviceInformationCustomPairing_AddPairingRequested = 9
    IDeviceInformationCustomPairing_RemovePairingRequested = 10
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceInformationCustomPairing As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceInformationCustomPairing)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceInformationCustomPairing
End Property

Public Property Let Ifc(ByVal pIDeviceInformationCustomPairing As Long)
    Call ReleaseIfc(m_pIDeviceInformationCustomPairing)
    m_pIDeviceInformationCustomPairing = pIDeviceInformationCustomPairing
End Property

' ----==== Functions ====----
' IDeviceInformationCustomPairing
Public Function PairAsync(ByVal PairingKindsSupported As DevicePairingKinds) As DevicePairingResult
    Dim Ret As DevicePairingResult
    If m_pIDeviceInformationCustomPairing <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationCustomPairing, _
                  IDeviceInformationCustomPairing_PairAsync, _
                  PairingKindsSupported, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DevicePairingResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set PairAsync = Ret
End Function

Public Function PairWithProtectionLevelAsync(ByVal PairingKindsSupported As DevicePairingKinds, _
                                             ByVal MinProtectionLevel As DevicePairingProtectionLevel) As DevicePairingResult
    Dim Ret As DevicePairingResult
    If m_pIDeviceInformationCustomPairing <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationCustomPairing, _
                  IDeviceInformationCustomPairing_PairWithProtectionLevelAsync, _
                  PairingKindsSupported, _
                  MinProtectionLevel, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DevicePairingResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set PairWithProtectionLevelAsync = Ret
End Function

Public Function PairWithProtectionLevelAndSettingsAsync(ByVal PairingKindsSupported As DevicePairingKinds, _
                                                        ByVal MinProtectionLevel As DevicePairingProtectionLevel, _
                                                        ByVal DevPairingSettings As DevicePairingSettings) As DevicePairingResult
    Dim Ret As DevicePairingResult
    If m_pIDeviceInformationCustomPairing <> 0& And _
       IsNotNothing(DevPairingSettings) Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationCustomPairing, _
                  IDeviceInformationCustomPairing_PairWithProtectionLevelAndSettingsAsync, _
                  PairingKindsSupported, _
                  MinProtectionLevel, _
                  DevPairingSettings.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DevicePairingResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set PairWithProtectionLevelAndSettingsAsync = Ret
End Function

Public Function AddPairingRequested(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDeviceInformationCustomPairing <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIDeviceInformationCustomPairing, _
                  IDeviceInformationCustomPairing_AddPairingRequested, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddPairingRequested = Ret
End Function

Public Function RemovePairingRequested(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformationCustomPairing <> 0& And token <> 0@ Then
        If Invoke(m_pIDeviceInformationCustomPairing, _
                  IDeviceInformationCustomPairing_RemovePairingRequested, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemovePairingRequested = Ret
End Function
                        
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


