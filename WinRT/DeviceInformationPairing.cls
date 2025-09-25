VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceInformationPairing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceInformationPairing As String = "Windows.Devices.Enumeration.DeviceInformationPairing"
Private Const IID_IDeviceInformationPairing As String = "{2c4769f5-f684-40d5-8469-e8dbaab70485}"
Private Const IID_IDeviceInformationPairing2 As String = "{f68612fd-0aee-4328-85cc-1c742bb1790d}"
Private Const IID_IDeviceInformationPairingStatics As String = "{e915c408-36d4-49a1-bf13-514173799b6b}"
Private Const IID_IDeviceInformationPairingStatics2 As String = "{04de5372-b7b7-476b-a74f-c5836a704d98}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDeviceInformationPairingStatics
    IDeviceInformationPairingStatics_TryRegisterForAllInboundPairingRequests = 6
                        
    ' IDeviceInformationPairingStatics2
    IDeviceInformationPairingStatics2_TryRegisterForAllInboundPairingRequestsWithProtectionLevel = 6
                        
    ' IDeviceInformationPairing
    IDeviceInformationPairing_GetIsPaired = 6
    IDeviceInformationPairing_GetCanPair = 7
    IDeviceInformationPairing_PairAsync = 8
    IDeviceInformationPairing_PairWithProtectionLevelAsync = 9
                        
    ' IDeviceInformationPairing2
    IDeviceInformationPairing2_GetProtectionLevel = 6
    IDeviceInformationPairing2_GetCustom = 7
    IDeviceInformationPairing2_PairWithProtectionLevelAndSettingsAsync = 8
    IDeviceInformationPairing2_UnpairAsync = 9
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceInformationPairing As Long
Private m_pIDeviceInformationPairing2 As Long
Private m_pIDeviceInformationPairingStatics As Long
Private m_pIDeviceInformationPairingStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Devices_Enumeration_DeviceInformationPairing, _
                            IID_IDeviceInformationPairingStatics, _
                            m_pIDeviceInformationPairingStatics) Then
        Call QueryIfc(m_pIDeviceInformationPairingStatics, _
                      IID_IDeviceInformationPairingStatics2, _
                      m_pIDeviceInformationPairingStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceInformationPairing2)
    Call ReleaseIfc(m_pIDeviceInformationPairing)
    Call ReleaseIfc(m_pIDeviceInformationPairingStatics2)
    Call ReleaseIfc(m_pIDeviceInformationPairingStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceInformationPairing
End Property

Public Property Let Ifc(ByVal pIDeviceInformationPairing As Long)
    Call ReleaseIfc(m_pIDeviceInformationPairing2)
    Call ReleaseIfc(m_pIDeviceInformationPairing)
    m_pIDeviceInformationPairing = pIDeviceInformationPairing
    Call QueryIfc(m_pIDeviceInformationPairing, _
                  IID_IDeviceInformationPairing2, _
                  m_pIDeviceInformationPairing2)
End Property

' IDeviceInformationPairing
Public Property Get IsPaired() As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformationPairing <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformationPairing, _
                  IDeviceInformationPairing_GetIsPaired, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPaired = Ret
End Property

Public Property Get CanPair() As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformationPairing <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformationPairing, _
                  IDeviceInformationPairing_GetCanPair, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CanPair = Ret
End Property
                        
' IDeviceInformationPairing2
Public Property Get ProtectionLevel() As DevicePairingProtectionLevel
    Dim Ret As DevicePairingProtectionLevel
    If m_pIDeviceInformationPairing2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformationPairing2, _
                  IDeviceInformationPairing2_GetProtectionLevel, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ProtectionLevel = Ret
End Property

Public Property Get Custom() As DeviceInformationCustomPairing
    Dim Ret As DeviceInformationCustomPairing
    If m_pIDeviceInformationPairing2 <> 0& Then
        Dim pIDeviceInformationCustomPairing As Long
        If Invoke(m_pIDeviceInformationPairing2, _
                  IDeviceInformationPairing2_GetCustom, _
                  VarPtr(pIDeviceInformationCustomPairing)) = S_OK Then
            If pIDeviceInformationCustomPairing <> 0& Then
                Set Ret = New DeviceInformationCustomPairing
                Ret.Ifc = pIDeviceInformationCustomPairing
            End If
        End If
    End If
    Set Custom = Ret
End Property

' ----==== Functions ====----
' IDeviceInformationPairingStatics
Public Function TryRegisterForAllInboundPairingRequests(ByVal PairingKindsSupported As DevicePairingKinds) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformationPairingStatics <> 0& Then
        Dim result As Long
        If Invoke(m_pIDeviceInformationPairingStatics, _
                  IDeviceInformationPairingStatics_TryRegisterForAllInboundPairingRequests, _
                  PairingKindsSupported, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
    End If
    TryRegisterForAllInboundPairingRequests = Ret
End Function
                        
' IDeviceInformationPairingStatics2
Public Function TryRegisterForAllInboundPairingRequestsWithProtectionLevel(ByVal PairingKindsSupported As DevicePairingKinds, _
                                                                           ByVal MinProtectionLevel As DevicePairingProtectionLevel) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformationPairingStatics2 <> 0& Then
        Dim result As Long
        If Invoke(m_pIDeviceInformationPairingStatics2, _
                  IDeviceInformationPairingStatics2_TryRegisterForAllInboundPairingRequestsWithProtectionLevel, _
                  PairingKindsSupported, _
                  MinProtectionLevel, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
    End If
    TryRegisterForAllInboundPairingRequestsWithProtectionLevel = Ret
End Function
                        
' IDeviceInformationPairing
Public Function PairAsync() As DevicePairingResult
    Dim Ret As DevicePairingResult
    If m_pIDeviceInformationPairing <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationPairing, _
                  IDeviceInformationPairing_PairAsync, _
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

Public Function PairWithProtectionLevelAsync(ByVal MinProtectionLevel As DevicePairingProtectionLevel) As DevicePairingResult
    Dim Ret As DevicePairingResult
    If m_pIDeviceInformationPairing <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationPairing, _
                  IDeviceInformationPairing_PairWithProtectionLevelAsync, _
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
                        
' IDeviceInformationPairing2
Public Function PairWithProtectionLevelAndSettingsAsync(ByVal MinProtectionLevel As DevicePairingProtectionLevel, _
                                                        ByVal DevPairingSettings As DevicePairingSettings) As DevicePairingResult
    Dim Ret As DevicePairingResult
    If m_pIDeviceInformationPairing2 <> 0& And _
       IsNotNothing(DevPairingSettings) Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationPairing2, _
                  IDeviceInformationPairing2_PairWithProtectionLevelAndSettingsAsync, _
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

Public Function UnpairAsync() As DeviceUnpairingResult
    Dim Ret As DeviceUnpairingResult
    If m_pIDeviceInformationPairing2 <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationPairing2, _
                  IDeviceInformationPairing2_UnpairAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceUnpairingResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set UnpairAsync = Ret
End Function
                        
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
