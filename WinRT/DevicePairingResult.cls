VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevicePairingResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DevicePairingResult As String = "Windows.Devices.Enumeration.DevicePairingResult"
Private Const IID_IDevicePairingResult As String = "{072b02bf-dd95-4025-9b37-de51adba37b7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDevicePairingResult
    IDevicePairingResult_GetStatus = 6
    IDevicePairingResult_GetProtectionLevelUsed = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIDevicePairingResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDevicePairingResult)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDevicePairingResult
End Property

Public Property Let Ifc(ByVal pIDevicePairingResult As Long)
    Call ReleaseIfc(m_pIDevicePairingResult)
    m_pIDevicePairingResult = pIDevicePairingResult
End Property

' IDevicePairingResult
Public Property Get Status() As DevicePairingResultStatus
    Dim Ret As DevicePairingResultStatus
    If m_pIDevicePairingResult <> 0& Then
        Dim value As Long
        If Invoke(m_pIDevicePairingResult, _
                  IDevicePairingResult_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
End Property

Public Property Get ProtectionLevelUsed() As DevicePairingProtectionLevel
    Dim Ret As DevicePairingProtectionLevel
    If m_pIDevicePairingResult <> 0& Then
        Dim value As Long
        If Invoke(m_pIDevicePairingResult, _
                  IDevicePairingResult_GetProtectionLevelUsed, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ProtectionLevelUsed = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


