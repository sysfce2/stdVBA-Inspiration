VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceUnpairingResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceUnpairingResult As String = "Windows.Devices.Enumeration.DeviceUnpairingResult"
Private Const IID_IDeviceUnpairingResult As String = "{66f44ad3-79d9-444b-92cf-a92ef72571c7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDeviceUnpairingResult
    IDeviceUnpairingResult_GetStatus = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceUnpairingResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceUnpairingResult)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceUnpairingResult
End Property

Public Property Let Ifc(ByVal pIDeviceUnpairingResult As Long)
    Call ReleaseIfc(m_pIDeviceUnpairingResult)
    m_pIDeviceUnpairingResult = pIDeviceUnpairingResult
End Property

' IDeviceUnpairingResult
Public Property Get Status() As DeviceUnpairingResultStatus
    Dim Ret As DeviceUnpairingResultStatus
    If m_pIDeviceUnpairingResult <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceUnpairingResult, _
                  IDeviceUnpairingResult_GetStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Status = Ret
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


