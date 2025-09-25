VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevicePairingSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DevicePairingSettings As String = "Windows.Devices.Enumeration.DevicePairingSettings"
Private Const IID_IDevicePairingSettings As String = "{482cb27c-83bb-420e-be51-6602b222de54}"

' ----==== Enums ====----
'Private Enum vtb_Interfaces
'
'    ' IDevicePairingSettings
'    ' has no member
'
'End Enum

' ----==== Variablen ====----
Private m_pIDevicePairingSettings As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDevicePairingSettings)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDevicePairingSettings
End Property

Public Property Let Ifc(ByVal pIDevicePairingSettings As Long)
    Call ReleaseIfc(m_pIDevicePairingSettings)
    m_pIDevicePairingSettings = pIDevicePairingSettings
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
'Private Function Invoke(ByVal pInterface As Long, _
'                        ByVal vtb As vtb_Interfaces, _
'                        ParamArray var()) As Variant
'    If pInterface <> 0& Then
'        Invoke = OleInvoke(pInterface, vtb, var)
'    End If
'End Function




