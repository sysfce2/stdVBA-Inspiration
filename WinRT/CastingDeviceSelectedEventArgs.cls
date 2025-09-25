VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CastingDeviceSelectedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_CastingDeviceSelectedEventArgs As String = "Windows.Devices.Enumeration.CastingDeviceSelectedEventArgs"
Private Const IID_ICastingDeviceSelectedEventArgs As String = "{dc439e86-dd57-4d0d-9400-af45e4fb3663}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICastingDeviceSelectedEventArgs
    ICastingDeviceSelectedEventArgs_GetSelectedCastingDevice = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pICastingDeviceSelectedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICastingDeviceSelectedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICastingDeviceSelectedEventArgs
End Property

Public Property Let Ifc(ByVal pICastingDeviceSelectedEventArgs As Long)
    Call ReleaseIfc(m_pICastingDeviceSelectedEventArgs)
    m_pICastingDeviceSelectedEventArgs = pICastingDeviceSelectedEventArgs
End Property

Public Property Get SelectedCastingDevice() As CastingDevice
    Dim Ret As CastingDevice
    If m_pICastingDeviceSelectedEventArgs <> 0& Then
        Dim pICastingDevice As Long
        If Invoke(m_pICastingDeviceSelectedEventArgs, _
                  ICastingDeviceSelectedEventArgs_GetSelectedCastingDevice, _
                  VarPtr(pICastingDevice)) = S_OK Then
            If pICastingDevice <> 0& Then
                Set Ret = New CastingDevice
                Ret.Ifc = pICastingDevice
            End If
        End If
    End If
    Set SelectedCastingDevice = Ret
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


