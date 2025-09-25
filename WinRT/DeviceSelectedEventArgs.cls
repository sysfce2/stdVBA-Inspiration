VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceSelectedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceSelectedEventArgs As String = "Windows.Devices.Enumeration.DeviceSelectedEventArgs"
Private Const IID_IDeviceSelectedEventArgs As String = "{269edade-1d2f-4940-8402-4156b81d3c77}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDeviceSelectedEventArgs
    IDeviceSelectedEventArgs_GetSelectedDevice = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceSelectedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceSelectedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceSelectedEventArgs
End Property

Public Property Let Ifc(ByVal pIDeviceSelectedEventArgs As Long)
    Call ReleaseIfc(m_pIDeviceSelectedEventArgs)
    m_pIDeviceSelectedEventArgs = pIDeviceSelectedEventArgs
End Property

Public Property Get SelectedDevice() As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIDeviceSelectedEventArgs <> 0& Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pIDeviceSelectedEventArgs, _
                  IDeviceSelectedEventArgs_GetSelectedDevice, _
                  VarPtr(pIDeviceInformation)) = S_OK Then
            If pIDeviceInformation <> 0& Then
                Set Ret = New DeviceInformation
                Ret.Ifc = pIDeviceInformation
            End If
        End If
    End If
    Set SelectedDevice = Ret
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


