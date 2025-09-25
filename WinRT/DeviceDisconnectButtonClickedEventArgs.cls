VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceDisconnectButtonClickedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceDisconnectButtonClickedEventArgs As String = "Windows.Devices.Enumeration.DeviceDisconnectButtonClickedEventArgs"
Private Const IID_IDeviceDisconnectButtonClickedEventArgs As String = "{8e44b56d-f902-4a00-b536-f37992e6a2a7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDeviceDisconnectButtonClickedEventArgs
    IDeviceDisconnectButtonClickedEventArgs_GetDevice = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceDisconnectButtonClickedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceDisconnectButtonClickedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceDisconnectButtonClickedEventArgs
End Property

Public Property Let Ifc(ByVal pIDeviceDisconnectButtonClickedEventArgs As Long)
    Call ReleaseIfc(m_pIDeviceDisconnectButtonClickedEventArgs)
    m_pIDeviceDisconnectButtonClickedEventArgs = pIDeviceDisconnectButtonClickedEventArgs
End Property

Public Property Get Device() As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIDeviceDisconnectButtonClickedEventArgs <> 0& Then
        Dim pIDeviceInformation As Long
        If Invoke(m_pIDeviceDisconnectButtonClickedEventArgs, _
                  IDeviceDisconnectButtonClickedEventArgs_GetDevice, _
                  VarPtr(pIDeviceInformation)) = S_OK Then
            If pIDeviceInformation <> 0& Then
                Set Ret = New DeviceInformation
                Ret.Ifc = pIDeviceInformation
            End If
        End If
    End If
    Set Device = Ret
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


