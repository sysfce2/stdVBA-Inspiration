VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceInformationUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceInformationUpdate As String = "Windows.Devices.Enumeration.DeviceInformationUpdate"
Private Const IID_IDeviceInformationUpdate As String = "{8f315305-d972-44b7-a37e-9e822c78213b}"
Private Const IID_IDeviceInformationUpdate2 As String = "{5d9d148c-a873-485e-baa6-aa620788e3cc}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDeviceInformationUpdate
    IDeviceInformationUpdate_GetId = 6
    IDeviceInformationUpdate_GetProperties = 7
    
    ' IDeviceInformationUpdate2
    IDeviceInformationUpdate2_GetKind = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceInformationUpdate As Long
Private m_pIDeviceInformationUpdate2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceInformationUpdate2)
    Call ReleaseIfc(m_pIDeviceInformationUpdate)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceInformationUpdate
End Property

Public Property Let Ifc(ByVal pIDeviceInformationUpdate As Long)
    Call ReleaseIfc(m_pIDeviceInformationUpdate2)
    Call ReleaseIfc(m_pIDeviceInformationUpdate)
    m_pIDeviceInformationUpdate = pIDeviceInformationUpdate
    Call QueryIfc(m_pIDeviceInformationUpdate, _
                  IID_IDeviceInformationUpdate2, _
                  m_pIDeviceInformationUpdate2)
End Property

' IDeviceInformationUpdate
Public Property Get Id() As String
    Dim Ret As String
    If m_pIDeviceInformationUpdate <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformationUpdate, _
                  IDeviceInformationUpdate_GetId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Id = Ret
End Property

Public Property Get Properties() As ReadOnlyList_2 'ReadOnlyList_String_Inspectable
    Dim Ret As ReadOnlyList_2 'ReadOnlyList_String_Inspectable
    If m_pIDeviceInformationUpdate <> 0& Then
        Dim pIMapView As Long
        If Invoke(m_pIDeviceInformationUpdate, _
                  IDeviceInformationUpdate_GetProperties, _
                  VarPtr(pIMapView)) = S_OK Then
            If pIMapView <> 0& Then
                Set Ret = New_ReadOnlyList_2(Of_String_Inspectable, pIMapView) 'ReadOnlyList_String_Inspectable
            End If
        End If
    End If
    Set Properties = Ret
End Property
    
' IDeviceInformationUpdate2
Public Property Get Kind() As DeviceInformationKind
    Dim Ret As DeviceInformationKind
    If m_pIDeviceInformationUpdate2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformationUpdate2, _
                  IDeviceInformationUpdate2_GetKind, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Kind = Ret
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


