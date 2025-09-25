VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevicePickerFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DevicePickerFilter As String = "Windows.Devices.Enumeration.DevicePickerFilter"
Private Const IID_IDevicePickerFilter As String = "{91db92a2-57cb-48f1-9b59-a59b7a1f02a2}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDevicePickerFilter
    IDevicePickerFilter_GetSupportedDeviceClasses = 6
    IDevicePickerFilter_GetSupportedDeviceSelectors = 7

End Enum

' ----==== Variablen ====----
Private m_pIDevicePickerFilter As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDevicePickerFilter)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDevicePickerFilter
End Property

Public Property Let Ifc(ByVal pIDevicePickerFilterAppearance As Long)
    Call ReleaseIfc(m_pIDevicePickerFilter)
    m_pIDevicePickerFilter = pIDevicePickerFilterAppearance
End Property

' IDevicePickerFilter
Public Property Get SupportedDeviceClasses() As List_DeviceClass
    Dim Ret As List_DeviceClass
    If m_pIDevicePickerFilter <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIDevicePickerFilter, _
                  IDevicePickerFilter_GetSupportedDeviceClasses, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_DeviceClass
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set SupportedDeviceClasses = Ret
End Property

Public Property Get SupportedDeviceSelectors() As List_String
    Dim Ret As List_String
    If m_pIDevicePickerFilter <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIDevicePickerFilter, _
                  IDevicePickerFilter_GetSupportedDeviceSelectors, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set SupportedDeviceSelectors = Ret
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


