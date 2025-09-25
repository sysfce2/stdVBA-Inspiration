VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CastingDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Casting_CastingDevice As String = "Windows.Media.Casting.CastingDevice"
Private Const IID_ICastingDevice As String = "{de721c83-4a43-4ad1-a6d2-2492a796c3f2}"
Private Const IID_ICastingDeviceStatics As String = "{e7d958d7-4d13-4237-a365-4c4f6a4cfd2f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICastingDeviceStatics
    ICastingDeviceStatics_GetDeviceSelector = 6
    ICastingDeviceStatics_GetDeviceSelectorFromCastingSourceAsync = 7
    ICastingDeviceStatics_FromIdAsync = 8
    ICastingDeviceStatics_DeviceInfoSupportsCastingAsync = 9
                        
    ' ICastingDevice
    ICastingDevice_GetId = 6
    ICastingDevice_GetFriendlyName = 7
    ICastingDevice_GetIcon = 8
    ICastingDevice_GetSupportedCastingPlaybackTypesAsync = 9
    ICastingDevice_CreateCastingConnection = 10
                        
End Enum

' ----==== Variablen ====----
Private m_pICastingDevice As Long
Private m_pICastingDeviceStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Casting_CastingDevice, _
                            IID_ICastingDeviceStatics, _
                            m_pICastingDeviceStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICastingDevice)
    Call ReleaseIfc(m_pICastingDeviceStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICastingDevice
End Property

Public Property Let Ifc(ByVal pICastingDevice As Long)
    Call ReleaseIfc(m_pICastingDevice)
    m_pICastingDevice = pICastingDevice
End Property

' ICastingDevice
Public Property Get Id() As String
    Dim Ret As String
    If m_pICastingDevice <> 0& Then
        Dim value As Long
        If Invoke(m_pICastingDevice, _
                  ICastingDevice_GetId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Id = Ret
End Property

Public Property Get FriendlyName() As String
    Dim Ret As String
    If m_pICastingDevice <> 0& Then
        Dim value As Long
        If Invoke(m_pICastingDevice, _
                  ICastingDevice_GetFriendlyName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    FriendlyName = Ret
End Property

Public Property Get Icon() As RandomAccessStreamWithContentType
    Dim Ret As RandomAccessStreamWithContentType
    If m_pICastingDevice <> 0& Then
        Dim pIRandomAccessStreamWithContentType As Long
        If Invoke(m_pICastingDevice, _
                  ICastingDevice_GetIcon, _
                  VarPtr(pIRandomAccessStreamWithContentType)) = S_OK Then
            If pIRandomAccessStreamWithContentType <> 0& Then
                Set Ret = New RandomAccessStreamWithContentType
                Ret.Ifc = pIRandomAccessStreamWithContentType
            End If
        End If
    End If
    Set Icon = Ret
End Property

' ----==== Functions ====----
' ICastingDeviceStatics
Public Function GetDeviceSelector(ByVal typ As CastingPlaybackTypes) As String
    Dim Ret As String
    If m_pICastingDeviceStatics <> 0& Then
        Dim value As Long
        If Invoke(m_pICastingDeviceStatics, _
                  ICastingDeviceStatics_GetDeviceSelector, _
                  typ, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GetDeviceSelector = Ret
End Function

Public Function GetDeviceSelectorFromCastingSourceAsync(ByVal Source As CastingSource) As String
    Dim Ret As String
    If m_pICastingDeviceStatics <> 0& And _
       IsNotNothing(Source) Then
        Dim value As Long
        If Invoke(m_pICastingDeviceStatics, _
                  ICastingDeviceStatics_GetDeviceSelectorFromCastingSourceAsync, _
                  Source.Ifc, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Ret = GetWindowsString(value)
                End If
            End If
        End If
    End If
    GetDeviceSelectorFromCastingSourceAsync = Ret
End Function

Public Function FromIdAsync(ByVal value As String) As CastingDevice
    Dim Ret As CastingDevice
    If m_pICastingDeviceStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        Dim pICastingDevice As Long
        If Invoke(m_pICastingDeviceStatics, _
                  ICastingDeviceStatics_FromIdAsync, _
                  hString, _
                  VarPtr(pICastingDevice)) = S_OK Then
            If pICastingDevice <> 0& Then
                If Await(pICastingDevice) Then
                    Set Ret = New CastingDevice
                    Ret.Ifc = pICastingDevice
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set FromIdAsync = Ret
End Function

Public Function DeviceInfoSupportsCastingAsync(ByVal Device As DeviceInformation) As Boolean
    Dim Ret As Boolean
    If m_pICastingDeviceStatics <> 0& And _
       IsNotNothing(Device) Then
        Dim value As Long
        If Invoke(m_pICastingDeviceStatics, _
                  ICastingDeviceStatics_DeviceInfoSupportsCastingAsync, _
                  Device.Ifc, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Ret = CBool(value)
                End If
            End If
        End If
    End If
    DeviceInfoSupportsCastingAsync = Ret
End Function

' ICastingDevice
Public Function GetSupportedCastingPlaybackTypesAsync() As CastingPlaybackTypes
    Dim Ret As CastingPlaybackTypes
    If m_pICastingDevice <> 0& Then
        Dim value As Long
        If Invoke(m_pICastingDevice, _
                  ICastingDevice_GetSupportedCastingPlaybackTypesAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Ret = value
                End If
            End If
        End If
    End If
    GetSupportedCastingPlaybackTypesAsync = Ret
End Function

Public Function CreateCastingConnection() As CastingConnection
    Dim Ret As CastingConnection
    If m_pICastingDevice <> 0& Then
        Dim pICastingConnection As Long
        If Invoke(m_pICastingDevice, _
                  ICastingDevice_CreateCastingConnection, _
                  VarPtr(pICastingConnection)) = S_OK Then
            If pICastingConnection <> 0& Then
                Set Ret = New CastingConnection
                Ret.Ifc = pICastingConnection
            End If
        End If
    End If
    Set CreateCastingConnection = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


