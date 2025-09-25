VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_DeviceInformation As String = "Windows.Devices.Enumeration.DeviceInformation"
Private Const IID_IDeviceInformation As String = "{aba0fb95-4398-489d-8e44-e6130927011f}"
Private Const IID_IDeviceInformation2 As String = "{f156a638-7997-48d9-a10c-269d46533f48}"
Private Const IID_IDeviceInformationStatics As String = "{c17f100e-3a46-4a78-8013-769dc9b97390}"
Private Const IID_IDeviceInformationStatics2 As String = "{493b4f34-a84f-45fd-9167-15d1cb1bd1f9}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDeviceInformationStatics
    IDeviceInformationStatics_CreateFromIdAsync = 6
    IDeviceInformationStatics_CreateFromIdAsyncAdditionalProperties = 7
    IDeviceInformationStatics_FindAllAsync = 8
    IDeviceInformationStatics_FindAllAsyncDeviceClass = 9
    IDeviceInformationStatics_FindAllAsyncAqsFilter = 10
    IDeviceInformationStatics_FindAllAsyncAqsFilterAndAdditionalProperties = 11
    IDeviceInformationStatics_CreateWatcher = 12
    IDeviceInformationStatics_CreateWatcherDeviceClass = 13
    IDeviceInformationStatics_CreateWatcherAqsFilter = 14
    IDeviceInformationStatics_CreateWatcherAqsFilterAndAdditionalProperties = 15
                        
    ' IDeviceInformationStatics2
    IDeviceInformationStatics2_GetAqsFilterFromDeviceClass = 6
    IDeviceInformationStatics2_CreateFromIdAsyncWithKindAndAdditionalProperties = 7
    IDeviceInformationStatics2_FindAllAsyncWithKindAqsFilterAndAdditionalProperties = 8
    IDeviceInformationStatics2_CreateWatcherWithKindAqsFilterAndAdditionalProperties = 9
                        
    ' IDeviceInformation
    IDeviceInformation_GetId = 6
    IDeviceInformation_GetName = 7
    IDeviceInformation_GetIsEnabled = 8
    IDeviceInformation_GetIsDefault = 9
    IDeviceInformation_GetEnclosureLocation = 10
    IDeviceInformation_GetProperties = 11
    IDeviceInformation_Update = 12
    IDeviceInformation_GetThumbnailAsync = 13
    IDeviceInformation_GetGlyphThumbnailAsync = 14
                        
    ' IDeviceInformation2
    IDeviceInformation2_GetKind = 6
    IDeviceInformation2_GetPairing = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIDeviceInformation As Long
Private m_pIDeviceInformation2 As Long
Private m_pIDeviceInformationStatics As Long
Private m_pIDeviceInformationStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Devices_Enumeration_DeviceInformation, _
                            IID_IDeviceInformationStatics, _
                            m_pIDeviceInformationStatics) Then
        Call QueryIfc(m_pIDeviceInformationStatics, _
                      IID_IDeviceInformationStatics2, _
                      m_pIDeviceInformationStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDeviceInformation2)
    Call ReleaseIfc(m_pIDeviceInformation)
    Call ReleaseIfc(m_pIDeviceInformationStatics2)
    Call ReleaseIfc(m_pIDeviceInformationStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDeviceInformation
End Property

Public Property Let Ifc(ByVal pIDeviceInformation As Long)
    Call ReleaseIfc(m_pIDeviceInformation2)
    Call ReleaseIfc(m_pIDeviceInformation)
    m_pIDeviceInformation = pIDeviceInformation
    Call QueryIfc(m_pIDeviceInformation, _
                  IID_IDeviceInformation2, _
                  m_pIDeviceInformation2)
End Property

' IDeviceInformation
Public Property Get Id() As String
    Dim Ret As String
    If m_pIDeviceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Id = Ret
End Property

Public Property Get Name() As String
    Dim Ret As String
    If m_pIDeviceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Name = Ret
End Property

Public Property Get IsEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetIsEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsEnabled = Ret
End Property

Public Property Get IsDefault() As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetIsDefault, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsDefault = Ret
End Property

Public Property Get EnclosureLocation() As EnclosureLocation
    Dim Ret As EnclosureLocation
    If m_pIDeviceInformation <> 0& Then
        Dim pIEnclosureLocation As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetEnclosureLocation, _
                  VarPtr(pIEnclosureLocation)) = S_OK Then
            If pIEnclosureLocation <> 0& Then
                Set Ret = New EnclosureLocation
                Ret.Ifc = pIEnclosureLocation
            End If
        End If
    End If
    Set EnclosureLocation = Ret
End Property

Public Property Get Properties() As ReadOnlyList_2 'ReadOnlyList_String_Inspectable
    Dim Ret As ReadOnlyList_2 'ReadOnlyList_String_Inspectable
    If m_pIDeviceInformation <> 0& Then
        Dim pIMapView As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetProperties, _
                  VarPtr(pIMapView)) = S_OK Then
            If pIMapView <> 0& Then
                Set Ret = New_ReadOnlyList_2(Of_String_Inspectable, pIMapView) 'ReadOnlyList_String_Inspectable
            End If
        End If
    End If
    Set Properties = Ret
End Property
                        
' IDeviceInformation2
Public Property Get Kind() As DeviceInformationKind
    Dim Ret As DeviceInformationKind
    If m_pIDeviceInformation2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIDeviceInformation2, _
                  IDeviceInformation2_GetKind, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Kind = Ret
End Property

Public Property Get Pairing() As DeviceInformationPairing
    Dim Ret As DeviceInformationPairing
    If m_pIDeviceInformation2 <> 0& Then
        Dim pIDeviceInformationPairing As Long
        If Invoke(m_pIDeviceInformation2, _
                  IDeviceInformation2_GetPairing, _
                  VarPtr(pIDeviceInformationPairing)) = S_OK Then
            If pIDeviceInformationPairing <> 0& Then
                Set Ret = New DeviceInformationPairing
                Ret.Ifc = pIDeviceInformationPairing
            End If
        End If
    End If
    Set Pairing = Ret
End Property

' ----==== Functions ====----
' IDeviceInformationStatics
Public Function CreateFromIdAsync(ByVal DeviceId As String) As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIDeviceInformationStatics <> 0& Then
        Dim operation As Long
        Dim hDeviceId As Long
        hDeviceId = CreateWindowsString(DeviceId)
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_CreateFromIdAsync, _
                  hDeviceId, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformation
                    Ret.Ifc = operation
                End If
            End If
        End If
        Call DeleteWindowsString(hDeviceId)
    End If
    Set CreateFromIdAsync = Ret
End Function

Public Function CreateFromIdAsyncAdditionalProperties(ByVal DeviceId As String, _
                                                      ByVal AdditionalProperties As List_String) As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIDeviceInformationStatics <> 0& And _
       IsNotNothing(AdditionalProperties) Then
        Dim operation As Long
        Dim hDeviceId As Long
        hDeviceId = CreateWindowsString(DeviceId)
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_CreateFromIdAsyncAdditionalProperties, _
                  hDeviceId, _
                  AdditionalProperties.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformation
                    Ret.Ifc = operation
                End If
            End If
        End If
        Call DeleteWindowsString(hDeviceId)
    End If
    Set CreateFromIdAsyncAdditionalProperties = Ret
End Function

Public Function FindAllAsync() As DeviceInformationCollection
    Dim Ret As DeviceInformationCollection
    If m_pIDeviceInformationStatics <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_FindAllAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformationCollection
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set FindAllAsync = Ret
End Function

Public Function FindAllAsyncDeviceClass(ByVal DevClass As DeviceClass) As DeviceInformationCollection
    Dim Ret As DeviceInformationCollection
    If m_pIDeviceInformationStatics <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_FindAllAsyncDeviceClass, _
                  DevClass, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformationCollection
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set FindAllAsyncDeviceClass = Ret
End Function

Public Function FindAllAsyncAqsFilter(ByVal AqsFilter As String) As DeviceInformationCollection
    Dim Ret As DeviceInformationCollection
    If m_pIDeviceInformationStatics <> 0& Then
        Dim operation As Long
        Dim hAqsFilter As Long
        hAqsFilter = CreateWindowsString(AqsFilter)
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_FindAllAsyncAqsFilter, _
                  hAqsFilter, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformationCollection
                    Ret.Ifc = operation
                End If
            End If
        End If
        Call DeleteWindowsString(hAqsFilter)
    End If
    Set FindAllAsyncAqsFilter = Ret
End Function

Public Function FindAllAsyncAqsFilterAndAdditionalProperties(ByVal AqsFilter As String, _
                                                             ByVal AdditionalProperties As List_String) As DeviceInformationCollection
    Dim Ret As DeviceInformationCollection
    If m_pIDeviceInformationStatics <> 0& And _
       IsNotNothing(AdditionalProperties) Then
        Dim operation As Long
        Dim hAqsFilter As Long
        hAqsFilter = CreateWindowsString(AqsFilter)
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_FindAllAsyncAqsFilterAndAdditionalProperties, _
                  hAqsFilter, _
                  AdditionalProperties.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformationCollection
                    Ret.Ifc = operation
                End If
            End If
        End If
        Call DeleteWindowsString(hAqsFilter)
    End If
    Set FindAllAsyncAqsFilterAndAdditionalProperties = Ret
End Function

Public Function CreateWatcher() As DeviceWatcher
    Dim Ret As DeviceWatcher
    If m_pIDeviceInformationStatics <> 0& Then
        Dim pIDeviceWatcher As Long
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_CreateWatcher, _
                  VarPtr(pIDeviceWatcher)) = S_OK Then
            If pIDeviceWatcher <> 0& Then
                Set Ret = New DeviceWatcher
                Ret.Ifc = pIDeviceWatcher
            End If
        End If
    End If
    Set CreateWatcher = Ret
End Function

Public Function CreateWatcherDeviceClass(ByVal DevClass As DeviceClass) As DeviceWatcher
    Dim Ret As DeviceWatcher
    If m_pIDeviceInformationStatics <> 0& Then
        Dim pIDeviceWatcher As Long
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_CreateWatcherDeviceClass, _
                  DevClass, _
                  VarPtr(pIDeviceWatcher)) = S_OK Then
            If pIDeviceWatcher <> 0& Then
                Set Ret = New DeviceWatcher
                Ret.Ifc = pIDeviceWatcher
            End If
        End If
    End If
    Set CreateWatcherDeviceClass = Ret
End Function

Public Function CreateWatcherAqsFilter(ByVal AqsFilter As String) As DeviceWatcher
    Dim Ret As DeviceWatcher
    If m_pIDeviceInformationStatics <> 0& Then
        Dim pIDeviceWatcher As Long
        Dim hAqsFilter As Long
        hAqsFilter = CreateWindowsString(AqsFilter)
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_CreateWatcherAqsFilter, _
                  hAqsFilter, _
                  VarPtr(pIDeviceWatcher)) = S_OK Then
            If pIDeviceWatcher <> 0& Then
                Set Ret = New DeviceWatcher
                Ret.Ifc = pIDeviceWatcher
            End If
        End If
        Call DeleteWindowsString(hAqsFilter)
    End If
    Set CreateWatcherAqsFilter = Ret
End Function

Public Function CreateWatcherAqsFilterAndAdditionalProperties(ByVal AqsFilter As String, _
                                                              ByVal AdditionalProperties As List_String) As DeviceWatcher
    Dim Ret As DeviceWatcher
    If m_pIDeviceInformationStatics <> 0& And _
       IsNotNothing(AdditionalProperties) Then
        Dim pIDeviceWatcher As Long
        Dim hAqsFilter As Long
        hAqsFilter = CreateWindowsString(AqsFilter)
        If Invoke(m_pIDeviceInformationStatics, _
                  IDeviceInformationStatics_CreateWatcherAqsFilterAndAdditionalProperties, _
                  hAqsFilter, _
                  AdditionalProperties.Ifc, _
                  VarPtr(pIDeviceWatcher)) = S_OK Then
            If pIDeviceWatcher <> 0& Then
                Set Ret = New DeviceWatcher
                Ret.Ifc = pIDeviceWatcher
            End If
        End If
        Call DeleteWindowsString(hAqsFilter)
    End If
    Set CreateWatcherAqsFilterAndAdditionalProperties = Ret
End Function
    
' IDeviceInformationStatics2
Public Function GetAqsFilterFromDeviceClass(ByVal DevClass As DeviceClass) As String
    Dim Ret As String
    If m_pIDeviceInformationStatics2 <> 0& Then
        Dim hAqsFilter As Long
        If Invoke(m_pIDeviceInformationStatics2, _
                  IDeviceInformationStatics2_GetAqsFilterFromDeviceClass, _
                  DevClass, _
                  VarPtr(hAqsFilter)) = S_OK Then
            Ret = GetWindowsString(hAqsFilter)
        End If
    End If
    GetAqsFilterFromDeviceClass = Ret
End Function

Public Function CreateFromIdAsyncWithKindAndAdditionalProperties(ByVal DeviceId As String, _
                                                                 ByVal AdditionalProperties As List_String, _
                                                                 ByVal Kind As DeviceInformationKind) As DeviceInformation
    Dim Ret As DeviceInformation
    If m_pIDeviceInformationStatics2 <> 0& And _
       IsNotNothing(AdditionalProperties) Then
        Dim operation As Long
        Dim hDeviceId As Long
        hDeviceId = CreateWindowsString(DeviceId)
        If Invoke(m_pIDeviceInformationStatics2, _
                  IDeviceInformationStatics2_CreateFromIdAsyncWithKindAndAdditionalProperties, _
                  hDeviceId, _
                  AdditionalProperties.Ifc, _
                  Kind, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformation
                    Ret.Ifc = operation
                End If
            End If
        End If
        Call DeleteWindowsString(hDeviceId)
    End If
    Set CreateFromIdAsyncWithKindAndAdditionalProperties = Ret
End Function

Public Function FindAllAsyncWithKindAqsFilterAndAdditionalProperties(ByVal AqsFilter As String, _
                                                                     ByVal AdditionalProperties As List_String, _
                                                                     ByVal Kind As DeviceInformationKind) As DeviceInformationCollection
    Dim Ret As DeviceInformationCollection
    If m_pIDeviceInformationStatics2 <> 0& And _
       IsNotNothing(AdditionalProperties) Then
        Dim operation As Long
        Dim hAqsFilter As Long
        hAqsFilter = CreateWindowsString(AqsFilter)
        If Invoke(m_pIDeviceInformationStatics2, _
                  IDeviceInformationStatics2_FindAllAsyncWithKindAqsFilterAndAdditionalProperties, _
                  hAqsFilter, _
                  AdditionalProperties.Ifc, _
                  Kind, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceInformationCollection
                    Ret.Ifc = operation
                End If
            End If
        End If
        Call DeleteWindowsString(hAqsFilter)
    End If
    Set FindAllAsyncWithKindAqsFilterAndAdditionalProperties = Ret
End Function

Public Function CreateWatcherWithKindAqsFilterAndAdditionalProperties(ByVal AqsFilter As String, _
                                                                      ByVal AdditionalProperties As List_String, _
                                                                      ByVal Kind As DeviceInformationKind) As DeviceWatcher
    Dim Ret As DeviceWatcher
    If m_pIDeviceInformationStatics2 <> 0& And _
       IsNotNothing(AdditionalProperties) Then
        Dim pIDeviceWatcher As Long
        Dim hAqsFilter As Long
        hAqsFilter = CreateWindowsString(AqsFilter)
        If Invoke(m_pIDeviceInformationStatics2, _
                  IDeviceInformationStatics2_CreateWatcherWithKindAqsFilterAndAdditionalProperties, _
                  hAqsFilter, _
                  AdditionalProperties.Ifc, _
                  Kind, _
                  VarPtr(pIDeviceWatcher)) = S_OK Then
            If pIDeviceWatcher <> 0& Then
                Set Ret = New DeviceWatcher
                Ret.Ifc = pIDeviceWatcher
            End If
        End If
        Call DeleteWindowsString(hAqsFilter)
    End If
    Set CreateWatcherWithKindAqsFilterAndAdditionalProperties = Ret
End Function
    
' IDeviceInformation
Public Function Update(ByVal UpdateInfo As DeviceInformationUpdate) As Boolean
    Dim Ret As Boolean
    If m_pIDeviceInformation <> 0& And _
       IsNotNothing(UpdateInfo) Then
        Dim value As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_Update, _
                  UpdateInfo.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    Update = Ret
End Function

Public Function GetThumbnailAsync() As DeviceThumbnail
    Dim Ret As DeviceThumbnail
    If m_pIDeviceInformation <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetThumbnailAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceThumbnail
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set GetThumbnailAsync = Ret
End Function

Public Function GetGlyphThumbnailAsync() As DeviceThumbnail
    Dim Ret As DeviceThumbnail
    If m_pIDeviceInformation <> 0& Then
        Dim operation As Long
        If Invoke(m_pIDeviceInformation, _
                  IDeviceInformation_GetGlyphThumbnailAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New DeviceThumbnail
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set GetGlyphThumbnailAsync = Ret
End Function
                        
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


