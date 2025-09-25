VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_FileProperties_ImageProperties As String = "Windows.Storage.FileProperties.ImageProperties"
Private Const IID_IImageProperties As String = "{523c9424-fcff-4275-afee-ecdb9ab47973}"
Private Const IID_IStorageItemExtraProperties As String = "{c54361b2-54cd-432b-bdbc-4b19c4b470d7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IImageProperties
    IImageProperties_GetRating = 6
    IImageProperties_PutRating = 7
    IImageProperties_GetKeywords = 8
    IImageProperties_GetDateTaken = 9
    IImageProperties_PutDateTaken = 10
    IImageProperties_GetWidth = 11
    IImageProperties_GetHeight = 12
    IImageProperties_GetTitle = 13
    IImageProperties_PutTitle = 14
    IImageProperties_GetLatitude = 15
    IImageProperties_GetLongitude = 16
    IImageProperties_GetCameraManufacturer = 17
    IImageProperties_PutCameraManufacturer = 18
    IImageProperties_GetCameraModel = 19
    IImageProperties_PutCameraModel = 20
    IImageProperties_GetOrientation = 21
    IImageProperties_GetPeopleNames = 22

    ' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7
    IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault = 8

    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Vars ====----
Private m_pIImageProperties As Long
Private m_pIStorageItemExtraProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIImageProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIImageProperties
End Property

Public Property Let Ifc(ByVal pIImageProperties As Long)
    Call ReleaseIfc(m_pIImageProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
    m_pIImageProperties = pIImageProperties
    Call QueryIfc(m_pIImageProperties, _
                  IID_IStorageItemExtraProperties, _
                  m_pIStorageItemExtraProperties)
End Property

' IImageProperties
Public Property Get Rating() As Long
    Dim Ret As Long
    If m_pIImageProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetRating, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Rating = Ret
End Property

Public Property Let Rating(ByVal value As Long)
    If m_pIImageProperties <> 0& Then
        If Invoke(m_pIImageProperties, _
                  IImageProperties_PutRating, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Keywords() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIImageProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetKeywords, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Keywords = Ret
End Property

Public Property Get DateTaken() As DateTime
    Dim Ret As DateTime
    If m_pIImageProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetDateTaken, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set DateTaken = Ret
End Property

Public Property Let DateTaken(ByVal value As DateTime)
    If m_pIImageProperties <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIImageProperties, _
                  IImageProperties_PutDateTaken, _
                  value.DateTimeValue) = S_OK Then
        End If
    End If
End Property

Public Property Get Width() As Long
    Dim Ret As Long
    If m_pIImageProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetWidth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Width = Ret
End Property

Public Property Get Height() As Long
    Dim Ret As Long
    If m_pIImageProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetHeight, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Height = Ret
End Property

Public Property Get Title() As String
    Dim Ret As String
    If m_pIImageProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pIImageProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIImageProperties, _
                  IImageProperties_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Latitude() As Double
    Dim Ret As Double
    If m_pIImageProperties <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetLatitude, _
                  VarPtr(pIReference)) = S_OK Then
            Dim value As Double
            If Invoke(pIReference, _
                      IReference_GetValue, _
                      VarPtr(value)) = S_OK Then
                Ret = value
            End If
            Call ReleaseIfc(pIReference)
        End If
    End If
    Latitude = Ret
End Property

Public Property Get Longitude() As Double
    Dim Ret As Double
    If m_pIImageProperties <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetLongitude, _
                  VarPtr(pIReference)) = S_OK Then
            Dim value As Double
            If Invoke(pIReference, _
                      IReference_GetValue, _
                      VarPtr(value)) = S_OK Then
                Ret = value
            End If
            Call ReleaseIfc(pIReference)
        End If
    End If
    Longitude = Ret
End Property

Public Property Get CameraManufacturer() As String
    Dim Ret As String
    If m_pIImageProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetCameraManufacturer, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CameraManufacturer = Ret
End Property

Public Property Let CameraManufacturer(ByVal value As String)
    If m_pIImageProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIImageProperties, _
                  IImageProperties_PutCameraManufacturer, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get CameraModel() As String
    Dim Ret As String
    If m_pIImageProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetCameraModel, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CameraModel = Ret
End Property

Public Property Let CameraModel(ByVal value As String)
    If m_pIImageProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIImageProperties, _
                  IImageProperties_PutCameraModel, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Orientation() As PhotoOrientation
    Dim Ret As PhotoOrientation
    If m_pIImageProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetOrientation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Orientation = Ret
End Property

Public Property Get PeopleNames() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIImageProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIImageProperties, _
                  IImageProperties_GetPeopleNames, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set PeopleNames = Ret
End Property

' ----==== Functions ====----
' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7

Public Function SavePropertiesAsyncOverloadDefault() As Boolean
    Dim Ret As Boolean
    If m_pIStorageItemExtraProperties <> 0& Then
        Dim operation As Long
        If Invoke(m_pIStorageItemExtraProperties, _
                  IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    SavePropertiesAsyncOverloadDefault = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

