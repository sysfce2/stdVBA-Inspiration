VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Geopoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_Geopoint As String = "Windows.Devices.Geolocation.Geopoint"
Private Const IID_IGeopoint As String = "{6bfa00eb-e56e-49bb-9caf-cbaa78a8bcef}"
Private Const IID_IGeoshape As String = "{c99ca2af-c729-43c1-8fab-d6dec914df7e}"
Private Const IID_IGeopointFactory As String = "{db6b8d33-76bd-4e30-8af7-a844dc37b7a0}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
        
    ' IGeopointFactory
    IGeopointFactory_Create = 6
    IGeopointFactory_CreateWithAltitudeReferenceSystem = 7
    IGeopointFactory_CreateWithAltitudeReferenceSystemAndSpatialReferenceId = 8
                        
    ' IGeopoint
    IGeopoint_GetPosition = 6
    
    ' IGeoshape
    IGeoshape_GetGeoshapeType = 6
    IGeoshape_GetSpatialReferenceId = 7
    IGeoshape_GetAltitudeReferenceSystem = 8

End Enum

' ----==== Type ====----
Private Type BasicGeoposition_
    Latitude As Double
    Longitude As Double
    Altitude As Double
End Type
                
' ----==== Vars ====----
Private m_pIGeopoint As Long
Private m_pIGeoshape As Long
Private m_pIGeopointFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Devices_Geolocation_Geopoint, _
                            IID_IGeopointFactory, _
                            m_pIGeopointFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGeoshape)
    Call ReleaseIfc(m_pIGeopoint)
    Call ReleaseIfc(m_pIGeopointFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGeopoint
End Property

Public Property Let Ifc(ByVal pIGeopoint As Long)
    Call ReleaseIfc(m_pIGeoshape)
    Call ReleaseIfc(m_pIGeopoint)
    m_pIGeopoint = pIGeopoint
    Call QueryIfc(m_pIGeopoint, IID_IGeoshape, m_pIGeoshape)
End Property

' IGeopoint
Public Property Get Position() As BasicGeoposition
    Dim Ret As BasicGeoposition
    If m_pIGeopoint <> 0& Then
        Dim Value As BasicGeoposition_
        If Invoke(m_pIGeopoint, _
                  IGeopoint_GetPosition, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New BasicGeoposition
            Ret.Latitude = Value.Latitude
            Ret.Longitude = Value.Longitude
            Ret.Altitude = Value.Altitude
        End If
    End If
    Set Position = Ret
End Property
    
' IGeoshape
Public Property Get GeoshapeType() As GeoshapeType
    Dim Ret As GeoshapeType
    If m_pIGeoshape <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeoshape, _
                  IGeoshape_GetGeoshapeType, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GeoshapeType = Ret
End Property

Public Property Get SpatialReferenceId() As Long
    Dim Ret As Long
    If m_pIGeoshape <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeoshape, _
                  IGeoshape_GetSpatialReferenceId, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    SpatialReferenceId = Ret
End Property

Public Property Get AltitudeReferenceSystem() As AltitudeReferenceSystem
    Dim Ret As GeoshapeType
    If m_pIGeoshape <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeoshape, _
                  IGeoshape_GetAltitudeReferenceSystem, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    AltitudeReferenceSystem = Ret
End Property

' ----==== Functions ====----
' IGeopointFactory
Public Function Create(ByVal Position As BasicGeoposition) As Geopoint
    Dim Ret As Geopoint
    If m_pIGeopointFactory <> 0& And _
       IsNotNothing(Position) Then
        Dim Value As Long
        If Invoke(m_pIGeopointFactory, _
                  IGeopointFactory_Create, _
                  Position.Latitude, _
                  Position.Longitude, _
                  Position.Altitude, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Geopoint
                Ret.Ifc = Value
            End If
        End If
    End If
    Set Create = Ret
End Function

Public Function CreateWithAltitudeReferenceSystem(ByVal Position As BasicGeoposition, _
                                                  ByVal altReferenceSystem As AltitudeReferenceSystem) As Geopoint
    Dim Ret As Geopoint
    If m_pIGeopointFactory <> 0& And _
       IsNotNothing(Position) Then
        Dim Value As Long
        If Invoke(m_pIGeopointFactory, _
                  IGeopointFactory_CreateWithAltitudeReferenceSystem, _
                  Position.Latitude, _
                  Position.Longitude, _
                  Position.Altitude, _
                  altReferenceSystem, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Geopoint
                Ret.Ifc = Value
            End If
        End If
    End If
    Set CreateWithAltitudeReferenceSystem = Ret
End Function

Public Function CreateWithAltitudeReferenceSystemAndSpatialReferenceId(ByVal Position As BasicGeoposition, _
                                                                       ByVal altReferenceSystem As AltitudeReferenceSystem, _
                                                                       ByVal SpatialReferenceId As Long) As Geopoint
    Dim Ret As Geopoint
    If m_pIGeopointFactory <> 0& And _
       IsNotNothing(Position) Then
        Dim Value As Long
        If Invoke(m_pIGeopointFactory, _
                  IGeopointFactory_CreateWithAltitudeReferenceSystemAndSpatialReferenceId, _
                  Position.Latitude, _
                  Position.Longitude, _
                  Position.Altitude, _
                  altReferenceSystem, _
                  SpatialReferenceId, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Geopoint
                Ret.Ifc = Value
            End If
        End If
    End If
    Set CreateWithAltitudeReferenceSystemAndSpatialReferenceId = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


