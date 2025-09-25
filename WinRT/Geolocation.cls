VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Geolocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

' Namespace Geolocation

Option Explicit

' ----==== Enums Windows.Devices.Geolocation.xxx ====----
Public Enum GeolocationAccessStatus
    GeolocationAccessStatus_Unspecified = 0
    GeolocationAccessStatus_Allowed = 1
    GeolocationAccessStatus_Denied = 2
End Enum

Public Enum PositionAccuracy
    PositionAccuracy_Default = 0
    PositionAccuracy_High = 1
End Enum
                
Public Enum PositionStatus
    PositionStatus_Ready = 0
    PositionStatus_Initializing = 1
    PositionStatus_NoData = 2
    PositionStatus_Disabled = 3
    PositionStatus_NotInitialized = 4
    PositionStatus_NotAvailable = 5
End Enum

Public Enum PositionSource
    PositionSource_Cellular = 0
    PositionSource_Satellite = 1
    PositionSource_WiFi = 2
    PositionSource_IPAddress = 3
    PositionSource_Unknown = 4
    PositionSource_Default = 5
    PositionSource_Obfuscated = 6
End Enum

Public Enum GeoshapeType
    GeoshapeType_Geopoint = 0
    GeoshapeType_Geocircle = 1
    GeoshapeType_Geopath = 2
    GeoshapeType_GeoboundingBox = 3
End Enum
                
Public Enum AltitudeReferenceSystem
    AltitudeReferenceSystem_Unspecified = 0
    AltitudeReferenceSystem_Terrain = 1
    AltitudeReferenceSystem_Ellipsoid = 2
    AltitudeReferenceSystem_Geoid = 3
    AltitudeReferenceSystem_Surface = 4
End Enum
                
' ----==== Namespace Geolocation.Geolocator ====----
Public Property Get Geolocator() As Geolocator
    Set Geolocator = New Geolocator
End Property

' ----==== Namespace Geolocation.Geoposition ====----
Public Property Get Geoposition() As Geoposition
    Set Geoposition = New Geoposition
End Property

' ----==== Namespace Geolocation.Geocoordinate ====----
Public Property Get Geocoordinate() As Geocoordinate
    Set Geocoordinate = New Geocoordinate
End Property

' ----==== Namespace Geolocation.Geopoint ====----
Public Property Get Geopoint() As Geopoint
    Set Geopoint = New Geopoint
End Property

' ----==== Namespace Geolocation.GeocoordinateSatelliteData ====----
Public Property Get GeocoordinateSatelliteData() As GeocoordinateSatelliteData
    Set GeocoordinateSatelliteData = New GeocoordinateSatelliteData
End Property

' ----==== Namespace Geolocation.CivicAddress ====----
Public Property Get CivicAddress() As CivicAddress
    Set CivicAddress = New CivicAddress
End Property

' ----==== Namespace Geolocation.VenueData ====----
Public Property Get VenueData() As VenueData
    Set VenueData = New VenueData
End Property

' ----==== Namespace Geolocation.BasicGeoposition ====----
Public Property Get BasicGeoposition(Optional ByVal Latitude As Double = 0, _
                                     Optional ByVal Longitude As Double = 0, _
                                     Optional ByVal Altitude As Double = 0) As BasicGeoposition
    Set BasicGeoposition = New BasicGeoposition
    BasicGeoposition.Latitude = Latitude
    BasicGeoposition.Longitude = Longitude
    BasicGeoposition.Altitude = Altitude
End Property

