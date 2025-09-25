VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Geocoordinate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_Geocoordinate As String = "Windows.Devices.Geolocation.Geocoordinate"
Private Const IID_IGeocoordinate As String = "{ee21a3aa-976a-4c70-803d-083ea55bcbc4}"
Private Const IID_IGeocoordinateWithPoint As String = "{feea0525-d22c-4d46-b527-0b96066fc7db}"
Private Const IID_IGeocoordinateWithPositionData As String = "{95e634be-dbd6-40ac-b8f2-a65c0340d9a6}"
Private Const IID_IGeocoordinateWithRemoteSource As String = "{397cebd7-ee38-5f3b-8900-c4a7bc9cf953}"
Private Const IID_IGeocoordinateWithPositionSourceTimestamp As String = "{8543fc02-c9f1-4610-afe0-8bc3a6a87036}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
        
    ' IGeocoordinate
    IGeocoordinate_GetLatitude = 6 ' DEPRECATED ("Latitude may be altered or unavailable after Windows 8.1. Instead, use Point.Position.Latitude")
    IGeocoordinate_GetLongitude = 7 ' DEPRECATED ("Longitude may be altered or unavailable after Windows 8.1. Instead, use Point.Position.Longitude")
    IGeocoordinate_GetAltitude = 8 ' DEPRECATED ("Altitude may be altered or unavailable after Windows 8.1. Instead, use Point.Position.Altitude")
    IGeocoordinate_GetAccuracy = 9
    IGeocoordinate_GetAltitudeAccuracy = 10
    IGeocoordinate_GetHeading = 11
    IGeocoordinate_GetSpeed = 12
    IGeocoordinate_GetTimestamp = 13

    ' IGeocoordinateWithPoint
    IGeocoordinateWithPoint_GetPoint = 6

    ' IGeocoordinateWithPositionData
    IGeocoordinateWithPositionData_GetPositionSource = 6
    IGeocoordinateWithPositionData_GetSatelliteData = 7
                        
    ' IGeocoordinateWithRemoteSource
    IGeocoordinateWithRemoteSource_GetIsRemoteSource = 6
    
    ' IGeocoordinateWithPositionSourceTimestamp
    IGeocoordinateWithPositionSourceTimestamp_GetPositionSourceTimestamp = 6
    
    ' IReference_xxx
    IReference_GetValue = 6

End Enum

' ----==== Vars ====----
Private m_pIGeocoordinate As Long
Private m_pIGeocoordinateWithPoint As Long
Private m_pIGeocoordinateWithPositionData As Long
Private m_pIGeocoordinateWithRemoteSource As Long
Private m_pIGeocoordinateWithPositionSourceTimestamp As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGeocoordinateWithPositionSourceTimestamp)
    Call ReleaseIfc(m_pIGeocoordinateWithRemoteSource)
    Call ReleaseIfc(m_pIGeocoordinateWithPositionData)
    Call ReleaseIfc(m_pIGeocoordinateWithPoint)
    Call ReleaseIfc(m_pIGeocoordinate)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGeocoordinate
End Property

Public Property Let Ifc(ByVal pIGeocoordinate As Long)
    Call ReleaseIfc(m_pIGeocoordinateWithPositionSourceTimestamp)
    Call ReleaseIfc(m_pIGeocoordinateWithRemoteSource)
    Call ReleaseIfc(m_pIGeocoordinateWithPositionData)
    Call ReleaseIfc(m_pIGeocoordinateWithPoint)
    Call ReleaseIfc(m_pIGeocoordinate)
    m_pIGeocoordinate = pIGeocoordinate
    Call QueryIfc(m_pIGeocoordinate, IID_IGeocoordinateWithPoint, m_pIGeocoordinateWithPoint)
    Call QueryIfc(m_pIGeocoordinate, IID_IGeocoordinateWithPositionData, m_pIGeocoordinateWithPositionData)
    Call QueryIfc(m_pIGeocoordinate, IID_IGeocoordinateWithRemoteSource, m_pIGeocoordinateWithRemoteSource)
    Call QueryIfc(m_pIGeocoordinate, IID_IGeocoordinateWithPositionSourceTimestamp, m_pIGeocoordinateWithPositionSourceTimestamp)
End Property

' IGeocoordinate
' DEPRECATED ("Latitude may be altered or unavailable after Windows 8.1. Instead, use Point.Position.Latitude")
Public Property Get Latitude() As Double
    Dim Ret As Double
    If m_pIGeocoordinate <> 0& Then
        Dim Value As Double
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetLatitude, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Latitude = Ret
End Property

' DEPRECATED ("Longitude may be altered or unavailable after Windows 8.1. Instead, use Point.Position.Longitude")
Public Property Get Longitude() As Double
    Dim Ret As Double
    If m_pIGeocoordinate <> 0& Then
        Dim Value As Double
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetLongitude, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Longitude = Ret
End Property
    
' DEPRECATED ("Altitude may be altered or unavailable after Windows 8.1. Instead, use Point.Position.Altitude")
Public Property Get Altitude() As Double
    Dim Ret As Double
    If m_pIGeocoordinate <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetAltitude, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim Value As Double
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(Value)) = S_OK Then
                    Ret = Value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Altitude = Ret
End Property

Public Property Get Accuracy() As Double
    Dim Ret As Double
    If m_pIGeocoordinate <> 0& Then
        Dim Value As Double
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetAccuracy, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Accuracy = Ret
End Property

Public Property Get AltitudeAccuracy() As Double
    Dim Ret As Double
    If m_pIGeocoordinate <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetAltitudeAccuracy, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim Value As Double
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(Value)) = S_OK Then
                    Ret = Value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    AltitudeAccuracy = Ret
End Property

Public Property Get Heading() As Double
    Dim Ret As Double
    If m_pIGeocoordinate <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetHeading, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim Value As Double
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(Value)) = S_OK Then
                    Ret = Value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Heading = Ret
End Property

Public Property Get Speed() As Double
    Dim Ret As Double
    If m_pIGeocoordinate <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetSpeed, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim Value As Double
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(Value)) = S_OK Then
                    Ret = Value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Speed = Ret
End Property

Public Property Get Timestamp() As DateTime
    Dim Ret As DateTime
    If m_pIGeocoordinate <> 0& Then
        Dim Value As Double
        If Invoke(m_pIGeocoordinate, _
                  IGeocoordinate_GetTimestamp, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = Value
        End If
    End If
    Set Timestamp = Ret
End Property

' IGeocoordinateWithPoint
Public Property Get Point() As Geopoint
    Dim Ret As Geopoint
    If m_pIGeocoordinateWithPoint <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeocoordinateWithPoint, _
                  IGeocoordinateWithPoint_GetPoint, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Geopoint
                Ret.Ifc = Value
            End If
        End If
    End If
    Set Point = Ret
End Property

' IGeocoordinateWithPositionData
Public Property Get PositionSource() As PositionSource
    Dim Ret As PositionSource
    If m_pIGeocoordinateWithPositionData <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeocoordinateWithPositionData, _
                  IGeocoordinateWithPositionData_GetPositionSource, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    PositionSource = Ret
End Property

Public Property Get SatelliteData() As GeocoordinateSatelliteData
    Dim Ret As GeocoordinateSatelliteData
    If m_pIGeocoordinateWithPositionData <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeocoordinateWithPositionData, _
                  IGeocoordinateWithPositionData_GetSatelliteData, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New GeocoordinateSatelliteData
                Ret.Ifc = Value
            End If
        End If
    End If
    Set SatelliteData = Ret
End Property
                        
' IGeocoordinateWithRemoteSource
Public Property Get IsRemoteSource() As Boolean
    Dim Ret As Boolean
    If m_pIGeocoordinateWithRemoteSource <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeocoordinateWithRemoteSource, _
                  IGeocoordinateWithRemoteSource_GetIsRemoteSource, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    IsRemoteSource = Ret
End Property
    
' IGeocoordinateWithPositionSourceTimestamp
Public Property Get PositionSourceTimestamp() As DateTime
    Dim Ret As DateTime
    If m_pIGeocoordinateWithPositionSourceTimestamp <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinateWithPositionSourceTimestamp, _
                  IGeocoordinateWithPositionSourceTimestamp_GetPositionSourceTimestamp, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim Value As Currency
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(Value)) = S_OK Then
                    Set Ret = New DateTime
                    Ret.DateTimeValue = Value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Set PositionSourceTimestamp = Ret
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
