VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Geoposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_Geoposition As String = "Windows.Devices.Geolocation.Geoposition"
Private Const IID_IGeoposition As String = "{c18d0454-7d41-4ff7-a957-9dffb4ef7f5b}"
Private Const IID_IGeoposition2 As String = "{7f62f697-8671-4b0d-86f8-474a8496187c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
        
    ' IGeoposition
    IGeoposition_GetCoordinate = 6
    IGeoposition_GetCivicAddress = 7
         
    ' IGeoposition2
    IGeoposition2_GetVenueData = 6
    
End Enum

' ----==== Vars ====----
Private m_pIGeoposition As Long
Private m_pIGeoposition2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGeoposition2)
    Call ReleaseIfc(m_pIGeoposition)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGeoposition
End Property

Public Property Let Ifc(ByVal pIGeoposition As Long)
    Call ReleaseIfc(m_pIGeoposition2)
    Call ReleaseIfc(m_pIGeoposition)
    m_pIGeoposition = pIGeoposition
    Call QueryIfc(m_pIGeoposition, IID_IGeoposition2, m_pIGeoposition2)
End Property

' IGeoposition
Public Property Get Coordinate() As Geocoordinate
    Dim Ret As Geocoordinate
    If m_pIGeoposition <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeoposition, _
                  IGeoposition_GetCoordinate, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Geocoordinate
                Ret.Ifc = Value
            End If
        End If
    End If
    Set Coordinate = Ret
End Property

Public Property Get CivicAddress() As CivicAddress
    Dim Ret As CivicAddress
    If m_pIGeoposition <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeoposition, _
                  IGeoposition_GetCivicAddress, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New CivicAddress
                Ret.Ifc = Value
            End If
        End If
    End If
    Set CivicAddress = Ret
End Property
         
' IGeoposition2
Public Property Get VenueData() As VenueData
    Dim Ret As VenueData
    If m_pIGeoposition2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIGeoposition2, _
                  IGeoposition2_GetVenueData, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New VenueData
                Ret.Ifc = Value
            End If
        End If
    End If
    Set VenueData = Ret
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

