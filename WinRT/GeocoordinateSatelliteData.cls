VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeocoordinateSatelliteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_GeocoordinateSatelliteData As String = "Windows.Devices.Geolocation.GeocoordinateSatelliteData"
Private Const IID_IGeocoordinateSatelliteData As String = "{c32a74d9-2608-474c-912c-06dd490f4af7}"
Private Const IID_IGeocoordinateSatelliteData2 As String = "{761c8cfd-a19d-5a51-80f5-71676115483e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
        
    ' IGeocoordinateSatelliteData
    IGeocoordinateSatelliteData_GetPositionDilutionOfPrecision = 6
    IGeocoordinateSatelliteData_GetHorizontalDilutionOfPrecision = 7
    IGeocoordinateSatelliteData_GetVerticalDilutionOfPrecision = 8
                        
    ' IGeocoordinateSatelliteData2
    IGeocoordinateSatelliteData2_GetGeometricDilutionOfPrecision = 6
    IGeocoordinateSatelliteData2_GetTimeDilutionOfPrecision = 7

    ' IReference_xxx
    IReference_GetValue = 6
    
End Enum

' ----==== Vars ====----
Private m_pIGeocoordinateSatelliteData As Long
Private m_pIGeocoordinateSatelliteData2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGeocoordinateSatelliteData2)
    Call ReleaseIfc(m_pIGeocoordinateSatelliteData)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGeocoordinateSatelliteData
End Property

Public Property Let Ifc(ByVal pIGeocoordinateSatelliteData As Long)
    Call ReleaseIfc(m_pIGeocoordinateSatelliteData)
    Call ReleaseIfc(m_pIGeocoordinateSatelliteData2)
    m_pIGeocoordinateSatelliteData = pIGeocoordinateSatelliteData
    Call QueryIfc(m_pIGeocoordinateSatelliteData, IID_IGeocoordinateSatelliteData2, m_pIGeocoordinateSatelliteData2)
End Property

' IGeocoordinateSatelliteData
Public Property Get PositionDilutionOfPrecision() As Double
    Dim Ret As Double
    If m_pIGeocoordinateSatelliteData <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinateSatelliteData, _
                  IGeocoordinateSatelliteData_GetPositionDilutionOfPrecision, _
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
    PositionDilutionOfPrecision = Ret
End Property

Public Property Get HorizontalDilutionOfPrecision() As Double
    Dim Ret As Double
    If m_pIGeocoordinateSatelliteData <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinateSatelliteData, _
                  IGeocoordinateSatelliteData_GetHorizontalDilutionOfPrecision, _
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
    HorizontalDilutionOfPrecision = Ret
End Property

Public Property Get VerticalDilutionOfPrecision() As Double
    Dim Ret As Double
    If m_pIGeocoordinateSatelliteData <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinateSatelliteData, _
                  IGeocoordinateSatelliteData_GetVerticalDilutionOfPrecision, _
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
    VerticalDilutionOfPrecision = Ret
End Property
                        
' IGeocoordinateSatelliteData2
Public Property Get GeometricDilutionOfPrecision() As Double
    Dim Ret As Double
    If m_pIGeocoordinateSatelliteData2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinateSatelliteData2, _
                  IGeocoordinateSatelliteData2_GetGeometricDilutionOfPrecision, _
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
    GeometricDilutionOfPrecision = Ret
End Property

Public Property Get TimeDilutionOfPrecision() As Double
    Dim Ret As Double
    If m_pIGeocoordinateSatelliteData2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeocoordinateSatelliteData2, _
                  IGeocoordinateSatelliteData2_GetTimeDilutionOfPrecision, _
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
    TimeDilutionOfPrecision = Ret
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


