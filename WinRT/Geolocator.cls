VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Geolocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_Geolocator As String = "Windows.Devices.Geolocation.Geolocator"
Private Const IID_IGeolocator As String = "{a9c3bf62-4524-4989-8aa9-de019d2e551f}"
Private Const IID_IGeolocator2 As String = "{d1b42e6d-8891-43b4-ad36-27c6fe9a97b1}"
Private Const IID_IGeolocatorStatics As String = "{9a8e7571-2df5-4591-9f87-eb5fd894e9b7}"
Private Const IID_IGeolocatorStatics2 As String = "{993011a2-fa1c-4631-a71d-0dbeb1250d9c}"
Private Const IID_IGeolocatorWithScalarAccuracy As String = "{96f5d3c1-b80f-460a-994d-a96c47a51aa4}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGeolocatorStatics
    IGeolocatorStatics_RequestAccessAsync = 6
    IGeolocatorStatics_GetGeopositionHistoryAsync = 7
    IGeolocatorStatics_GetGeopositionHistoryWithDurationAsync = 8
                        
    ' IGeolocatorStatics2
    IGeolocatorStatics2_GetIsDefaultGeopositionRecommended = 6
'    IGeolocatorStatics2_PutDefaultGeoposition = 7
    IGeolocatorStatics2_GetDefaultGeoposition = 8
                        
    ' IGeolocator
    IGeolocator_GetDesiredAccuracy = 6
    IGeolocator_PutDesiredAccuracy = 7
    IGeolocator_GetMovementThreshold = 8
    IGeolocator_PutMovementThreshold = 9
    IGeolocator_GetReportInterval = 10
    IGeolocator_PutReportInterval = 11
    IGeolocator_GetLocationStatus = 12
    IGeolocator_GetGeopositionAsync = 13
    IGeolocator_GetGeopositionAsyncWithAgeAndTimeout = 14
    IGeolocator_AddPositionChanged = 15
    IGeolocator_RemovePositionChanged = 16
    IGeolocator_AddStatusChanged = 17
    IGeolocator_RemoveStatusChanged = 18

    ' IGeolocator2
    IGeolocator2_AllowFallbackToConsentlessPositions = 6

    ' IGeolocatorWithScalarAccuracy
    IGeolocatorWithScalarAccuracy_GetDesiredAccuracyInMeters = 6
    IGeolocatorWithScalarAccuracy_PutDesiredAccuracyInMeters = 7
    
    ' IReference_xxx
    IReference_GetValue = 6
    
End Enum

' ----==== Type ====----
Private Type BasicGeoposition_
    Latitude As Double
    Longitude As Double
    Altitude As Double
End Type

' ----==== Vars ====----
Private m_pIGeolocator As Long
Private m_pIGeolocator2 As Long
Private m_pIGeolocatorStatics As Long
Private m_pIGeolocatorStatics2 As Long
Private m_pIGeolocatorWithScalarAccuracy As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Devices_Geolocation_Geolocator, _
                            IID_IGeolocatorStatics, _
                            m_pIGeolocatorStatics) Then
        Call QueryIfc(m_pIGeolocatorStatics, IID_IGeolocatorStatics2, m_pIGeolocatorStatics2)
    End If
    Dim pIGeolocator As Long
    If GetActivateInstance(Windows_Devices_Geolocation_Geolocator, _
                           IID_IGeolocator, _
                           pIGeolocator) Then
        Me.Ifc = pIGeolocator
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGeolocatorWithScalarAccuracy)
    Call ReleaseIfc(m_pIGeolocator2)
    Call ReleaseIfc(m_pIGeolocator)
    Call ReleaseIfc(m_pIGeolocatorStatics2)
    Call ReleaseIfc(m_pIGeolocatorStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGeolocator
End Property

Public Property Let Ifc(ByVal pIGeolocator As Long)
    Call ReleaseIfc(m_pIGeolocatorWithScalarAccuracy)
    Call ReleaseIfc(m_pIGeolocator2)
    Call ReleaseIfc(m_pIGeolocator)
    m_pIGeolocator = pIGeolocator
    Call QueryIfc(m_pIGeolocator, IID_IGeolocator2, m_pIGeolocator2)
    Call QueryIfc(m_pIGeolocator, IID_IGeolocatorWithScalarAccuracy, m_pIGeolocatorWithScalarAccuracy)
End Property

' IGeolocatorStatics2
Public Property Get IsDefaultGeopositionRecommended() As Boolean
    Dim Ret As Boolean
    If m_pIGeolocatorStatics2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeolocatorStatics2, _
                  IGeolocatorStatics2_GetIsDefaultGeopositionRecommended, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsDefaultGeopositionRecommended = Ret
End Property

'    IGeolocatorStatics2_put_DefaultGeoposition = 7

Public Property Get DefaultGeoposition() As BasicGeoposition
    Dim Ret As BasicGeoposition
    If m_pIGeolocatorStatics2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeolocatorStatics2, _
                  IGeolocatorStatics2_GetDefaultGeoposition, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As BasicGeoposition_
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Set Ret = New BasicGeoposition
                    Ret.Latitude = value.Latitude
                    Ret.Longitude = value.Longitude
                    Ret.Altitude = value.Altitude
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Set DefaultGeoposition = Ret
End Property

' IGeolocator
Public Property Get DesiredAccuracy() As PositionAccuracy
    Dim Ret As PositionAccuracy
    If m_pIGeolocator <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeolocator, _
                  IGeolocator_GetDesiredAccuracy, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DesiredAccuracy = Ret
End Property

Public Property Let DesiredAccuracy(ByVal value As PositionAccuracy)
    Dim Ret As Double
    If m_pIGeolocator <> 0& Then
        If Invoke(m_pIGeolocator, _
                  IGeolocator_PutDesiredAccuracy, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get MovementThreshold() As Double
    Dim Ret As Double
    If m_pIGeolocator <> 0& Then
        Dim value As Double
        If Invoke(m_pIGeolocator, _
                  IGeolocator_GetMovementThreshold, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    MovementThreshold = Ret
End Property

Public Property Let MovementThreshold(ByVal value As Double)
    Dim Ret As Double
    If m_pIGeolocator <> 0& Then
        If Invoke(m_pIGeolocator, _
                  IGeolocator_PutMovementThreshold, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get ReportInterval() As Long
    Dim Ret As Long
    If m_pIGeolocator <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeolocator, _
                  IGeolocator_GetReportInterval, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ReportInterval = Ret
End Property

Public Property Let ReportInterval(ByVal value As Long)
    Dim Ret As Double
    If m_pIGeolocator <> 0& Then
        If Invoke(m_pIGeolocator, _
                  IGeolocator_PutReportInterval, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get LocationStatus() As PositionStatus
    Dim Ret As PositionStatus
    If m_pIGeolocator <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeolocator, _
                  IGeolocator_GetLocationStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LocationStatus = Ret
End Property

' IGeolocatorWithScalarAccuracy
Public Property Get DesiredAccuracyInMeters() As Long
    Dim Ret As Long
    If m_pIGeolocatorWithScalarAccuracy <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGeolocatorWithScalarAccuracy, _
                  IGeolocatorWithScalarAccuracy_GetDesiredAccuracyInMeters, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    DesiredAccuracyInMeters = Ret
End Property

Public Property Let DesiredAccuracyInMeters(ByVal value As Long)
    If m_pIGeolocatorWithScalarAccuracy <> 0& Then
        Dim PropertyValue As New PropertyValue
        Dim pIReference As Long
        pIReference = PropertyValue.CreateUInt32(value)
        If pIReference <> 0& Then
            If Invoke(m_pIGeolocatorWithScalarAccuracy, _
                      IGeolocatorWithScalarAccuracy_PutDesiredAccuracyInMeters, _
                      pIReference) = S_OK Then
            End If
            Call ReleaseIfc(pIReference)
        End If
    End If
End Property

' ----==== Functions ====----
' IGeolocatorStatics
Public Function RequestAccessAsync() As GeolocationAccessStatus
    Dim Ret As GeolocationAccessStatus
    If m_pIGeolocatorStatics <> 0& Then
        Dim Result As Long
        If Invoke(m_pIGeolocatorStatics, _
                  IGeolocatorStatics_RequestAccessAsync, _
                  VarPtr(Result)) = S_OK Then
            If Result <> 0& Then
                If Await(Result) Then
                    Ret = Result
                End If
            End If
        End If
    End If
    RequestAccessAsync = Ret
End Function

Public Function GetGeopositionHistoryAsync(ByVal StartTime As DateTime) As ReadOnlyList_1 'ReadOnlyList_Geoposition
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_Geoposition
    If m_pIGeolocatorStatics <> 0& And _
       IsNotNothing(StartTime) Then
        Dim pIVectorView As Long
        If Invoke(m_pIGeolocatorStatics, _
                  IGeolocatorStatics_GetGeopositionHistoryAsync, _
                  StartTime.DateTimeValue, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_Geoposition, pIVectorView) 'ReadOnlyList_Geoposition
                End If
            End If
        End If
    End If
    Set GetGeopositionHistoryAsync = Ret
End Function

Public Function GetGeopositionHistoryWithDurationAsync(ByVal StartTime As DateTime, _
                                                       ByVal Duration As TimeSpan) As ReadOnlyList_1 'ReadOnlyList_Geoposition
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_Geoposition
    If m_pIGeolocatorStatics <> 0& And _
       IsNotNothing(StartTime) And _
       IsNotNothing(Duration) Then
        Dim pIVectorView As Long
        If Invoke(m_pIGeolocatorStatics, _
                  IGeolocatorStatics_GetGeopositionHistoryWithDurationAsync, _
                  StartTime.DateTimeValue, _
                  Duration.TimeSpanValue, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_Geoposition, pIVectorView) 'ReadOnlyList_Geoposition
                End If
            End If
        End If
    End If
    Set GetGeopositionHistoryWithDurationAsync = Ret
End Function

' IGeolocator
Public Function GetGeopositionAsync() As Geoposition
    Dim Ret As Geoposition
    If m_pIGeolocator <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeolocator, _
                  IGeolocator_GetGeopositionAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Set Ret = New Geoposition
                    Ret.Ifc = value
                End If
            End If
        End If
    End If
    Set GetGeopositionAsync = Ret
End Function

Public Function GetGeopositionAsyncWithAgeAndTimeout(ByVal maximumAge As TimeSpan, _
                                                     ByVal timeout As TimeSpan) As Geoposition
    Dim Ret As Geoposition
    If m_pIGeolocator <> 0& And _
       IsNotNothing(maximumAge) And _
       IsNotNothing(timeout) Then
        Dim value As Long
        If Invoke(m_pIGeolocator, _
                  IGeolocator_GetGeopositionAsyncWithAgeAndTimeout, _
                  maximumAge.TimeSpanValue, _
                  timeout.TimeSpanValue, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Set Ret = New Geoposition
                    Ret.Ifc = value
                End If
            End If
        End If
    End If
    Set GetGeopositionAsyncWithAgeAndTimeout = Ret
End Function

Public Function AddPositionChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGeolocator <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIGeolocator, _
                  IGeolocator_AddPositionChanged, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddPositionChanged = Ret
End Function

Public Function RemovePositionChanged(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGeolocator <> 0& Then
        If Invoke(m_pIGeolocator, _
                  IGeolocator_RemovePositionChanged, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemovePositionChanged = Ret
End Function

Public Function AddStatusChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGeolocator <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIGeolocator, _
                  IGeolocator_AddStatusChanged, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddStatusChanged = Ret
End Function

Public Function RemoveStatusChanged(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGeolocator <> 0& Then
        If Invoke(m_pIGeolocator, _
                  IGeolocator_RemoveStatusChanged, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStatusChanged = Ret
End Function

' IGeolocator2
Public Function AllowFallbackToConsentlessPositions() As Boolean
    Dim Ret As Boolean
    If m_pIGeolocator2 <> 0& Then
        If Invoke(m_pIGeolocator2, _
                  IGeolocator2_AllowFallbackToConsentlessPositions) = S_OK Then
            Ret = True
        End If
    End If
    AllowFallbackToConsentlessPositions = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


