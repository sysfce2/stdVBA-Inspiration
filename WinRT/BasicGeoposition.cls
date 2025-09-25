VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasicGeoposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_Latitude As Double
Private m_Longitude As Double
Private m_Altitude As Double

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get Latitude() As Double
    Latitude = m_Latitude
End Property

Public Property Let Latitude(ByVal value As Double)
    m_Latitude = value
End Property

Public Property Get Longitude() As Double
    Longitude = m_Longitude
End Property

Public Property Let Longitude(ByVal value As Double)
    m_Longitude = value
End Property

Public Property Get Altitude() As Double
    Altitude = m_Altitude
End Property

Public Property Let Altitude(ByVal value As Double)
    m_Altitude = value
End Property

Public Property Get ToString() As String
     ToString = "{Latitude=" & CStr(m_Latitude) & ",Longitude=" & CStr(m_Longitude) & ",Altitude=" & CStr(m_Altitude) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----


