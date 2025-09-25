VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VenueData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_VenueData As String = "Windows.Devices.Geolocation.VenueData"
Private Const IID_IVenueData As String = "{66f39187-60e3-4b2f-b527-4f53f1c3c677}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
        
    ' IVenueData
    IVenueData_GetId = 6
    IVenueData_GetLevel = 7
                        
End Enum

' ----==== Vars ====----
Private m_pIVenueData As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVenueData)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVenueData
End Property

Public Property Let Ifc(ByVal pIVenueData As Long)
    Call ReleaseIfc(m_pIVenueData)
    m_pIVenueData = pIVenueData
End Property

' IVenueData
Public Property Get Id() As String
    Dim Ret As String
    If m_pIVenueData <> 0& Then
        Dim hString As Long
        If Invoke(m_pIVenueData, _
                  IVenueData_GetId, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Id = Ret
End Property

Public Property Get Level() As String
    Dim Ret As String
    If m_pIVenueData <> 0& Then
        Dim hString As Long
        If Invoke(m_pIVenueData, _
                  IVenueData_GetLevel, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Level = Ret
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


