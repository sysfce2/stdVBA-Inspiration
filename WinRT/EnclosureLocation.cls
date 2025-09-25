VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnclosureLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Enumeration_EnclosureLocation As String = "Windows.Devices.Enumeration.EnclosureLocation"
Private Const IID_IEnclosureLocation As String = "{42340a27-5810-459c-aabb-c65e1f813ecf}"
Private Const IID_IEnclosureLocation2 As String = "{2885995b-e07d-485d-8a9e-bdf29aef4f66}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IEnclosureLocation
    IEnclosureLocation_GetInDock = 6
    IEnclosureLocation_GetInLid = 7
    IEnclosureLocation_GetPanel = 8
                        
    ' IEnclosureLocation2
    IEnclosureLocation2_GetRotationAngleInDegreesClockwise = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIEnclosureLocation As Long
Private m_pIEnclosureLocation2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIEnclosureLocation2)
    Call ReleaseIfc(m_pIEnclosureLocation)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIEnclosureLocation
End Property

Public Property Let Ifc(ByVal pIEnclosureLocation As Long)
    Call ReleaseIfc(m_pIEnclosureLocation2)
    Call ReleaseIfc(m_pIEnclosureLocation)
    m_pIEnclosureLocation = pIEnclosureLocation
    Call QueryIfc(m_pIEnclosureLocation, IID_IEnclosureLocation2, m_pIEnclosureLocation2)
End Property

' IEnclosureLocation
Public Property Get InDock() As Boolean
    Dim Ret As Boolean
    If m_pIEnclosureLocation <> 0& Then
        Dim value As Long
        If Invoke(m_pIEnclosureLocation, _
                  IEnclosureLocation_GetInDock, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    InDock = Ret
End Property

Public Property Get InLid() As Boolean
    Dim Ret As Boolean
    If m_pIEnclosureLocation <> 0& Then
        Dim value As Long
        If Invoke(m_pIEnclosureLocation, _
                  IEnclosureLocation_GetInLid, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    InLid = Ret
End Property

Public Property Get Panel() As Panel
    Dim Ret As Panel
    If m_pIEnclosureLocation <> 0& Then
        Dim value As Long
        If Invoke(m_pIEnclosureLocation, _
                  IEnclosureLocation_GetPanel, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Panel = Ret
End Property
                        
' IEnclosureLocation2
Public Property Get RotationAngleInDegreesClockwise() As Long
    Dim Ret As Long
    If m_pIEnclosureLocation2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIEnclosureLocation2, _
                  IEnclosureLocation2_GetRotationAngleInDegreesClockwise, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    RotationAngleInDegreesClockwise = Ret
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


