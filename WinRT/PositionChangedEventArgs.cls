VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionChangedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_PositionChangedEventArgs As String = "Windows.Devices.Geolocation.IPositionChangedEventArgs"
Private Const IID_IPositionChangedEventArgs As String = "{37859ce5-9d1e-46c5-bf3b-6ad8cac1a093}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IPositionChangedEventArgs
    IPositionChangedEventArgs_GetPosition = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIPositionChangedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPositionChangedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPositionChangedEventArgs
End Property

Public Property Let Ifc(ByVal pIPositionChangedEventArgs As Long)
    Call ReleaseIfc(m_pIPositionChangedEventArgs)
    m_pIPositionChangedEventArgs = pIPositionChangedEventArgs
End Property

' IPositionChangedEventArgs
Public Property Get Position() As Geoposition
    Dim Ret As Geoposition
    If m_pIPositionChangedEventArgs <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPositionChangedEventArgs, _
                  IPositionChangedEventArgs_GetPosition, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Geoposition
                Ret.Ifc = Value
            End If
        End If
    End If
    Set Position = Ret
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


