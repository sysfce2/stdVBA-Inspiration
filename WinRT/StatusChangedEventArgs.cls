VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusChangedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_StatusChangedEventArgs As String = "Windows.Devices.Geolocation.IStatusChangedEventArgs"
Private Const IID_IStatusChangedEventArgs As String = "{3453d2da-8c93-4111-a205-9aecfc9be5c0}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IStatusChangedEventArgs
    IStatusChangedEventArgs_GetStatus = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIStatusChangedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStatusChangedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStatusChangedEventArgs
End Property

Public Property Let Ifc(ByVal pIStatusChangedEventArgs As Long)
    Call ReleaseIfc(m_pIStatusChangedEventArgs)
    m_pIStatusChangedEventArgs = pIStatusChangedEventArgs
End Property

' IStatusChangedEventArgs
Public Property Get Status() As PositionStatus
    Dim Ret As PositionStatus
    If m_pIStatusChangedEventArgs <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStatusChangedEventArgs, _
                  IStatusChangedEventArgs_GetStatus, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Status = Ret
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


