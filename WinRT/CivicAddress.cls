VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CivicAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Geolocation_CivicAddress As String = "Windows.Devices.Geolocation.CivicAddress"
Private Const IID_ICivicAddress As String = "{a8567a1a-64f4-4d48-bcea-f6b008eca34c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
        
    ' ICivicAddress
    ICivicAddress_GetCountry = 6
    ICivicAddress_GetState = 7
    ICivicAddress_GetCity = 8
    ICivicAddress_GetPostalCode = 9
    ICivicAddress_GetTimestamp = 10
                        
End Enum

' ----==== Vars ====----
Private m_pICivicAddress As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICivicAddress)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICivicAddress
End Property

Public Property Let Ifc(ByVal pICivicAddress As Long)
    Call ReleaseIfc(m_pICivicAddress)
    m_pICivicAddress = pICivicAddress
End Property

' ICivicAddress
Public Property Get Country() As String
    Dim Ret As String
    If m_pICivicAddress <> 0& Then
        Dim hString As Long
        If Invoke(m_pICivicAddress, _
                  ICivicAddress_GetCountry, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Country = Ret
End Property

Public Property Get State() As String
    Dim Ret As String
    If m_pICivicAddress <> 0& Then
        Dim hString As Long
        If Invoke(m_pICivicAddress, _
                  ICivicAddress_GetState, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    State = Ret
End Property

Public Property Get City() As String
    Dim Ret As String
    If m_pICivicAddress <> 0& Then
        Dim hString As Long
        If Invoke(m_pICivicAddress, _
                  ICivicAddress_GetCity, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    City = Ret
End Property

Public Property Get PostalCode() As String
    Dim Ret As String
    If m_pICivicAddress <> 0& Then
        Dim hString As Long
        If Invoke(m_pICivicAddress, _
                  ICivicAddress_GetPostalCode, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    PostalCode = Ret
End Property

Public Property Get Timestamp() As DateTime
    Dim Ret As DateTime
    If m_pICivicAddress <> 0& Then
        Dim Value As Currency
        If Invoke(m_pICivicAddress, _
                  ICivicAddress_GetTimestamp, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = Value
        End If
    End If
    Set Timestamp = Ret
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


