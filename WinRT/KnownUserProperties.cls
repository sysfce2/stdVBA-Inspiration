VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnownUserProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_KnownUserProperties As String = "Windows.System.KnownUserProperties"
Private Const IID_IKnownUserPropertiesStatics As String = "{7755911a-70c5-48e5-b637-5ba3441e4ee4}"
Private Const IID_IKnownUserPropertiesStatics2 As String = "{5b450782-f620-577e-b1b3-dd56644d79b1}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IKnownUserPropertiesStatics
    IKnownUserPropertiesStatics_GetDisplayName = 6
    IKnownUserPropertiesStatics_GetFirstName = 7
    IKnownUserPropertiesStatics_GetLastName = 8
    IKnownUserPropertiesStatics_GetProviderName = 9
    IKnownUserPropertiesStatics_GetAccountName = 10
    IKnownUserPropertiesStatics_GetGuestHost = 11
    IKnownUserPropertiesStatics_GetPrincipalName = 12
    IKnownUserPropertiesStatics_GetDomainName = 13
    IKnownUserPropertiesStatics_GetSessionInitiationProtocolUri = 14
                    
    ' IKnownUserPropertiesStatics2
    IKnownUserPropertiesStatics2_GetAgeEnforcementRegion = 6

End Enum

' ----==== Variablen ====----
Private m_pIKnownUserPropertiesStatics As Long
Private m_pIKnownUserPropertiesStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_System_KnownUserProperties, _
                            IID_IKnownUserPropertiesStatics, _
                            m_pIKnownUserPropertiesStatics) Then
        Call QueryIfc(m_pIKnownUserPropertiesStatics, _
                      IID_IKnownUserPropertiesStatics2, _
                      m_pIKnownUserPropertiesStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIKnownUserPropertiesStatics2)
    Call ReleaseIfc(m_pIKnownUserPropertiesStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIKnownUserPropertiesStatics
End Property

' IKnownUserPropertiesStatics
Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetDisplayName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get FirstName() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetFirstName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    FirstName = Ret
End Property

Public Property Get LastName() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetLastName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    LastName = Ret
End Property

Public Property Get ProviderName() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetProviderName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    ProviderName = Ret
End Property

Public Property Get AccountName() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetAccountName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    AccountName = Ret
End Property

Public Property Get GuestHost() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetGuestHost, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    GuestHost = Ret
End Property

Public Property Get PrincipalName() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetPrincipalName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    PrincipalName = Ret
End Property

Public Property Get DomainName() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetDomainName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DomainName = Ret
End Property

Public Property Get SessionInitiationProtocolUri() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics, _
                  IKnownUserPropertiesStatics_GetSessionInitiationProtocolUri, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    SessionInitiationProtocolUri = Ret
End Property
                    
' IKnownUserPropertiesStatics2
Public Property Get AgeEnforcementRegion() As String
    Dim Ret As String
    If m_pIKnownUserPropertiesStatics2 <> 0& Then
        Dim hString As Long
        If Invoke(m_pIKnownUserPropertiesStatics2, _
                  IKnownUserPropertiesStatics2_GetAgeEnforcementRegion, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    AgeEnforcementRegion = Ret
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


