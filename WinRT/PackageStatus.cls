VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackageStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_ApplicationModel_PackageStatus As String = "Windows.ApplicationModel.PackageStatus"
Private Const IID_IPackageStatus As String = "{5fe74f71-a365-4c09-a02d-046d525ea1da}"
Private Const IID_IPackageStatus2 As String = "{f428fa93-7c56-4862-acfa-abaedcc0694d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IPackageStatus
    IPackageStatus_VerifyIsOK = 6
    IPackageStatus_GetNotAvailable = 7
    IPackageStatus_GetPackageOffline = 8
    IPackageStatus_GetDataOffline = 9
    IPackageStatus_GetDisabled = 10
    IPackageStatus_GetNeedsRemediation = 11
    IPackageStatus_GetLicenseIssue = 12
    IPackageStatus_GetModified = 13
    IPackageStatus_GetTampered = 14
    IPackageStatus_GetDependencyIssue = 15
    IPackageStatus_GetServicing = 16
    IPackageStatus_GetDeploymentInProgress = 17
  
    ' IPackageStatus2
    IPackageStatus2_GetIsPartiallyStaged = 6
  
End Enum

' ----==== Variablen ====----
Private m_pIPackageStatus As Long
Private m_pIPackageStatus2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPackageStatus2)
    Call ReleaseIfc(m_pIPackageStatus)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIPackageStatus As Long)
    Call ReleaseIfc(m_pIPackageStatus2)
    Call ReleaseIfc(m_pIPackageStatus)
    m_pIPackageStatus = pIPackageStatus
    Call QueryIfc(m_pIPackageStatus, IID_IPackageStatus2, m_pIPackageStatus2)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIPackageStatus
End Property

' IPackageStatus
Public Property Get NotAvailable() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetNotAvailable, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    NotAvailable = Ret
End Property

Public Property Get PackageOffline() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetPackageOffline, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    PackageOffline = Ret
End Property

Public Property Get DataOffline() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetDataOffline, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    DataOffline = Ret
End Property

Public Property Get Disabled() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetDisabled, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    Disabled = Ret
End Property

Public Property Get NeedsRemediation() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetNeedsRemediation, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    NeedsRemediation = Ret
End Property

Public Property Get LicenseIssue() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetLicenseIssue, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    LicenseIssue = Ret
End Property

Public Property Get Modified() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetModified, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    Modified = Ret
End Property

Public Property Get Tampered() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetTampered, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    Tampered = Ret
End Property

Public Property Get DependencyIssue() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetDependencyIssue, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    DependencyIssue = Ret
End Property

Public Property Get Servicing() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetServicing, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    Servicing = Ret
End Property

Public Property Get DeploymentInProgress() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_GetDeploymentInProgress, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    DeploymentInProgress = Ret
End Property
  
' IPackageStatus2
Public Property Get IsPartiallyStaged() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus2, _
                  IPackageStatus2_GetIsPartiallyStaged, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    IsPartiallyStaged = Ret
End Property

' ----==== Functions ====----
' IPackageStatus
Public Function VerifyIsOK() As Boolean
    Dim Ret As Boolean
    If m_pIPackageStatus <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageStatus, _
                  IPackageStatus_VerifyIsOK, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    VerifyIsOK = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

