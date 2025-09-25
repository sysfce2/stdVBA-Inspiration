VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_ApplicationModel_Package As String = "Windows.ApplicationModel.Package"
Private Const IID_IPackageStatics As String = "{4e534bdf-2960-4878-97a4-9624deb72f2d}"
Private Const IID_IPackage As String = "{163c792f-bd75-413c-bf23-b1fe7b95d825}"
Private Const IID_IPackage2 As String = "{a6612fb6-7688-4ace-95fb-359538e7aa01}"
Private Const IID_IPackage3 As String = "{5f738b61-f86a-4917-93d1-f1ee9d3b35d9}"
Private Const IID_IPackage4 As String = "{65aed1ae-b95b-450c-882b-6255187f397e}"
Private Const IID_IPackage5 As String = "{0e842dd4-d9ac-45ed-9a1e-74ce056b2635}"
Private Const IID_IPackage6 As String = "{8b1ad942-12d7-4754-ae4e-638cbc0e3a2e}"
Private Const IID_IPackage7 As String = "{86ff8d31-a2e4-45e0-9732-283a6d88fde1}"
Private Const IID_IPackage8 As String = "{2c584f7b-ce2a-4be6-a093-77cfbb2a7ea1}"
Private Const IID_IPackage9 As String = "{d5ab224f-d7e1-49ec-90ce-720cdbd02e9c}"
Private Const IID_IPackageWithMetadata As String = "{95949780-1de9-40f2-b452-0de9f1910012}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IPackageStatics
    IPackageStatics_GetCurrent = 6
    
    ' IPackage
    IPackage_GetId = 6
    IPackage_GetInstalledLocation = 7
    IPackage_GetIsFramework = 8
    IPackage_GetDependencies = 9

    ' IPackage2
    IPackage2_GetDisplayName = 6
    IPackage2_GetPublisherDisplayName = 7
    IPackage2_GetDescription = 8
    IPackage2_GetLogo = 9
    IPackage2_GetIsResourcePackage = 10
    IPackage2_GetIsBundle = 11
    IPackage2_GetIsDevelopmentMode = 12

    ' IPackage3
    IPackage3_GetStatus = 6
    IPackage3_GetInstalledDate = 7
'    IPackage3_GetAppListEntriesAsync = 8

    ' IPackage4
    IPackage4_GetSignatureKind = 6
    IPackage4_GetIsOptional = 7
    IPackage4_VerifyContentIntegrityAsync = 8
   
    ' IPackage5
'    IPackage5_GetContentGroupsAsync = 6
'    IPackage5_GetContentGroupAsync = 7
'    IPackage5_StageContentGroupsAsync = 8
'    IPackage5_StageContentGroupsWithPriorityAsync = 9
    IPackage5_SetInUseAsync = 10

    ' IPackage6
'    IPackage6_GetAppInstallerInfo = 6
'    IPackage6_CheckUpdateAvailabilityAsync = 7

    ' IPackage7
    IPackage7_GetMutableLocation = 6
    IPackage7_GetEffectiveLocation = 7

    ' IPackage8
    IPackage8_GetEffectiveExternalLocation = 6
    IPackage8_GetMachineExternalLocation = 7
    IPackage8_GetUserExternalLocation = 8
    IPackage8_GetInstalledPath = 9
    IPackage8_GetMutablePath = 10
    IPackage8_GetEffectivePath = 11
    IPackage8_GetEffectiveExternalPath = 12
    IPackage8_GetMachineExternalPath = 13
    IPackage8_GetUserExternalPath = 14
    IPackage8_GetLogoAsRandomAccessStreamReference = 15
'    IPackage8_GetAppListEntries = 16
    IPackage8_GetIsStub = 17
                    
    ' IPackage9
'    IPackage9_FindRelatedPackages = 6
    IPackage9_GetSourceUriSchemeName = 7

    ' IPackageWithMetadata
    IPackageWithMetadata_GetInstallDate = 6
    IPackageWithMetadata_GetThumbnailToken = 7

End Enum

' ----==== Variablen ====----
Private m_pIPackage As Long
Private m_pIPackage2 As Long
Private m_pIPackage3 As Long
Private m_pIPackage4 As Long
Private m_pIPackage5 As Long
Private m_pIPackage6 As Long
Private m_pIPackage7 As Long
Private m_pIPackage8 As Long
Private m_pIPackage9 As Long
Private m_pIPackageStatics As Long
Private m_pIPackageWithMetadata As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_ApplicationModel_Package, _
                            IID_IPackageStatics, _
                            m_pIPackageStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPackageWithMetadata)
    Call ReleaseIfc(m_pIPackage9)
    Call ReleaseIfc(m_pIPackage8)
    Call ReleaseIfc(m_pIPackage7)
    Call ReleaseIfc(m_pIPackage6)
    Call ReleaseIfc(m_pIPackage5)
    Call ReleaseIfc(m_pIPackage4)
    Call ReleaseIfc(m_pIPackage3)
    Call ReleaseIfc(m_pIPackage2)
    Call ReleaseIfc(m_pIPackage)
    Call ReleaseIfc(m_pIPackageStatics)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIPackage As Long)
    Call ReleaseIfc(m_pIPackageWithMetadata)
    Call ReleaseIfc(m_pIPackage9)
    Call ReleaseIfc(m_pIPackage8)
    Call ReleaseIfc(m_pIPackage7)
    Call ReleaseIfc(m_pIPackage6)
    Call ReleaseIfc(m_pIPackage5)
    Call ReleaseIfc(m_pIPackage4)
    Call ReleaseIfc(m_pIPackage3)
    Call ReleaseIfc(m_pIPackage2)
    Call ReleaseIfc(m_pIPackage)
    m_pIPackage = pIPackage
    Call QueryIfc(m_pIPackage, IID_IPackage2, m_pIPackage2)
    Call QueryIfc(m_pIPackage, IID_IPackage3, m_pIPackage3)
    Call QueryIfc(m_pIPackage, IID_IPackage4, m_pIPackage4)
    Call QueryIfc(m_pIPackage, IID_IPackage5, m_pIPackage5)
    Call QueryIfc(m_pIPackage, IID_IPackage6, m_pIPackage6)
    Call QueryIfc(m_pIPackage, IID_IPackage7, m_pIPackage7)
    Call QueryIfc(m_pIPackage, IID_IPackage8, m_pIPackage8)
    Call QueryIfc(m_pIPackage, IID_IPackage9, m_pIPackage9)
    Call QueryIfc(m_pIPackage, IID_IPackageWithMetadata, m_pIPackageWithMetadata)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIPackage
End Property

' IPackageStatics
Public Property Get Current() As Package
    Dim Ret As Package
    If m_pIPackageStatics <> 0& Then
        Dim pIPackage As Long
        If Invoke(m_pIPackageStatics, _
                  IPackageStatics_GetCurrent, _
                  VarPtr(pIPackage)) = S_OK Then
            If pIPackage <> 0& Then
                Set Ret = New Package
                Ret.Ifc = pIPackage
            End If
        End If
    End If
    Set Current = Ret
End Property

' IPackage
Public Property Get Id() As PackageId
    Dim Ret As PackageId
    If m_pIPackage <> 0& Then
        Dim pIPackageId As Long
        If Invoke(m_pIPackage, _
                  IPackage_GetId, _
                  VarPtr(pIPackageId)) = S_OK Then
            If pIPackageId <> 0& Then
                Set Ret = New PackageId
                Ret.Ifc = pIPackageId
            End If
        End If
    End If
    Set Id = Ret
End Property

Public Property Get InstalledLocation() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIPackage <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIPackage, _
                  IPackage_GetInstalledLocation, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set InstalledLocation = Ret
End Property

Public Property Get IsFramework() As Boolean
    Dim Ret As Boolean
    If m_pIPackage <> 0& Then
        Dim value As Boolean
        If Invoke(m_pIPackage, _
                  IPackage_GetIsFramework, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IsFramework = Ret
End Property

Public Property Get Dependencies() As ReadOnlyList_1 'ReadOnlyList_Package
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_Package
    If m_pIPackage <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIPackage, _
                  IPackage_GetDependencies, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_Package, pIVectorView) 'ReadOnlyList_Package
            End If
        End If
    End If
    Set Dependencies = Ret
End Property

' IPackage2
Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIPackage2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage2, _
                  IPackage2_GetDisplayName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get PublisherDisplayName() As String
    Dim Ret As String
    If m_pIPackage2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage2, _
                  IPackage2_GetPublisherDisplayName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    PublisherDisplayName = Ret
End Property

Public Property Get Description() As String
    Dim Ret As String
    If m_pIPackage2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage2, _
                  IPackage2_GetDescription, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Description = Ret
End Property

Public Property Get Logo() As Uri
    Dim Ret As Uri
    If m_pIPackage2 <> 0& Then
        Dim pIUriRuntimeClass As Long
        If Invoke(m_pIPackage2, _
                  IPackage2_GetLogo, _
                  VarPtr(pIUriRuntimeClass)) = S_OK Then
            If pIUriRuntimeClass <> 0& Then
                Set Ret = New Uri
                Ret.Ifc = pIUriRuntimeClass
            End If
        End If
    End If
    Set Logo = Ret
End Property

Public Property Get IsResourcePackage() As Boolean
    Dim Ret As Boolean
    If m_pIPackage2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage2, _
                  IPackage2_GetIsResourcePackage, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsResourcePackage = Ret
End Property

Public Property Get IsBundle() As Boolean
    Dim Ret As Boolean
    If m_pIPackage2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage2, _
                  IPackage2_GetIsBundle, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsBundle = Ret
End Property

Public Property Get IsDevelopmentMode() As Boolean
    Dim Ret As Boolean
    If m_pIPackage2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage2, _
                  IPackage2_GetIsDevelopmentMode, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsDevelopmentMode = Ret
End Property

' IPackage3
Public Property Get Status() As PackageStatus
    Dim Ret As PackageStatus
    If m_pIPackage3 <> 0& Then
        Dim pIPackageStatus As Long
        If Invoke(m_pIPackage3, _
                  IPackage3_GetStatus, _
                  VarPtr(pIPackageStatus)) = S_OK Then
            If pIPackageStatus <> 0& Then
                Set Ret = New PackageStatus
                Ret.Ifc = pIPackageStatus
            End If
        End If
    End If
    Set Status = Ret
End Property

Public Property Get InstalledDate() As DateTime
    Dim Ret As DateTime
    If m_pIPackage3 <> 0& Then
        Dim value As Currency
        If Invoke(m_pIPackage3, _
                  IPackage3_GetInstalledDate, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set InstalledDate = Ret
End Property

'IPackage4
Public Property Get SignatureKind() As PackageSignatureKind
    Dim Ret As PackageSignatureKind
    If m_pIPackage4 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage4, _
                  IPackage4_GetSignatureKind, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SignatureKind = Ret
End Property

Public Property Get IsOptional() As Boolean
    Dim Ret As Boolean
    If m_pIPackage4 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage4, _
                  IPackage4_GetIsOptional, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsOptional = Ret
End Property

' IPackage7
Public Property Get MutableLocation() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIPackage7 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIPackage7, _
                  IPackage7_GetMutableLocation, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set MutableLocation = Ret
End Property

Public Property Get EffectiveLocation() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIPackage7 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIPackage7, _
                  IPackage7_GetEffectiveLocation, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set EffectiveLocation = Ret
End Property

' IPackage8
Public Property Get EffectiveExternalLocation() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIPackage8 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetEffectiveExternalLocation, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set EffectiveExternalLocation = Ret
End Property

Public Property Get MachineExternalLocation() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIPackage8 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetMachineExternalLocation, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set MachineExternalLocation = Ret
End Property

Public Property Get UserExternalLocation() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIPackage8 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetUserExternalLocation, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set UserExternalLocation = Ret
End Property

Public Property Get InstalledPath() As String
    Dim Ret As String
    If m_pIPackage8 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetInstalledPath, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    InstalledPath = Ret
End Property

Public Property Get MutablePath() As String
    Dim Ret As String
    If m_pIPackage8 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetMutablePath, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    MutablePath = Ret
End Property

Public Property Get EffectivePath() As String
    Dim Ret As String
    If m_pIPackage8 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetEffectivePath, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    EffectivePath = Ret
End Property

Public Property Get EffectiveExternalPath() As String
    Dim Ret As String
    If m_pIPackage8 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetEffectiveExternalPath, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    EffectiveExternalPath = Ret
End Property

Public Property Get MachineExternalPath() As String
    Dim Ret As String
    If m_pIPackage8 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetMachineExternalPath, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    MachineExternalPath = Ret
End Property

Public Property Get UserExternalPath() As String
    Dim Ret As String
    If m_pIPackage8 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetUserExternalPath, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    UserExternalPath = Ret
End Property

Public Property Get IsStub() As Boolean
    Dim Ret As Boolean
    If m_pIPackage8 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetIsStub, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsStub = Ret
End Property

' IPackage9
Public Property Get SourceUriSchemeName() As String
    Dim Ret As String
    If m_pIPackage9 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage9, _
                  IPackage9_GetSourceUriSchemeName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SourceUriSchemeName = Ret
End Property

' IPackageWithMetadata
Public Property Get InstallDate() As DateTime
    Dim Ret As DateTime
    If m_pIPackageWithMetadata <> 0& Then
        Dim value As Currency
        If Invoke(m_pIPackageWithMetadata, _
                  IPackageWithMetadata_GetInstallDate, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set InstallDate = Ret
End Property

' ----==== Functions ====----
' IPackage3
'    IPackage3_GetAppListEntriesAsync = 8

'IPackage4
Public Function VerifyContentIntegrityAsync() As Boolean
    Dim Ret As Boolean
    If m_pIPackage4 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage4, _
                  IPackage4_VerifyContentIntegrityAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Ret = CBool(value)
                End If
            End If
        End If
    End If
    VerifyContentIntegrityAsync = Ret
End Function

' IPackage5
'    IPackage5_GetContentGroupsAsync = 6
'    IPackage5_GetContentGroupAsync = 7
'    IPackage5_StageContentGroupsAsync = 8
'    IPackage5_StageContentGroupsWithPriorityAsync = 9

Public Function SetInUseAsync() As Boolean
    Dim Ret As Boolean
    If m_pIPackage5 <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackage5, _
                  IPackage5_SetInUseAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Ret = CBool(value)
                End If
            End If
        End If
    End If
    SetInUseAsync = Ret
End Function

' IPackage6
'    IPackage6_GetAppInstallerInfo = 6
'    IPackage6_CheckUpdateAvailabilityAsync = 7

' IPackage8
Public Function GetLogoAsRandomAccessStreamReference(ByVal logoSize As Size) As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIPackage8 <> 0& And _
       IsNotNothing(logoSize) Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIPackage8, _
                  IPackage8_GetLogoAsRandomAccessStreamReference, _
                  logoSize.Width, _
                  logoSize.Height, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                Set Ret = New RandomAccessStreamReference
                Ret.Ifc = pIRandomAccessStreamReference
            End If
        End If
    End If
    Set GetLogoAsRandomAccessStreamReference = Ret
End Function

'    IPackage8_GetAppListEntries = 16

' IPackage9
'    IPackage9_FindRelatedPackages = 6

' IPackageWithMetadata
Public Function GetThumbnailToken() As String
    Dim Ret As String
    If m_pIPackageWithMetadata <> 0& Then
        Dim value As Long
        If Invoke(m_pIPackageWithMetadata, _
                  IPackageWithMetadata_GetThumbnailToken, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GetThumbnailToken = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


