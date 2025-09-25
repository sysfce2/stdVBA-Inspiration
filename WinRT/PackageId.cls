VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackageId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 01/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_ApplicationModel_PackageId As String = "Windows.ApplicationModel.PackageId"
Private Const IID_IPackageId As String = "{1adb665e-37c7-4790-9980-dd7ae74e8bb2}"
Private Const IID_IPackageIdWithMetadata As String = "{40577a7c-0c9e-443d-9074-855f5ce0a08d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IPackageId
    IPackageId_GetName = 6
    IPackageId_GetVersion = 7
    IPackageId_GetArchitecture = 8
    IPackageId_GetResourceId = 9
    IPackageId_GetPublisher = 10
    IPackageId_GetPublisherId = 11
    IPackageId_GetFullName = 12
    IPackageId_GetFamilyName = 13
    
    ' IPackageIdWithMetadata
    IPackageIdWithMetadata_GetProductId = 6
    IPackageIdWithMetadata_GetAuthor = 7

End Enum

Private Type PackageVersion_
    Major As Integer
    Minor As Integer
    Build As Integer
    Revision As Integer
End Type

' ----==== Variablen ====----
Private m_pIPackageId As Long
Private m_pIPackageIdWithMetadata As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPackageIdWithMetadata)
    Call ReleaseIfc(m_pIPackageId)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIPackageId As Long)
    Call ReleaseIfc(m_pIPackageIdWithMetadata)
    Call ReleaseIfc(m_pIPackageId)
    m_pIPackageId = pIPackageId
    Call QueryIfc(m_pIPackageId, _
                  IID_IPackageIdWithMetadata, _
                  m_pIPackageIdWithMetadata)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIPackageId
End Property

' IPackageId
Public Property Get Name() As String
    Dim Ret As String
    If m_pIPackageId <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageId, _
                  IPackageId_GetName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Name = Ret
End Property

Public Property Get Version() As PackageVersion
    Dim Ret As PackageVersion
    If m_pIPackageId <> 0& Then
        Dim tPackageVersion As PackageVersion_
        If Invoke(m_pIPackageId, _
                  IPackageId_GetVersion, _
                  VarPtr(tPackageVersion)) = S_OK Then
            Set Ret = New PackageVersion
            Ret.Major = tPackageVersion.Major
            Ret.Minor = tPackageVersion.Minor
            Ret.Build = tPackageVersion.Build
            Ret.Revision = tPackageVersion.Revision
        End If
    End If
    Set Version = Ret
End Property

Public Property Get Architecture() As ProcessorArchitecture
    Dim Ret As ProcessorArchitecture
    If m_pIPackageId <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageId, _
                  IPackageId_GetArchitecture, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Architecture = Ret
End Property

Public Property Get ResourceId() As String
    Dim Ret As String
    If m_pIPackageId <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageId, _
                  IPackageId_GetResourceId, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    ResourceId = Ret
End Property

Public Property Get Publisher() As String
    Dim Ret As String
    If m_pIPackageId <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageId, _
                  IPackageId_GetPublisher, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Publisher = Ret
End Property

Public Property Get PublisherId() As String
    Dim Ret As String
    If m_pIPackageId <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageId, _
                  IPackageId_GetPublisherId, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    PublisherId = Ret
End Property

Public Property Get FullName() As String
    Dim Ret As String
    If m_pIPackageId <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageId, _
                  IPackageId_GetFullName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    FullName = Ret
End Property

Public Property Get FamilyName() As String
    Dim Ret As String
    If m_pIPackageId <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageId, _
                  IPackageId_GetFamilyName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    FamilyName = Ret
End Property
    
' IPackageIdWithMetadata
Public Property Get ProductId() As String
    Dim Ret As String
    If m_pIPackageIdWithMetadata <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageIdWithMetadata, _
                  IPackageIdWithMetadata_GetProductId, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    ProductId = Ret
End Property

Public Property Get Author() As String
    Dim Ret As String
    If m_pIPackageIdWithMetadata <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPackageIdWithMetadata, _
                  IPackageIdWithMetadata_GetAuthor, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Author = Ret
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


