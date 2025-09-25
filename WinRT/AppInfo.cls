VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_ApplicationModel_AppInfo As String = "Windows.ApplicationModel.AppInfo"
Private Const IID_IAppInfo As String = "{cf7f59b3-6a09-4de8-a6c0-5792d56880d1}"
Private Const IID_IAppInfo2 As String = "{be4b1f5a-2098-431b-bd25-b30878748d47}"
Private Const IID_IAppInfo3 As String = "{09a78e46-93a4-46de-9397-0843b57115ea}"
Private Const IID_IAppInfo4 As String = "{2f34bdeb-1609-4554-9f33-12e1e803e0d4}"
Private Const IID_IAppInfoStatics As String = "{cf1f782a-e48b-4f0c-9b0b-79c3f8957dd7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAppInfoStatics
    IAppInfoStatics_GetCurrent = 6
    IAppInfoStatics_GetFromAppUserModelId = 7
    IAppInfoStatics_GetFromAppUserModelIdForUser = 8
    
    ' IAppInfo
    IAppInfo_GetId = 6
    IAppInfo_GetAppUserModelId = 7
    IAppInfo_GetDisplayInfo = 8
    IAppInfo_GetPackageFamilyName = 9

    ' IAppInfo2
    IAppInfo2_GetPackage = 6

    ' IAppInfo3
    IAppInfo3_GetExecutionContext = 6

    ' IAppInfo4
    IAppInfo4_GetSupportedFileExtensions = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIAppInfo As Long
Private m_pIAppInfo2 As Long
Private m_pIAppInfo3 As Long
Private m_pIAppInfo4 As Long
Private m_pIAppInfoStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_ApplicationModel_AppInfo, _
                            IID_IAppInfoStatics, _
                            m_pIAppInfoStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAppInfo4)
    Call ReleaseIfc(m_pIAppInfo3)
    Call ReleaseIfc(m_pIAppInfo2)
    Call ReleaseIfc(m_pIAppInfo)
    Call ReleaseIfc(m_pIAppInfoStatics)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIAppInfo As Long)
    Call ReleaseIfc(m_pIAppInfo4)
    Call ReleaseIfc(m_pIAppInfo3)
    Call ReleaseIfc(m_pIAppInfo2)
    Call ReleaseIfc(m_pIAppInfo)
    m_pIAppInfo = pIAppInfo
    Call QueryIfc(m_pIAppInfo, IID_IAppInfo2, m_pIAppInfo2)
    Call QueryIfc(m_pIAppInfo, IID_IAppInfo3, m_pIAppInfo3)
    Call QueryIfc(m_pIAppInfo, IID_IAppInfo4, m_pIAppInfo4)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIAppInfo
End Property

' IAppInfoStatics
Public Property Get Current() As AppInfo
    Dim Ret As AppInfo
    If m_pIAppInfoStatics <> 0& Then
        Dim pIAppInfo As Long
        If Invoke(m_pIAppInfoStatics, _
                  IAppInfoStatics_GetCurrent, _
                  VarPtr(pIAppInfo)) = S_OK Then
            If pIAppInfo <> 0& Then
                Set Ret = New AppInfo
                Ret.Ifc = pIAppInfo
            End If
        End If
    End If
    Set Current = Ret
End Property

' IAppInfo
Public Property Get Id() As String
    Dim Ret As String
    If m_pIAppInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIAppInfo, _
                  IAppInfo_GetId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Id = Ret
End Property

Public Property Get AppUserModelId() As String
    Dim Ret As String
    If m_pIAppInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIAppInfo, _
                  IAppInfo_GetAppUserModelId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AppUserModelId = Ret
End Property

Public Property Get DisplayInfo() As AppDisplayInfo
    Dim Ret As AppDisplayInfo
    If m_pIAppInfo <> 0& Then
        Dim pIAppDisplayInfo As Long
        If Invoke(m_pIAppInfo, _
                  IAppInfo_GetDisplayInfo, _
                  VarPtr(pIAppDisplayInfo)) = S_OK Then
            If pIAppDisplayInfo <> 0& Then
                Set Ret = New AppDisplayInfo
                Ret.Ifc = pIAppDisplayInfo
            End If
        End If
    End If
    Set DisplayInfo = Ret
End Property

Public Property Get PackageFamilyName() As String
    Dim Ret As String
    If m_pIAppInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIAppInfo, _
                  IAppInfo_GetPackageFamilyName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    PackageFamilyName = Ret
End Property

' IAppInfo2
Public Property Get Package() As Package
    Dim Ret As Package
    If m_pIAppInfo2 <> 0& Then
        Dim pIPackage As Long
        If Invoke(m_pIAppInfo2, _
                  IAppInfo2_GetPackage, _
                  VarPtr(pIPackage)) = S_OK Then
            If pIPackage <> 0& Then
                Set Ret = New Package
                Ret.Ifc = pIPackage
            End If
        End If
    End If
    Set Package = Ret
End Property

' IAppInfo3
Public Property Get ExecutionContext() As AppExecutionContext
    Dim Ret As AppExecutionContext
    If m_pIAppInfo3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIAppInfo3, _
                  IAppInfo3_GetExecutionContext, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ExecutionContext = Ret
End Property

' IAppInfo4
Public Property Get SupportedFileExtensions(ByRef valueLength As Long) As String()
    Dim Ret() As String
    If m_pIAppInfo4 <> 0& Then
        Dim value As Long
        If Invoke(m_pIAppInfo4, _
                  IAppInfo4_GetSupportedFileExtensions, _
                  VarPtr(valueLength), _
                  VarPtr(value)) = S_OK Then
            If valueLength > 0& Then
                Dim valueItem As Long
                ReDim Ret(valueLength - 1)
                For valueItem = 0 To valueLength - 1
                    Dim hString As Long
                    Call CopyMemory(hString, ByVal value + (4 * valueItem), 4)
                    Ret(valueItem) = GetWindowsString(hString)
                Next
            End If
        End If
    End If
    SupportedFileExtensions = Ret
End Property

' ----==== Functions ====----
' IAppInfoStatics
Public Function GetFromAppUserModelId(ByVal AppUserModelId As String) As AppInfo
    Dim Ret As AppInfo
    If m_pIAppInfoStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(AppUserModelId)
        Dim pIAppInfo As Long
        If Invoke(m_pIAppInfoStatics, _
                  IAppInfoStatics_GetFromAppUserModelId, _
                  hString, _
                  VarPtr(pIAppInfo)) = S_OK Then
            If pIAppInfo <> 0& Then
                Set Ret = New AppInfo
                Ret.Ifc = pIAppInfo
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetFromAppUserModelId = Ret
End Function

Public Function GetFromAppUserModelIdForUser(ByVal value As User, _
                                             ByVal AppUserModelId As String) As AppInfo
    Dim Ret As AppInfo
    If m_pIAppInfoStatics <> 0& And _
       IsNotNothing(value) Then
        Dim hString As Long
        hString = CreateWindowsString(AppUserModelId)
        Dim pIAppInfo As Long
        If Invoke(m_pIAppInfoStatics, _
                  IAppInfoStatics_GetFromAppUserModelIdForUser, _
                  value.Ifc, _
                  hString, _
                  VarPtr(pIAppInfo)) = S_OK Then
            If pIAppInfo <> 0& Then
                Set Ret = New AppInfo
                Ret.Ifc = pIAppInfo
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetFromAppUserModelIdForUser = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



