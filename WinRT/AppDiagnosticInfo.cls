VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppDiagnosticInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_AppDiagnosticInfo As String = "Windows.System.AppDiagnosticInfo"
Private Const IID_IAppDiagnosticInfo As String = "{e348a69a-8889-4ca3-be07-d5ffff5f0804}"
Private Const IID_IAppDiagnosticInfo2 As String = "{df46fbd7-191a-446c-9473-8fbc2374a354}"
Private Const IID_IAppDiagnosticInfo3 As String = "{c895c63d-dd61-4c65-babd-81a10b4f9815}"
Private Const IID_IAppDiagnosticInfoStatics As String = "{ce6925bf-10ca-40c8-a9ca-c5c96501866e}"
Private Const IID_IAppDiagnosticInfoStatics2 As String = "{05b24b86-1000-4c90-bb9f-7235071c50fe}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAppDiagnosticInfoStatics
    IAppDiagnosticInfoStatics_RequestInfoAsync = 6

    ' IAppDiagnosticInfoStatics2
    IAppDiagnosticInfoStatics2_CreateWatcher = 6
    IAppDiagnosticInfoStatics2_RequestAccessAsync = 7
    IAppDiagnosticInfoStatics2_RequestInfoForPackageAsync = 8
    IAppDiagnosticInfoStatics2_RequestInfoForAppAsync = 9
    IAppDiagnosticInfoStatics2_RequestInfoForAppUserModelId = 10
    
    ' IAppDiagnosticInfo
    IAppDiagnosticInfo_GetAppInfo = 6
    
    ' IAppDiagnosticInfo2
    IAppDiagnosticInfo2_GetResourceGroups = 6
    IAppDiagnosticInfo2_CreateResourceGroupWatcher = 7
    
    ' IAppDiagnosticInfo3
    IAppDiagnosticInfo3_LaunchAsync = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIAppDiagnosticInfo As Long
Private m_pIAppDiagnosticInfo2 As Long
Private m_pIAppDiagnosticInfo3 As Long
Private m_IAppDiagnosticInfoStatics As Long
Private m_IAppDiagnosticInfoStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_System_AppDiagnosticInfo, _
                            IID_IAppDiagnosticInfoStatics, _
                            m_IAppDiagnosticInfoStatics) Then
        Call QueryIfc(m_IAppDiagnosticInfoStatics, _
                      IID_IAppDiagnosticInfoStatics2, _
                      m_IAppDiagnosticInfoStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAppDiagnosticInfo3)
    Call ReleaseIfc(m_pIAppDiagnosticInfo2)
    Call ReleaseIfc(m_pIAppDiagnosticInfo)
    Call ReleaseIfc(m_IAppDiagnosticInfoStatics2)
    Call ReleaseIfc(m_IAppDiagnosticInfoStatics)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIAppDiagnosticInfo As Long)
    Call ReleaseIfc(m_pIAppDiagnosticInfo3)
    Call ReleaseIfc(m_pIAppDiagnosticInfo2)
    Call ReleaseIfc(m_pIAppDiagnosticInfo)
    m_pIAppDiagnosticInfo = pIAppDiagnosticInfo
    Call QueryIfc(m_pIAppDiagnosticInfo, IID_IAppDiagnosticInfo2, m_pIAppDiagnosticInfo2)
    Call QueryIfc(m_pIAppDiagnosticInfo, IID_IAppDiagnosticInfo3, m_pIAppDiagnosticInfo3)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIAppDiagnosticInfo
End Property

' IAppDiagnosticInfo
Public Property Get AppInfo() As AppInfo
    Dim Ret As AppInfo
        If m_pIAppDiagnosticInfo <> 0& Then
            Dim pIAppInfo As Long
            If Invoke(m_pIAppDiagnosticInfo, _
                      IAppDiagnosticInfo_GetAppInfo, _
                      VarPtr(pIAppInfo)) = S_OK Then
                If pIAppInfo <> 0& Then
                    Set Ret = New AppInfo
                    Ret.Ifc = pIAppInfo
                End If
            End If
        End If
    Set AppInfo = Ret
End Property

' ----==== Functions ====----
' IAppDiagnosticInfoStatics
Public Function RequestInfoAsync() As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    If m_IAppDiagnosticInfoStatics <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_IAppDiagnosticInfoStatics, _
                  IAppDiagnosticInfoStatics_RequestInfoAsync, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_AppDiagnosticInfo, pIVectorView) 'ReadOnlyList_AppDiagnosticInfo
                End If
            End If
        End If
    End If
    Set RequestInfoAsync = Ret
End Function

' IAppDiagnosticInfoStatics2
Public Function CreateWatcher() As AppDiagnosticInfoWatcher
    Dim Ret As AppDiagnosticInfoWatcher
    If m_IAppDiagnosticInfoStatics2 <> 0& Then
        Dim pIAppDiagnosticInfoWatcher As Long
        If Invoke(m_IAppDiagnosticInfoStatics2, _
                  IAppDiagnosticInfoStatics2_CreateWatcher, _
                  VarPtr(pIAppDiagnosticInfoWatcher)) = S_OK Then
            If pIAppDiagnosticInfoWatcher <> 0& Then
                Set Ret = New AppDiagnosticInfoWatcher
                Ret.Ifc = pIAppDiagnosticInfoWatcher
            End If
        End If
    End If
    Set CreateWatcher = Ret
End Function

Public Function RequestAccessAsync() As DiagnosticAccessStatus
    Dim Ret As DiagnosticAccessStatus
    If m_IAppDiagnosticInfoStatics2 <> 0& Then
        Dim value As Long
        If Invoke(m_IAppDiagnosticInfoStatics2, _
                  IAppDiagnosticInfoStatics2_RequestAccessAsync, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                If Await(value) Then
                    Ret = value
                End If
            End If
        End If
    End If
    RequestAccessAsync = Ret
End Function

Public Function RequestInfoForPackageAsync(ByVal PackageFamilyName As String) As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    If m_IAppDiagnosticInfoStatics2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(PackageFamilyName)
        Dim pIVectorView As Long
        If Invoke(m_IAppDiagnosticInfoStatics2, _
                  IAppDiagnosticInfoStatics2_RequestInfoForPackageAsync, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_AppDiagnosticInfo, pIVectorView) 'ReadOnlyList_AppDiagnosticInfo
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set RequestInfoForPackageAsync = Ret
End Function

Public Function RequestInfoForAppAsync() As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    If m_IAppDiagnosticInfoStatics2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_IAppDiagnosticInfoStatics2, _
                  IAppDiagnosticInfoStatics2_RequestInfoForAppAsync, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_AppDiagnosticInfo, pIVectorView) 'ReadOnlyList_AppDiagnosticInfo
                End If
            End If
        End If
    End If
    Set RequestInfoForAppAsync = Ret
End Function

Public Function RequestInfoForAppUserModelId(ByVal AppUserModelId As String) As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    If m_IAppDiagnosticInfoStatics2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(AppUserModelId)
        Dim pIVectorView As Long
        If Invoke(m_IAppDiagnosticInfoStatics2, _
                  IAppDiagnosticInfoStatics2_RequestInfoForAppUserModelId, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_AppDiagnosticInfo, pIVectorView) 'ReadOnlyList_AppDiagnosticInfo
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set RequestInfoForAppUserModelId = Ret
End Function

' IAppDiagnosticInfo2
Public Function GetResourceGroups() As ReadOnlyList_1 'ReadOnlyList_AppResourceGroupInfo
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AppResourceGroupInfo
    If m_pIAppDiagnosticInfo2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIAppDiagnosticInfo2, _
                  IAppDiagnosticInfo2_GetResourceGroups, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_AppResourceGroupInfo, pIVectorView) 'ReadOnlyList_AppResourceGroupInfo
            End If
        End If
    End If
    Set GetResourceGroups = Ret
End Function

Public Function CreateResourceGroupWatcher() As AppResourceGroupInfoWatcher
    Dim Ret As AppResourceGroupInfoWatcher
        If m_pIAppDiagnosticInfo2 <> 0& Then
            Dim pIAppResourceGroupInfoWatcher As Long
            If Invoke(m_pIAppDiagnosticInfo2, _
                      IAppDiagnosticInfo2_CreateResourceGroupWatcher, _
                      VarPtr(pIAppResourceGroupInfoWatcher)) = S_OK Then
                If pIAppResourceGroupInfoWatcher <> 0& Then
                    Set Ret = New AppResourceGroupInfoWatcher
                    Ret.Ifc = pIAppResourceGroupInfoWatcher
                End If
            End If
        End If
    Set CreateResourceGroupWatcher = Ret
End Function

' IAppDiagnosticInfo3
Public Function LaunchAsync() As AppActivationResult
    Dim Ret As AppActivationResult
        If m_pIAppDiagnosticInfo3 <> 0& Then
            Dim pIAppActivationResult As Long
            If Invoke(m_pIAppDiagnosticInfo3, _
                      IAppDiagnosticInfo3_LaunchAsync, _
                      VarPtr(pIAppActivationResult)) = S_OK Then
                If pIAppActivationResult <> 0& Then
                    If Await(pIAppActivationResult) Then
                        If pIAppActivationResult <> 0& Then
                            Set Ret = New AppActivationResult
                            Ret.Ifc = pIAppActivationResult
                        End If
                    End If
                End If
            End If
        End If
    Set LaunchAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


