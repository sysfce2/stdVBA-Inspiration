VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LauncherOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_LauncherOptions As String = "Windows.System.LauncherOptions"
Private Const IID_ILauncherOptions As String = "{bafa21d8-b071-4cd8-853e-341203e557d3}"
Private Const IID_ILauncherOptions2 As String = "{3ba08eb4-6e40-4dce-a1a3-2f53950afb49}"
Private Const IID_ILauncherOptions3 As String = "{f0770655-4b63-4e3a-9107-4e687841923a}"
Private Const IID_ILauncherOptions4 As String = "{ef6fd10e-e6fb-4814-a44e-57e8b9d9a01b}"
Private Const IID_ILauncherViewOptions As String = "{8a9b29f1-7ca7-49de-9bd3-3c5b7184f616}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ILauncherOptions
    ILauncherOptions_GetTreatAsUntrusted = 6
    ILauncherOptions_PutTreatAsUntrusted = 7
    ILauncherOptions_GetDisplayApplicationPicker = 8
    ILauncherOptions_PutDisplayApplicationPicker = 9
    ILauncherOptions_GetUI = 10
    ILauncherOptions_GetPreferredApplicationPackageFamilyName = 11
    ILauncherOptions_PutPreferredApplicationPackageFamilyName = 12
    ILauncherOptions_GetPreferredApplicationDisplayName = 13
    ILauncherOptions_PutPreferredApplicationDisplayName = 14
    ILauncherOptions_GetFallbackUri = 15
    ILauncherOptions_PutFallbackUri = 16
    ILauncherOptions_GetContentType = 17
    ILauncherOptions_PutContentType = 18
                    
    ' ILauncherOptions2
    ILauncherOptions2_GetTargetApplicationPackageFamilyName = 6
    ILauncherOptions2_PutTargetApplicationPackageFamilyName = 7
    ILauncherOptions2_GetNeighboringFilesQuery = 8
    ILauncherOptions2_PutNeighboringFilesQuery = 9
     
    ' ILauncherOptions3
    ILauncherOptions3_GetIgnoreAppUriHandlers = 6
    ILauncherOptions3_PutIgnoreAppUriHandlers = 7

    ' ILauncherOptions4
    ILauncherOptions4_GetLimitPickerToCurrentAppAndAppUriHandlers = 6
    ILauncherOptions4_PutLimitPickerToCurrentAppAndAppUriHandlers = 7
    
    ' ILauncherViewOptions
    ILauncherViewOptions_GetDesiredRemainingView = 6
    ILauncherViewOptions_PutDesiredRemainingView = 7
    
End Enum

' ----==== Variablen ====----
Private m_pILauncherOptions As Long
Private m_pILauncherOptions2 As Long
Private m_pILauncherOptions3 As Long
Private m_pILauncherOptions4 As Long
Private m_pILauncherViewOptions As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_System_LauncherOptions, _
                           IID_ILauncherOptions, _
                           m_pILauncherOptions) Then
        Call QueryIfc(m_pILauncherOptions, _
                      IID_ILauncherOptions2, _
                      m_pILauncherOptions2)
        Call QueryIfc(m_pILauncherOptions, _
                      IID_ILauncherOptions3, _
                      m_pILauncherOptions3)
        Call QueryIfc(m_pILauncherOptions, _
                      IID_ILauncherOptions4, _
                      m_pILauncherOptions4)
        Call QueryIfc(m_pILauncherOptions, _
                      IID_ILauncherViewOptions, _
                      m_pILauncherViewOptions)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILauncherViewOptions)
    Call ReleaseIfc(m_pILauncherOptions4)
    Call ReleaseIfc(m_pILauncherOptions3)
    Call ReleaseIfc(m_pILauncherOptions2)
    Call ReleaseIfc(m_pILauncherOptions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pILauncherOptions
End Property

' ILauncherOptions
Public Property Get TreatAsUntrusted() As Boolean
    Dim Ret As Boolean
    If m_pILauncherOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_GetTreatAsUntrusted, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    TreatAsUntrusted = Ret
End Property

Public Property Let TreatAsUntrusted(ByVal value As Boolean)
    If m_pILauncherOptions <> 0& Then
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_PutTreatAsUntrusted, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get DisplayApplicationPicker() As Boolean
    Dim Ret As Boolean
    If m_pILauncherOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_GetDisplayApplicationPicker, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    DisplayApplicationPicker = Ret
End Property

Public Property Let DisplayApplicationPicker(ByVal value As Boolean)
    If m_pILauncherOptions <> 0& Then
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_PutDisplayApplicationPicker, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get UI() As LauncherUIOptions
    Dim Ret As LauncherUIOptions
    If m_pILauncherOptions <> 0& Then
        Dim pILauncherUIOptions As Long
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_GetUI, _
                  VarPtr(pILauncherUIOptions)) = S_OK Then
            If pILauncherUIOptions <> 0& Then
                Set Ret = New LauncherUIOptions
                Ret.Ifc = pILauncherUIOptions
            End If
        End If
    End If
    Set UI = Ret
End Property

Public Property Get PreferredApplicationPackageFamilyName() As String
    Dim Ret As String
    If m_pILauncherOptions <> 0& Then
        Dim hString As Long
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_GetPreferredApplicationPackageFamilyName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    PreferredApplicationPackageFamilyName = Ret
End Property

Public Property Let PreferredApplicationPackageFamilyName(ByVal value As String)
    If m_pILauncherOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_PutPreferredApplicationPackageFamilyName, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get PreferredApplicationDisplayName() As String
    Dim Ret As String
    If m_pILauncherOptions <> 0& Then
        Dim hString As Long
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_GetPreferredApplicationDisplayName, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    PreferredApplicationDisplayName = Ret
End Property

Public Property Let PreferredApplicationDisplayName(ByVal value As String)
    If m_pILauncherOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_PutPreferredApplicationDisplayName, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get FallbackUri() As Uri
    Dim Ret As New Uri
    If m_pILauncherOptions <> 0& Then
        Dim pIUriRuntimeClass As Long
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_GetFallbackUri, _
                  VarPtr(pIUriRuntimeClass)) = S_OK Then
            If pIUriRuntimeClass <> 0& Then
                Set Ret = New Uri
                Ret.Ifc = pIUriRuntimeClass
            End If
        End If
    End If
    Set FallbackUri = Ret
End Property

Public Property Let FallbackUri(ByVal value As Uri)
    If m_pILauncherOptions <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_PutFallbackUri, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get ContentType() As String
    Dim Ret As String
    If m_pILauncherOptions <> 0& Then
        Dim hString As Long
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_GetContentType, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    ContentType = Ret
End Property

Public Property Let ContentType(ByVal value As String)
    If m_pILauncherOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pILauncherOptions, _
                  ILauncherOptions_PutContentType, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property
                    
' ILauncherOptions2
Public Property Get TargetApplicationPackageFamilyName() As String
    Dim Ret As String
    If m_pILauncherOptions2 <> 0& Then
        Dim value As Long
        If Invoke(m_pILauncherOptions2, _
                  ILauncherOptions2_GetTargetApplicationPackageFamilyName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    TargetApplicationPackageFamilyName = Ret
End Property

Public Property Let TargetApplicationPackageFamilyName(ByVal value As String)
    If m_pILauncherOptions2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If hString <> 0& Then
            If Invoke(m_pILauncherOptions2, _
                      ILauncherOptions2_PutTargetApplicationPackageFamilyName, _
                      hString) = S_OK Then
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
End Property

Public Property Get NeighboringFilesQuery() As StorageFileQueryResult
    Dim Ret As StorageFileQueryResult
    If m_pILauncherOptions2 <> 0& Then
        Dim pIStorageFileQueryResult As Long
        If Invoke(m_pILauncherOptions2, _
                  ILauncherOptions2_GetNeighboringFilesQuery, _
                  VarPtr(pIStorageFileQueryResult)) = S_OK Then
            If pIStorageFileQueryResult <> 0& Then
                Set Ret = New StorageFileQueryResult
                Ret.Ifc = pIStorageFileQueryResult
            End If
        End If
    End If
    Set NeighboringFilesQuery = Ret
End Property

Public Property Let NeighboringFilesQuery(ByVal value As StorageFileQueryResult)
    If m_pILauncherOptions2 <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pILauncherOptions2, _
                  ILauncherOptions2_PutNeighboringFilesQuery, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property
     
' ILauncherOptions3
Public Property Get IgnoreAppUriHandlers() As Boolean
    Dim Ret As Boolean
    If m_pILauncherOptions3 <> 0& Then
        Dim value As Long
        If Invoke(m_pILauncherOptions3, _
                  ILauncherOptions3_GetIgnoreAppUriHandlers, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IgnoreAppUriHandlers = Ret
End Property

Public Property Let IgnoreAppUriHandlers(ByVal value As Boolean)
    If m_pILauncherOptions3 <> 0& Then
        If Invoke(m_pILauncherOptions3, _
                  ILauncherOptions3_PutIgnoreAppUriHandlers, _
                  value) = S_OK Then
            End If
    End If
End Property

' ILauncherOptions4
Public Property Get LimitPickerToCurrentAppAndAppUriHandlers() As Boolean
    Dim Ret As Boolean
    If m_pILauncherOptions4 <> 0& Then
        Dim value As Long
        If Invoke(m_pILauncherOptions4, _
                  ILauncherOptions4_GetLimitPickerToCurrentAppAndAppUriHandlers, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IgnoreAppUriHandlers = Ret
End Property

Public Property Let LimitPickerToCurrentAppAndAppUriHandlers(ByVal value As Boolean)
    If m_pILauncherOptions4 <> 0& Then
        If Invoke(m_pILauncherOptions4, _
                  ILauncherOptions4_PutLimitPickerToCurrentAppAndAppUriHandlers, _
                  value) = S_OK Then
        End If
    End If
End Property
    
' ILauncherViewOptions
Public Property Get DesiredRemainingView() As ViewSizePreference
    Dim Ret As ViewSizePreference
    If m_pILauncherViewOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pILauncherViewOptions, _
                  ILauncherViewOptions_GetDesiredRemainingView, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DesiredRemainingView = Ret
End Property
    
Public Property Let DesiredRemainingView(ByVal value As ViewSizePreference)
    If m_pILauncherViewOptions <> 0& Then
        If Invoke(m_pILauncherViewOptions, _
                  ILauncherViewOptions_PutDesiredRemainingView, _
                  value) = S_OK Then
        End If
    End If
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pILauncherOptions <> 0& Then
        If InitWithWindow(m_pILauncherOptions, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Function ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


