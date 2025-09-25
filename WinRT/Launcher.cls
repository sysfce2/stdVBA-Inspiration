VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Launcher As String = "Windows.System.Launcher"
Private Const IID_ILauncherStatics As String = "{277151c3-9e3e-42f6-91a4-5dfdeb232451}"
Private Const IID_ILauncherStatics2 As String = "{59ba2fbb-24cb-4c02-a4c4-8294569d54f1}"
Private Const IID_ILauncherStatics3 As String = "{234261a8-9db3-4683-aa42-dc6f51d33847}"
Private Const IID_ILauncherStatics4 As String = "{b9ec819f-b5a5-41c6-b3b3-dd1b3178bcf2}"
Private Const IID_ILauncherStatics5 As String = "{5b24ef84-d895-5fea-9153-1ac49aed9ba9}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ILauncherStatics
    ILauncherStatics_LaunchFileAsync = 6
    ILauncherStatics_LaunchFileWithOptionsAsync = 7
    ILauncherStatics_LaunchUriAsync = 8
    ILauncherStatics_LaunchUriWithOptionsAsync = 9
                    
    ' ILauncherStatics2
    ILauncherStatics2_LaunchUriForResultsAsync = 6
'    ILauncherStatics2_LaunchUriForResultsWithDataAsync = 7
'    ILauncherStatics2_LaunchUriWithDataAsync = 8
'    ILauncherStatics2_QueryUriSupportAsync = 9
'    ILauncherStatics2_QueryUriSupportWithPackageFamilyNameAsync = 10
'    ILauncherStatics2_QueryFileSupportAsync = 11
'    ILauncherStatics2_QueryFileSupportWithPackageFamilyNameAsync = 12
'    ILauncherStatics2_FindUriSchemeHandlersAsync = 13
'    ILauncherStatics2_FindUriSchemeHandlersWithLaunchUriTypeAsync = 14
'    ILauncherStatics2_FindFileHandlersAsync = 15
                    
    ' ILauncherStatics3
    ILauncherStatics3_LaunchFolderAsync = 6
    ILauncherStatics3_LaunchFolderWithOptionsAsync = 7
                    
    ' ILauncherStatics4
'    ILauncherStatics4_QueryAppUriSupportAsync = 6
'    ILauncherStatics4_QueryAppUriSupportWithPackageFamilyNameAsync = 7
'    ILauncherStatics4_FindAppUriHandlersAsync = 8
'    ILauncherStatics4_LaunchUriForUserAsync = 9
'    ILauncherStatics4_LaunchUriWithOptionsForUserAsync = 10
'    ILauncherStatics4_LaunchUriWithDataForUserAsync = 11
'    ILauncherStatics4_LaunchUriForResultsForUserAsync = 12
'    ILauncherStatics4_LaunchUriForResultsWithDataForUserAsync = 13
                    
    ' ILauncherStatics5
    ILauncherStatics5_LaunchFolderPathAsync = 6
    ILauncherStatics5_LaunchFolderPathWithOptionsAsync = 7
'    ILauncherStatics5_LaunchFolderPathForUserAsync = 8
'    ILauncherStatics5_LaunchFolderPathWithOptionsForUserAsync = 9
                    
End Enum

' ----==== Variablen ====----
Private m_pILauncherStatics As Long
Private m_pILauncherStatics2 As Long
Private m_pILauncherStatics3 As Long
Private m_pILauncherStatics4 As Long
Private m_pILauncherStatics5 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_System_Launcher, _
                            IID_ILauncherStatics, _
                            m_pILauncherStatics) Then
        Call QueryIfc(m_pILauncherStatics, IID_ILauncherStatics2, m_pILauncherStatics2)
        Call QueryIfc(m_pILauncherStatics, IID_ILauncherStatics3, m_pILauncherStatics3)
        Call QueryIfc(m_pILauncherStatics, IID_ILauncherStatics4, m_pILauncherStatics4)
        Call QueryIfc(m_pILauncherStatics, IID_ILauncherStatics5, m_pILauncherStatics5)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILauncherStatics5)
    Call ReleaseIfc(m_pILauncherStatics4)
    Call ReleaseIfc(m_pILauncherStatics3)
    Call ReleaseIfc(m_pILauncherStatics2)
    Call ReleaseIfc(m_pILauncherStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pILauncherStatics
End Property

' ----==== Functions ====----
' ILauncherStatics
Public Function LaunchFileAsync(ByVal file As StorageFile) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics <> 0& And _
       IsNotNothing(file) Then
        Dim operation As Long
        If Invoke(m_pILauncherStatics, _
                  ILauncherStatics_LaunchFileAsync, _
                  file.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
    End If
    LaunchFileAsync = Ret
End Function

Public Function LaunchFileWithOptionsAsync(ByVal file As StorageFile, _
                                           ByVal options As LauncherOptions) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics <> 0& And _
       IsNotNothing(file) And _
       IsNotNothing(options) Then
        Dim operation As Long
        If Invoke(m_pILauncherStatics, _
                  ILauncherStatics_LaunchFileWithOptionsAsync, _
                  file.Ifc, _
                  options.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
    End If
    LaunchFileWithOptionsAsync = Ret
End Function

Public Function LaunchUriAsync(ByVal Uri As Uri) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics <> 0& And _
       IsNotNothing(Uri) Then
        Dim operation As Long
        If Invoke(m_pILauncherStatics, _
                  ILauncherStatics_LaunchUriAsync, _
                  Uri.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
    End If
    LaunchUriAsync = Ret
End Function

Public Function LaunchUriWithOptionsAsync(ByVal Uri As Uri, _
                                          ByVal options As LauncherOptions) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics <> 0& And _
       IsNotNothing(Uri) And _
       IsNotNothing(options) Then
        Dim operation As Long
        If Invoke(m_pILauncherStatics, _
                  ILauncherStatics_LaunchUriWithOptionsAsync, _
                  Uri.Ifc, _
                  options.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
    End If
    LaunchUriWithOptionsAsync = Ret
End Function
                    
' ILauncherStatics2
Public Function LaunchUriForResultsAsync(ByVal Uri As Uri, _
                                         ByVal options As LauncherOptions) As LaunchUriResult
    Dim Ret As LaunchUriResult
    If m_pILauncherStatics2 <> 0& And _
       IsNotNothing(Uri) And _
       IsNotNothing(options) Then
        Dim operation As Long
        If Invoke(m_pILauncherStatics2, _
                  ILauncherStatics2_LaunchUriForResultsAsync, _
                  Uri.Ifc, _
                  options.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New LaunchUriResult
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set LaunchUriForResultsAsync = Ret
End Function

'    ILauncherStatics2_LaunchUriForResultsWithDataAsync = 7
'    ILauncherStatics2_LaunchUriWithDataAsync = 8
'    ILauncherStatics2_QueryUriSupportAsync = 9
'    ILauncherStatics2_QueryUriSupportWithPackageFamilyNameAsync = 10
'    ILauncherStatics2_QueryFileSupportAsync = 11
'    ILauncherStatics2_QueryFileSupportWithPackageFamilyNameAsync = 12
'    ILauncherStatics2_FindUriSchemeHandlersAsync = 13
'    ILauncherStatics2_FindUriSchemeHandlersWithLaunchUriTypeAsync = 14
'    ILauncherStatics2_FindFileHandlersAsync = 15
                    
' ILauncherStatics3
Public Function LaunchFolderAsync(ByVal Folder As StorageFolder) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics3 <> 0& And _
       IsNotNothing(Folder) Then
        Dim operation As Long
        If Invoke(m_pILauncherStatics3, _
                  ILauncherStatics3_LaunchFolderAsync, _
                  Folder.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
    End If
    LaunchFolderAsync = Ret
End Function

Public Function LaunchFolderWithOptionsAsync(ByVal Folder As StorageFolder, _
                                             ByVal options As FolderLauncherOptions) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics3 <> 0& And _
       IsNotNothing(Folder) And _
       IsNotNothing(options) Then
        Dim operation As Long
        If Invoke(m_pILauncherStatics3, _
                  ILauncherStatics3_LaunchFolderWithOptionsAsync, _
                  Folder.Ifc, _
                  options.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
    End If
    LaunchFolderWithOptionsAsync = Ret
End Function
                    
' ILauncherStatics4
'    ILauncherStatics4_QueryAppUriSupportAsync = 6
'    ILauncherStatics4_QueryAppUriSupportWithPackageFamilyNameAsync = 7
'    ILauncherStatics4_FindAppUriHandlersAsync = 8
'    ILauncherStatics4_LaunchUriForUserAsync = 9
'    ILauncherStatics4_LaunchUriWithOptionsForUserAsync = 10
'    ILauncherStatics4_LaunchUriWithDataForUserAsync = 11
'    ILauncherStatics4_LaunchUriForResultsForUserAsync = 12
'    ILauncherStatics4_LaunchUriForResultsWithDataForUserAsync = 13
                    
' ILauncherStatics5
Public Function LaunchFolderPathAsync(ByVal Path As String) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics5 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Path)
        Dim operation As Long
        If Invoke(m_pILauncherStatics5, _
                  ILauncherStatics5_LaunchFolderPathAsync, _
                  hString, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    LaunchFolderPathAsync = Ret
End Function

Public Function LaunchFolderPathWithOptionsAsync(ByVal Path As String, _
                                                 ByVal options As FolderLauncherOptions) As Boolean
    Dim Ret As Boolean
    If m_pILauncherStatics5 <> 0& And _
       IsNotNothing(options) Then
        Dim hString As Long
        hString = CreateWindowsString(Path)
        Dim operation As Long
        If Invoke(m_pILauncherStatics5, _
                  ILauncherStatics5_LaunchFolderPathWithOptionsAsync, _
                  hString, _
                  options.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If Await(operation) Then
                Ret = CBool(operation)
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    LaunchFolderPathWithOptionsAsync = Ret
End Function

'    ILauncherStatics5_LaunchFolderPathForUserAsync = 8
'    ILauncherStatics5_LaunchFolderPathWithOptionsForUserAsync = 9

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


