VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDataPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_UserDataPaths As String = "Windows.Storage.UserDataPaths"
Private Const IID_IUserDataPaths As String = "{f9c53912-abc4-46ff-8a2b-dc9d7fa6e52f}"
Private Const IID_IUserDataPathsInterop As String = "{f12ef0b5-befb-46c9-9ee5-8a0681f13f8c}"
Private Const IID_IUserDataPathsStatics As String = "{01b29def-e062-48a1-8b0c-f2c7a9ca56c0}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IUserDataPathsStatics
    IUserDataPathsStatics_GetForUser = 6
    IUserDataPathsStatics_GetDefault = 7
    
    ' IUserDataPaths
    IUserDataPaths_GetCameraRoll = 6
    IUserDataPaths_GetCookies = 7
    IUserDataPaths_GetDesktop = 8
    IUserDataPaths_GetDocuments = 9
    IUserDataPaths_GetDownloads = 10
    IUserDataPaths_GetFavorites = 11
    IUserDataPaths_GetHistory = 12
    IUserDataPaths_GetInternetCache = 13
    IUserDataPaths_GetLocalAppData = 14
    IUserDataPaths_GetLocalAppDataLow = 15
    IUserDataPaths_GetMusic = 16
    IUserDataPaths_GetPictures = 17
    IUserDataPaths_GetProfile = 18
    IUserDataPaths_GetRecent = 19
    IUserDataPaths_GetRoamingAppData = 20
    IUserDataPaths_GetSavedPictures = 21
    IUserDataPaths_GetScreenshots = 22
    IUserDataPaths_GetTemplates = 23
    IUserDataPaths_GetVideos = 24

    ' IUserDataPathsInterop
    IUserDataPathsInterop_GetAdminTools = 3
    IUserDataPathsInterop_GetCDBurning = 4
    IUserDataPathsInterop_GetNetworkShortcuts = 5
    IUserDataPathsInterop_GetPrograms = 6
    IUserDataPathsInterop_GetSendTo = 7
    IUserDataPathsInterop_GetStartMenu = 8
    IUserDataPathsInterop_GetStartup = 9
    
End Enum

' ----==== Variablen ====----
Private m_pIUserDataPaths As Long
Private m_pIUserDataPathsInterop As Long
Private m_pIUserDataPathsStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_UserDataPaths, _
                            IID_IUserDataPathsStatics, _
                            m_pIUserDataPathsStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUserDataPaths)
    Call ReleaseIfc(m_pIUserDataPathsInterop)
    Call ReleaseIfc(m_pIUserDataPathsStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUserDataPaths
End Property

Public Property Let Ifc(ByVal pIUserDataPaths As Long)
    Call ReleaseIfc(m_pIUserDataPaths)
    m_pIUserDataPaths = pIUserDataPaths
    Call QueryIfc(m_pIUserDataPaths, IID_IUserDataPathsInterop, m_pIUserDataPathsInterop)
End Property

' IUserDataPaths
Public Property Get CameraRoll() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetCameraRoll, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CameraRoll = Ret
End Property

Public Property Get Cookies() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetCookies, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Cookies = Ret
End Property

Public Property Get Desktop() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetDesktop, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Desktop = Ret
End Property

Public Property Get Documents() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetDocuments, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Documents = Ret
End Property

Public Property Get Downloads() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetDownloads, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Downloads = Ret
End Property

Public Property Get Favorites() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetFavorites, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Favorites = Ret
End Property

Public Property Get History() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetHistory, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    History = Ret
End Property

Public Property Get InternetCache() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetInternetCache, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    InternetCache = Ret
End Property

Public Property Get LocalAppData() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetLocalAppData, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    LocalAppData = Ret
End Property

Public Property Get LocalAppDataLow() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetLocalAppDataLow, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    LocalAppDataLow = Ret
End Property

Public Property Get Music() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetMusic, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Music = Ret
End Property

Public Property Get Pictures() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetPictures, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Pictures = Ret
End Property

Public Property Get Profile() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetProfile, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Profile = Ret
End Property

Public Property Get Recent() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetRecent, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Recent = Ret
End Property

Public Property Get RoamingAppData() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetRoamingAppData, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    RoamingAppData = Ret
End Property

Public Property Get SavedPictures() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetSavedPictures, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SavedPictures = Ret
End Property

Public Property Get Screenshots() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetScreenshots, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Screenshots = Ret
End Property

Public Property Get Templates() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetTemplates, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Templates = Ret
End Property

Public Property Get Videos() As String
    Dim Ret As String
    If m_pIUserDataPaths <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPaths, _
                  IUserDataPaths_GetVideos, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Videos = Ret
End Property

' IUserDataPathsInterop
Public Property Get AdminTools() As String
    Dim Ret As String
    If m_pIUserDataPathsInterop <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPathsInterop, _
                  IUserDataPathsInterop_GetAdminTools, _
                  VarPtr(value)) = S_OK Then
            Ret = Ptr2Str(value)
        End If
    End If
    AdminTools = Ret
End Property

Public Property Get CDBurning() As String
    Dim Ret As String
    If m_pIUserDataPathsInterop <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPathsInterop, _
                  IUserDataPathsInterop_GetCDBurning, _
                  VarPtr(value)) = S_OK Then
            Ret = Ptr2Str(value)
        End If
    End If
    CDBurning = Ret
End Property

Public Property Get NetworkShortcuts() As String
    Dim Ret As String
    If m_pIUserDataPathsInterop <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPathsInterop, _
                  IUserDataPathsInterop_GetNetworkShortcuts, _
                  VarPtr(value)) = S_OK Then
            Ret = Ptr2Str(value)
        End If
    End If
    NetworkShortcuts = Ret
End Property

Public Property Get Programs() As String
    Dim Ret As String
    If m_pIUserDataPathsInterop <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPathsInterop, _
                  IUserDataPathsInterop_GetPrograms, _
                  VarPtr(value)) = S_OK Then
            Ret = Ptr2Str(value)
        End If
    End If
    Programs = Ret
End Property

Public Property Get SendTo() As String
    Dim Ret As String
    If m_pIUserDataPathsInterop <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPathsInterop, _
                  IUserDataPathsInterop_GetSendTo, _
                  VarPtr(value)) = S_OK Then
            Ret = Ptr2Str(value)
        End If
    End If
    SendTo = Ret
End Property

Public Property Get StartMenu() As String
    Dim Ret As String
    If m_pIUserDataPathsInterop <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPathsInterop, _
                  IUserDataPathsInterop_GetStartMenu, _
                  VarPtr(value)) = S_OK Then
            Ret = Ptr2Str(value)
        End If
    End If
    StartMenu = Ret
End Property

Public Property Get Startup() As String
    Dim Ret As String
    If m_pIUserDataPathsInterop <> 0& Then
        Dim value As Long
        If Invoke(m_pIUserDataPathsInterop, _
                  IUserDataPathsInterop_GetStartup, _
                  VarPtr(value)) = S_OK Then
            Ret = Ptr2Str(value)
        End If
    End If
    Startup = Ret
End Property

' ----==== Functions ====----
' IUserDataPathsStatics
Public Function GetForUser(ByVal User As User) As UserDataPaths
    Dim Ret As UserDataPaths
    If m_pIUserDataPathsStatics <> 0& And _
       IsNotNothing(User) Then
        Dim pIUserDataPaths As Long
        If Invoke(m_pIUserDataPathsStatics, _
                  IUserDataPathsStatics_GetForUser, _
                  User.Ifc, _
                  VarPtr(pIUserDataPaths)) = S_OK Then
            If pIUserDataPaths <> 0& Then
                Set Ret = New UserDataPaths
                Ret.Ifc = pIUserDataPaths
            End If
        End If
    End If
    Set GetForUser = Ret
End Function

Public Function GetDefault() As UserDataPaths
    Dim Ret As UserDataPaths
    If m_pIUserDataPathsStatics <> 0& Then
        Dim pIUserDataPaths As Long
        If Invoke(m_pIUserDataPathsStatics, _
                  IUserDataPathsStatics_GetDefault, _
                  VarPtr(pIUserDataPaths)) = S_OK Then
            If pIUserDataPaths <> 0& Then
                Set Ret = New UserDataPaths
                Ret.Ifc = pIUserDataPaths
            End If
        End If
    End If
    Set GetDefault = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


