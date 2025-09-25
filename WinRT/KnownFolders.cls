VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnownFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_KnownFolders As String = "Windows.Storage.KnownFolders"
Private Const IID_IKnownFoldersStatics As String = "{5a2a7520-4802-452d-9ad9-4351ada7ec35}"
Private Const IID_IKnownFoldersStatics2 As String = "{194bd0cd-cf6e-4d07-9d53-e9163a2536e9}"
Private Const IID_IKnownFoldersStatics3 As String = "{c5194341-9742-4ed5-823d-fc1401148764}"
Private Const IID_IKnownFoldersStatics4 As String = "{1722e6bf-9ff9-4b21-bed5-90ecb13a192e}"
Private Const IID_IKnownFoldersPlaylistsStatics As String = "{dad5ecd6-306f-4d6a-b496-46ba8eb106ce}"
Private Const IID_IKnownFoldersCameraRollStatics As String = "{5d115e66-27e8-492f-b8e5-2f90896cd4cd}"
Private Const IID_IKnownFoldersSavedPicturesStatics As String = "{055c93ea-253d-467c-b6ca-a97da1e9a18d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IKnownFoldersStatics
    IKnownFoldersStatics_GetMusicLibrary = 6
    IKnownFoldersStatics_GetPicturesLibrary = 7
    IKnownFoldersStatics_GetVideosLibrary = 8
    IKnownFoldersStatics_GetDocumentsLibrary = 9
    IKnownFoldersStatics_GetHomeGroup = 10
    IKnownFoldersStatics_GetRemovableDevices = 11
    IKnownFoldersStatics_GetMediaServerDevices = 12
                    
    ' IKnownFoldersStatics2
    IKnownFoldersStatics2_GetObjects3D = 6
    IKnownFoldersStatics2_GetAppCaptures = 7
    IKnownFoldersStatics2_GetRecordedCalls = 8
                
    ' IKnownFoldersStatics3
    IKnownFoldersStatics3_GetFolderForUserAsync = 6
    
    ' IKnownFoldersStatics4
    IKnownFoldersStatics4_RequestAccessAsync = 6
    IKnownFoldersStatics4_RequestAccessForUserAsync = 7
    IKnownFoldersStatics4_GetFolderAsync = 8
    
    ' IKnownFoldersPlaylistsStatics
    IKnownFoldersPlaylistsStatics_GetPlaylists = 6
    
    ' IKnownFoldersCameraRollStatics
    IKnownFoldersCameraRollStatics_GetCameraRoll = 6
    
    ' IKnownFoldersSavedPicturesStatics
    IKnownFoldersSavedPicturesStatics_GetSavedPictures = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIKnownFoldersStatics As Long
Private m_pIKnownFoldersStatics2 As Long
Private m_pIKnownFoldersStatics3 As Long
Private m_pIKnownFoldersStatics4 As Long
Private m_pIKnownFoldersPlaylistsStatics As Long
Private m_pIKnownFoldersCameraRollStatics As Long
Private m_pIKnownFoldersSavedPicturesStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_KnownFolders, IID_IKnownFoldersStatics, m_pIKnownFoldersStatics) Then
        Call QueryIfc(m_pIKnownFoldersStatics, IID_IKnownFoldersStatics2, m_pIKnownFoldersStatics2)
        Call QueryIfc(m_pIKnownFoldersStatics, IID_IKnownFoldersStatics3, m_pIKnownFoldersStatics3)
        Call QueryIfc(m_pIKnownFoldersStatics, IID_IKnownFoldersStatics4, m_pIKnownFoldersStatics4)
        Call QueryIfc(m_pIKnownFoldersStatics, IID_IKnownFoldersPlaylistsStatics, m_pIKnownFoldersPlaylistsStatics)
        Call QueryIfc(m_pIKnownFoldersStatics, IID_IKnownFoldersCameraRollStatics, m_pIKnownFoldersCameraRollStatics)
        Call QueryIfc(m_pIKnownFoldersStatics, IID_IKnownFoldersSavedPicturesStatics, m_pIKnownFoldersSavedPicturesStatics)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIKnownFoldersSavedPicturesStatics)
    Call ReleaseIfc(m_pIKnownFoldersCameraRollStatics)
    Call ReleaseIfc(m_pIKnownFoldersPlaylistsStatics)
    Call ReleaseIfc(m_pIKnownFoldersStatics4)
    Call ReleaseIfc(m_pIKnownFoldersStatics3)
    Call ReleaseIfc(m_pIKnownFoldersStatics2)
    Call ReleaseIfc(m_pIKnownFoldersStatics)
End Sub

' ----==== Properties ====----
' IKnownFoldersStatics
Public Property Get MusicLibrary() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics, _
                  IKnownFoldersStatics_GetMusicLibrary, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set MusicLibrary = Ret
End Property

Public Property Get PicturesLibrary() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics, _
                  IKnownFoldersStatics_GetPicturesLibrary, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set PicturesLibrary = Ret
End Property

Public Property Get VideosLibrary() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics, _
                  IKnownFoldersStatics_GetVideosLibrary, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set VideosLibrary = Ret
End Property

Public Property Get DocumentsLibrary() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics, _
                  IKnownFoldersStatics_GetDocumentsLibrary, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set DocumentsLibrary = Ret
End Property

Public Property Get HomeGroup() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics, _
                  IKnownFoldersStatics_GetHomeGroup, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set HomeGroup = Ret
End Property

Public Property Get RemovableDevices() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics, _
                  IKnownFoldersStatics_GetRemovableDevices, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set RemovableDevices = Ret
End Property

Public Property Get MediaServerDevices() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics, _
                  IKnownFoldersStatics_GetMediaServerDevices, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set MediaServerDevices = Ret
End Property

' IKnownFoldersStatics2
Public Property Get Objects3D() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics2 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics2, _
                  IKnownFoldersStatics2_GetObjects3D, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set Objects3D = Ret
End Property

Public Property Get AppCaptures() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics2 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics2, _
                  IKnownFoldersStatics2_GetAppCaptures, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set AppCaptures = Ret
End Property

Public Property Get RecordedCalls() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics2 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersStatics2, _
                  IKnownFoldersStatics2_GetRecordedCalls, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set RecordedCalls = Ret
End Property

' IKnownFoldersPlaylistsStatics
Public Property Get Playlists() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersPlaylistsStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersPlaylistsStatics, _
                  IKnownFoldersPlaylistsStatics_GetPlaylists, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set Playlists = Ret
End Property

' IKnownFoldersCameraRollStatics
Public Property Get CameraRoll() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersCameraRollStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersCameraRollStatics, _
                  IKnownFoldersCameraRollStatics_GetCameraRoll, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set CameraRoll = Ret
End Property

' IKnownFoldersSavedPicturesStatics
Public Property Get SavedPictures() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersSavedPicturesStatics <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIKnownFoldersSavedPicturesStatics, _
                  IKnownFoldersSavedPicturesStatics_GetSavedPictures, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set SavedPictures = Ret
End Property

' ----==== Functions ====----
' IKnownFoldersStatics3
Public Function GetFolderForUserAsync(ByVal User As User, _
                                      ByVal FolderID As KnownFolderId) As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics3 <> 0& And _
       IsNotNothing(User) Then
        Dim operation As Long
        If Invoke(m_pIKnownFoldersStatics3, _
                  IKnownFoldersStatics3_GetFolderForUserAsync, _
                  User.Ifc, _
                  FolderID, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set GetFolderForUserAsync = Ret
End Function

    ' IKnownFoldersStatics4
Public Function RequestAccessAsync(ByVal FolderID As KnownFolderId) As KnownFoldersAccessStatus
    Dim Ret As KnownFoldersAccessStatus
    If m_pIKnownFoldersStatics4 <> 0& Then
        Dim operation As Long
        If Invoke(m_pIKnownFoldersStatics4, _
                  IKnownFoldersStatics4_RequestAccessAsync, _
                  FolderID, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = operation
                End If
            End If
        End If
    End If
    RequestAccessAsync = Ret
End Function

Public Function RequestAccessForUserAsync(ByVal User As User, _
                                          ByVal FolderID As KnownFolderId) As KnownFoldersAccessStatus
    Dim Ret As KnownFoldersAccessStatus
    If m_pIKnownFoldersStatics4 <> 0& And _
       IsNotNothing(User) Then
        Dim operation As Long
        If Invoke(m_pIKnownFoldersStatics4, _
                  IKnownFoldersStatics4_RequestAccessForUserAsync, _
                  User.Ifc, _
                  FolderID, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = operation
                End If
            End If
        End If
    End If
    RequestAccessForUserAsync = Ret
End Function

Public Function GetFolderAsync(ByVal FolderID As KnownFolderId) As StorageFolder
    Dim Ret As StorageFolder
    If m_pIKnownFoldersStatics4 <> 0& Then
        Dim operation As Long
        If Invoke(m_pIKnownFoldersStatics4, _
                  IKnownFoldersStatics4_GetFolderAsync, _
                  FolderID, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set GetFolderAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

