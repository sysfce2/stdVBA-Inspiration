VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Storage

Option Explicit

' ----==== Enums Windows.Storage.xxx ====----
Public Enum FileAccessMode
    FileAccessMode_Read = 0
    FileAccessMode_ReadWrite = 1
End Enum

Public Enum StorageOpenOptions
    StorageOpenOptions_None = 0
    StorageOpenOptions_AllowOnlyReaders = 1
    StorageOpenOptions_AllowReadersAndWriters = 2
End Enum

Public Enum NameCollisionOption
    NameCollisionOption_GenerateUniqueName = 0
    NameCollisionOption_ReplaceExisting = 1
    NameCollisionOption_FailIfExists = 2
End Enum

Public Enum StorageDeleteOption
    StorageDeleteOption_Default = 0
    StorageDeleteOption_PermanentDelete = 1
End Enum

Public Enum FileAttributes
    FileAttributes_Normal = &H0&
    FileAttributes_ReadOnly = &H1&
    FileAttributes_Directory = &H10&
    FileAttributes_Archive = &H20&
    FileAttributes_Temporary = &H100&
    FileAttributes_LocallyIncomplete = &H200&
End Enum

Public Enum StorageItemTypes
    StorageItemTypes_None = &H0&
    StorageItemTypes_File = &H1&
    StorageItemTypes_Folder = &H2&
End Enum

Public Enum CreationCollisionOption
    CreationCollisionOption_GenerateUniqueName = 0
    CreationCollisionOption_ReplaceExisting = 1
    CreationCollisionOption_FailIfExists = 2
    CreationCollisionOption_OpenIfExists = 3
End Enum

Public Enum StorageLibraryChangeType
    StorageLibraryChangeType_Created = 0
    StorageLibraryChangeType_Deleted = 1
    StorageLibraryChangeType_MovedOrRenamed = 2
    StorageLibraryChangeType_ContentsChanged = 3
    StorageLibraryChangeType_MovedOutOfLibrary = 4
    StorageLibraryChangeType_MovedIntoLibrary = 5
    StorageLibraryChangeType_ContentsReplaced = 6
    StorageLibraryChangeType_IndexingStatusChanged = 7
    StorageLibraryChangeType_EncryptionChanged = 8
    StorageLibraryChangeType_ChangeTrackingLost = 9
End Enum

Public Enum KnownFolderId
    KnownFolderId_AppCaptures = 0
    KnownFolderId_CameraRoll = 1
    KnownFolderId_DocumentsLibrary = 2
    KnownFolderId_HomeGroup = 3
    KnownFolderId_MediaServerDevices = 4
    KnownFolderId_MusicLibrary = 5
    KnownFolderId_Objects3D = 6
    KnownFolderId_PicturesLibrary = 7
    KnownFolderId_Playlists = 8
    KnownFolderId_RecordedCalls = 9
    KnownFolderId_RemovableDevices = 10
    KnownFolderId_SavedPictures = 11
    KnownFolderId_Screenshots = 12
    KnownFolderId_VideosLibrary = 13
    KnownFolderId_AllAppMods = 14
    KnownFolderId_CurrentAppMods = 15
    KnownFolderId_DownloadsFolder = 16
End Enum

Public Enum KnownFoldersAccessStatus
    KnownFoldersAccessStatus_DeniedBySystem = 0
    KnownFoldersAccessStatus_NotDeclaredByApp = 1
    KnownFoldersAccessStatus_DeniedByUser = 2
    KnownFoldersAccessStatus_UserPromptRequired = 3
    KnownFoldersAccessStatus_Allowed = 4
    KnownFoldersAccessStatus_AllowedPerAppFolder = 5
End Enum

' ----==== Namespace Storage.Search ====----
Public Property Get Search() As Search
    Set Search = New Search
End Property

' ----==== Namespace Storage.Streams ====----
Public Property Get Streams() As Streams
    Set Streams = New Streams
End Property

' ----==== Namespace Storage.Pickers ====----
Public Property Get Pickers() As Pickers
    Set Pickers = New Pickers
End Property

' ----==== Namespace Storage.StorageFile ====----
Public Property Get StorageItem() As StorageItem
    Set StorageItem = New StorageItem
End Property

' ----==== Namespace Storage.StorageFile ====----
Public Property Get StorageFile() As StorageFile
    Set StorageFile = New StorageFile
End Property

' ----==== Namespace Storage.StorageFolder ====----
Public Property Get StorageFolder() As StorageFolder
    Set StorageFolder = New StorageFolder
End Property

' ----==== Namespace Storage.StorageProvider ====----
Public Property Get StorageProvider() As StorageProvider
    Set StorageProvider = New StorageProvider
End Property

' ----==== Namespace Storage.FileProperties ====----
Public Property Get FileProperties() As FileProperties
    Set FileProperties = New FileProperties
End Property

' ----==== Namespace Storage.StorageLibraryChange ====----
Public Property Get StorageLibraryChange() As StorageLibraryChange
    Set StorageLibraryChange = New StorageLibraryChange
End Property

' ----==== Namespace Storage.StorageStreamTransaction ====----
Public Property Get StorageStreamTransaction() As StorageStreamTransaction
    Set StorageStreamTransaction = New StorageStreamTransaction
End Property

' ----==== Namespace Storage.StorageLibraryChangeReader ====----
Public Property Get StorageLibraryChangeReader() As StorageLibraryChangeReader
    Set StorageLibraryChangeReader = New StorageLibraryChangeReader
End Property

' ----==== Namespace Storage.StorageLibraryChangeTracker ====----
Public Property Get StorageLibraryChangeTracker() As StorageLibraryChangeTracker
    Set StorageLibraryChangeTracker = New StorageLibraryChangeTracker
End Property

' ----==== Namespace Storage.StorageLibraryChangeTrackerOptions ====----
Public Property Get StorageLibraryChangeTrackerOptions() As StorageLibraryChangeTrackerOptions
    Set StorageLibraryChangeTrackerOptions = New StorageLibraryChangeTrackerOptions
End Property

' ----==== Namespace Storage.KnownFolders ====----
Public Property Get KnownFolders() As KnownFolders
    Set KnownFolders = New KnownFolders
End Property

' ----==== Namespace Storage.UserDataPaths ====----
Public Property Get UserDataPaths() As UserDataPaths
    Set UserDataPaths = New UserDataPaths
End Property

