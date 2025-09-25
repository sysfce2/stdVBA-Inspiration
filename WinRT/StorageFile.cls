VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_StorageFile As String = "Windows.Storage.StorageFile"
Private Const IID_IStorageFile As String = "{fa3f6186-4214-428c-a64c-14c9ac7315ea}"
Private Const IID_IStorageFile2 As String = "{954e4bcf-0a77-42fb-b777-c2ed58a52e44}"
Private Const IID_IStorageItem As String = "{4207a996-ca2f-42f7-bde8-8b10457a7f30}"
Private Const IID_IStorageItem2 As String = "{53f926d2-083c-4283-b45b-81c007237e44}"
Private Const IID_IStorageFileStatics As String = "{5984c710-daf2-43c8-8bb4-a4d3eacfd03f}"
Private Const IID_IStorageFileStatics2 As String = "{5c76a781-212e-4af9-8f04-740cae108974}"
Private Const IID_IStorageItemProperties As String = "{86664478-8029-46fe-a789-1c2f3e2ffb5c}"
Private Const IID_IStorageItemProperties2 As String = "{8e86a951-04b9-4bd2-929d-fef3f71621d0}"
Private Const IID_IStorageItemPropertiesWithProvider As String = "{861bf39b-6368-4dee-b40e-74684a5ce714}"
Private Const IID_IStorageFilePropertiesWithAvailability As String = "{afcbbe9b-582b-4133-9648-e44ca46ee491}"
Private Const IID_IInputStreamReference As String = "{43929d18-5ec9-4b5a-919c-4205b0c804b6}"
Private Const IID_IRandomAccessStreamReference As String = "{33ee3134-1dd6-4e3a-8067-d1c162e8642b}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IStorageFileStatics
    IStorageFileStatics_GetFileFromPathAsync = 6

    ' IStorageFileStatics2
    IStorageFileStatics2_GetFileFromPathForUserAsync = 6

    ' IStorageFile
    IStorageFile_GetFileType = 6
    IStorageFile_GetContentType = 7
    IStorageFile_OpenAsync = 8
    IStorageFile_OpenTransactedWriteAsync = 9
    IStorageFile_CopyOverloadDefaultNameAndOptions = 10
    IStorageFile_CopyOverloadDefaultOptions = 11
    IStorageFile_CopyOverload = 12
    IStorageFile_CopyAndReplaceAsync = 13
    IStorageFile_MoveOverloadDefaultNameAndOptions = 14
    IStorageFile_MoveOverloadDefaultOptions = 15
    IStorageFile_MoveOverload = 16
    IStorageFile_MoveAndReplaceAsync = 17

    ' IStorageFile2
    IStorageFile2_OpenWithOptionsAsync = 6
    IStorageFile2_OpenTransactedWriteWithOptionsAsync = 6
                    
    ' IStorageItem
    IStorageItem_RenameAsyncOverloadDefaultOptions = 6
    IStorageItem_RenameAsync = 7
    IStorageItem_DeleteAsyncOverloadDefaultOptions = 8
    IStorageItem_DeleteAsync = 9
    IStorageItem_GetBasicPropertiesAsync = 10
    IStorageItem_GetName = 11
    IStorageItem_GetPath = 12
    IStorageItem_GetAttributes = 13
    IStorageItem_GetDateCreated = 14
    IStorageItem_IsOfType = 15

    ' IStorageItem2
    IStorageItem2_GetParentAsync = 6
    IStorageItem2_IsEqual = 7

    ' IStorageItemProperties
    IStorageItemProperties_GetThumbnailAsyncOverloadDefaultSizeDefaultOptions = 6
    IStorageItemProperties_GetThumbnailAsyncOverloadDefaultOptions = 7
    IStorageItemProperties_GetThumbnailAsync = 8
    IStorageItemProperties_GetDisplayName = 9
    IStorageItemProperties_GetDisplayType = 10
    IStorageItemProperties_GetFolderRelativeId = 11
    IStorageItemProperties_GetProperties = 12
                    
    ' IStorageItemProperties2
    IStorageItemProperties2_GetScaledImageAsThumbnailAsyncOverloadDefaultSizeDefaultOptions = 6
    IStorageItemProperties2_GetScaledImageAsThumbnailAsyncOverloadDefaultOptions = 7
    IStorageItemProperties2_GetScaledImageAsThumbnailAsync = 8

    ' IStorageItemPropertiesWithProvider
    IStorageItemPropertiesWithProvider_GetProvider = 6

    ' IStorageFilePropertiesWithAvailability
    IStorageFilePropertiesWithAvailability_GetIsAvailable = 6

    ' IInputStreamReference
    IInputStreamReference_OpenSequentialReadAsync = 6
    
    ' IRandomAccessStreamReference
    IRandomAccessStreamReference_OpenReadAsync = 6

End Enum

' ----==== Vars ====----
Private m_pIStorageFile As Long
Private m_pIStorageFile2 As Long
Private m_pIStorageItem As Long
Private m_pIStorageItem2 As Long
Private m_pIStorageFileStatics As Long
Private m_pIStorageFileStatics2 As Long
Private m_pIStorageItemProperties As Long
Private m_pIStorageItemProperties2 As Long
Private m_pIStorageItemPropertiesWithProvider As Long
Private m_pIStorageFilePropertiesWithAvailability As Long
Private m_pIInputStreamReference As Long
Private m_pIRandomAccessStreamReference As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_StorageFile, _
                            IID_IStorageFileStatics, _
                            m_pIStorageFileStatics) Then
        Call QueryIfc(m_pIStorageFileStatics, _
                      IID_IStorageFileStatics2, _
                      m_pIStorageFileStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageItem)
    Call ReleaseIfc(m_pIStorageItem2)
    Call ReleaseIfc(m_pIStorageFile)
    Call ReleaseIfc(m_pIStorageFile2)
    Call ReleaseIfc(m_pIStorageFileStatics)
    Call ReleaseIfc(m_pIStorageFileStatics2)
    Call ReleaseIfc(m_pIStorageItemProperties)
    Call ReleaseIfc(m_pIStorageItemProperties2)
    Call ReleaseIfc(m_pIStorageItemPropertiesWithProvider)
    Call ReleaseIfc(m_pIStorageFilePropertiesWithAvailability)
    Call ReleaseIfc(m_pIInputStreamReference)
    Call ReleaseIfc(m_pIRandomAccessStreamReference)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageFile
End Property

Public Property Let Ifc(ByVal pIStorageFile As Long)
    Call ReleaseIfc(m_pIStorageItem)
    Call ReleaseIfc(m_pIStorageItem2)
    Call ReleaseIfc(m_pIStorageFile)
    Call ReleaseIfc(m_pIStorageFile2)
    Call ReleaseIfc(m_pIStorageItemProperties)
    Call ReleaseIfc(m_pIStorageItemProperties2)
    Call ReleaseIfc(m_pIStorageItemPropertiesWithProvider)
    Call ReleaseIfc(m_pIStorageFilePropertiesWithAvailability)
    Call ReleaseIfc(m_pIInputStreamReference)
    Call ReleaseIfc(m_pIRandomAccessStreamReference)
    m_pIStorageFile = pIStorageFile
    Call QueryIfc(m_pIStorageFile, _
                  IID_IStorageItem, _
                  m_pIStorageItem)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IStorageItem2, _
                  m_pIStorageItem2)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IStorageFile2, _
                  m_pIStorageFile2)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IStorageItemProperties, _
                  m_pIStorageItemProperties)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IStorageItemProperties2, _
                  m_pIStorageItemProperties2)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IStorageItemPropertiesWithProvider, _
                  m_pIStorageItemPropertiesWithProvider)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IStorageFilePropertiesWithAvailability, _
                  m_pIStorageFilePropertiesWithAvailability)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IInputStreamReference, _
                  m_pIInputStreamReference)
    Call QueryIfc(m_pIStorageFile, _
                  IID_IRandomAccessStreamReference, _
                  m_pIRandomAccessStreamReference)
End Property

' IStorageFile
Public Property Get FileType() As String
    Dim Ret As String
    If m_pIStorageFile <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_GetFileType, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    FileType = Ret
End Property

Public Property Get ContentType() As String
    Dim Ret As String
    If m_pIStorageFile <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_GetContentType, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ContentType = Ret
End Property

' IStorageItem
Public Property Get Name() As String
    Dim Ret As String
    If m_pIStorageItem <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Name = Ret
End Property

Public Property Get Path() As String
    Dim Ret As String
    If m_pIStorageItem <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetPath, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Path = Ret
End Property

Public Property Get Attributes() As FileAttributes
    Dim Ret As FileAttributes
    If m_pIStorageItem <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetAttributes, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Attributes = Ret
End Property

Public Property Get DateCreated() As DateTime
    Dim Ret As DateTime
    If m_pIStorageItem <> 0& Then
        Dim value As Currency
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetDateCreated, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        Else
            Ret = Nothing
        End If
    End If
    Set DateCreated = Ret
End Property

' IStorageItemProperties
Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIStorageItemProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItemProperties, _
                  IStorageItemProperties_GetDisplayName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get DisplayType() As String
    Dim Ret As String
    If m_pIStorageItemProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItemProperties, _
                  IStorageItemProperties_GetDisplayType, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DisplayType = Ret
End Property

Public Property Get FolderRelativeId() As String
    Dim Ret As String
    If m_pIStorageItemProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItemProperties, _
                  IStorageItemProperties_GetFolderRelativeId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    FolderRelativeId = Ret
End Property

Public Property Get Properties() As StorageItemContentProperties
    Dim Ret As StorageItemContentProperties
    If m_pIStorageItemProperties <> 0& Then
        Dim pIStorageItemContentProperties As Long
        If Invoke(m_pIStorageItemProperties, _
                  IStorageItemProperties_GetProperties, _
                  VarPtr(pIStorageItemContentProperties)) = S_OK Then
            If pIStorageItemContentProperties <> 0& Then
                Set Ret = New StorageItemContentProperties
                Ret.Ifc = pIStorageItemContentProperties
            End If
        End If
    End If
    Set Properties = Ret
End Property

' IStorageItemPropertiesWithProvider
Public Property Get Provider() As StorageProvider
    Dim Ret As StorageProvider
    If m_pIStorageItemPropertiesWithProvider <> 0& Then
        Dim pIStorageProvider As Long
        If Invoke(m_pIStorageItemPropertiesWithProvider, _
                  IStorageItemPropertiesWithProvider_GetProvider, _
                  VarPtr(pIStorageProvider)) = S_OK Then
            If pIStorageProvider <> 0& Then
                Set Ret = New StorageProvider
                Ret.Ifc = pIStorageProvider
            End If
        End If
    End If
    Set Provider = Ret
End Property

' IStorageFilePropertiesWithAvailability
Public Property Get IsAvailable() As Boolean
    Dim Ret As Boolean
    If m_pIStorageFilePropertiesWithAvailability <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageFilePropertiesWithAvailability, _
                  IStorageFilePropertiesWithAvailability_GetIsAvailable, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsAvailable = Ret
End Property

' ----==== Functions ====----
' IStorageFileStatics
Public Function GetFileFromPathAsync(ByVal Path As String) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFileStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Path)
        Dim pIStorageFile As Long
        If Invoke(m_pIStorageFileStatics, _
                  IStorageFileStatics_GetFileFromPathAsync, _
                  hString, VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                If Await(pIStorageFile) Then
                    Set Ret = New StorageFile
                    Ret.Ifc = pIStorageFile
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetFileFromPathAsync = Ret
End Function

' IStorageFileStatics2
Public Function GetFileFromPathForUserAsync(ByVal value As User, _
                                            ByVal Path As String) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFileStatics2 <> 0& And _
       IsNotNothing(value) Then
        Dim hString As Long
        hString = CreateWindowsString(Path)
        Dim pIStorageFile As Long
        If Invoke(m_pIStorageFileStatics2, _
                  IStorageFileStatics2_GetFileFromPathForUserAsync, _
                  value.Ifc, _
                  hString, VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                If Await(pIStorageFile) Then
                    Set Ret = New StorageFile
                    Ret.Ifc = pIStorageFile
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetFileFromPathForUserAsync = Ret
End Function

' IStorageFile
Public Function OpenAsync(ByVal accessMode As FileAccessMode) As RandomAccessStream
    Dim Ret As RandomAccessStream
    If m_pIStorageFile <> 0& Then
        Dim pIRandomAccessStream As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_OpenAsync, _
                  accessMode, _
                  VarPtr(pIRandomAccessStream)) = S_OK Then
            If pIRandomAccessStream <> 0& Then
                If Await(pIRandomAccessStream) Then
                    Set Ret = New RandomAccessStream
                    Ret.Ifc = pIRandomAccessStream
                End If
            End If
        End If
    End If
    Set OpenAsync = Ret
End Function

Public Function OpenTransactedWriteAsync() As StorageStreamTransaction
    Dim Ret As StorageStreamTransaction
    If m_pIStorageFile <> 0& Then
        Dim pIStorageStreamTransaction As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_OpenTransactedWriteAsync, _
                  VarPtr(pIStorageStreamTransaction)) = S_OK Then
            If pIStorageStreamTransaction <> 0& Then
                If Await(pIStorageStreamTransaction) Then
                    Set Ret = New StorageStreamTransaction
                    Ret.Ifc = pIStorageStreamTransaction
                End If
            End If
        End If
    End If
    Set OpenTransactedWriteAsync = Ret
End Function

Public Function CopyOverloadDefaultNameAndOptions(ByVal destinationFolder As StorageFolder) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFile <> 0& And _
       IsNotNothing(destinationFolder) Then
        Dim pIStorageFile As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_CopyOverloadDefaultNameAndOptions, _
                  destinationFolder.Ifc, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                If Await(pIStorageFile) Then
                    Set Ret = New StorageFile
                    Ret.Ifc = pIStorageFile
                End If
            End If
        End If
    End If
    Set CopyOverloadDefaultNameAndOptions = Ret
End Function

Public Function CopyOverloadDefaultOptions(ByVal destinationFolder As StorageFolder, _
                                           ByVal desiredNewName As String) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFile <> 0& And _
       IsNotNothing(destinationFolder) Then
        Dim pIStorageFile As Long
        Dim hString As Long
        hString = CreateWindowsString(desiredNewName)
        If Invoke(m_pIStorageFile, _
                  IStorageFile_CopyOverloadDefaultOptions, _
                  destinationFolder.Ifc, _
                  hString, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                If Await(pIStorageFile) Then
                    Set Ret = New StorageFile
                    Ret.Ifc = pIStorageFile
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CopyOverloadDefaultOptions = Ret
End Function

Public Function CopyOverload(ByVal destinationFolder As StorageFolder, _
                             ByVal desiredNewName As String, _
                             ByVal opt As NameCollisionOption) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFile <> 0& And _
       IsNotNothing(destinationFolder) Then
        Dim pIStorageFile As Long
        Dim hString As Long
        hString = CreateWindowsString(desiredNewName)
        If Invoke(m_pIStorageFile, _
                  IStorageFile_CopyOverload, _
                  destinationFolder.Ifc, _
                  hString, _
                  opt, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                If Await(pIStorageFile) Then
                    Set Ret = New StorageFile
                    Ret.Ifc = pIStorageFile
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CopyOverload = Ret
End Function

Public Function CopyAndReplaceAsync(ByVal fileToReplace As StorageFile) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFile <> 0& And _
       IsNotNothing(fileToReplace) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_CopyAndReplaceAsync, _
                  fileToReplace.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
                Call ReleaseIfc(pIAsyncAction)
            End If
        End If
    End If
    CopyAndReplaceAsync = Ret
End Function

Public Function MoveOverloadDefaultNameAndOptions(ByVal destinationFolder As StorageFolder) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFile <> 0& And _
       IsNotNothing(destinationFolder) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_MoveOverloadDefaultNameAndOptions, _
                  destinationFolder.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
                Call ReleaseIfc(pIAsyncAction)
            End If
        End If
    End If
    MoveOverloadDefaultNameAndOptions = Ret
End Function

Public Function MoveOverloadDefaultOptions(ByVal destinationFolder As StorageFolder, _
                                           ByVal desiredNewName As String) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFile <> 0& And _
       IsNotNothing(destinationFolder) Then
        Dim pIAsyncAction As Long
        Dim hString As Long
        hString = CreateWindowsString(desiredNewName)
        If Invoke(m_pIStorageFile, _
                  IStorageFile_MoveOverloadDefaultOptions, _
                  destinationFolder.Ifc, _
                  hString, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
                Call ReleaseIfc(pIAsyncAction)
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    MoveOverloadDefaultOptions = Ret
End Function

Public Function MoveOverload(ByVal destinationFolder As StorageFolder, _
                             ByVal desiredNewName As String, _
                             ByVal opt As NameCollisionOption) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFile <> 0& And _
       IsNotNothing(destinationFolder) Then
        Dim pIAsyncAction As Long
        Dim hString As Long
        hString = CreateWindowsString(desiredNewName)
        If Invoke(m_pIStorageFile, _
                  IStorageFile_MoveOverload, _
                  destinationFolder.Ifc, _
                  hString, _
                  opt, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
                Call ReleaseIfc(pIAsyncAction)
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    MoveOverload = Ret
End Function

Public Function MoveAndReplaceAsync(ByVal fileToReplace As StorageFile) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFile <> 0& And _
       IsNotNothing(fileToReplace) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIStorageFile, _
                  IStorageFile_MoveAndReplaceAsync, _
                  fileToReplace.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
                Call ReleaseIfc(pIAsyncAction)
            End If
        End If
    End If
    MoveAndReplaceAsync = Ret
End Function

' IStorageFile2
Public Function OpenWithOptionsAsync(ByVal accessMode As FileAccessMode, _
                                     ByVal Options As StorageOpenOptions) As RandomAccessStream
    Dim Ret As RandomAccessStream
    If m_pIStorageFile2 <> 0& Then
        Dim pIRandomAccessStream As Long
        If Invoke(m_pIStorageFile2, _
                  IStorageFile2_OpenWithOptionsAsync, _
                  accessMode, _
                  Options, _
                  VarPtr(pIRandomAccessStream)) = S_OK Then
            If pIRandomAccessStream <> 0& Then
                If Await(pIRandomAccessStream) Then
                    Set Ret = New RandomAccessStream
                    Ret.Ifc = pIRandomAccessStream
                End If
            End If
        End If
    End If
    Set OpenWithOptionsAsync = Ret
End Function

Public Function OpenTransactedWriteWithOptionsAsync(ByVal Options As StorageOpenOptions) As StorageStreamTransaction
    Dim Ret As StorageStreamTransaction
    If m_pIStorageFile2 <> 0& Then
        Dim pIStorageStreamTransaction As Long
        If Invoke(m_pIStorageFile2, _
                  IStorageFile2_OpenTransactedWriteWithOptionsAsync, _
                  Options, _
                  VarPtr(pIStorageStreamTransaction)) = S_OK Then
            If pIStorageStreamTransaction <> 0& Then
                If Await(pIStorageStreamTransaction) Then
                    Set Ret = New StorageStreamTransaction
                    Ret.Ifc = pIStorageStreamTransaction
                End If
            End If
        End If
    End If
    Set OpenTransactedWriteWithOptionsAsync = Ret
End Function

' IStorageItem
Public Function RenameAsyncOverloadDefaultOptions(ByVal desiredName As String) As Boolean
    Dim Ret As Boolean
    If m_pIStorageItem <> 0& Then
        Dim hString As Long
        Dim pIAsyncAction As Long
        hString = CreateWindowsString(desiredName)
        If Invoke(m_pIStorageItem, _
                  IStorageItem_RenameAsyncOverloadDefaultOptions, _
                  hString, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = CBool(pIAsyncAction)
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    RenameAsyncOverloadDefaultOptions = Ret
End Function

Public Function RenameAsync(ByVal desiredName As String, _
                            Optional ByVal collisionOption As NameCollisionOption = NameCollisionOption_GenerateUniqueName) As Boolean
    Dim Ret As Boolean
    If m_pIStorageItem <> 0& Then
        Dim hString As Long
        Dim pIAsyncAction As Long
        hString = CreateWindowsString(desiredName)
        If Invoke(m_pIStorageItem, _
                  IStorageItem_RenameAsync, _
                  hString, _
                  collisionOption, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = CBool(pIAsyncAction)
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    RenameAsync = Ret
End Function

Public Function DeleteAsyncOverloadDefaultOptions() As Boolean
    Dim Ret As Boolean
    If m_pIStorageItem <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_DeleteAsyncOverloadDefaultOptions, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = CBool(pIAsyncAction)
            End If
        End If
    End If
    DeleteAsyncOverloadDefaultOptions = Ret
End Function

Public Function DeleteAsync(Optional ByVal deleteOption As StorageDeleteOption = StorageDeleteOption_Default) As Boolean
    Dim Ret As Boolean
    If m_pIStorageItem <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_DeleteAsync, _
                  deleteOption, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = CBool(pIAsyncAction)
            End If
        End If
    End If
    DeleteAsync = Ret
End Function

Public Function GetBasicPropertiesAsync() As BasicProperties
    Dim Ret As BasicProperties
    If m_pIStorageItem <> 0& Then
        Dim pIBasicProperties As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetBasicPropertiesAsync, _
                  VarPtr(pIBasicProperties)) = S_OK Then
            If pIBasicProperties <> 0& Then
                If Await(pIBasicProperties) Then
                    Set Ret = New BasicProperties
                    Ret.Ifc = pIBasicProperties
                End If
            End If
        End If
    End If
    Set GetBasicPropertiesAsync = Ret
End Function

Public Function IsOfType(ByVal itemType As StorageItemTypes) As Boolean
    Dim Ret As Boolean
    If m_pIStorageItem <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_IsOfType, _
                  itemType, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsOfType = Ret
End Function

' IStorageItem2
Public Function GetParentAsync() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIStorageItem2 <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIStorageItem2, _
                  IStorageItem2_GetParentAsync, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                If Await(pIStorageFolder) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = pIStorageFolder
                End If
            End If
        End If
    End If
    Set GetParentAsync = Ret
End Function

Public Function IsEqual(ByVal Item As Object) As Boolean
    Dim Ret As Boolean
    If m_pIStorageItem2 <> 0& And _
       Item.IStorageItem <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageItem2, _
                  IStorageItem2_IsEqual, _
                  Item.IStorageItem, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsEqual = Ret
End Function

' IStorageItemProperties
Public Function GetThumbnailAsyncOverloadDefaultSizeDefaultOptions(ByVal Mode As ThumbnailMode) As StorageItemThumbnail
    Dim Ret As StorageItemThumbnail
    If m_pIStorageItemProperties <> 0& Then
        Dim pIStorageItemThumbnail As Long
        If Invoke(m_pIStorageItemProperties, _
                  IStorageItemProperties_GetThumbnailAsyncOverloadDefaultSizeDefaultOptions, _
                  Mode, _
                  VarPtr(pIStorageItemThumbnail)) = S_OK Then
            If pIStorageItemThumbnail <> 0& Then
                If Await(pIStorageItemThumbnail) Then
                    Set Ret = New StorageItemThumbnail
                    Ret.Ifc = pIStorageItemThumbnail
                End If
            End If
        End If
    End If
    Set GetThumbnailAsyncOverloadDefaultSizeDefaultOptions = Ret
End Function

Public Function GetThumbnailAsyncOverloadDefaultOptions(ByVal Mode As ThumbnailMode, _
                                                        ByVal requestedSize As Long) As StorageItemThumbnail
    Dim Ret As StorageItemThumbnail
    If m_pIStorageItemProperties <> 0& Then
        Dim pIStorageItemThumbnail As Long
        If Invoke(m_pIStorageItemProperties, _
                  IStorageItemProperties_GetThumbnailAsyncOverloadDefaultOptions, _
                  Mode, _
                  requestedSize, _
                  VarPtr(pIStorageItemThumbnail)) = S_OK Then
            If pIStorageItemThumbnail <> 0& Then
                If Await(pIStorageItemThumbnail) Then
                    Set Ret = New StorageItemThumbnail
                    Ret.Ifc = pIStorageItemThumbnail
                End If
            End If
        End If
    End If
    Set GetThumbnailAsyncOverloadDefaultOptions = Ret
End Function

Public Function GetThumbnailAsync(ByVal Mode As ThumbnailMode, _
                                  ByVal requestedSize As Long, _
                                  ByVal Options As ThumbnailOptions) As StorageItemThumbnail
    Dim Ret As StorageItemThumbnail
    If m_pIStorageItemProperties <> 0& Then
        Dim pIStorageItemThumbnail As Long
        If Invoke(m_pIStorageItemProperties, _
                  IStorageItemProperties_GetThumbnailAsync, _
                  Mode, _
                  requestedSize, _
                  Options, _
                  VarPtr(pIStorageItemThumbnail)) = S_OK Then
            If pIStorageItemThumbnail <> 0& Then
                If Await(pIStorageItemThumbnail) Then
                    Set Ret = New StorageItemThumbnail
                    Ret.Ifc = pIStorageItemThumbnail
                End If
            End If
        End If
    End If
    Set GetThumbnailAsync = Ret
End Function

' IStorageItemProperties2
Public Function GetScaledImageAsThumbnailAsyncOverloadDefaultSizeDefaultOptions(ByVal Mode As ThumbnailMode) As StorageItemThumbnail
    Dim Ret As StorageItemThumbnail
    If m_pIStorageItemProperties2 <> 0& Then
        Dim pIStorageItemThumbnail As Long
        If Invoke(m_pIStorageItemProperties2, _
                  IStorageItemProperties2_GetScaledImageAsThumbnailAsyncOverloadDefaultSizeDefaultOptions, _
                  Mode, _
                  VarPtr(pIStorageItemThumbnail)) = S_OK Then
            If pIStorageItemThumbnail <> 0& Then
                If Await(pIStorageItemThumbnail) Then
                    Set Ret = New StorageItemThumbnail
                    Ret.Ifc = pIStorageItemThumbnail
                End If
            End If
        End If
    End If
    Set GetScaledImageAsThumbnailAsyncOverloadDefaultSizeDefaultOptions = Ret
End Function

Public Function GetScaledImageAsThumbnailAsyncOverloadDefaultOptions(ByVal Mode As ThumbnailMode, _
                                                                     ByVal requestedSize As Long) As StorageItemThumbnail
    Dim Ret As StorageItemThumbnail
    If m_pIStorageItemProperties2 <> 0& Then
        Dim pIStorageItemThumbnail As Long
        If Invoke(m_pIStorageItemProperties2, _
                  IStorageItemProperties2_GetScaledImageAsThumbnailAsyncOverloadDefaultOptions, _
                  Mode, _
                  requestedSize, _
                  VarPtr(pIStorageItemThumbnail)) = S_OK Then
            If pIStorageItemThumbnail <> 0& Then
                If Await(pIStorageItemThumbnail) Then
                    Set Ret = New StorageItemThumbnail
                    Ret.Ifc = pIStorageItemThumbnail
                End If
            End If
        End If
    End If
    Set GetScaledImageAsThumbnailAsyncOverloadDefaultOptions = Ret
End Function

Public Function GetScaledImageAsThumbnailAsync(ByVal Mode As ThumbnailMode, _
                                               ByVal requestedSize As Long, _
                                               ByVal Options As ThumbnailOptions) As StorageItemThumbnail
    Dim Ret As StorageItemThumbnail
    If m_pIStorageItemProperties2 <> 0& Then
        Dim pIStorageItemThumbnail As Long
        If Invoke(m_pIStorageItemProperties2, _
                  IStorageItemProperties2_GetScaledImageAsThumbnailAsync, _
                  Mode, _
                  requestedSize, _
                  Options, _
                  VarPtr(pIStorageItemThumbnail)) = S_OK Then
            If pIStorageItemThumbnail <> 0& Then
                If Await(pIStorageItemThumbnail) Then
                    Set Ret = New StorageItemThumbnail
                    Ret.Ifc = pIStorageItemThumbnail
                End If
            End If
        End If
    End If
    Set GetScaledImageAsThumbnailAsync = Ret
End Function

' IInputStreamReference
Public Function OpenSequentialReadAsync() As InputStream
    Dim Ret As InputStream
    If m_pIInputStreamReference <> 0& Then
        Dim pIInputStream As Long
        If Invoke(m_pIInputStreamReference, _
                  IInputStreamReference_OpenSequentialReadAsync, _
                  VarPtr(pIInputStream)) = S_OK Then
            If pIInputStream <> 0& Then
                If Await(pIInputStream) Then
                    Set Ret = New InputStream
                    Ret.Ifc = pIInputStream
                End If
            End If
        End If
    End If
    Set OpenSequentialReadAsync = Ret
End Function
    
' IRandomAccessStreamReference
Public Function OpenReadAsync() As RandomAccessStreamWithContentType
    Dim Ret As RandomAccessStreamWithContentType
    If m_pIRandomAccessStreamReference <> 0& Then
        Dim pIRandomAccessStreamWithContentType As Long
        If Invoke(m_pIRandomAccessStreamReference, _
                  IRandomAccessStreamReference_OpenReadAsync, _
                  VarPtr(pIRandomAccessStreamWithContentType)) = S_OK Then
            If pIRandomAccessStreamWithContentType <> 0& Then
                If Await(pIRandomAccessStreamWithContentType) Then
                    Set Ret = New RandomAccessStreamWithContentType
                    Ret.Ifc = pIRandomAccessStreamWithContentType
                End If
            End If
        End If
    End If
    Set OpenReadAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


