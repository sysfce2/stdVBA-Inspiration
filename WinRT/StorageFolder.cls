VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_StorageFolder As String = "Windows.Storage.StorageFolder"
Private Const IID_IStorageItem As String = "{4207a996-ca2f-42f7-bde8-8b10457a7f30}"
Private Const IID_IStorageItem2 As String = "{53f926d2-083c-4283-b45b-81c007237e44}"
Private Const IID_IStorageFolder As String = "{72d1cb78-b3ef-4f75-a80b-6fd9dae2944b}"
Private Const IID_IStorageFolder2 As String = "{e827e8b9-08d9-4a8e-a0ac-fe5ed3cbbbd3}"
Private Const IID_IStorageFolder3 As String = "{9f617899-bde1-4124-aeb3-b06ad96f98d4}"
Private Const IID_IStorageFolderStatics As String = "{08f327ff-85d5-48b9-aee9-28511e339f9f}"
Private Const IID_IStorageFolderStatics2 As String = "{b4656dc3-71d2-467d-8b29-371f0f62bf6f}"
Private Const IID_IStorageItemProperties As String = "{86664478-8029-46fe-a789-1c2f3e2ffb5c}"
Private Const IID_IStorageItemProperties2 As String = "{8e86a951-04b9-4bd2-929d-fef3f71621d0}"
Private Const IID_IStorageFolderQueryOperations As String = "{cb43ccc9-446b-4a4f-be97-757771be5203}"
Private Const IID_IStorageItemPropertiesWithProvider As String = "{861bf39b-6368-4dee-b40e-74684a5ce714}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IStorageFolderStatics
    IStorageFolderStatics_GetFolderFromPathAsync = 6

    ' IStorageFolderStatics2
    IStorageFolderStatics2_GetFolderFromPathForUserAsync = 6

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

    ' IStorageFolder
    IStorageFolder_CreateFileAsyncOverloadDefaultOptions = 6
    IStorageFolder_CreateFileAsync = 7
    IStorageFolder_CreateFolderAsyncOverloadDefaultOptions = 8
    IStorageFolder_CreateFolderAsync = 9
    IStorageFolder_GetFileAsync = 10
    IStorageFolder_GetFolderAsync = 11
    IStorageFolder_GetItemAsync = 12
    IStorageFolder_GetFilesAsyncOverloadDefaultOptionsStartAndCount = 13
    IStorageFolder_GetFoldersAsyncOverloadDefaultOptionsStartAndCount = 14
    IStorageFolder_GetItemsAsyncOverloadDefaultStartAndCount = 15
    
    ' IStorageFolder2
    IStorageFolder2_TryGetItemAsync = 6
    
    ' IStorageFolder3
    IStorageFolder3_TryGetChangeTracker = 6
    
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

    ' IStorageFolderQueryOperations
    IStorageFolderQueryOperations_GetIndexedStateAsync = 6
    IStorageFolderQueryOperations_CreateFileQueryOverloadDefault = 7
    IStorageFolderQueryOperations_CreateFileQuery = 8
    IStorageFolderQueryOperations_CreateFileQueryWithOptions = 9
    IStorageFolderQueryOperations_CreateFolderQueryOverloadDefault = 10
    IStorageFolderQueryOperations_CreateFolderQuery = 11
    IStorageFolderQueryOperations_CreateFolderQueryWithOptions = 12
    IStorageFolderQueryOperations_CreateItemQuery = 13
    IStorageFolderQueryOperations_CreateItemQueryWithOptions = 14
    IStorageFolderQueryOperations_GetFilesAsync = 15
    IStorageFolderQueryOperations_GetFilesAsyncOverloadDefaultStartAndCount = 16
    IStorageFolderQueryOperations_GetFoldersAsync = 17
    IStorageFolderQueryOperations_GetFoldersAsyncOverloadDefaultStartAndCount = 18
    IStorageFolderQueryOperations_GetItemsAsync = 19
    IStorageFolderQueryOperations_AreQueryOptionsSupported = 20
    IStorageFolderQueryOperations_IsCommonFolderQuerySupported = 21
    IStorageFolderQueryOperations_IsCommonFileQuerySupported = 22

    ' IStorageItemPropertiesWithProvider
    IStorageItemPropertiesWithProvider_GetProvider = 6

End Enum

' ----==== Vars ====----
Private m_pIStorageItem As Long
Private m_pIStorageItem2 As Long
Private m_pIStorageFolder As Long
Private m_pIStorageFolder2 As Long
Private m_pIStorageFolder3 As Long
Private m_pIStorageFolderStatics As Long
Private m_pIStorageFolderStatics2 As Long
Private m_pIStorageItemProperties As Long
Private m_pIStorageItemProperties2 As Long
Private m_pIStorageFolderQueryOperations As Long
Private m_pIStorageItemPropertiesWithProvider As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_StorageFolder, _
                            IID_IStorageFolderStatics, _
                            m_pIStorageFolderStatics) Then
        Call QueryIfc(m_pIStorageFolderStatics, _
                      IID_IStorageFolderStatics2, _
                      m_pIStorageFolderStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageItem)
    Call ReleaseIfc(m_pIStorageItem2)
    Call ReleaseIfc(m_pIStorageFolder)
    Call ReleaseIfc(m_pIStorageFolder2)
    Call ReleaseIfc(m_pIStorageFolder3)
    Call ReleaseIfc(m_pIStorageFolderStatics)
    Call ReleaseIfc(m_pIStorageFolderStatics2)
    Call ReleaseIfc(m_pIStorageItemProperties)
    Call ReleaseIfc(m_pIStorageItemProperties2)
    Call ReleaseIfc(m_pIStorageFolderQueryOperations)
    Call ReleaseIfc(m_pIStorageItemPropertiesWithProvider)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageFolder
End Property

Public Property Let Ifc(ByVal pIStorageFolder As Long)
    Call ReleaseIfc(m_pIStorageItem)
    Call ReleaseIfc(m_pIStorageItem2)
    Call ReleaseIfc(m_pIStorageFolder)
    Call ReleaseIfc(m_pIStorageFolder2)
    Call ReleaseIfc(m_pIStorageFolder3)
    Call ReleaseIfc(m_pIStorageItemProperties)
    Call ReleaseIfc(m_pIStorageItemProperties2)
    Call ReleaseIfc(m_pIStorageFolderQueryOperations)
    Call ReleaseIfc(m_pIStorageItemPropertiesWithProvider)
    m_pIStorageFolder = pIStorageFolder
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageItem, _
                  m_pIStorageItem)
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageItem2, _
                  m_pIStorageItem2)
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageFolder2, _
                  m_pIStorageFolder2)
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageFolder3, _
                  m_pIStorageFolder3)
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageItemProperties, _
                  m_pIStorageItemProperties)
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageItemProperties2, _
                  m_pIStorageItemProperties2)
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageFolderQueryOperations, _
                  m_pIStorageFolderQueryOperations)
    Call QueryIfc(m_pIStorageFolder, _
                  IID_IStorageItemPropertiesWithProvider, _
                  m_pIStorageItemPropertiesWithProvider)
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

' ----==== Functions ====----
' IStorageFolderStatics
Public Function GetFolderFromPathAsync(ByVal Path As String) As StorageFolder
    Dim Ret As StorageFolder
    If m_pIStorageFolderStatics <> 0& Then
        Dim hString As Long
        Dim pIStorageFolder As Long
        hString = CreateWindowsString(Path)
        If Invoke(m_pIStorageFolderStatics, _
                  IStorageFolderStatics_GetFolderFromPathAsync, _
                  hString, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                If Await(pIStorageFolder) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = pIStorageFolder
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetFolderFromPathAsync = Ret
End Function

' IStorageFolderStatics2
Public Function GetFolderFromPathForUserAsync(ByVal value As User, _
                                              ByVal Path As String) As StorageFolder
    Dim Ret As StorageFolder
    If m_pIStorageFolderStatics2 <> 0& And _
       IsNotNothing(value) Then
        Dim hString As Long
        Dim pIStorageFolder As Long
        hString = CreateWindowsString(Path)
        If Invoke(m_pIStorageFolderStatics2, _
                  IStorageFolderStatics2_GetFolderFromPathForUserAsync, _
                  value.Ifc, _
                  hString, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                If Await(pIStorageFolder) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = pIStorageFolder
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetFolderFromPathForUserAsync = Ret
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
       IsNotNothing(Item) Then
        Dim value As Long
        If Invoke(m_pIStorageItem2, _
                  IStorageItem2_IsEqual, _
                  Item.Ifc, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsEqual = Ret
End Function

' IStorageFolder
Public Function CreateFileAsyncOverloadDefaultOptions(ByVal fileName As String) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFolder <> 0& Then
        Dim hString As Long
        Dim pIStorageFile As Long
        hString = CreateWindowsString(fileName)
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_CreateFileAsyncOverloadDefaultOptions, _
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
    Set CreateFileAsyncOverloadDefaultOptions = Ret
End Function

Public Function CreateFileAsync(ByVal fileName As String, _
                                ByVal Options As CreationCollisionOption) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFolder <> 0& Then
        Dim hString As Long
        Dim pIStorageFile As Long
        hString = CreateWindowsString(fileName)
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_CreateFileAsync, _
                  hString, _
                  Options, _
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
    Set CreateFileAsync = Ret
End Function

Public Function CreateFolderAsyncOverloadDefaultOptions(ByVal folderName As String) As StorageFolder
    Dim Ret As StorageFolder
    If m_pIStorageFolder <> 0& Then
        Dim hString As Long
        Dim pIStorageFolder As Long
        hString = CreateWindowsString(folderName)
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_CreateFolderAsyncOverloadDefaultOptions, _
                  hString, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                If Await(pIStorageFolder) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = pIStorageFolder
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateFolderAsyncOverloadDefaultOptions = Ret
End Function

Public Function CreateFolderAsync(ByVal folderName As String, _
                                  ByVal Options As CreationCollisionOption) As StorageFolder
    Dim Ret As StorageFolder
    If m_pIStorageFolder <> 0& Then
        Dim hString As Long
        Dim pIStorageFolder As Long
        hString = CreateWindowsString(folderName)
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_CreateFolderAsync, _
                  hString, _
                  Options, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                If Await(pIStorageFolder) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = pIStorageFolder
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateFolderAsync = Ret
End Function

Public Function GetFileAsync(ByVal fileName As String) As StorageFile
    Dim Ret As StorageFile
    If m_pIStorageFolder <> 0& Then
        Dim hString As Long
        Dim pIStorageFile As Long
        hString = CreateWindowsString(fileName)
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_GetFileAsync, _
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
    Set GetFileAsync = Ret
End Function

Public Function GetFolderAsync(ByVal folderName As String) As StorageFolder
    Dim Ret As StorageFolder
    If m_pIStorageFolder <> 0& Then
        Dim hString As Long
        Dim pIStorageFolder As Long
        hString = CreateWindowsString(folderName)
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_GetFolderAsync, _
                  hString, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                If Await(pIStorageFolder) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = pIStorageFolder
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetFolderAsync = Ret
End Function

Public Function GetItemAsync(ByVal itemName As String) As StorageItem
    Dim Ret As StorageItem
    If m_pIStorageFolder <> 0& Then
        Dim hString As Long
        Dim pIStorageItem As Long
        hString = CreateWindowsString(itemName)
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_GetItemAsync, _
                  hString, _
                  VarPtr(pIStorageItem)) = S_OK Then
            If pIStorageItem <> 0& Then
                If Await(pIStorageItem) Then
                    Set Ret = New StorageItem
                    Ret.Ifc = pIStorageItem
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetItemAsync = Ret
End Function

Public Function GetFilesAsyncOverloadDefaultOptionsStartAndCount() As ReadOnlyList_1 'ReadOnlyList_StorageFile
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageFile
    If m_pIStorageFolder <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_GetFilesAsyncOverloadDefaultOptionsStartAndCount, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageFile, pIVectorView) 'ReadOnlyList_StorageFile
                End If
            End If
        End If
    End If
    Set GetFilesAsyncOverloadDefaultOptionsStartAndCount = Ret
End Function

Public Function GetFoldersAsyncOverloadDefaultOptionsStartAndCount() As ReadOnlyList_1 'ReadOnlyList_StorageFolder
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageFolder
    If m_pIStorageFolder <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_GetFoldersAsyncOverloadDefaultOptionsStartAndCount, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageFolder, pIVectorView) 'ReadOnlyList_StorageFolder
                End If
            End If
        End If
    End If
    Set GetFoldersAsyncOverloadDefaultOptionsStartAndCount = Ret
End Function

Public Function GetItemsAsyncOverloadDefaultStartAndCount() As ReadOnlyList_1 'ReadOnlyList_StorageItem
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageItem
    If m_pIStorageFolder <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolder, _
                  IStorageFolder_GetItemsAsyncOverloadDefaultStartAndCount, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageItem, pIVectorView) 'ReadOnlyList_StorageItem
                End If
            End If
        End If
    End If
    Set GetItemsAsyncOverloadDefaultStartAndCount = Ret
End Function

' IStorageFolder2
Public Function TryGetItemAsync(ByVal itemName As String) As StorageItem
    Dim Ret As StorageItem
    If m_pIStorageFolder2 <> 0& Then
        Dim hString As Long
        Dim pIStorageItem As Long
        hString = CreateWindowsString(itemName)
        If Invoke(m_pIStorageFolder2, _
                  IStorageFolder2_TryGetItemAsync, _
                  hString, _
                  VarPtr(pIStorageItem)) = S_OK Then
            If pIStorageItem <> 0& Then
                If Await(pIStorageItem) Then
                    Set Ret = New StorageItem
                    Ret.Ifc = pIStorageItem
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set TryGetItemAsync = Ret
End Function

' IStorageFolder3
Public Function TryGetChangeTracker() As StorageLibraryChangeTracker
    Dim Ret As StorageLibraryChangeTracker
    If m_pIStorageFolder3 <> 0& Then
        Dim pIStorageLibraryChangeTracker As Long
        If Invoke(m_pIStorageFolder3, _
                  IStorageFolder3_TryGetChangeTracker, _
                  VarPtr(pIStorageLibraryChangeTracker)) = S_OK Then
            If pIStorageLibraryChangeTracker <> 0& Then
                    Set Ret = New StorageLibraryChangeTracker
                    Ret.Ifc = pIStorageLibraryChangeTracker
            End If
        End If
    End If
    Set TryGetChangeTracker = Ret
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
                Set Ret = New StorageItemThumbnail
                Ret.Ifc = pIStorageItemThumbnail
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
                Set Ret = New StorageItemThumbnail
                Ret.Ifc = pIStorageItemThumbnail
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
                Set Ret = New StorageItemThumbnail
                Ret.Ifc = pIStorageItemThumbnail
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
                Set Ret = New StorageItemThumbnail
                Ret.Ifc = pIStorageItemThumbnail
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
                Set Ret = New StorageItemThumbnail
                Ret.Ifc = pIStorageItemThumbnail
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
                Set Ret = New StorageItemThumbnail
                Ret.Ifc = pIStorageItemThumbnail
            End If
        End If
    End If
    Set GetScaledImageAsThumbnailAsync = Ret
End Function

' IStorageFolderQueryOperations
Public Function GetIndexedStateAsync() As IndexedState
    Dim Ret As IndexedState
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_GetIndexedStateAsync, _
                  VarPtr(value)) = S_OK Then
            If Await(value) Then
                Ret = value
            End If
        End If
    End If
    GetIndexedStateAsync = Ret
End Function

Public Function CreateFileQuery(ByVal Query As CommonFileQuery) As StorageFileQueryResult
    Dim Ret As StorageFileQueryResult
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIStorageFileQueryResult As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_CreateFileQuery, _
                  Query, _
                  VarPtr(pIStorageFileQueryResult)) = S_OK Then
            If pIStorageFileQueryResult <> 0& Then
                Set Ret = New StorageFileQueryResult
                Ret.Ifc = pIStorageFileQueryResult
            End If
        End If
    End If
    Set CreateFileQuery = Ret
End Function

Public Function CreateFileQueryWithOptions(ByVal Options As QueryOptions) As StorageFileQueryResult
    Dim Ret As StorageFileQueryResult
    If m_pIStorageFolderQueryOperations <> 0& And _
       IsNotNothing(Options) Then
        Dim pIStorageFileQueryResult As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_CreateFileQueryWithOptions, _
                  Options.Ifc, _
                  VarPtr(pIStorageFileQueryResult)) = S_OK Then
            If pIStorageFileQueryResult <> 0& Then
                Set Ret = New StorageFileQueryResult
                Ret.Ifc = pIStorageFileQueryResult
            End If
        End If
    End If
    Set CreateFileQueryWithOptions = Ret
End Function

Public Function CreateFolderQueryOverloadDefault() As StorageFolderQueryResult
    Dim Ret As StorageFolderQueryResult
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIStorageFolderQueryResult As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_CreateFolderQueryOverloadDefault, _
                  VarPtr(pIStorageFolderQueryResult)) = S_OK Then
            If pIStorageFolderQueryResult <> 0& Then
                Set Ret = New StorageFolderQueryResult
                Ret.Ifc = pIStorageFolderQueryResult
            End If
        End If
    End If
    Set CreateFolderQueryOverloadDefault = Ret
End Function

Public Function CreateFolderQuery(ByVal Query As CommonFolderQuery) As StorageFolderQueryResult
    Dim Ret As StorageFolderQueryResult
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIStorageFolderQueryResult As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_CreateFolderQuery, _
                  Query, _
                  VarPtr(pIStorageFolderQueryResult)) = S_OK Then
            If pIStorageFolderQueryResult <> 0& Then
                Set Ret = New StorageFolderQueryResult
                Ret.Ifc = pIStorageFolderQueryResult
            End If
        End If
    End If
    Set CreateFolderQuery = Ret
End Function

Public Function CreateFolderQueryWithOptions(ByVal Options As QueryOptions) As StorageFolderQueryResult
    Dim Ret As StorageFolderQueryResult
    If m_pIStorageFolderQueryOperations <> 0& And _
       IsNotNothing(Options) Then
        Dim pIStorageFolderQueryResult As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_CreateFolderQueryWithOptions, _
                  Options.Ifc, _
                  VarPtr(pIStorageFolderQueryResult)) = S_OK Then
            If pIStorageFolderQueryResult <> 0& Then
                Set Ret = New StorageFolderQueryResult
                Ret.Ifc = pIStorageFolderQueryResult
            End If
        End If
    End If
    Set CreateFolderQueryWithOptions = Ret
End Function

Public Function CreateItemQuery() As StorageItemQueryResult
    Dim Ret As StorageItemQueryResult
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIStorageItemQueryResult As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_CreateItemQuery, _
                  VarPtr(pIStorageItemQueryResult)) = S_OK Then
            If pIStorageItemQueryResult <> 0& Then
                Set Ret = New StorageItemQueryResult
                Ret.Ifc = pIStorageItemQueryResult
            End If
        End If
    End If
    Set CreateItemQuery = Ret
End Function

Public Function CreateItemQueryWithOptions(ByVal Options As QueryOptions) As StorageItemQueryResult
    Dim Ret As StorageItemQueryResult
    If m_pIStorageFolderQueryOperations <> 0& And _
       IsNotNothing(Options) Then
        Dim pIStorageItemQueryResult As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_CreateItemQueryWithOptions, _
                  Options.Ifc, _
                  VarPtr(pIStorageItemQueryResult)) = S_OK Then
            If pIStorageItemQueryResult <> 0& Then
               Set Ret = New StorageItemQueryResult
                Ret.Ifc = pIStorageItemQueryResult
            End If
        End If
    End If
    Set CreateItemQueryWithOptions = Ret
End Function

Public Function GetFilesAsync(ByVal Query As CommonFileQuery, _
                              ByVal startIndex As Long, _
                              ByVal maxItemsToRetriev As Long) As ReadOnlyList_1 'ReadOnlyList_StorageFile
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageFile
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_GetFilesAsync, _
                  Query, _
                  startIndex, _
                  maxItemsToRetriev, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageFile, pIVectorView) 'ReadOnlyList_StorageFile
                End If
            End If
        End If
    End If
    Set GetFilesAsync = Ret
End Function

Public Function GetFilesAsyncOverloadDefaultStartAndCount(ByVal Query As CommonFileQuery) As ReadOnlyList_1 'ReadOnlyList_StorageFile
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageFile
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_GetFilesAsyncOverloadDefaultStartAndCount, _
                  Query, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageFile, pIVectorView) 'ReadOnlyList_StorageFile
                End If
            End If
        End If
    End If
    Set GetFilesAsyncOverloadDefaultStartAndCount = Ret
End Function

Public Function GetFoldersAsync(ByVal Query As CommonFolderQuery, _
                                ByVal startIndex As Long, _
                                ByVal maxItemsToRetriev As Long) As ReadOnlyList_1 'ReadOnlyList_StorageFolder
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageFolder
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_GetFoldersAsync, _
                  Query, _
                  startIndex, _
                  maxItemsToRetriev, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageFolder, pIVectorView) 'ReadOnlyList_StorageFolder
                End If
            End If
        End If
    End If
    Set GetFoldersAsync = Ret
End Function

Public Function GetFoldersAsyncOverloadDefaultStartAndCount(ByVal Query As CommonFolderQuery) As ReadOnlyList_1 'ReadOnlyList_StorageFolder
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageFolder
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_GetFoldersAsyncOverloadDefaultStartAndCount, _
                  Query, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageFolder, pIVectorView) 'ReadOnlyList_StorageFolder
                End If
            End If
        End If
    End If
    Set GetFoldersAsyncOverloadDefaultStartAndCount = Ret
End Function

Public Function GetItemsAsync(ByVal startIndex As Long, _
                              ByVal maxItemsToRetriev As Long) As ReadOnlyList_1 'ReadOnlyList_StorageItem
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageItem
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_GetItemsAsync, _
                  startIndex, _
                  maxItemsToRetriev, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageItem, pIVectorView) 'ReadOnlyList_StorageItem
                End If
            End If
        End If
    End If
    Set GetItemsAsync = Ret
End Function

Public Function AreQueryOptionsSupported(ByVal Options As QueryOptions) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFolderQueryOperations <> 0& And _
       IsNotNothing(Options) Then
        Dim value As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_AreQueryOptionsSupported, _
                  Options.Ifc, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    AreQueryOptionsSupported = Ret
End Function

Public Function IsCommonFolderQuerySupported(ByVal Query As CommonFolderQuery) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_IsCommonFolderQuerySupported, _
                  Query, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsCommonFolderQuerySupported = Ret
End Function

Public Function IsCommonFileQuerySupported(ByVal Query As CommonFileQuery) As Boolean
    Dim Ret As Boolean
    If m_pIStorageFolderQueryOperations <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageFolderQueryOperations, _
                  IStorageFolderQueryOperations_IsCommonFileQuerySupported, _
                  Query, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsCommonFileQuerySupported = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



