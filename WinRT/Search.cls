VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Search

Option Explicit

' ----==== Enums Windows.Storage.Search.xxx ====----
Public Enum IndexedState
    IndexedState_Unknown = 0
    IndexedState_NotIndexed = 1
    IndexedState_PartiallyIndexed = 2
    IndexedState_FullyIndexed = 3
End Enum

Public Enum FolderDepth
    FolderDepth_Shallow = 0
    FolderDepth_Deep = 1
End Enum

Public Enum DateStackOption
    DateStackOption_None = 0
    DateStackOption_Year = 1
    DateStackOption_Month = 2
End Enum

Public Enum IndexerOption
    IndexerOption_UseIndexerWhenAvailable = 0
    IndexerOption_OnlyUseIndexer = 1
    IndexerOption_DoNotUseIndexer = 2
    IndexerOption_OnlyUseIndexerAndOptimizeForIndexedProperties = 3
End Enum

Public Enum CommonFileQuery
    CommonFileQuery_DefaultQuery = 0
    CommonFileQuery_OrderByName = 1
    CommonFileQuery_OrderByTitle = 2
    CommonFileQuery_OrderByMusicProperties = 3
    CommonFileQuery_OrderBySearchRank = 4
    CommonFileQuery_OrderByDate = 5
End Enum

Public Enum CommonFolderQuery
    CommonFolderQuery_DefaultQuery = 0
    CommonFolderQuery_GroupByYear = 100
    CommonFolderQuery_GroupByMonth = 101
    CommonFolderQuery_GroupByArtist = 102
    CommonFolderQuery_GroupByAlbum = 103
    CommonFolderQuery_GroupByAlbumArtist = 104
    CommonFolderQuery_GroupByComposer = 105
    CommonFolderQuery_GroupByGenre = 106
    CommonFolderQuery_GroupByPublishedYear = 107
    CommonFolderQuery_GroupByRating = 108
    CommonFolderQuery_GroupByTag = 109
    CommonFolderQuery_GroupByAuthor = 110
    CommonFolderQuery_GroupByType = 111
End Enum

' ----==== Namespace Search.SortEntry ====----
Public Property Get SortEntry() As SortEntry
    Set SortEntry = New SortEntry
End Property

' ----==== Namespace Search.QueryOptions ====----
Public Property Get QueryOptions() As QueryOptions
    Set QueryOptions = New QueryOptions
End Property

' ----==== Namespace Search.StorageItemQueryResult ====----
Public Property Get StorageItemQueryResult() As StorageItemQueryResult
    Set StorageItemQueryResult = New StorageItemQueryResult
End Property

' ----==== Namespace Search.StorageFileQueryResult ====----
Public Property Get StorageFileQueryResult() As StorageFileQueryResult
    Set StorageFileQueryResult = New StorageFileQueryResult
End Property

' ----==== Namespace Search.StorageFolderQueryResult ====----
Public Property Get StorageFolderQueryResult() As StorageFolderQueryResult
    Set StorageFolderQueryResult = New StorageFolderQueryResult
End Property


