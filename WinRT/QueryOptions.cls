VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Search_QueryOptions As String = "Windows.Storage.Search.QueryOptions"
Private Const IID_IQueryOptions As String = "{1e5e46ee-0f45-4838-a8e9-d0479d446c30}"
Private Const IID_IQueryOptionsFactory As String = "{032e1f8c-a9c1-4e71-8011-0dee9d4811a3}"
Private Const IID_IQueryOptionsWithProviderFilter As String = "{5b9d1026-15c4-44dd-b89a-47a59b7d7c4f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IQueryOptionsFactory
    IQueryOptionsFactory_CreateCommonFileQuery = 6
    IQueryOptionsFactory_CreateCommonFolderQuery = 7
    
    ' IQueryOptions
    IQueryOptions_GetFileTypeFilter = 6
    IQueryOptions_GetFolderDepth = 7
    IQueryOptions_PutFolderDepth = 8
    IQueryOptions_GetApplicationSearchFilter = 9
    IQueryOptions_PutApplicationSearchFilter = 10
    IQueryOptions_GetUserSearchFilter = 11
    IQueryOptions_PutUserSearchFilter = 12
    IQueryOptions_GetLanguage = 13
    IQueryOptions_PutLanguage = 14
    IQueryOptions_GetIndexerOption = 15
    IQueryOptions_PutIndexerOption = 16
    IQueryOptions_GetSortOrder = 17
    IQueryOptions_GetGroupPropertyName = 18
    IQueryOptions_GetDateStackOption = 19
    IQueryOptions_SaveToString = 20
    IQueryOptions_LoadFromString = 21
    IQueryOptions_SetThumbnailPrefetch = 22
'    IQueryOptions_SetPropertyPrefetch = 23
    
    ' IQueryOptionsWithProviderFilter
    IQueryOptionsWithProviderFilter_GetStorageProviderIdFilter = 6
        
    ' IVector_xxx
    IVector_GetAt = 6
    IVector_GetSize = 7
    IVector_GetView = 8
    IVector_IndexOf = 9
    IVector_SetAt = 10
    IVector_InsertAt = 11
    IVector_RemoveAt = 12
    IVector_Append = 13
    IVector_RemoveAtEnd = 14
    IVector_Clear = 15
    IVector_GetMany = 16
        
End Enum

' ----==== Variablen ====----
Private m_pIQueryOptions As Long
Private m_pIQueryOptionsFactory As Long
Private m_pIQueryOptionsWithProviderFilter As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_Search_QueryOptions, _
                            IID_IQueryOptionsFactory, _
                            m_pIQueryOptionsFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIQueryOptions)
    Call ReleaseIfc(m_pIQueryOptionsFactory)
    Call ReleaseIfc(m_pIQueryOptionsWithProviderFilter)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIQueryOptions
End Property

Public Property Let Ifc(ByVal pIQueryOptions As Long)
    Call ReleaseIfc(m_pIQueryOptions)
    Call ReleaseIfc(m_pIQueryOptionsWithProviderFilter)
    m_pIQueryOptions = pIQueryOptions
    Call QueryIfc(m_pIQueryOptions, _
                  IID_IQueryOptionsWithProviderFilter, _
                  m_pIQueryOptionsWithProviderFilter)
End Property

' IQueryOptions
Public Property Get FileTypeFilter() As List_String
    Dim Ret As List_String
    If m_pIQueryOptions <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetFileTypeFilter, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIVector
            End If
        End If
    Else
        If m_pIQueryOptionsFactory <> 0& Then
            Dim pIQueryOptions As Long
            If Invoke(m_pIQueryOptionsFactory, _
                      IQueryOptionsFactory_CreateCommonFileQuery, _
                      0&, _
                      0&, _
                      VarPtr(pIQueryOptions)) = S_OK Then
                Dim pIVector_HSTRING As Long
                If Invoke(pIQueryOptions, _
                          IQueryOptions_GetFileTypeFilter, _
                          VarPtr(pIVector_HSTRING)) = S_OK Then
                    If pIVector_HSTRING <> 0& Then
                        Set Ret = New List_String
                        Ret.Ifc = pIVector_HSTRING
                    End If
                End If
                Call ReleaseIfc(pIQueryOptions)
            End If
        End If
    End If
    Set FileTypeFilter = Ret
End Property

Public Property Get FolderDepth() As FolderDepth
    Dim Ret As FolderDepth
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetFolderDepth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FolderDepth = Ret
End Property

Public Property Let FolderDepth(ByVal value As FolderDepth)
    If m_pIQueryOptions <> 0& Then
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_PutFolderDepth, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get ApplicationSearchFilter() As String
    Dim Ret As String
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetApplicationSearchFilter, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ApplicationSearchFilter = Ret
End Property

Public Property Let ApplicationSearchFilter(ByVal value As String)
    If m_pIQueryOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_PutApplicationSearchFilter, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get UserSearchFilter() As String
    Dim Ret As String
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetUserSearchFilter, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    UserSearchFilter = Ret
End Property

Public Property Let UserSearchFilter(ByVal value As String)
    If m_pIQueryOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_PutUserSearchFilter, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Language() As String
    Dim Ret As String
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Language = Ret
End Property

Public Property Let Language(ByVal value As String)
    If m_pIQueryOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_PutLanguage, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get IndexerOption() As IndexerOption
    Dim Ret As IndexerOption
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetIndexerOption, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IndexerOption = Ret
End Property

Public Property Let IndexerOption(ByVal value As IndexerOption)
    If m_pIQueryOptions <> 0& Then
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_PutIndexerOption, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get SortOrder() As ReadOnlyList_1 'ReadOnlyList_SortEntry
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_SortEntry
    If m_pIQueryOptions <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetSortOrder, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_SortEntry, pIVectorView) 'ReadOnlyList_SortEntry
            End If
        End If
    End If
    Set SortOrder = Ret
End Property

Public Property Get GroupPropertyName() As String
    Dim Ret As String
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetGroupPropertyName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GroupPropertyName = Ret
End Property

Public Property Get DateStackOption() As DateStackOption
    Dim Ret As DateStackOption
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_GetDateStackOption, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DateStackOption = Ret
End Property

' IQueryOptionsWithProviderFilter
Public Property Get StorageProviderIdFilter() As List_String
    Dim Ret As List_String
    If m_pIQueryOptionsWithProviderFilter <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIQueryOptionsWithProviderFilter, _
                  IQueryOptionsWithProviderFilter_GetStorageProviderIdFilter, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set StorageProviderIdFilter = Ret
End Property

' ----==== Functions ====----
' IQueryOptionsFactory
Public Function CreateCommonFileQuery(ByVal Query As CommonFileQuery, _
                                      ByVal Filter As List_String) As Boolean
    Dim Ret As Boolean
    If m_pIQueryOptionsFactory <> 0& And _
       Filter.Ifc Then
        Dim pIQueryOptions As Long
        If Invoke(m_pIQueryOptionsFactory, _
                  IQueryOptionsFactory_CreateCommonFileQuery, _
                  Query, _
                  Filter.Ifc, _
                  VarPtr(pIQueryOptions)) = S_OK Then
            Me.Ifc = pIQueryOptions
            Ret = True
        End If
    End If
    CreateCommonFileQuery = Ret
End Function

Public Function CreateCommonFolderQuery(ByVal Query As CommonFolderQuery) As Boolean
    Dim Ret As Boolean
    If m_pIQueryOptionsFactory <> 0& Then
        Dim pIQueryOptions As Long
        If Invoke(m_pIQueryOptionsFactory, _
                  IQueryOptionsFactory_CreateCommonFolderQuery, _
                  Query, _
                  VarPtr(pIQueryOptions)) = S_OK Then
            Me.Ifc = pIQueryOptions
            Ret = True
        End If
    End If
    CreateCommonFolderQuery = Ret
End Function

' IQueryOptions
Public Function SaveToString() As String
    Dim Ret As String
    If m_pIQueryOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_SaveToString, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SaveToString = Ret
End Function

Public Function LoadFromString(ByVal value As String) As Boolean
    Dim Ret As Boolean
    If m_pIQueryOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_LoadFromString, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    LoadFromString = Ret
End Function

Public Function SetThumbnailPrefetch(ByVal Mode As ThumbnailMode, _
                                     ByVal requestedSize As Long, _
                                     ByVal Options As ThumbnailOptions) As Boolean
    Dim Ret As Boolean
    If m_pIQueryOptions <> 0& Then
        If Invoke(m_pIQueryOptions, _
                  IQueryOptions_SetThumbnailPrefetch, _
                  Mode, _
                  requestedSize, _
                  Options) = S_OK Then
            Ret = True
        End If
    End If
    SetThumbnailPrefetch = Ret
End Function

'    IQueryOptions_SetPropertyPrefetch = 23

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


