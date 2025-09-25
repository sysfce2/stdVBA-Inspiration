VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageItemQueryResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Search_StorageItemQueryResult As String = "Windows.Storage.Search.StorageItemQueryResult"
Private Const IID_IStorageQueryResultBase As String = "{c297d70d-7353-47ab-ba58-8c61425dc54b}"
Private Const IID_IStorageItemQueryResult As String = "{e8948079-9d58-47b8-b2b2-41b07f4795f9}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IStorageItemQueryResult
    IStorageItemQueryResult_GetItemsAsync = 6
    IStorageItemQueryResult_GetItemsAsyncDefaultStartAndCount = 7
    
    ' IStorageQueryResultBase
    IStorageQueryResultBase_GetItemCountAsync = 6
    IStorageQueryResultBase_GetFolder = 7
    IStorageQueryResultBase_AddContentsChanged = 8
    IStorageQueryResultBase_RemoveContentsChanged = 9
    IStorageQueryResultBase_AddOptionsChanged = 10
    IStorageQueryResultBase_RemoveOptionsChanged = 11
    IStorageQueryResultBase_FindStartIndexAsync = 12
    IStorageQueryResultBase_GetCurrentQueryOptions = 13
    IStorageQueryResultBase_ApplyNewQueryOptions = 14
    
End Enum

' ----==== Variablen ====----
Private m_pIStorageQueryResultBase As Long
Private m_pIStorageItemQueryResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageQueryResultBase)
    Call ReleaseIfc(m_pIStorageItemQueryResult)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageItemQueryResult
End Property

Public Property Let Ifc(ByVal pIStorageItemQueryResult As Long)
    Call ReleaseIfc(m_pIStorageQueryResultBase)
    Call ReleaseIfc(m_pIStorageItemQueryResult)
    m_pIStorageItemQueryResult = pIStorageItemQueryResult
    Call QueryIfc(m_pIStorageItemQueryResult, _
                  IID_IStorageQueryResultBase, _
                  m_pIStorageQueryResultBase)
End Property

' IStorageQueryResultBase
Public Property Get Folder() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIStorageQueryResultBase <> 0& Then
        Dim pIStorageFolder As Long
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_GetFolder, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                Set Ret = New StorageFolder
                Ret.Ifc = pIStorageFolder
            End If
        End If
    End If
    Set Folder = Ret
End Property

' ----==== Functions ====----
' IStorageItemQueryResult
Public Function GetItemsAsync(ByVal startIndex As Long, _
                              ByVal maxNumberOfItems As Long) As ReadOnlyList_1 'ReadOnlyList_StorageItem
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageItem
    If m_pIStorageItemQueryResult <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageItemQueryResult, _
                  IStorageItemQueryResult_GetItemsAsync, _
                  startIndex, _
                  maxNumberOfItems, _
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

Public Function GetItemsAsyncDefaultStartAndCount() As ReadOnlyList_1 'ReadOnlyList_StorageItem
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageItem
    If m_pIStorageItemQueryResult <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageItemQueryResult, _
                  IStorageItemQueryResult_GetItemsAsyncDefaultStartAndCount, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageItem, pIVectorView) 'ReadOnlyList_StorageItem
                End If
            End If
        End If
    End If
    Set GetItemsAsyncDefaultStartAndCount = Ret
End Function

' IStorageQueryResultBase
Public Function GetItemCountAsync() As Long
    Dim Ret As Long
    If m_pIStorageQueryResultBase <> 0& Then
        Dim value As Long
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_GetItemCountAsync, _
                  VarPtr(value)) = S_OK Then
            If Await(value) Then
                Ret = value
            End If
        End If
    End If
    GetItemCountAsync = Ret
End Function

Public Function AddContentsChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIStorageQueryResultBase <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_AddContentsChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddContentsChanged = Ret
End Function

Public Function RemoveContentsChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIStorageQueryResultBase <> 0& And token <> 0@ Then
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_RemoveContentsChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveContentsChanged = Ret
End Function

Public Function AddOptionsChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIStorageQueryResultBase <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_AddOptionsChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddOptionsChanged = Ret
End Function

Public Function RemoveOptionsChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIStorageQueryResultBase <> 0& And token <> 0@ Then
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_RemoveOptionsChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveOptionsChanged = Ret
End Function

Public Function FindStartIndexAsync(ByVal value As Inspectable) As Long
    Dim Ret As Long
    If m_pIStorageQueryResultBase <> 0& And _
       IsNotNothing(value) Then
        Dim operation As Long
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_FindStartIndexAsync, _
                  value.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = operation
                End If
            End If
        End If
    End If
    FindStartIndexAsync = Ret
End Function

Public Function GetCurrentQueryOptions() As QueryOptions
    Dim Ret As QueryOptions
    If m_pIStorageQueryResultBase <> 0& Then
        Dim pIQueryOptions As Long
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_GetCurrentQueryOptions, _
                  VarPtr(pIQueryOptions)) = S_OK Then
            If pIQueryOptions <> 0& Then
                Set Ret = New QueryOptions
                Ret.Ifc = pIQueryOptions
            End If
        End If
    End If
    Set GetCurrentQueryOptions = Ret
End Function

Public Function ApplyNewQueryOptions(ByVal newQueryOptions As QueryOptions) As Boolean
    Dim Ret As Boolean
    If m_pIStorageQueryResultBase <> 0& And _
       IsNotNothing(newQueryOptions) Then
        If Invoke(m_pIStorageQueryResultBase, _
                  IStorageQueryResultBase_ApplyNewQueryOptions, _
                  newQueryOptions.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    ApplyNewQueryOptions = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


