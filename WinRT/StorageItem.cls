VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_IStorageItem As String = "Windows.Storage.IStorageItem"
Private Const IID_IStorageItem As String = "{4207a996-ca2f-42f7-bde8-8b10457a7f30}"
Private Const IID_IStorageItem2 As String = "{53f926d2-083c-4283-b45b-81c007237e44}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
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

End Enum

' ----==== Vars ====----
Private m_pIStorageItem As Long
Private m_pIStorageItem2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageItem)
    Call ReleaseIfc(m_pIStorageItem2)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageItem
End Property

Public Property Let Ifc(ByVal pIStorageItem As Long)
    Call ReleaseIfc(m_pIStorageItem)
    Call ReleaseIfc(m_pIStorageItem2)
    m_pIStorageItem = pIStorageItem
    Call QueryIfc(m_pIStorageItem, _
                  IID_IStorageItem2, _
                  m_pIStorageItem2)
End Property

' IStorageItem
Public Property Get Name() As String
    Dim Ret As String
    If m_pIStorageItem <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Name = Ret
End Property

Public Property Get Path() As String
    Dim Ret As String
    If m_pIStorageItem <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetPath, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Path = Ret
End Property

Public Property Get Attributes() As FileAttributes
    Dim Ret As FileAttributes
    If m_pIStorageItem <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetAttributes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Attributes = Ret
End Property

Public Property Get DateCreated() As DateTime
    Dim Ret As DateTime
    If m_pIStorageItem <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIStorageItem, _
                  IStorageItem_GetDateCreated, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = Value
        End If
    End If
    Set DateCreated = Ret
End Property

' ----==== Functions ====----
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
        Dim Value As Long
        If Invoke(m_pIStorageItem, _
                  IStorageItem_IsOfType, _
                  itemType, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
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
        Dim Value As Long
        If Invoke(m_pIStorageItem2, _
                  IStorageItem2_IsEqual, _
                  Item.IStorageItem, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    IsEqual = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


