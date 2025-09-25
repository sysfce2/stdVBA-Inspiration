VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileOpenPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Pickers_FileOpenPicker As String = "Windows.Storage.Pickers.FileOpenPicker"
Private Const IID_IFileOpenPicker As String = "{2ca8278a-12c5-4c5f-8977-94547793c241}"
Private Const IID_IFileOpenPicker3 As String = "{d9a5c5b3-c5dc-5b98-bd80-a8d0ca0584d8}"
Private Const IID_IFileOpenPickerStatics2 As String = "{e8917415-eddd-5c98-b6f3-366fdfcad392}"
Private Const IID_IFileOpenPickerWithOperationId As String = "{3f57b569-2522-4ca5-aa73-a15509f1fcbf}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IFileOpenPickerStatics2
    IFileOpenPickerStatics2_CreateForUser = 6

    ' IFileOpenPicker
    IFileOpenPicker_GetViewMode = 6
    IFileOpenPicker_PutViewMode = 7
    IFileOpenPicker_GetSettingsIdentifier = 8
    IFileOpenPicker_PutSettingsIdentifier = 9
    IFileOpenPicker_GetSuggestedStartLocation = 10
    IFileOpenPicker_PutSuggestedStartLocation = 11
    IFileOpenPicker_GetCommitButtonText = 12
    IFileOpenPicker_PutCommitButtonText = 13
    IFileOpenPicker_GetFileTypeFilter = 14
    IFileOpenPicker_PickSingleFileAsync = 15
    IFileOpenPicker_PickMultipleFilesAsync = 16
    
    ' IFileOpenPicker3
    IFileOpenPicker_GetUser = 6
    
    ' IFileOpenPickerWithOperationId
    IFileOpenPickerWithOperationId_PickSingleFileAsync = 6
    
    ' IVectorView
    IVectorView_GetAt = 6
    IVectorView_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIFileOpenPicker As Long
Private m_pIFileOpenPicker3 As Long
Private m_pIFileOpenPickerStatics2 As Long
Private m_pIFileOpenPickerWithOperationId As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_Pickers_FileOpenPicker, _
                            IID_IFileOpenPickerStatics2, _
                            m_pIFileOpenPickerStatics2) Then
    End If
    Dim pIFileOpenPicker As Long
    If GetActivateInstance(Windows_Storage_Pickers_FileOpenPicker, _
                           IID_IFileOpenPicker, _
                           pIFileOpenPicker) Then
        Me.Ifc = pIFileOpenPicker
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIFileOpenPickerWithOperationId)
    Call ReleaseIfc(m_pIFileOpenPicker3)
    Call ReleaseIfc(m_pIFileOpenPicker)
    Call ReleaseIfc(m_pIFileOpenPickerStatics2)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIFileOpenPicker
End Property

Public Property Let Ifc(ByVal pIFileOpenPicker As Long)
    Call ReleaseIfc(m_pIFileOpenPickerWithOperationId)
    Call ReleaseIfc(m_pIFileOpenPicker3)
    Call ReleaseIfc(m_pIFileOpenPicker)
    m_pIFileOpenPicker = pIFileOpenPicker
    Call QueryIfc(m_pIFileOpenPicker, IID_IFileOpenPicker3, m_pIFileOpenPicker3)
    Call QueryIfc(m_pIFileOpenPicker, IID_IFileOpenPickerWithOperationId, m_pIFileOpenPickerWithOperationId)
End Property

' IFileOpenPicker
Public Property Get ViewMode() As PickerViewMode
    Dim Ret As PickerViewMode
    If m_pIFileOpenPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_GetViewMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ViewMode = Ret
End Property

Public Property Let ViewMode(ByVal value As PickerViewMode)
    If m_pIFileOpenPicker <> 0& Then
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_PutViewMode, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get SettingsIdentifier() As String
    Dim Ret As String
    If m_pIFileOpenPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_GetSettingsIdentifier, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SettingsIdentifier = Ret
End Property

Public Property Let SettingsIdentifier(ByVal value As String)
    If m_pIFileOpenPicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_PutSettingsIdentifier, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get SuggestedStartLocation() As PickerLocationId
    Dim Ret As PickerLocationId
    If m_pIFileOpenPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_GetSuggestedStartLocation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SuggestedStartLocation = Ret
End Property

Public Property Let SuggestedStartLocation(ByVal value As PickerLocationId)
    If m_pIFileOpenPicker <> 0& Then
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_PutSuggestedStartLocation, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get CommitButtonText() As String
    Dim Ret As String
    If m_pIFileOpenPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_GetCommitButtonText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CommitButtonText = Ret
End Property

Public Property Let CommitButtonText(ByVal value As String)
    If m_pIFileOpenPicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_PutCommitButtonText, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get FileTypeFilter() As List_String
    Dim Ret As List_String
    If m_pIFileOpenPicker <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_GetFileTypeFilter, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set FileTypeFilter = Ret
End Property

' IFileOpenPicker3
Public Property Get User() As User
    Dim Ret As User
    If m_pIFileOpenPicker <> 0& Then
        Dim pIUser As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_GetUser, _
                  VarPtr(pIUser)) = S_OK Then
            If pIUser <> 0& Then
                Set Ret = New User
                Ret.Ifc = pIUser
            End If
        End If
    End If
    Set User = Ret
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pIFileOpenPicker <> 0& Then
        If InitWithWindow(m_pIFileOpenPicker, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' IFileOpenPickerStatics2
Public Function CreateForUser(ByVal value As User) As FileOpenPicker
    Dim Ret As FileOpenPicker
    If m_pIFileOpenPickerStatics2 <> 0& Then
        Dim pIFileOpenPicker As Long
        If Invoke(m_pIFileOpenPickerStatics2, _
                  IFileOpenPickerStatics2_CreateForUser, _
                  value.Ifc, _
                  VarPtr(pIFileOpenPicker)) = S_OK Then
            If pIFileOpenPicker <> 0& Then
                Set Ret = New FileOpenPicker
                Ret.Ifc = pIFileOpenPicker
            End If
        End If
    End If
    Set CreateForUser = Ret
End Function

' IFileOpenPicker
Public Function PickSingleFileAsync() As StorageFile
    Dim Ret As StorageFile
    If m_pIFileOpenPicker <> 0& Then
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = Me.FileTypeFilter
        If IsNotNothing(FileTypeFilter) Then
            If FileTypeFilter.Size = 0 Then
                Call FileTypeFilter.Append("*")
            End If
            Set FileTypeFilter = Nothing
        End If
        Dim pIStorageFile As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_PickSingleFileAsync, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                If Await(pIStorageFile) Then
                    Set Ret = New StorageFile
                    Ret.Ifc = pIStorageFile
                End If
            End If
        End If
    End If
    Set PickSingleFileAsync = Ret
End Function

Public Function PickMultipleFilesAsync() As ReadOnlyList_1 'ReadOnlyList_StorageFile
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageFile
    If m_pIFileOpenPicker <> 0& Then
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = Me.FileTypeFilter
        If IsNotNothing(FileTypeFilter) Then
            If FileTypeFilter.Size = 0 Then
                Call FileTypeFilter.Append("*")
            End If
            Set FileTypeFilter = Nothing
        End If
        Dim pIVectorView As Long
        If Invoke(m_pIFileOpenPicker, _
                  IFileOpenPicker_PickMultipleFilesAsync, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageFile, pIVectorView) 'ReadOnlyList_StorageFile
                End If
            End If
        End If
    End If
    Set PickMultipleFilesAsync = Ret
End Function

' IFileOpenPickerWithOperationId
Public Function PickSingleFileWithOperationIdAsync(ByVal pickerOperationId As String) As StorageFile
    Dim Ret As StorageFile
    If m_pIFileOpenPickerWithOperationId <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(pickerOperationId)
        Dim pIStorageFile As Long
        If Invoke(m_pIFileOpenPickerWithOperationId, _
                  IFileOpenPickerWithOperationId_PickSingleFileAsync, _
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
    Set PickSingleFileWithOperationIdAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


