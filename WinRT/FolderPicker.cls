VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Pickers_FolderPicker As String = "Windows.Storage.Pickers.FolderPicker"
Private Const IID_IFolderPicker As String = "{084f7799-f3fb-400a-99b1-7b4a772fd60d}"
Private Const IID_IFolderPicker2 As String = "{8eb3ba97-dc85-4616-be94-9660881f2f5d}"
Private Const IID_IFolderPicker3 As String = "{673b1e29-d326-53c0-bd24-a25c714cee36}"
Private Const IID_IFolderPickerStatics As String = "{9be34740-7ca1-5942-a3c8-46f2551ecff3}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IFolderPickerStatics
    IFolderPickerStatics_CreateForUser = 6
    
    ' IFolderPicker
    IFolderPicker_GetViewMode = 6
    IFolderPicker_PutViewMode = 7
    IFolderPicker_GetSettingsIdentifier = 8
    IFolderPicker_PutSettingsIdentifier = 9
    IFolderPicker_GetSuggestedStartLocation = 10
    IFolderPicker_PutSuggestedStartLocation = 11
    IFolderPicker_GetCommitButtonText = 12
    IFolderPicker_PutCommitButtonText = 13
    IFolderPicker_GetFileTypeFilter = 14
    IFolderPicker_PickSingleFolderAsync = 15

    ' IFolderPicker2
'    IFolderPicker2_GetContinuationData = 6

    ' IFolderPicker3
    IFolderPicker3_GetUser = 6

End Enum

' ----==== Variablen ====----
Private m_pIFolderPicker As Long
Private m_pIFolderPicker2 As Long
Private m_pIFolderPicker3 As Long
Private m_pIFolderPickerStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()

    If GetActivationFactory(Windows_Storage_Pickers_FolderPicker, _
                            IID_IFolderPickerStatics, _
                            m_pIFolderPickerStatics) Then
    End If
    Dim pIFolderPicker As Long
    If GetActivateInstance(Windows_Storage_Pickers_FolderPicker, _
                           IID_IFolderPicker, _
                           pIFolderPicker) Then
        Me.Ifc = pIFolderPicker
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIFolderPicker3)
    Call ReleaseIfc(m_pIFolderPicker2)
    Call ReleaseIfc(m_pIFolderPicker)
    Call ReleaseIfc(m_pIFolderPickerStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIFolderPicker
End Property

Public Property Let Ifc(ByVal pIFolderPicker As Long)
    Call ReleaseIfc(m_pIFolderPicker3)
    Call ReleaseIfc(m_pIFolderPicker2)
    Call ReleaseIfc(m_pIFolderPicker)
    m_pIFolderPicker = pIFolderPicker
    Call QueryIfc(m_pIFolderPicker, IID_IFolderPicker2, m_pIFolderPicker2)
    Call QueryIfc(m_pIFolderPicker, IID_IFolderPicker3, m_pIFolderPicker3)
End Property

' IFolderPicker
Public Property Get ViewMode() As PickerViewMode
    Dim Ret As PickerViewMode
    If m_pIFolderPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_GetViewMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ViewMode = Ret
End Property

Public Property Let ViewMode(ByVal value As PickerViewMode)
    If m_pIFolderPicker <> 0& Then
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_PutViewMode, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get SettingsIdentifier() As String
    Dim Ret As String
    If m_pIFolderPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_GetSettingsIdentifier, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SettingsIdentifier = Ret
End Property

Public Property Let SettingsIdentifier(ByVal value As String)
    If m_pIFolderPicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_PutSettingsIdentifier, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get SuggestedStartLocation() As PickerLocationId
    Dim Ret As PickerLocationId
    If m_pIFolderPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_GetSuggestedStartLocation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SuggestedStartLocation = Ret
End Property

Public Property Let SuggestedStartLocation(ByVal value As PickerLocationId)
    If m_pIFolderPicker <> 0& Then
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_PutSuggestedStartLocation, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get CommitButtonText() As String
    Dim Ret As String
    If m_pIFolderPicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_GetCommitButtonText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CommitButtonText = Ret
End Property

Public Property Let CommitButtonText(ByVal value As String)
    If m_pIFolderPicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_PutCommitButtonText, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get FileTypeFilter() As List_String
    Dim Ret As List_String
    If m_pIFolderPicker <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_GetFileTypeFilter, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set FileTypeFilter = Ret
End Property

' IFolderPicker2
'    IFolderPicker2_GetContinuationData = 6

' IFolderPicker3
Public Property Get User() As User
    Dim Ret As User
    If m_pIFolderPicker3 <> 0& Then
        Dim pIUser As Long
        If Invoke(m_pIFolderPicker3, _
                  IFolderPicker3_GetUser, _
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
    If m_pIFolderPicker <> 0& Then
        If InitWithWindow(m_pIFolderPicker, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' IFolderPickerStatics
Public Function CreateForUser(ByVal value As User) As FolderPicker
    Dim Ret As FolderPicker
    If m_pIFolderPickerStatics <> 0& And _
       IsNotNothing(value) Then
        Dim pIFolderPicker As Long
        If Invoke(m_pIFolderPickerStatics, _
                  IFolderPickerStatics_CreateForUser, _
                  value.Ifc, _
                  VarPtr(pIFolderPicker)) = S_OK Then
            If pIFolderPicker <> 0& Then
                Set Ret = New FolderPicker
                Ret.Ifc = pIFolderPicker
            End If
        End If
    End If
    Set CreateForUser = Ret
End Function

' IFolderPicker
Public Function PickSingleFolderAsync() As StorageFolder
    Dim Ret As StorageFolder
    If m_pIFolderPicker <> 0& Then
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = Me.FileTypeFilter
        If IsNotNothing(FileTypeFilter) Then
            If FileTypeFilter.Size = 0 Then
                Call FileTypeFilter.Append("*")
            End If
            Set FileTypeFilter = Nothing
        End If
        Dim pIStorageFolder As Long
        If Invoke(m_pIFolderPicker, _
                  IFolderPicker_PickSingleFolderAsync, _
                  VarPtr(pIStorageFolder)) = S_OK Then
            If pIStorageFolder <> 0& Then
                If Await(pIStorageFolder) Then
                    Set Ret = New StorageFolder
                    Ret.Ifc = pIStorageFolder
                End If
            End If
        End If
    End If
    Set PickSingleFolderAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


