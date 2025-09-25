VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSavePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Pickers_FileSavePicker As String = "Windows.Storage.Pickers.FileSavePicker"
Private Const IID_IFileSavePicker As String = "{3286ffcb-617f-4cc5-af6a-b3fdf29ad145}"
Private Const IID_IFileSavePicker2 As String = "{0ec313a2-d24b-449a-8197-e89104fd42cc}"
Private Const IID_IFileSavePicker3 As String = "{698aec69-ba3c-4e51-bd90-4abcbbf4cfaf}"
Private Const IID_IFileSavePicker4 As String = "{e7d83a5a-ddfa-5de0-8b70-c842c21988ec}"
Private Const IID_IFileSavePickerStatics As String = "{28e3cf9e-961c-5e2c-aed7-e64737f4ce37}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IFileSavePickerStatics
    IFileSavePickerStatics_CreateForUser = 6
    
    ' IFileSavePicker
    IFileSavePicker_GetSettingsIdentifier = 6
    IFileSavePicker_PutSettingsIdentifier = 7
    IFileSavePicker_GetSuggestedStartLocation = 8
    IFileSavePicker_PutSuggestedStartLocation = 9
    IFileSavePicker_GetCommitButtonText = 10
    IFileSavePicker_PutCommitButtonText = 11
    IFileSavePicker_GetFileTypeChoices = 12
    IFileSavePicker_GetDefaultFileExtension = 13
    IFileSavePicker_PutDefaultFileExtension = 14
    IFileSavePicker_GetSuggestedSaveFile = 15
    IFileSavePicker_PutSuggestedSaveFile = 16
    IFileSavePicker_GetSuggestedFileName = 17
    IFileSavePicker_PutSuggestedFileName = 18
    IFileSavePicker_PickSaveFileAsync = 19

    ' IFileSavePicker2
'    IFileSavePicker2_GetContinuationData = 6
    
    ' IFileSavePicker3
'    IFileSavePicker3_GetEnterpriseId = 6
'    IFileSavePicker3_PuttEnterpriseId = 7
    
    ' IFileSavePicker4
    IFileSavePicker4_GetUser = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIFileSavePicker As Long
Private m_pIFileSavePicker2 As Long
Private m_pIFileSavePicker3 As Long
Private m_pIFileSavePicker4 As Long
Private m_pIFileSavePickerStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_Pickers_FileSavePicker, _
                            IID_IFileSavePickerStatics, _
                            m_pIFileSavePickerStatics) Then
    End If
    Dim pIFileSavePicker As Long
    If GetActivateInstance(Windows_Storage_Pickers_FileSavePicker, _
                           IID_IFileSavePicker, _
                           pIFileSavePicker) Then
        Me.Ifc = pIFileSavePicker
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIFileSavePicker4)
    Call ReleaseIfc(m_pIFileSavePicker3)
    Call ReleaseIfc(m_pIFileSavePicker2)
    Call ReleaseIfc(m_pIFileSavePicker)
    Call ReleaseIfc(m_pIFileSavePickerStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIFileSavePicker
End Property

Public Property Let Ifc(ByVal pIFileSavePicker As Long)
    Call ReleaseIfc(m_pIFileSavePicker4)
    Call ReleaseIfc(m_pIFileSavePicker3)
    Call ReleaseIfc(m_pIFileSavePicker2)
    Call ReleaseIfc(m_pIFileSavePicker)
    m_pIFileSavePicker = pIFileSavePicker
    Call QueryIfc(m_pIFileSavePicker, IID_IFileSavePicker2, m_pIFileSavePicker2)
    Call QueryIfc(m_pIFileSavePicker, IID_IFileSavePicker3, m_pIFileSavePicker3)
    Call QueryIfc(m_pIFileSavePicker, IID_IFileSavePicker4, m_pIFileSavePicker4)
End Property

' IFileSavePicker
Public Property Get SettingsIdentifier() As String
    Dim Ret As String
    If m_pIFileSavePicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_GetSettingsIdentifier, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SettingsIdentifier = Ret
End Property

Public Property Let SettingsIdentifier(ByVal value As String)
    If m_pIFileSavePicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If hString <> 0& Then
            If Invoke(m_pIFileSavePicker, _
                      IFileSavePicker_PutSettingsIdentifier, _
                      hString) = S_OK Then
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
End Property

Public Property Get SuggestedStartLocation() As PickerLocationId
    Dim Ret As PickerLocationId
    If m_pIFileSavePicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_GetSuggestedStartLocation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SuggestedStartLocation = Ret
End Property

Public Property Let SuggestedStartLocation(ByVal value As PickerLocationId)
    If m_pIFileSavePicker <> 0& Then
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_PutSuggestedStartLocation, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get CommitButtonText() As String
    Dim Ret As String
    If m_pIFileSavePicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_GetCommitButtonText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CommitButtonText = Ret
End Property

Public Property Let CommitButtonText(ByVal value As String)
    If m_pIFileSavePicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If hString <> 0& Then
            If Invoke(m_pIFileSavePicker, _
                      IFileSavePicker_PutCommitButtonText, _
                      hString) = S_OK Then
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
End Property

Public Property Get FileTypeChoices() As FilePickerFileTypesOrderedMap
    Dim Ret As FilePickerFileTypesOrderedMap
    If m_pIFileSavePicker <> 0& Then
        Dim pIMap_HSTRING_IVector_HSTRING As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_GetFileTypeChoices, _
                  VarPtr(pIMap_HSTRING_IVector_HSTRING)) = S_OK Then
            If pIMap_HSTRING_IVector_HSTRING <> 0& Then
                Set Ret = New FilePickerFileTypesOrderedMap
                Ret.Ifc = pIMap_HSTRING_IVector_HSTRING
            End If
        End If
    End If
    Set FileTypeChoices = Ret
End Property

Public Property Get DefaultFileExtension() As String
    Dim Ret As String
    If m_pIFileSavePicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_GetDefaultFileExtension, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DefaultFileExtension = Ret
End Property

Public Property Let DefaultFileExtension(ByVal value As String)
    If m_pIFileSavePicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If hString <> 0& Then
            If Invoke(m_pIFileSavePicker, _
                      IFileSavePicker_PutDefaultFileExtension, _
                      hString) = S_OK Then
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
End Property

Public Property Get SuggestedSaveFile() As StorageFile
    Dim Ret As StorageFile
    If m_pIFileSavePicker <> 0& Then
        Dim pIStorageFile As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_GetSuggestedSaveFile, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                Set Ret = New StorageFile
                Ret.Ifc = pIStorageFile
            End If
        End If
    End If
    Set SuggestedSaveFile = Ret
End Property

Public Property Let SuggestedSaveFile(ByVal value As StorageFile)
    If m_pIFileSavePicker <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_PutSuggestedSaveFile, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get SuggestedFileName() As String
    Dim Ret As String
    If m_pIFileSavePicker <> 0& Then
        Dim value As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_GetSuggestedFileName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SuggestedFileName = Ret
End Property

Public Property Let SuggestedFileName(ByVal value As String)
    If m_pIFileSavePicker <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If hString <> 0& Then
            If Invoke(m_pIFileSavePicker, _
                      IFileSavePicker_PutSuggestedFileName, _
                      hString) = S_OK Then
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
End Property

' IFileSavePicker2
'    IFileSavePicker2_GetContinuationData=6

' IFileSavePicker3
'    IFileSavePicker3_GetEnterpriseId = 6
'    IFileSavePicker3_PuttEnterpriseId = 7

' IFileSavePicker4
Public Property Get User() As User
    Dim Ret As User
    If m_pIFileSavePicker4 <> 0& Then
        Dim pIUser As Long
        If Invoke(m_pIFileSavePicker4, _
                  IFileSavePicker4_GetUser, _
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
    If m_pIFileSavePicker <> 0& Then
        If InitWithWindow(m_pIFileSavePicker, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' IFileSavePickerStatics
Public Function CreateForUser(ByVal value As User) As FileSavePicker
    Dim Ret As FileSavePicker
    If m_pIFileSavePickerStatics <> 0& And _
       IsNotNothing(value) Then
        Dim pIFileSavePicker As Long
        If Invoke(m_pIFileSavePickerStatics, _
                  IFileSavePickerStatics_CreateForUser, _
                  value.Ifc, _
                  VarPtr(pIFileSavePicker)) = S_OK Then
            If pIFileSavePicker <> 0& Then
                Set Ret = New FileSavePicker
                Ret.Ifc = pIFileSavePicker
            End If
        End If
    End If
    Set CreateForUser = Ret
End Function

' IFileSavePicker
Public Function PickSaveFileAsync() As StorageFile
    Dim Ret As StorageFile
    If m_pIFileSavePicker <> 0& Then
        Dim pIStorageFile As Long
        If Invoke(m_pIFileSavePicker, _
                  IFileSavePicker_PickSaveFileAsync, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                If Await(pIStorageFile) Then
                    Set Ret = New StorageFile
                    Ret.Ifc = pIStorageFile
                End If
            End If
        End If
    End If
    Set PickSaveFileAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


