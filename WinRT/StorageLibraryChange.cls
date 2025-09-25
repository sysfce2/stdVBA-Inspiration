VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageLibraryChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_StorageLibraryChange As String = "Windows.Storage.StorageLibraryChange"
Private Const IID_IStorageLibraryChange As String = "{00980b23-2be2-4909-aa48-159f5203a51e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IStorageLibraryChange
    IStorageLibraryChange_GetChangeType = 6
    IStorageLibraryChange_GetPath = 7
    IStorageLibraryChange_GetPreviousPath = 8
    IStorageLibraryChange_IsOfType = 9
    IStorageLibraryChange_GetStorageItemAsync = 10

End Enum

' ----==== Variablen ====----
Private m_pIStorageLibraryChange As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageLibraryChange)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageLibraryChange
End Property

Public Property Let Ifc(ByVal pIStorageLibraryChange As Long)
    Call ReleaseIfc(m_pIStorageLibraryChange)
    m_pIStorageLibraryChange = pIStorageLibraryChange
End Property

' IStorageLibraryChange
Public Property Get ChangeType() As StorageLibraryChangeType
    Dim Ret As StorageLibraryChangeType
    If m_pIStorageLibraryChange <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageLibraryChange, _
                  IStorageLibraryChange_GetChangeType, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    ChangeType = Ret
End Property

Public Property Get Path() As String
    Dim Ret As String
    If m_pIStorageLibraryChange <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageLibraryChange, _
                  IStorageLibraryChange_GetPath, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Path = Ret
End Property

Public Property Get PreviousPath() As String
    Dim Ret As String
    If m_pIStorageLibraryChange <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageLibraryChange, _
                  IStorageLibraryChange_GetPreviousPath, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    PreviousPath = Ret
End Property

' ----==== Functions ====----
' IStorageLibraryChange
Public Function IsOfType() As StorageItemTypes
    Dim Ret As StorageItemTypes
    If m_pIStorageLibraryChange <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageLibraryChange, _
                  IStorageLibraryChange_IsOfType, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    IsOfType = Ret
End Function

Public Function GetStorageItemAsync() As StorageItem
    Dim Ret As StorageItem
    If m_pIStorageLibraryChange <> 0& Then
        Dim pIStorageItem As Long
        If Invoke(m_pIStorageLibraryChange, _
                  IStorageLibraryChange_GetStorageItemAsync, _
                  VarPtr(pIStorageItem)) = S_OK Then
            If pIStorageItem <> 0& Then
                If Await(pIStorageItem) Then
                    Set Ret = New StorageItem
                    Ret.Ifc = pIStorageItem
                End If
            End If
        End If
    End If
    Set GetStorageItemAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


