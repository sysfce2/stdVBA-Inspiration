VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageLibraryChangeReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_StorageLibraryChangeReader As String = "Windows.Storage.StorageLibraryChangeReader"
Private Const IID_IStorageLibraryChangeReader As String = "{f205bc83-fca2-41f9-8954-ee2e991eb96f}"
Private Const IID_IStorageLibraryChangeReader2 As String = "{abf4868b-fbcc-4a4f-999e-e7ab7c646dbe}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IStorageLibraryChangeReader
    IStorageLibraryChangeReader_ReadBatchAsync = 6
    IStorageLibraryChangeReader_AcceptChangesAsync = 7
    
    ' IStorageLibraryChangeReader2
    IStorageLibraryChangeReader2_GetLastChangeId = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIStorageLibraryChangeReader As Long
Private m_pIStorageLibraryChangeReader2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageLibraryChangeReader2)
    Call ReleaseIfc(m_pIStorageLibraryChangeReader)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageLibraryChangeReader
End Property

Public Property Let Ifc(ByVal pIStorageLibraryChangeReader As Long)
    Call ReleaseIfc(m_pIStorageLibraryChangeReader2)
    Call ReleaseIfc(m_pIStorageLibraryChangeReader)
    m_pIStorageLibraryChangeReader = pIStorageLibraryChangeReader
    Call QueryIfc(m_pIStorageLibraryChangeReader, _
                  IID_IStorageLibraryChangeReader2, _
                  m_pIStorageLibraryChangeReader2)
End Property

' ----==== Functions ====----
' IStorageLibraryChangeReader
Public Function ReadBatchAsync() As ReadOnlyList_1 'ReadOnlyList_StorageLibraryChange
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_StorageLibraryChange
    If m_pIStorageLibraryChangeReader <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIStorageLibraryChangeReader, _
                  IStorageLibraryChangeReader_ReadBatchAsync, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_StorageLibraryChange, pIVectorView) 'ReadOnlyList_StorageLibraryChange
                End If
            End If
        End If
    End If
    Set ReadBatchAsync = Ret
End Function

Public Function AcceptChangesAsync() As Boolean
    Dim Ret As Boolean
    If m_pIStorageLibraryChangeReader <> 0& Then
        Dim operation As Long
        If Invoke(m_pIStorageLibraryChangeReader, _
                  IStorageLibraryChangeReader_AcceptChangesAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    AcceptChangesAsync = Ret
End Function
    
' IStorageLibraryChangeReader2
Public Function GetLastChangeId() As Currency
    Dim Ret As Currency
    If m_pIStorageLibraryChangeReader2 <> 0& Then
        Dim value As Currency
        If Invoke(m_pIStorageLibraryChangeReader2, _
                  IStorageLibraryChangeReader2_GetLastChangeId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetLastChangeId = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


