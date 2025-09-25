VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageLibraryChangeTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_StorageLibraryChangeTracker As String = "Windows.Storage.StorageLibraryChangeTracker"
Private Const IID_IStorageLibraryChangeTracker As String = "{9e157316-6073-44f6-9681-7492d1286c90}"
Private Const IID_IStorageLibraryChangeTracker2 As String = "{cd051c3b-0f9f-42f9-8fb3-158d82e13821}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IStorageLibraryChangeTracker
    IStorageLibraryChangeTracker_GetChangeReader = 6
    IStorageLibraryChangeTracker_Enable = 7
    IStorageLibraryChangeTracker_Reset = 8
    
    ' IStorageLibraryChangeTracker2
    IStorageLibraryChangeTracker2_EnableWithOptions = 6
    IStorageLibraryChangeTracker2_Disable = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIStorageLibraryChangeTracker As Long
Private m_pIStorageLibraryChangeTracker2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageLibraryChangeTracker2)
    Call ReleaseIfc(m_pIStorageLibraryChangeTracker)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageLibraryChangeTracker
End Property

Public Property Let Ifc(ByVal pIStorageLibraryChangeTracker As Long)
    Call ReleaseIfc(m_pIStorageLibraryChangeTracker2)
    Call ReleaseIfc(m_pIStorageLibraryChangeTracker)
    m_pIStorageLibraryChangeTracker = pIStorageLibraryChangeTracker
    Call QueryIfc(m_pIStorageLibraryChangeTracker, _
                  IID_IStorageLibraryChangeTracker2, _
                  m_pIStorageLibraryChangeTracker2)
End Property

' ----==== Functions ====----
' IStorageLibraryChangeTracker
Public Function GetChangeReader() As StorageLibraryChangeReader
    Dim Ret As StorageLibraryChangeReader
    If m_pIStorageLibraryChangeTracker <> 0& Then
        Dim pIStorageLibraryChangeReader As Long
        If Invoke(m_pIStorageLibraryChangeTracker, _
                  IStorageLibraryChangeTracker_GetChangeReader, _
                  VarPtr(pIStorageLibraryChangeReader)) = S_OK Then
            If pIStorageLibraryChangeReader <> 0& Then
                Set Ret = New StorageLibraryChangeReader
                Ret.Ifc = pIStorageLibraryChangeReader
            End If
        End If
    End If
    Set GetChangeReader = Ret
End Function

Public Function Enable() As Boolean
    Dim Ret As Boolean
    If m_pIStorageLibraryChangeTracker <> 0& Then
        If Invoke(m_pIStorageLibraryChangeTracker, _
                  IStorageLibraryChangeTracker_Enable) = S_OK Then
            Ret = True
        End If
    End If
    Enable = Ret
End Function

Public Function Reset() As Boolean
    Dim Ret As Boolean
    If m_pIStorageLibraryChangeTracker <> 0& Then
        If Invoke(m_pIStorageLibraryChangeTracker, _
                  IStorageLibraryChangeTracker_Reset) = S_OK Then
            Ret = True
        End If
    End If
    Reset = Ret
End Function
    
' IStorageLibraryChangeTracker2
Public Function EnableWithOptions(ByVal options As StorageLibraryChangeTrackerOptions) As Boolean
    Dim Ret As Boolean
    If m_pIStorageLibraryChangeTracker2 <> 0& And _
       IsNotNothing(options) Then
        If Invoke(m_pIStorageLibraryChangeTracker2, _
                  IStorageLibraryChangeTracker2_EnableWithOptions, _
                  options) = S_OK Then
            Ret = True
        End If
    End If
    EnableWithOptions = Ret
End Function

Public Function Disable() As Boolean
    Dim Ret As Boolean
    If m_pIStorageLibraryChangeTracker2 <> 0& Then
        If Invoke(m_pIStorageLibraryChangeTracker2, _
                  IStorageLibraryChangeTracker2_Disable) = S_OK Then
            Ret = True
        End If
    End If
    Disable = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


