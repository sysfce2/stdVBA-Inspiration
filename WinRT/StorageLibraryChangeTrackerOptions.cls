VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageLibraryChangeTrackerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_StorageLibraryChangeTrackerOptions As String = "Windows.Storage.StorageLibraryChangeTrackerOptions"
Private Const IID_IStorageLibraryChangeTrackerOptions As String = "{bb52bcd4-1a6d-59c0-ad2a-823a20532483}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IStorageLibraryChangeTrackerOptions
    IStorageLibraryChangeTrackerOptions_GetTrackChangeDetails = 6
    IStorageLibraryChangeTrackerOptions_PutTrackChangeDetails = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIStorageLibraryChangeTrackerOptions As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_Storage_StorageLibraryChangeTrackerOptions, _
                           IID_IStorageLibraryChangeTrackerOptions, _
                           m_pIStorageLibraryChangeTrackerOptions) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageLibraryChangeTrackerOptions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageLibraryChangeTrackerOptions
End Property

Public Property Let Ifc(ByVal pIStorageLibraryChangeTrackerOptions As Long)
    Call ReleaseIfc(m_pIStorageLibraryChangeTrackerOptions)
    m_pIStorageLibraryChangeTrackerOptions = pIStorageLibraryChangeTrackerOptions
End Property

' IStorageLibraryChangeTrackerOptions
Public Property Get TrackChangeDetails() As Boolean
    Dim Ret As Boolean
    If m_pIStorageLibraryChangeTrackerOptions <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageLibraryChangeTrackerOptions, _
                  IStorageLibraryChangeTrackerOptions_GetTrackChangeDetails, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    TrackChangeDetails = Ret
End Property

Public Property Let TrackChangeDetails(ByVal Value As Boolean)
    If m_pIStorageLibraryChangeTrackerOptions <> 0& Then
        If Invoke(m_pIStorageLibraryChangeTrackerOptions, _
                  IStorageLibraryChangeTrackerOptions_PutTrackChangeDetails, _
                  Value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


