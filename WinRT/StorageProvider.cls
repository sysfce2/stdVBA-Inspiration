VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Provider As String = "Windows.Storage.StorageProvider"
Private Const IID_IStorageProvider As String = "{e705eed4-d478-47d6-ba46-1a8ebe114a20}"
Private Const IID_IStorageProvider2 As String = "{010d1917-3404-414b-9fd7-cd44472eaa39}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IStorageProvider
    IStorageProvider_GetId = 6
    IStorageProvider_GetDisplayName = 7

    ' IStorageProvider2
    IStorageProvider2_IsPropertySupportedForPartialFileAsync = 6

End Enum

' ----==== Vars ====----
Private m_pIStorageProvider As Long
Private m_pIStorageProvider2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageProvider)
    Call ReleaseIfc(m_pIStorageProvider2)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageProvider
End Property

Public Property Let Ifc(ByVal pIStorageProvider As Long)
    Call ReleaseIfc(m_pIStorageProvider)
    Call ReleaseIfc(m_pIStorageProvider2)
    m_pIStorageProvider = pIStorageProvider
    Call QueryIfc(m_pIStorageProvider, _
                  IID_IStorageProvider2, _
                  m_pIStorageProvider2)
End Property

' IStorageProvider
Public Property Get Id() As String
    Dim Ret As String
    If m_pIStorageProvider <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageProvider, _
                  IStorageProvider_GetId, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Id = Ret
End Property

Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIStorageProvider <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStorageProvider, _
                  IStorageProvider_GetDisplayName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    DisplayName = Ret
End Property

' ----==== Functions ====----
' IStorageProvider2
Public Function IsPropertySupportedForPartialFileAsync(ByVal propertyCanonicalName As String) As Boolean
    Dim Ret As Boolean
    If m_pIStorageProvider2 <> 0& Then
        Dim hString As Long
        Dim Value As Long
        hString = CreateWindowsString(propertyCanonicalName)
        If Invoke(m_pIStorageProvider2, _
                  IStorageProvider2_IsPropertySupportedForPartialFileAsync, _
                  hString, _
                  VarPtr(Value)) = S_OK Then
            If Await(Value) Then
                Ret = CBool(Value)
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    IsPropertySupportedForPartialFileAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


