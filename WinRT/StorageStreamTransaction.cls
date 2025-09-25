VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageStreamTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_IStorageStreamTransaction As String = "Windows.Storage.IStorageStreamTransaction"
Private Const IID_IStorageStreamTransaction As String = "{f67cf363-a53d-4d94-ae2c-67232d93acdd}"

' * Any object which implements this interface must also implement the following interfaces:
' *     Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IStorageStreamTransaction
    IStorageStreamTransaction_GetStream = 6
'    IStorageStreamTransaction_CommitAsync = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIStorageStreamTransaction As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIStorageStreamTransaction) '<-IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageStreamTransaction
End Property

Public Property Let Ifc(ByVal pIStorageStreamTransaction As Long)
    Call DisposeIfc(m_pIStorageStreamTransaction) '<-IClosable
    m_pIStorageStreamTransaction = pIStorageStreamTransaction
End Property

' IStorageStreamTransaction
Public Property Get Stream() As RandomAccessStream
    Dim Ret As RandomAccessStream
    If m_pIStorageStreamTransaction <> 0& Then
        Dim pIRandomAccessStream As Long
        If Invoke(m_pIStorageStreamTransaction, _
                  IStorageStreamTransaction_GetStream, _
                  VarPtr(pIRandomAccessStream)) = S_OK Then
            If pIRandomAccessStream <> 0& Then
                Set Ret = New RandomAccessStream
                Ret.Ifc = pIRandomAccessStream
            End If
        End If
    End If
    Set Stream = Ret
End Property

' ----==== Functions ====----
' IStorageStreamTransaction
'    IStorageStreamTransaction_CommitAsync = 7

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


