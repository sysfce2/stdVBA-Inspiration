VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutputStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Streams_IOutputStream As String = "Windows.Storage.Streams.IOutputStream"
Private Const IID_IOutputStream As String = "{905a0fe6-bc53-11df-8c49-001e4fc686da}"

' * Any object which implements this interface must also implement the following interfaces:
' *     Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IOutputStream
'    IOutputStream_WriteAsync = 6
    IOutputStream_FlushAsync = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIOutputStream As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIOutputStream) '<-IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIOutputStream
End Property

Public Property Let Ifc(ByVal pIOutputStream As Long)
    Call DisposeIfc(m_pIOutputStream) '<-IClosable
    m_pIOutputStream = pIOutputStream
End Property

' ----==== Functions ====----
' IOutputStream
'    IOutputStream_WriteAsync = 6

Public Function FlushAsync() As Boolean
    Dim Ret As Boolean
    If m_pIOutputStream <> 0& Then
        Dim operation As Long
        If Invoke(m_pIOutputStream, _
                  IOutputStream_FlushAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    FlushAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


