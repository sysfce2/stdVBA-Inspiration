VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Streams_IInputStream As String = "Windows.Storage.Streams.IInputStream"
Private Const IID_IInputStream As String = "{905a0fe2-bc53-11df-8c49-001e4fc686da}"

' * Any object which implements this interface must also implement the following interfaces:
' *     Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IInputStream
    IInputStream_ReadAsync = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIInputStream As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIInputStream) '<-IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIInputStream
End Property

Public Property Let Ifc(ByVal pIInputStream As Long)
    Call DisposeIfc(m_pIInputStream) '<-IClosable
    m_pIInputStream = pIInputStream
End Property

' ----==== Functions ====----
' IInputStream
'    IInputStream_ReadAsync = 6

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

