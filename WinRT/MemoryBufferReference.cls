VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryBufferReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 03/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Foundation_IMemoryBufferReference As String = "Windows.Foundation.IMemoryBufferReference"
Private Const IID_IMemoryBufferReference As String = "{fbc4dd29-245b-11e4-af98-689423260cf8}"

' * Class implements the following interface:
' *    Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMemoryBufferReference
    IMemoryBufferReference_GetCapacity = 6
    IMemoryBufferReference_AddClosed = 7
    IMemoryBufferReference_RemoveClosed = 8
    
End Enum

' ----==== Variablen ====----
Private m_pIMemoryBufferReference As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIMemoryBufferReference) '<- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMemoryBufferReference
End Property

Public Property Let Ifc(ByVal pIMemoryBufferReference As Long)
    Call DisposeIfc(m_pIMemoryBufferReference) '<- IClosable
    m_pIMemoryBufferReference = pIMemoryBufferReference
End Property

' IMemoryBufferReference
Public Property Get capacity() As Long
    Dim Ret As Long
    If m_pIMemoryBufferReference <> 0& Then
        Dim value As Long
        If Invoke(m_pIMemoryBufferReference, _
                  IMemoryBufferReference_GetCapacity, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    capacity = Ret
End Property

' ----==== Functions ====----
' IMemoryBufferReference
Public Function AddClosed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMemoryBufferReference <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIMemoryBufferReference, _
                  IMemoryBufferReference_AddClosed, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddClosed = Ret
End Function

Public Function RemoveClosed(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMemoryBufferReference <> 0& Then
        If Invoke(m_pIMemoryBufferReference, _
                  IMemoryBufferReference_RemoveClosed, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveClosed = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


