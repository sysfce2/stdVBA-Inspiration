VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Foundation_MemoryBuffer As String = "Windows.Foundation.MemoryBuffer"
Private Const IID_IMemoryBuffer As String = "{fbc4dd2a-245b-11e4-af98-689423260cf8}"
Private Const IID_IMemoryBufferFactory As String = "{fbc4dd2b-245b-11e4-af98-689423260cf8}"
 
' * Class implements the following interface:
' *    Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMemoryBufferFactory
    IMemoryBufferFactory_Create = 6
    
    ' IMemoryBuffer
    IMemoryBuffer_CreateReference = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMemoryBuffer As Long
Private m_pIMemoryBufferFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Foundation_MemoryBuffer, _
                            IID_IMemoryBufferFactory, _
                            m_pIMemoryBufferFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIMemoryBuffer) '<- IClosable
    Call ReleaseIfc(m_pIMemoryBufferFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMemoryBuffer
End Property

Public Property Let Ifc(ByVal pIMemoryBuffer As Long)
    Call DisposeIfc(m_pIMemoryBuffer) '<- IClosable
    m_pIMemoryBuffer = pIMemoryBuffer
End Property

' ----==== Functions ====----
' IMemoryBufferFactory
Public Function Create(ByVal Capacity As Long) As Boolean
    Dim Ret As Boolean
    If m_pIMemoryBufferFactory <> 0& Then
        Dim pIMemoryBuffer As Long
        If Invoke(m_pIMemoryBufferFactory, _
                  IMemoryBufferFactory_Create, _
                  Capacity, _
                  VarPtr(pIMemoryBuffer)) = S_OK Then
            Me.Ifc = pIMemoryBuffer
        End If
    End If
    Create = Ret
End Function

' IMemoryBuffer
Public Function CreateReference() As MemoryBufferReference
    Dim Ret As MemoryBufferReference
    If m_pIMemoryBuffer <> 0& Then
        Dim pIMemoryBufferReference As Long
        If Invoke(m_pIMemoryBuffer, _
                  IMemoryBuffer_CreateReference, _
                  VarPtr(pIMemoryBufferReference)) = S_OK Then
            If pIMemoryBufferReference <> 0& Then
                Set Ret = New MemoryBufferReference
                Ret.Ifc = pIMemoryBufferReference
            End If
        End If
    End If
    Set CreateReference = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


