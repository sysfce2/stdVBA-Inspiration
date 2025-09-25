VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Streams_Buffer As String = "Windows.Storage.Streams.Buffer"
Private Const IID_IBuffer As String = "{905a0fe0-bc53-11df-8c49-001e4fc686da}"
Private Const IID_IBufferFactory As String = "{71af914d-c10f-484b-bc50-14bc623b3a27}"
Private Const IID_IBufferStatics As String = "{e901e65b-d716-475a-a90a-af7229b1e741}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IBufferFactory
    IBufferFactory_Create = 6
    
    ' IBufferStatics
    IBufferStatics_CreateCopyFromMemoryBuffer = 6
    IBufferStatics_CreateMemoryBufferOverIBuffer = 7
    
    ' IBuffer
    IBuffer_GetCapacity = 6
    IBuffer_GetLength = 7
    IBuffer_PutLength = 8
    
End Enum

' ----==== Variablen ====----
Private m_pIBuffer As Long
Private m_pIBufferFactory As Long
Private m_pIBufferStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_Streams_Buffer, _
                            IID_IBufferFactory, _
                            m_pIBufferFactory) Then
        Call QueryIfc(m_pIBufferFactory, _
                      IID_IBufferStatics, _
                      m_pIBufferStatics)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBuffer)
    Call ReleaseIfc(m_pIBufferFactory)
    Call ReleaseIfc(m_pIBufferStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBuffer
End Property

Public Property Let Ifc(ByVal pIBuffer As Long)
    Call ReleaseIfc(m_pIBuffer)
    m_pIBuffer = pIBuffer
End Property

' IBuffer
Public Property Get capacity() As Long
    Dim Ret As Long
    If m_pIBuffer <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBuffer, _
                  IBuffer_GetCapacity, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    capacity = Ret
End Property

Public Property Get Length() As Long
    Dim Ret As Long
    If m_pIBuffer <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBuffer, _
                  IBuffer_GetLength, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Length = Ret
End Property

Public Property Let Length(ByVal Value As Long)
    If m_pIBuffer <> 0& Then
        If Invoke(m_pIBuffer, _
                  IBuffer_PutLength, _
                  Value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IBufferFactory
Public Function Create(ByVal capacity As Long) As Boolean
    Dim Ret As Boolean
    If m_pIBufferFactory <> 0& Then
        Dim pIBuffer As Long
        If Invoke(m_pIBufferFactory, _
                  IBufferFactory_Create, _
                  capacity, _
                  VarPtr(pIBuffer)) = S_OK Then
            Me.Ifc = pIBuffer
        End If
    End If
    Create = Ret
End Function
    
' IBufferStatics
Public Function CreateCopyFromMemoryBuffer(ByVal memBuffer As MemoryBuffer) As Buffer
    Dim Ret As Buffer
    If m_pIBufferStatics <> 0& And _
       IsNotNothing(memBuffer) Then
        Dim pIBuffer As Long
        If Invoke(m_pIBufferStatics, _
                  IBufferStatics_CreateCopyFromMemoryBuffer, _
                  memBuffer.Ifc, _
                  VarPtr(pIBuffer)) = S_OK Then
            If pIBuffer <> 0& Then
                Set Ret = New Buffer
                Ret.Ifc = pIBuffer
            End If
        End If
    End If
    Set CreateCopyFromMemoryBuffer = Ret
End Function

Public Function CreateMemoryBufferOverIBuffer(ByVal Buffer As Buffer) As MemoryBuffer
    Dim Ret As MemoryBuffer
    If m_pIBufferStatics <> 0& And _
       IsNotNothing(Buffer) Then
        Dim pIMemoryBuffer As Long
        If Invoke(m_pIBufferStatics, _
                  IBufferStatics_CreateMemoryBufferOverIBuffer, _
                  Buffer.Ifc, _
                  VarPtr(pIMemoryBuffer)) = S_OK Then
            If pIMemoryBuffer <> 0& Then
                Set Ret = New MemoryBuffer
                Ret.Ifc = pIMemoryBuffer
            End If
        End If
    End If
    Set CreateMemoryBufferOverIBuffer = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


