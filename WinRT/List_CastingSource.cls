VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List_CastingSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVector_CastingSource
    IVector_CastingSource_GetAt = 6
    IVector_CastingSource_GetSize = 7
'    IVector_CastingSource_GetView = 8
    IVector_CastingSource_IndexOf = 9
    IVector_CastingSource_SetAt = 10
    IVector_CastingSource_InsertAt = 11
    IVector_CastingSource_RemoveAt = 12
    IVector_CastingSource_Append = 13
    IVector_CastingSource_RemoveAtEnd = 14
    IVector_CastingSource_Clear = 15
'    IVector_CastingSource_GetMany = 16
'    IVector_CastingSource_ReplaceAll = 17

End Enum

' ----==== Variablen ====----
Private m_pIVector_CastingSource As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVector_CastingSource)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVector_CastingSource
End Property

Public Property Let Ifc(ByVal pIVector_CastingSource As Long)
    Call ReleaseIfc(m_pIVector_CastingSource)
    m_pIVector_CastingSource = pIVector_CastingSource
End Property

'IVector
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVector_CastingSource <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVector
Public Function GetAt(ByVal index As Long) As CastingSource
    Dim Ret As CastingSource
    If m_pIVector_CastingSource <> 0& Then
        Dim pICastingSource As Long
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_GetAt, _
                  index, _
                  VarPtr(pICastingSource)) = S_OK Then
            If pICastingSource <> 0& Then
                Set Ret = New CastingSource
                Ret.Ifc = pICastingSource
            End If
        End If
    End If
    GetAt = Ret
End Function

'    IVector_CastingSource_GetView = 8

Public Function IndexOf(ByVal Value As CastingSource) As Long
    Dim Ret As Long
    If m_pIVector_CastingSource <> 0& And _
       IsNotNothing(Value) Then
        Dim index As Long
        Dim result As Boolean
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_IndexOf, _
                  Value.Ifc, _
                  VarPtr(index), _
                  VarPtr(result)) = S_OK Then
            If result Then
                Ret = index
            Else
                Ret = -1
            End If
        End If
    End If
    IndexOf = Ret
End Function

Public Function SetAt(ByVal index As Long, _
                      ByVal Value As CastingSource) As Boolean
    Dim Ret As Boolean
    If m_pIVector_CastingSource <> 0& And _
       IsNotNothing(Value) And _
       index >= 0& And _
       index < Me.Size Then
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_SetAt, _
                  index, _
                  Value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    SetAt = Ret
End Function

Public Function InsertAt(ByVal index As Long, _
                         ByVal Value As CastingSource) As Boolean
    Dim Ret As Boolean
    If m_pIVector_CastingSource <> 0& And _
       IsNotNothing(Value) And _
       index >= 0& And _
       index < Me.Size Then
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_InsertAt, _
                  index, _
                  Value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    InsertAt = Ret
End Function

Public Function RemoveAt(ByVal index As Long) As Boolean
    Dim Ret As Boolean
    If m_pIVector_CastingSource <> 0& And _
       index >= 0& And _
       index < Me.Size Then
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_RemoveAt, _
                  index) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAt = Ret
End Function

Public Function Append(ByVal Value As CastingSource) As Boolean
    Dim Ret As Boolean
    If m_pIVector_CastingSource <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_Append, _
                  Value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    Append = Ret
End Function

Public Function RemoveAtEnd() As Boolean
    Dim Ret As Boolean
    If m_pIVector_CastingSource <> 0& Then
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_RemoveAtEnd) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAtEnd = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIVector_CastingSource <> 0& Then
        If Invoke(m_pIVector_CastingSource, _
                  IVector_CastingSource_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

'    IVector_CastingSource_GetMany = 16
'    IVector_CastingSource_ReplaceAll = 17

' ----==== Extra Functions ====----
Public Function GetAll() As CastingSource()
    Dim Ret() As CastingSource
    Dim count As Long
    count = Me.Size
    If count > 0& Then
        ReDim Ret(count - 1)
        Dim Item As Long
        For Item = 0 To count - 1
            Ret(Item) = Me.GetAt(Item)
        Next
    End If
    GetAll = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


