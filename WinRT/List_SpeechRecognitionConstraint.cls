VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List_SpeechRecognitionConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVector_ISpeechRecognitionConstraint
    IVector_ISpeechRecognitionConstraint_GetAt = 6
    IVector_ISpeechRecognitionConstraint_GetSize = 7
'    IVector_ISpeechRecognitionConstraint_GetView = 8
    IVector_ISpeechRecognitionConstraint_IndexOf = 9
    IVector_ISpeechRecognitionConstraint_SetAt = 10
    IVector_ISpeechRecognitionConstraint_InsertAt = 11
    IVector_ISpeechRecognitionConstraint_RemoveAt = 12
    IVector_ISpeechRecognitionConstraint_Append = 13
    IVector_ISpeechRecognitionConstraint_RemoveAtEnd = 14
    IVector_ISpeechRecognitionConstraint_Clear = 15
'    IVector_ISpeechRecognitionConstraint_GetMany = 16
'    IVector_ISpeechRecognitionConstraint_ReplaceAll = 17

End Enum

' ----==== Variablen ====----
Private m_pIVector_ISpeechRecognitionConstraint As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVector_ISpeechRecognitionConstraint)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVector_ISpeechRecognitionConstraint
End Property

Public Property Let Ifc(ByVal pIVector_ISpeechRecognitionConstraint As Long)
    Call ReleaseIfc(m_pIVector_ISpeechRecognitionConstraint)
    m_pIVector_ISpeechRecognitionConstraint = pIVector_ISpeechRecognitionConstraint
End Property

'IVector
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVector_ISpeechRecognitionConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVector
Public Function GetAt(ByVal Index As Long) As SpeechRecognitionConstraint
    Dim Ret As SpeechRecognitionConstraint
    If m_pIVector_ISpeechRecognitionConstraint <> 0& Then
        Dim value As Long
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_GetAt, _
                  Index, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New SpeechRecognitionConstraint
                Ret.Ifc = value
            End If
        End If
    End If
    Set GetAt = Ret
End Function

'    IVector_ISpeechRecognitionConstraint_GetView = 8

Public Function IndexOf(ByVal value As SpeechRecognitionConstraint) As Long
    Dim Ret As Long
    If m_pIVector_ISpeechRecognitionConstraint <> 0& And _
       IsNotNothing(value) Then
        Dim Index As Long
        Dim Result As Boolean
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_IndexOf, _
                  value.Ifc, _
                  VarPtr(Index), _
                  VarPtr(Result)) = S_OK Then
            If Result Then
                Ret = Index
            Else
                Ret = -1
            End If
        End If
    End If
    IndexOf = Ret
End Function

Public Function SetAt(ByVal Index As Long, _
                      ByVal value As SpeechRecognitionConstraint) As Boolean
    Dim Ret As Boolean
    If m_pIVector_ISpeechRecognitionConstraint <> 0& And _
       Index >= 0& And _
       Index < Me.Size And _
       IsNotNothing(value) Then
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_SetAt, _
                  Index, _
                  value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    SetAt = Ret
End Function

Public Function InsertAt(ByVal Index As Long, _
                         ByVal value As SpeechRecognitionConstraint) As Boolean
    Dim Ret As Boolean
    If m_pIVector_ISpeechRecognitionConstraint <> 0& And _
       Index >= 0& And _
       Index < Me.Size And _
       IsNotNothing(value) Then
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_InsertAt, _
                  Index, _
                  value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    InsertAt = Ret
End Function

Public Function RemoveAt(ByVal Index As Long) As Boolean
    Dim Ret As Boolean
    If m_pIVector_ISpeechRecognitionConstraint <> 0& And _
       Index >= 0& And _
       Index < Me.Size Then
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_RemoveAt, _
                  Index) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAt = Ret
End Function

Public Function Append(ByVal value As Object) As Boolean
    Dim Ret As Boolean
    If m_pIVector_ISpeechRecognitionConstraint <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_Append, _
                  value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    Append = Ret
End Function

Public Function RemoveAtEnd() As Boolean
    Dim Ret As Boolean
    If m_pIVector_ISpeechRecognitionConstraint <> 0& Then
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_RemoveAtEnd) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAtEnd = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIVector_ISpeechRecognitionConstraint <> 0& Then
        If Invoke(m_pIVector_ISpeechRecognitionConstraint, _
                  IVector_ISpeechRecognitionConstraint_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

'    IVector_ISpeechRecognitionConstraint_GetMany = 16
'    IVector_ISpeechRecognitionConstraint_ReplaceAll = 17

' ----==== Extra Functions ====----
Public Function GetAll() As SpeechRecognitionConstraint()
    Dim Ret() As SpeechRecognitionConstraint
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


