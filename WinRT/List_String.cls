VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List_String"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const IVector_NAME As String = "IVector<HSTRING>"
Private Const IVector_HSTRING As String = "{98b9acc1-4b56-532e-ac73-03d5291cca90}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVector_HSTRING
    IVector_HSTRING_GetAt = 6
    IVector_HSTRING_GetSize = 7
'    IVector_HSTRING_GetView = 8
    IVector_HSTRING_IndexOf = 9
    IVector_HSTRING_SetAt = 10
    IVector_HSTRING_InsertAt = 11
    IVector_HSTRING_RemoveAt = 12
    IVector_HSTRING_Append = 13
    IVector_HSTRING_RemoveAtEnd = 14
    IVector_HSTRING_Clear = 15
'    IVector_HSTRING_GetMany = 16
'    IVector_HSTRING_ReplaceAll = 17

End Enum

' ----==== Variablen ====----
Private m_pIVector_HSTRING As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    m_pIVector_HSTRING = IVector.Implement(IVector_NAME, IVector_HSTRING)
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVector_HSTRING)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVector_HSTRING
End Property

Public Property Let Ifc(ByVal pIVector_HSTRING As Long)
    Call ReleaseIfc(m_pIVector_HSTRING)
    m_pIVector_HSTRING = pIVector_HSTRING
End Property

'IVector
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVector_HSTRING <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVector
Public Function GetAt(ByVal index As Long) As String
    Dim Ret As String
    If m_pIVector_HSTRING <> 0& Then
        Dim hString As Long
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_GetAt, _
                  index, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    GetAt = Ret
End Function

'    IVector_HSTRING_GetView = 8

Public Function IndexOf(ByVal Value As String) As Long
    Dim Ret As Long
    If m_pIVector_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        Dim index As Long
        Dim result As Boolean
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_IndexOf, _
                  hString, _
                  VarPtr(index), _
                  VarPtr(result)) = S_OK Then
            If result Then
                Ret = index
            Else
                Ret = -1
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    IndexOf = Ret
End Function

Public Function SetAt(ByVal index As Long, _
                      ByVal Value As String) As Boolean
    Dim Ret As Boolean
    If m_pIVector_HSTRING <> 0& And _
       index >= 0& And _
       index < Me.Size Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_SetAt, _
                  index, _
                  hString) = S_OK Then
            Ret = True
        End If
        'Call DeleteWindowsString(hString)
    End If
    SetAt = Ret
End Function

Public Function InsertAt(ByVal index As Long, _
                         ByVal Value As String) As Boolean
    Dim Ret As Boolean
    If m_pIVector_HSTRING <> 0& And _
       index >= 0& And _
       index < Me.Size Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_InsertAt, _
                  index, _
                  hString) = S_OK Then
            Ret = True
        End If
        'Call DeleteWindowsString(hString)
    End If
    InsertAt = Ret
End Function

Public Function RemoveAt(ByVal index As Long) As Boolean
    Dim Ret As Boolean
    If m_pIVector_HSTRING <> 0& And _
       index >= 0& And _
       index < Me.Size Then
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_RemoveAt, _
                  index) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAt = Ret
End Function

Public Function Append(ByVal Value As String) As Boolean
    Dim Ret As Boolean
    If m_pIVector_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_Append, _
                  hString) = S_OK Then
            Ret = True
        End If
        'Call DeleteWindowsString(hString)
    End If
    Append = Ret
End Function

Public Function RemoveAtEnd() As Boolean
    Dim Ret As Boolean
    If m_pIVector_HSTRING <> 0& Then
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_RemoveAtEnd) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAtEnd = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIVector_HSTRING <> 0& Then
        If Invoke(m_pIVector_HSTRING, _
                  IVector_HSTRING_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

'    IVector_HSTRING_GetMany = 16
'    IVector_HSTRING_ReplaceAll = 17

' ----==== Extra Functions ====----
Public Function GetAll() As String()
    Dim Ret() As String
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


