VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_String"
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

    ' IVector_HSTRING
    IVector_HSTRING_GetAt = 6
    IVector_HSTRING_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIVector_HSTRING As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
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

' IVector_HSTRING
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
' IVector_HSTRING
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


