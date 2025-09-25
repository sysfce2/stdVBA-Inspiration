VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_SortEntry"
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

    ' IVector_SortEntry
    IVector_SortEntry_GetAt = 6
    IVector_SortEntry_GetSize = 7
    
End Enum

' ----==== Types ====----
Private Type Windows_Storage_Search_SortEntry
    PropertyName As Long
    AscendingOrder As Long
End Type

' ----==== Variablen ====----
Private m_pIVector_SortEntry As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVector_SortEntry)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVector_SortEntry
End Property

Public Property Let Ifc(ByVal pIVector_SortEntry As Long)
    Call ReleaseIfc(m_pIVector_SortEntry)
    m_pIVector_SortEntry = pIVector_SortEntry
End Property

' IVector_SortEntry
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVector_SortEntry <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVector_SortEntry, _
                  IVector_SortEntry_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
' IVector_SortEntry
Public Function GetAt(ByVal index As Long) As SortEntry
    Dim Ret As SortEntry
    If m_pIVector_SortEntry <> 0& Then
        Dim tSortEntry As Windows_Storage_Search_SortEntry
        If Invoke(m_pIVector_SortEntry, _
                  IVector_SortEntry_GetAt, _
                  index, _
                  VarPtr(tSortEntry)) = S_OK Then
            Set Ret = New SortEntry
            Ret.PropertyName = GetWindowsString(tSortEntry.PropertyName)
            Ret.AscendingOrder = CBool(tSortEntry.AscendingOrder)
        End If
    End If
    Set GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As SortEntry()
    Dim Ret() As SortEntry
    Dim count As Long
    count = Me.Size
    If count > 0& Then
        ReDim Ret(count - 1)
        Dim Item As Long
        For Item = 0 To count - 1
            Set Ret(Item) = Me.GetAt(Item)
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


