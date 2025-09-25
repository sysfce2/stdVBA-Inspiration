VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_BitmapPixelFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVector_BitmapPixelFormat
    IVector_BitmapPixelFormat_GetAt = 6
    IVector_BitmapPixelFormat_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIVector_BitmapPixelFormat As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVector_BitmapPixelFormat)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVector_BitmapPixelFormat
End Property

Public Property Let Ifc(ByVal pIVector_BitmapPixelFormat As Long)
    Call ReleaseIfc(m_pIVector_BitmapPixelFormat)
    m_pIVector_BitmapPixelFormat = pIVector_BitmapPixelFormat
End Property

' IVector_BitmapPixelFormat
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVector_BitmapPixelFormat <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVector_BitmapPixelFormat, _
                  IVector_BitmapPixelFormat_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
' IVector_BitmapPixelFormat
Public Function GetAt(ByVal index As Long) As BitmapPixelFormat
    Dim Ret As BitmapPixelFormat
    If m_pIVector_BitmapPixelFormat <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVector_BitmapPixelFormat, _
                  IVector_BitmapPixelFormat_GetAt, _
                  index, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As BitmapPixelFormat()
    Dim Ret() As BitmapPixelFormat
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


