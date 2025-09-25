VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FilePickerFileTypesOrderedMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Pickers_FilePickerFileTypesOrderedMap As String = "Windows.Storage.Pickers.FilePickerFileTypesOrderedMap"
Private Const IMap_HSTRING_IVector_HSTRING As String = "{e475ca9d-6afb-5992-993e-53e6ef7a9ecd}"
Private Const IIterable_IKeyValuePair_HSTRING_IVector_HSTRING As String = "{4fed2669-d0d3-59f6-91d9-95902d728d6a}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

  ' IMap_HSTRING_IVector_HSTRING
    IMap_HSTRING_IVector_HSTRING_Lookup = 6
    IMap_HSTRING_IVector_HSTRING_GetSize = 7
    IMap_HSTRING_IVector_HSTRING_HasKey = 8
    IMap_HSTRING_IVector_HSTRING_GetView = 9
    IMap_HSTRING_IVector_HSTRING_Insert = 10
    IMap_HSTRING_IVector_HSTRING_Remove = 11
    IMap_HSTRING_IVector_HSTRING_Clear = 12
    
'    ' IIterable_IKeyValuePair_HSTRING_IVector_HSTRING
'    IIterable_IKeyValuePair_HSTRING_IVector_HSTRING_First = 6
'
'    ' IIterator_IKeyValuePair_HSTRING_IVector_HSTRING
'    IIterator_IKeyValuePair_HSTRING_IVector_HSTRING_GetCurrent = 6
'    IIterator_IKeyValuePair_HSTRING_IVector_HSTRING_GetHasCurrent = 7
'    IIterator_IKeyValuePair_HSTRING_IVector_HSTRING_MoveNext = 8
'    IIterator_IKeyValuePair_HSTRING_IVector_HSTRING_GetMany = 9
'
'    ' IKeyValuePair_HSTRING_IVector_HSTRING
'    IKeyValuePair_HSTRING_IVector_HSTRING_GetKey = 6
'    IKeyValuePair_HSTRING_IVector_HSTRING_GetValue = 7
End Enum

' ----==== Variablen ====----
Private m_pIMap_HSTRING_IVector_HSTRING As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMap_HSTRING_IVector_HSTRING)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMap_HSTRING_IVector_HSTRING
End Property

Public Property Let Ifc(ByVal pIMap_HSTRING_IVector_HSTRING As Long)
    Call ReleaseIfc(m_pIMap_HSTRING_IVector_HSTRING)
    m_pIMap_HSTRING_IVector_HSTRING = pIMap_HSTRING_IVector_HSTRING
End Property

' IMap_HSTRING_IVector_HSTRING
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIMap_HSTRING_IVector_HSTRING <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMap_HSTRING_IVector_HSTRING, _
                  IMap_HSTRING_IVector_HSTRING_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
' IMap_HSTRING_IVector_HSTRING
Public Function Lookup(ByVal Key As String) As List_String
    Dim Ret As List_String
    If m_pIMap_HSTRING_IVector_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim pIVector As Long
        If Invoke(m_pIMap_HSTRING_IVector_HSTRING, _
                  IMap_HSTRING_IVector_HSTRING_Lookup, _
                  hString, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_String
                Ret.Ifc = pIVector
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Lookup = Ret
End Function

Public Function HasKey(ByVal Key As String) As Boolean
    Dim Ret As Boolean
    If m_pIMap_HSTRING_IVector_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim Value As Long
        If Invoke(m_pIMap_HSTRING_IVector_HSTRING, _
                  IMap_HSTRING_IVector_HSTRING_HasKey, _
                  hString, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
        Call DeleteWindowsString(hString)
    End If
    HasKey = Ret
End Function

'    IMap_HSTRING_IVector_HSTRING_GetView = 9
'Public Function GetView() As Boolean
'    Dim Ret As Boolean
'    If m_pIMap_HSTRING_IVector_HSTRING <> 0& Then
'        Dim value As Long
'        If Invoke(m_pIMap_HSTRING_IVector_HSTRING, _
'                  IMap_HSTRING_IVector_HSTRING_GetView, _
'                  VarPtr(value)) = S_OK Then
'            If value <> 0& Then
'                Call GetInspectableInfo(value, "value")
'
'                Ret = True
'            End If
'        End If
'    End If
'    GetView = Ret
'End Function

Public Function Insert(ByVal Key As String, _
                       ByVal Value As List_String, _
                       Optional ByRef replaced As Boolean) As Boolean
    Dim Ret As Boolean
    If m_pIMap_HSTRING_IVector_HSTRING <> 0& And _
       IsNotNothing(Value) Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        If Invoke(m_pIMap_HSTRING_IVector_HSTRING, _
                  IMap_HSTRING_IVector_HSTRING_Insert, _
                  hString, _
                  Value.Ifc, _
                  VarPtr(replaced)) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    Insert = Ret
End Function

Public Function Remove(ByVal Key As String) As Boolean
    Dim Ret As Boolean
    If m_pIMap_HSTRING_IVector_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        If Invoke(m_pIMap_HSTRING_IVector_HSTRING, _
                  IMap_HSTRING_IVector_HSTRING_Remove, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    Remove = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIMap_HSTRING_IVector_HSTRING <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMap_HSTRING_IVector_HSTRING, _
                  IMap_HSTRING_IVector_HSTRING_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

