VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List_String_String"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const IID_IIterable_IKeyValuePair_HSTRING_HSTRING As String = "{e9bdaaf0-cbf6-5c72-be90-29cbf3a1319b}"
Private Const IID_IIterator_IKeyValuePair_HSTRING_HSTRING As String = "{05eb86f1-7140-5517-b88d-cbaebe57e6b1}"
Private Const IID_IKeyValuePair_HSTRING_HSTRING As String = "{60310303-49c5-52e6-abc6-a9b36eccc716}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMap_HSTRING_HSTRING
    IMap_HSTRING_HSTRING_Lookup = 6
    IMap_HSTRING_HSTRING_GetSize = 7
    IMap_HSTRING_HSTRING_HasKey = 8
    IMap_HSTRING_HSTRING_GetView = 9
    IMap_HSTRING_HSTRING_Insert = 10
    IMap_HSTRING_HSTRING_Remove = 11
    IMap_HSTRING_HSTRING_Clear = 12
    
    ' IIterable_IKeyValuePair_HSTRING_HSTRING
    IIterable_IKeyValuePair_HSTRING_HSTRING_First = 6
    
    ' IIterator_IKeyValuePair_HSTRING_HSTRING
    IIterator_IKeyValuePair_HSTRING_HSTRING_GetCurrent = 6
    IIterator_IKeyValuePair_HSTRING_HSTRING_GetHasCurrent = 7
    IIterator_IKeyValuePair_HSTRING_HSTRING_MoveNext = 8
    IIterator_IKeyValuePair_HSTRING_HSTRING_GetMany = 9

    ' IKeyValuePair_HSTRING_HSTRING
    IKeyValuePair_HSTRING_HSTRING_GetKey = 6
    IKeyValuePair_HSTRING_HSTRING_GetValue = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIMap_HSTRING_HSTRING As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMap_HSTRING_HSTRING)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMap_HSTRING_HSTRING
End Property

Public Property Let Ifc(ByVal pIMap_HSTRING_HSTRING As Long)
    Call ReleaseIfc(m_pIMap_HSTRING_HSTRING)
    m_pIMap_HSTRING_HSTRING = pIMap_HSTRING_HSTRING
End Property

' IMap_HSTRING_HSTRING
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIMap_HSTRING_HSTRING <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMap_HSTRING_HSTRING, _
                  IMap_HSTRING_HSTRING_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
' IMap_HSTRING_HSTRING
Public Function Lookup(ByVal Key As String) As NotificationBinding
    Dim Ret As NotificationBinding
    If m_pIMap_HSTRING_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim pINotificationBinding As Long
        If Invoke(m_pIMap_HSTRING_HSTRING, _
                  IMap_HSTRING_HSTRING_Lookup, _
                  hString, _
                  VarPtr(pINotificationBinding)) = S_OK Then
            If pINotificationBinding <> 0& Then
                Set Ret = New NotificationBinding
                Ret.Ifc = pINotificationBinding
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Lookup = Ret
End Function

Public Function HasKey(ByVal Key As String) As Boolean
    Dim Ret As Boolean
    If m_pIMap_HSTRING_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim Value As Long
        If Invoke(m_pIMap_HSTRING_HSTRING, _
                  IMap_HSTRING_HSTRING_HasKey, _
                  hString, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
        Call DeleteWindowsString(hString)
    End If
    HasKey = Ret
End Function

Public Function Insert(ByVal Key As String, _
                       ByVal Value As NotificationBinding, _
                       Optional ByRef replaced As Boolean) As Boolean
    Dim Ret As Boolean
    If m_pIMap_HSTRING_HSTRING <> 0& And _
       IsNotNothing(Value) Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        If Invoke(m_pIMap_HSTRING_HSTRING, _
                  IMap_HSTRING_HSTRING_Insert, _
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
    If m_pIMap_HSTRING_HSTRING <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        If Invoke(m_pIMap_HSTRING_HSTRING, _
                  IMap_HSTRING_HSTRING_Remove, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    Remove = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIMap_HSTRING_HSTRING <> 0& Then
        If Invoke(m_pIMap_HSTRING_HSTRING, _
                  IMap_HSTRING_HSTRING_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

' IIterable_IKeyValuePair_HSTRING_HSTRING
' IIterator_IKeyValuePair_HSTRING_HSTRING
' IKeyValuePair_HSTRING_HSTRING
Public Function GetKeyValuePairs() As KeyValuePair_String_String()
    Dim Ret() As KeyValuePair_String_String
    If m_pIMap_HSTRING_HSTRING <> 0& Then
        Dim count As Long
        count = Me.Size
        If count > 0& Then
            ReDim Ret(count - 1)
            Dim pIIterable_IKeyValuePair_HSTRING_HSTRING As Long
            If QueryIfc(m_pIMap_HSTRING_HSTRING, _
                        IID_IIterable_IKeyValuePair_HSTRING_HSTRING, _
                        pIIterable_IKeyValuePair_HSTRING_HSTRING) Then
                Dim pIIterator_IKeyValuePair_HSTRING_HSTRING As Long
                If Invoke(pIIterable_IKeyValuePair_HSTRING_HSTRING, _
                          IIterable_IKeyValuePair_HSTRING_HSTRING_First, _
                          VarPtr(pIIterator_IKeyValuePair_HSTRING_HSTRING)) = S_OK Then
                    Dim Item As Long
                    For Item = 0 To count - 1
                        Dim Value As Long
                        Dim pIKeyValuePair_HSTRING_HSTRING As Long
                        Set Ret(Item) = New KeyValuePair_String_String
                        If Item = 0 Then
                            If Invoke(pIIterator_IKeyValuePair_HSTRING_HSTRING, _
                                      IIterator_IKeyValuePair_HSTRING_HSTRING_GetCurrent, _
                                      VarPtr(pIKeyValuePair_HSTRING_HSTRING)) = S_OK Then
                                If Invoke(pIKeyValuePair_HSTRING_HSTRING, _
                                          IKeyValuePair_HSTRING_HSTRING_GetKey, _
                                          VarPtr(Value)) = S_OK Then
                                    If Value <> 0& Then
                                        Ret(Item).Key = GetWindowsString(Value)
                                    End If
                                End If
                                If Invoke(pIKeyValuePair_HSTRING_HSTRING, _
                                          IKeyValuePair_HSTRING_HSTRING_GetValue, _
                                          VarPtr(Value)) = S_OK Then
                                    If Value <> 0& Then
                                        Ret(Item).Value = GetWindowsString(Value)
                                    End If
                                End If
                                Call ReleaseIfc(pIKeyValuePair_HSTRING_HSTRING)
                            End If
                        Else
                            Dim moveNext As Long
                            If Invoke(pIIterator_IKeyValuePair_HSTRING_HSTRING, _
                                      IIterator_IKeyValuePair_HSTRING_HSTRING_MoveNext, _
                                      VarPtr(moveNext)) = S_OK Then
                                If CBool(moveNext) Then
                                    If Invoke(pIIterator_IKeyValuePair_HSTRING_HSTRING, _
                                              IIterator_IKeyValuePair_HSTRING_HSTRING_GetCurrent, _
                                              VarPtr(pIKeyValuePair_HSTRING_HSTRING)) = S_OK Then
                                        If Invoke(pIKeyValuePair_HSTRING_HSTRING, _
                                                  IKeyValuePair_HSTRING_HSTRING_GetKey, _
                                                  VarPtr(Value)) = S_OK Then
                                            If Value <> 0& Then
                                                Ret(Item).Key = GetWindowsString(Value)
                                            End If
                                        End If
                                        If Invoke(pIKeyValuePair_HSTRING_HSTRING, _
                                                  IKeyValuePair_HSTRING_HSTRING_GetValue, _
                                                  VarPtr(Value)) = S_OK Then
                                            If Value <> 0& Then
                                                Ret(Item).Value = GetWindowsString(Value)
                                            End If
                                        End If
                                        Call ReleaseIfc(pIKeyValuePair_HSTRING_HSTRING)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    Call ReleaseIfc(pIIterator_IKeyValuePair_HSTRING_HSTRING)
                End If
                Call ReleaseIfc(pIIterable_IKeyValuePair_HSTRING_HSTRING)
            End If
        End If
    End If
    GetKeyValuePairs = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


