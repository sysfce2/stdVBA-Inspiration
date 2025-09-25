VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_String_Inspectable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const IID_IMapView_HSTRING_IInspectable As String = "{bb78502a-f79d-54fa-92c9-90c5039fdf7e}"
Private Const IID_IIterable_IKeyValuePair_HSTRING_IInspectable As String = "{fe2f3d47-5d47-5499-8374-430c7cda0204}"
Private Const IID_IIterator_IKeyValuePair_HSTRING_IInspectable As String = "{5db5fa32-707c-5849-a06b-91c8eb9d10e8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMapView_HSTRING_IInspectable
    IMapView_HSTRING_IInspectable_Lookup = 6
    IMapView_HSTRING_IInspectable_GetSize = 7
    IMapView_HSTRING_IInspectable_HasKey = 8
    IMapView_HSTRING_IInspectable_Split = 9
                        
    ' IIterable_IKeyValuePair_HSTRING_IInspectable
    IIterable_IKeyValuePair_HSTRING_IInspectable_First = 6
                        
    ' IIterator_IKeyValuePair_HSTRING_IInspectable
    IIterator_IKeyValuePair_HSTRING_IInspectable_GetCurrent = 6
    IIterator_IKeyValuePair_HSTRING_IInspectable_GetHasCurrent = 7
    IIterator_IKeyValuePair_HSTRING_IInspectable_MoveNext = 8
    IIterator_IKeyValuePair_HSTRING_IInspectable_GetMany = 9

    ' IKeyValuePair_HSTRING_IInspectable
    IKeyValuePair_HSTRING_IInspectable_GetKey = 6
    IKeyValuePair_HSTRING_IInspectable_GetValue = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIMapView_HSTRING_IInspectable As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMapView_HSTRING_IInspectable)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMapView_HSTRING_IInspectable
End Property

Public Property Let Ifc(ByVal pIMapView_HSTRING_IInspectable As Long)
    Call ReleaseIfc(m_pIMapView_HSTRING_IInspectable)
    m_pIMapView_HSTRING_IInspectable = pIMapView_HSTRING_IInspectable
End Property

' IMapView_HSTRING_IInspectable
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIMapView_HSTRING_IInspectable <> 0& Then
        Dim value As Long
        If Invoke(m_pIMapView_HSTRING_IInspectable, _
                  IMapView_HSTRING_IInspectable_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
' IMapView_HSTRING_IInspectable
Public Function Lookup(ByVal Key As String) As Inspectable
    Dim Ret As Inspectable
    If m_pIMapView_HSTRING_IInspectable <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim pIInspectable As Long
        If Invoke(m_pIMapView_HSTRING_IInspectable, _
                  IMapView_HSTRING_IInspectable_Lookup, _
                  hString, _
                  VarPtr(pIInspectable)) = S_OK Then
            If pIInspectable <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = pIInspectable
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Lookup = Ret
End Function

Public Function HasKey(ByVal Key As String) As Boolean
    Dim Ret As Boolean
    If m_pIMapView_HSTRING_IInspectable <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim value As Long
        If Invoke(m_pIMapView_HSTRING_IInspectable, _
                  IMapView_HSTRING_IInspectable_HasKey, _
                  hString, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
        Call DeleteWindowsString(hString)
    End If
    HasKey = Ret
End Function

Public Function Split(ByRef First As ReadOnlyList_String_Inspectable, _
                      ByRef second As ReadOnlyList_String_Inspectable) As Boolean
    Dim Ret As Boolean
    If m_pIMapView_HSTRING_IInspectable <> 0& Then
        Dim hFirst As Long
        Dim hSecond As Long
        If Invoke(m_pIMapView_HSTRING_IInspectable, _
                  IMapView_HSTRING_IInspectable_Split, _
                  VarPtr(hFirst), _
                  VarPtr(hSecond)) = S_OK Then
            If hFirst <> 0& Then
                Set First = New ReadOnlyList_String_Inspectable
                First.Ifc = hFirst
            End If
            If hSecond <> 0& Then
                Set second = New ReadOnlyList_String_Inspectable
                second.Ifc = hSecond
            End If
            Ret = True
        End If
    End If
    Split = Ret
End Function

' IIterable_IKeyValuePair_HSTRING_IInspectable
' IIterator_IKeyValuePair_HSTRING_IInspectable
' IKeyValuePair_HSTRING_IInspectable
Public Function GetKeyValuePairs() As KeyValuePair_String_Inspectable()
    Dim Ret() As KeyValuePair_String_Inspectable
    If m_pIMapView_HSTRING_IInspectable <> 0& Then
        Dim count As Long
        count = Me.Size
        If count > 0& Then
            ReDim Ret(count - 1)
            Dim pIIterable_IKeyValuePair_HSTRING_IInspectable As Long
            If QueryIfc(m_pIMapView_HSTRING_IInspectable, _
                        IID_IIterable_IKeyValuePair_HSTRING_IInspectable, _
                        pIIterable_IKeyValuePair_HSTRING_IInspectable) Then
                Dim pIIterator_IKeyValuePair_HSTRING_IInspectable As Long
                If Invoke(pIIterable_IKeyValuePair_HSTRING_IInspectable, _
                          IIterable_IKeyValuePair_HSTRING_IInspectable_First, _
                          VarPtr(pIIterator_IKeyValuePair_HSTRING_IInspectable)) = S_OK Then
                    Dim Item As Long
                    For Item = 0 To count - 1
                        Dim value As Long
                        Dim Inspectable As Inspectable
                        Dim pIKeyValuePair_HSTRING_IInspectable As Long
                        Set Ret(Item) = New KeyValuePair_String_Inspectable
                        If Item = 0 Then
                            If Invoke(pIIterator_IKeyValuePair_HSTRING_IInspectable, _
                                      IIterator_IKeyValuePair_HSTRING_IInspectable_GetCurrent, _
                                      VarPtr(pIKeyValuePair_HSTRING_IInspectable)) = S_OK Then
                                If Invoke(pIKeyValuePair_HSTRING_IInspectable, _
                                          IKeyValuePair_HSTRING_IInspectable_GetKey, _
                                          VarPtr(value)) = S_OK Then
                                    If value <> 0& Then
                                        Ret(Item).Key = GetWindowsString(value)
                                    End If
                                End If
                                If Invoke(pIKeyValuePair_HSTRING_IInspectable, _
                                          IKeyValuePair_HSTRING_IInspectable_GetValue, _
                                          VarPtr(value)) = S_OK Then
                                    If value <> 0& Then
                                        Set Inspectable = New Inspectable
                                        Inspectable.Ifc = value
                                        Ret(Item).value = Inspectable
                                    End If
                                End If
                                Call ReleaseIfc(pIKeyValuePair_HSTRING_IInspectable)
                            End If
                        Else
                            Dim moveNext As Long
                            If Invoke(pIIterator_IKeyValuePair_HSTRING_IInspectable, _
                                      IIterator_IKeyValuePair_HSTRING_IInspectable_MoveNext, _
                                      VarPtr(moveNext)) = S_OK Then
                                If CBool(moveNext) Then
                                    If Invoke(pIIterator_IKeyValuePair_HSTRING_IInspectable, _
                                              IIterator_IKeyValuePair_HSTRING_IInspectable_GetCurrent, _
                                              VarPtr(pIKeyValuePair_HSTRING_IInspectable)) = S_OK Then
                                        If Invoke(pIKeyValuePair_HSTRING_IInspectable, _
                                                  IKeyValuePair_HSTRING_IInspectable_GetKey, _
                                                  VarPtr(value)) = S_OK Then
                                            If value <> 0& Then
                                                Ret(Item).Key = GetWindowsString(value)
                                            End If
                                        End If
                                        If Invoke(pIKeyValuePair_HSTRING_IInspectable, _
                                                  IKeyValuePair_HSTRING_IInspectable_GetValue, _
                                                  VarPtr(value)) = S_OK Then
                                            If value <> 0& Then
                                                Set Inspectable = New Inspectable
                                                Inspectable.Ifc = value
                                                Ret(Item).value = Inspectable
                                            End If
                                        End If
                                        Call ReleaseIfc(pIKeyValuePair_HSTRING_IInspectable)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    Call ReleaseIfc(pIIterator_IKeyValuePair_HSTRING_IInspectable)
                End If
                Call ReleaseIfc(pIIterable_IKeyValuePair_HSTRING_IInspectable)
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

