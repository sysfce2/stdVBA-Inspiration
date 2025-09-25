VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const IID_IMapView_HSTRING_IInspectable As String = "{bb78502a-f79d-54fa-92c9-90c5039fdf7e}"
Private Const IID_IIterable_IKeyValuePair_HSTRING_IInspectable As String = "{fe2f3d47-5d47-5499-8374-430c7cda0204}"
Private Const IID_IIterator_IKeyValuePair_HSTRING_IInspectable As String = "{5db5fa32-707c-5849-a06b-91c8eb9d10e8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMapView_xxx_yyy
    IMapView_xxx_yyy_Lookup = 6
    IMapView_xxx_yyy_GetSize = 7
    IMapView_xxx_yyy_HasKey = 8
    IMapView_xxx_yyy_Split = 9
                        
    ' IIterable_IKeyValuePair_xxx_yyy
    IIterable_IKeyValuePair_xxx_yyy_First = 6
                        
    ' IIterator_IKeyValuePair_xxx_yyy
    IIterator_IKeyValuePair_xxx_yyy_GetCurrent = 6
    IIterator_IKeyValuePair_xxx_yyy_GetHasCurrent = 7
    IIterator_IKeyValuePair_xxx_yyy_MoveNext = 8
'    IIterator_IKeyValuePair_xxx_yyy_GetMany = 9

    ' IKeyValuePair_xxx_yyy
    IKeyValuePair_xxx_yyy_GetKey = 6
    IKeyValuePair_xxx_yyy_GetValue = 7
    
End Enum

Public Enum OfType_xxx_yyy
    Of_String_Inspectable
End Enum

' ----==== Variablen ====----
Private m_OfType_xxx_yyy As OfType_xxx_yyy
Private m_pIMapView_xxx_yyy As Long
Private m_pIIterable_IKeyValuePair_xxx_yyy As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    m_OfType_xxx_yyy = -1
    Call ReleaseIfc(m_pIIterable_IKeyValuePair_xxx_yyy)
    Call ReleaseIfc(m_pIMapView_xxx_yyy)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMapView_xxx_yyy
End Property

Public Property Let Ifc(ByVal pIMapView As Long)
    Call ReleaseIfc(m_pIIterable_IKeyValuePair_xxx_yyy)
    Call ReleaseIfc(m_pIMapView_xxx_yyy)
    m_pIMapView_xxx_yyy = pIMapView
    Dim IID_IIterable_xxx_yyy As String
    Select Case m_OfType_xxx_yyy
        Case OfType_xxx_yyy.Of_String_Inspectable
            IID_IIterable_xxx_yyy = IID_IIterable_IKeyValuePair_HSTRING_IInspectable
    End Select
    Call QueryIfc(m_pIMapView_xxx_yyy, IID_IIterable_xxx_yyy, m_pIIterable_IKeyValuePair_xxx_yyy)
End Property

Public Property Let Of(ByVal ReturnAs As OfType_xxx_yyy)
    m_OfType_xxx_yyy = ReturnAs
End Property

' IMapView_xxx_yyy
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIMapView_xxx_yyy <> 0& Then
        Dim value As Long
        If Invoke(m_pIMapView_xxx_yyy, _
                  IMapView_xxx_yyy_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
Public Function Lookup(ByVal Key As Variant) As Inspectable
    Dim Ret As Inspectable
    If m_pIMapView_xxx_yyy <> 0& Then
        Dim pValue As Long
        If VarType(Key) = vbString Then
            pValue = CreateWindowsString(Key)
        Else
            pValue = Key.Ifc
        End If
        Dim pIInspectable As Long
        If Invoke(m_pIMapView_xxx_yyy, _
                  IMapView_xxx_yyy_Lookup, _
                  pValue, _
                  VarPtr(pIInspectable)) = S_OK Then
            If pIInspectable <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = pIInspectable
            End If
        End If
        If VarType(Key) = vbString Then
            Call DeleteWindowsString(pValue)
        End If
    End If
    Set Lookup = Ret
End Function

Public Function HasKey(ByVal Key As Variant) As Boolean
    Dim Ret As Boolean
    If m_pIMapView_xxx_yyy <> 0& Then
        Dim pValue As Long
        If VarType(Key) = vbString Then
            pValue = CreateWindowsString(Key)
        Else
            pValue = Key.Ifc
        End If
        Dim value As Long
        If Invoke(m_pIMapView_xxx_yyy, _
                  IMapView_xxx_yyy_HasKey, _
                  pValue, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
        If VarType(Key) = vbString Then
            Call DeleteWindowsString(pValue)
        End If
    End If
    HasKey = Ret
End Function

Public Function Split(ByRef First As ReadOnlyList_2, _
                      ByRef Second As ReadOnlyList_2) As Boolean
    Dim Ret As Boolean
    If m_pIMapView_xxx_yyy <> 0& Then
        Dim hFirst As Long
        Dim hSecond As Long
        If Invoke(m_pIMapView_xxx_yyy, _
                  IMapView_xxx_yyy_Split, _
                  VarPtr(hFirst), _
                  VarPtr(hSecond)) = S_OK Then
            If hFirst <> 0& Then
                Set First = New ReadOnlyList_2
                First.Ifc = hFirst
                Select Case m_OfType_xxx_yyy
                    Case OfType_xxx_yyy.Of_String_Inspectable
                        First.Of = Of_String_Inspectable
                End Select
            End If
            If hSecond <> 0& Then
                Set Second = New ReadOnlyList_2
                Second.Ifc = hSecond
                Select Case m_OfType_xxx_yyy
                    Case OfType_xxx_yyy.Of_String_Inspectable
                        Second.Of = Of_String_Inspectable
                End Select
            End If
            Ret = True
        End If
    End If
    Split = Ret
End Function

' IIterable_IKeyValuePair_xxx_yyy
' IIterator_IKeyValuePair_xxx_yyy
' IKeyValuePair_xxx_yyy
Public Function GetKeyValuePairs() As Variant()
    Dim Ret() As Variant
    If m_pIMapView_xxx_yyy <> 0& Then
        Dim count As Long
        count = Me.Size
        If count > 0& Then
            ReDim Ret(count - 1)
            If m_pIIterable_IKeyValuePair_xxx_yyy Then
                Dim pIIterator_IKeyValuePair_xxx_yyy As Long
                If Invoke(m_pIIterable_IKeyValuePair_xxx_yyy, _
                          IIterable_IKeyValuePair_xxx_yyy_First, _
                          VarPtr(pIIterator_IKeyValuePair_xxx_yyy)) = S_OK Then
                    Dim Item As Long
                    For Item = 0 To count - 1
                        Dim value As Long
                        Dim Inspectable As Inspectable
                        Dim pIKeyValuePair_xxx_yyy As Long
                        
                        Select Case m_OfType_xxx_yyy
                            Case OfType_xxx_yyy.Of_String_Inspectable
                                Set Ret(Item) = New KeyValuePair_String_Inspectable
                        End Select
                        
                        If Item = 0 Then
                            If Invoke(pIIterator_IKeyValuePair_xxx_yyy, _
                                      IIterator_IKeyValuePair_xxx_yyy_GetCurrent, _
                                      VarPtr(pIKeyValuePair_xxx_yyy)) = S_OK Then
                                If Invoke(pIKeyValuePair_xxx_yyy, _
                                          IKeyValuePair_xxx_yyy_GetKey, _
                                          VarPtr(value)) = S_OK Then
                                    If value <> 0& Then
                                        Select Case m_OfType_xxx_yyy
                                            Case OfType_xxx_yyy.Of_String_Inspectable
                                                Ret(Item).Key = GetWindowsString(value)
                                        End Select
                                    End If
                                End If
                                If Invoke(pIKeyValuePair_xxx_yyy, _
                                          IKeyValuePair_xxx_yyy_GetValue, _
                                          VarPtr(value)) = S_OK Then
                                    If value <> 0& Then
                                        Select Case m_OfType_xxx_yyy
                                            Case OfType_xxx_yyy.Of_String_Inspectable
                                                Set Inspectable = New Inspectable
                                                Inspectable.Ifc = value
                                                Ret(Item).value = Inspectable
                                        End Select
                                    End If
                                End If
                                Call ReleaseIfc(pIKeyValuePair_xxx_yyy)
                            End If
                        Else
                            Dim MoveNext As Long
                            If Invoke(pIIterator_IKeyValuePair_xxx_yyy, _
                                      IIterator_IKeyValuePair_xxx_yyy_MoveNext, _
                                      VarPtr(MoveNext)) = S_OK Then
                                If CBool(MoveNext) Then
                                    If Invoke(pIIterator_IKeyValuePair_xxx_yyy, _
                                              IIterator_IKeyValuePair_xxx_yyy_GetCurrent, _
                                              VarPtr(pIKeyValuePair_xxx_yyy)) = S_OK Then
                                        If Invoke(pIKeyValuePair_xxx_yyy, _
                                                  IKeyValuePair_xxx_yyy_GetKey, _
                                                  VarPtr(value)) = S_OK Then
                                            If value <> 0& Then
                                                Select Case m_OfType_xxx_yyy
                                                    Case OfType_xxx_yyy.Of_String_Inspectable
                                                        Ret(Item).Key = GetWindowsString(value)
                                                End Select
                                            End If
                                        End If
                                        If Invoke(pIKeyValuePair_xxx_yyy, _
                                                  IKeyValuePair_xxx_yyy_GetValue, _
                                                  VarPtr(value)) = S_OK Then
                                            If value <> 0& Then
                                                Select Case m_OfType_xxx_yyy
                                                    Case OfType_xxx_yyy.Of_String_Inspectable
                                                        Set Inspectable = New Inspectable
                                                        Inspectable.Ifc = value
                                                        Ret(Item).value = Inspectable
                                                End Select
                                            End If
                                        End If
                                        Call ReleaseIfc(pIKeyValuePair_xxx_yyy)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    Call ReleaseIfc(pIIterator_IKeyValuePair_xxx_yyy)
                End If
                Call ReleaseIfc(m_pIIterable_IKeyValuePair_xxx_yyy)
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
