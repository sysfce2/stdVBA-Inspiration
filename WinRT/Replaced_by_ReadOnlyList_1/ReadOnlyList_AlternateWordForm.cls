VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_AlternateWordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVectorView_IAlternateWordForm
    IVectorView_IAlternateWordForm_GetAt = 6
    IVectorView_IAlternateWordForm_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIVectorView_IAlternateWordForm As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVectorView_IAlternateWordForm)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVectorView_IAlternateWordForm
End Property

Public Property Let Ifc(ByVal pIVectorView_IAlternateWordForm As Long)
    Call ReleaseIfc(m_pIVectorView_IAlternateWordForm)
    m_pIVectorView_IAlternateWordForm = pIVectorView_IAlternateWordForm
End Property

'IVectorView_IAlternateWordForm
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVectorView_IAlternateWordForm <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVectorView_IAlternateWordForm, _
                  IVectorView_IAlternateWordForm_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVectorView_IAlternateWordForm
Public Function GetAt(ByVal index As Long) As AlternateWordForm
    Dim Ret As AlternateWordForm
    If m_pIVectorView_IAlternateWordForm <> 0& Then
        Dim pIAlternateWordForm As Long
        If Invoke(m_pIVectorView_IAlternateWordForm, _
                  IVectorView_IAlternateWordForm_GetAt, _
                  index, _
                  VarPtr(pIAlternateWordForm)) = S_OK Then
            If pIAlternateWordForm <> 0& Then
                Set Ret = New AlternateWordForm
                Ret.Ifc = pIAlternateWordForm
            End If
        End If
    End If
    Set GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As AlternateWordForm()
    Dim Ret() As AlternateWordForm
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


