VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_Language"
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

    ' IVectorView_ILanguage
    IVectorView_ILanguage_GetAt = 6
    IVectorView_ILanguage_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIVectorView_ILanguage As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVectorView_ILanguage)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVectorView_ILanguage
End Property

Public Property Let Ifc(ByVal pIVectorView_ILanguage As Long)
    Call ReleaseIfc(m_pIVectorView_ILanguage)
    m_pIVectorView_ILanguage = pIVectorView_ILanguage
End Property

'IVectorView_ILanguage
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVectorView_ILanguage <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVectorView_ILanguage, _
                  IVectorView_ILanguage_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVectorView_ILanguage
Public Function GetAt(ByVal index As Long) As language
    Dim Ret As language
    If m_pIVectorView_ILanguage <> 0& Then
        Dim pILanguage As Long
        If Invoke(m_pIVectorView_ILanguage, _
                  IVectorView_ILanguage_GetAt, _
                  index, _
                  VarPtr(pILanguage)) = S_OK Then
            If pILanguage <> 0& Then
                Set Ret = New language
                Ret.Ifc = pILanguage
            End If
        End If
    End If
    Set GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As language()
    Dim Ret() As language
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

