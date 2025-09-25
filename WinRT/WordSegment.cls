VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WordSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Text_WordSegment As String = "Windows.Data.Text.WordSegment"
Private Const IID_IWordSegment As String = "{d2d4ba6d-987c-4cc0-b6bd-d49a11b38f9a}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IWordSegment
    IWordSegment_GetText = 6
    IWordSegment_GetSourceTextSegment = 7
    IWordSegment_GetAlternateForms = 8
                           
End Enum

' ----==== Types ====----
Private Type TextSegment_
    StartPosition As Long
    Length As Long
End Type

' ----==== Vars ====----
Private m_pIWordSegment As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIWordSegment)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIWordSegment
End Property

Public Property Let Ifc(ByVal pIWordSegment As Long)
    Call ReleaseIfc(m_pIWordSegment)
    m_pIWordSegment = pIWordSegment
End Property

' IWordSegment
Public Property Get Text() As String
    Dim Ret As String
    If m_pIWordSegment <> 0& Then
        Dim value As Long
        If Invoke(m_pIWordSegment, _
                  IWordSegment_GetText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Text = Ret
End Property

Public Property Get SourceTextSegment() As TextSegment
    Dim Ret As TextSegment
    If m_pIWordSegment <> 0& Then
        Dim value As TextSegment_
        If Invoke(m_pIWordSegment, _
                  IWordSegment_GetSourceTextSegment, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TextSegment
            Ret.StartPosition = value.StartPosition
            Ret.Length = value.Length
        End If
    End If
    Set SourceTextSegment = Ret
End Property

Public Property Get AlternateForms() As ReadOnlyList_1 'ReadOnlyList_AlternateWordForm
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AlternateWordForm
    If m_pIWordSegment <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIWordSegment, _
                  IWordSegment_GetAlternateForms, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_AlternateWordForm, pIVectorView) 'ReadOnlyList_AlternateWordForm
            End If
        End If
    End If
    Set AlternateForms = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

