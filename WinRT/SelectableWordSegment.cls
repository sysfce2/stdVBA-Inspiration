VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectableWordSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Text_SelectableWordSegment As String = "Windows.Data.Text.SelectableWordSegment"
Private Const IID_ISelectableWordSegment As String = "{916a4cb7-8aa7-4c78-b374-5dedb752e60b}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISelectableWordSegment
    ISelectableWordSegment_GetText = 6
    ISelectableWordSegment_GetSourceTextSegment = 7
                           
End Enum

' ----==== Types ====----
Private Type TextSegment_
    StartPosition As Long
    Length As Long
End Type

' ----==== Vars ====----
Private m_pISelectableWordSegment As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISelectableWordSegment)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISelectableWordSegment
End Property

Public Property Let Ifc(ByVal pISelectableWordSegment As Long)
    Call ReleaseIfc(m_pISelectableWordSegment)
    m_pISelectableWordSegment = pISelectableWordSegment
End Property

' ISelectableWordSegment
Public Property Get Text() As String
    Dim Ret As String
    If m_pISelectableWordSegment <> 0& Then
        Dim Value As Long
        If Invoke(m_pISelectableWordSegment, _
                  ISelectableWordSegment_GetText, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Text = Ret
End Property

Public Property Get SourceTextSegment() As TextSegment
    Dim Ret As TextSegment
    If m_pISelectableWordSegment <> 0& Then
        Dim Value As TextSegment_
        If Invoke(m_pISelectableWordSegment, _
                  ISelectableWordSegment_GetSourceTextSegment, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TextSegment
            Ret.StartPosition = Value.StartPosition
            Ret.Length = Value.Length
        End If
    End If
    Set SourceTextSegment = Ret
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


