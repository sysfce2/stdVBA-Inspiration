VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlternateWordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Text_AlternateWordForm As String = "Windows.Data.Text.AlternateWordForm"
Private Const IID_IAlternateWordForm As String = "{47396c1e-51b9-4207-9146-248e636a1d1d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAlternateWordForm
    IAlternateWordForm_GetSourceTextSegment = 6
    IAlternateWordForm_GetAlternateText = 7
    IAlternateWordForm_GetNormalizationFormat = 8
                           
End Enum

' ----==== Types ====----
Private Type TextSegment_
    StartPosition As Long
    Length As Long
End Type

' ----==== Vars ====----
Private m_pIAlternateWordForm As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAlternateWordForm)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIAlternateWordForm
End Property

Public Property Let Ifc(ByVal pIAlternateWordForm As Long)
    Call ReleaseIfc(m_pIAlternateWordForm)
    m_pIAlternateWordForm = pIAlternateWordForm
End Property

' IAlternateWordForm
Public Property Get SourceTextSegment() As TextSegment
    Dim Ret As TextSegment
    If m_pIAlternateWordForm <> 0& Then
        Dim Value As TextSegment_
        If Invoke(m_pIAlternateWordForm, _
                  IAlternateWordForm_GetSourceTextSegment, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TextSegment
            Ret.StartPosition = Value.StartPosition
            Ret.Length = Value.Length
        End If
    End If
    Set SourceTextSegment = Ret
End Property

Public Property Get AlternateText() As String
    Dim Ret As String
    If m_pIAlternateWordForm <> 0& Then
        Dim Value As Long
        If Invoke(m_pIAlternateWordForm, _
                  IAlternateWordForm_GetAlternateText, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    AlternateText = Ret
End Property

Public Property Get NormalizationFormat() As AlternateNormalizationFormat
    Dim Ret As AlternateNormalizationFormat
    If m_pIAlternateWordForm <> 0& Then
        Dim Value As Long
        If Invoke(m_pIAlternateWordForm, _
                  IAlternateWordForm_GetNormalizationFormat, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    NormalizationFormat = Ret
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


