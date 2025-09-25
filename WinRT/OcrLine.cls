VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OcrLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Ocr_OcrResult As String = "Windows.Media.Ocr.OcrLine"
Private Const IID_IOcrLine As String = "{0043a16f-e31f-3a24-899c-d444bd088124}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IOcrLine
    IOcrLine_GetWords = 6
    IOcrLine_GetText = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIOcrLine As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIOcrLine)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIOcrLine
End Property

Public Property Let Ifc(ByVal pIOcrLine As Long)
    Call ReleaseIfc(m_pIOcrLine)
    m_pIOcrLine = pIOcrLine
End Property

' IOcrLine
Public Property Get Words() As ReadOnlyList_1 'ReadOnlyList_OcrWord
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_OcrWord
    If m_pIOcrLine <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIOcrLine, _
                  IOcrLine_GetWords, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_OcrWord, pIVectorView) 'ReadOnlyList_OcrWord
            End If
        End If
    End If
    Set Words = Ret
End Property

Public Property Get Text() As String
    Dim Ret As String
    If m_pIOcrLine <> 0& Then
        Dim value As Long
        If Invoke(m_pIOcrLine, _
                  IOcrLine_GetText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Text = Ret
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


