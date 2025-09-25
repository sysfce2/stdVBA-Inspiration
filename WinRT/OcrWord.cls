VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OcrWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Ocr_OcrResult As String = "Windows.Media.Ocr.OcrWord"
Private Const IID_IOcrWord As String = "{3c2a477a-5cd9-3525-ba2a-23d1e0a68a1d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IOcrWord
    IOcrWord_GetBoundingRect = 6
    IOcrWord_GetText = 7
    
End Enum

' ----==== Types ====----
Private Type RECTF
    X As Single
    Y As Single
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pIOcrWord As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIOcrWord)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIOcrWord
End Property

Public Property Let Ifc(ByVal pIOcrWord As Long)
    Call ReleaseIfc(m_pIOcrWord)
    m_pIOcrWord = pIOcrWord
End Property

' IOcrWord
Public Property Get BoundingRect() As Rect
    Dim Ret As Rect
    If m_pIOcrWord <> 0& Then
        Dim tRECTF As RECTF
        If Invoke(m_pIOcrWord, _
                  IOcrWord_GetBoundingRect, _
                  VarPtr(tRECTF)) = S_OK Then
            Set Ret = New Rect
            Ret.X = tRECTF.X
            Ret.Y = tRECTF.Y
            Ret.Width = tRECTF.Width
            Ret.Height = tRECTF.Height
        End If
    End If
    Set BoundingRect = Ret
End Property

Public Property Get Text() As String
    Dim Ret As String
    If m_pIOcrWord <> 0& Then
        Dim Value As Long
        If Invoke(m_pIOcrWord, _
                  IOcrWord_GetText, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
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


