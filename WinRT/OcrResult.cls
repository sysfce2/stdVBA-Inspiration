VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OcrResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Ocr_OcrResult As String = "Windows.Media.Ocr.OcrResult"
Private Const IID_IOcrResult As String = "{9bd235b2-175b-3d6a-92e2-388c206e2f63}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IOcrResult
    IOcrResult_GetLines = 6
    IOcrResult_GetTextAngle = 7
    IOcrResult_GetText = 8
    
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIOcrResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIOcrResult)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIOcrResult
End Property

Public Property Let Ifc(ByVal pIOcrResult As Long)
    Call ReleaseIfc(m_pIOcrResult)
    m_pIOcrResult = pIOcrResult
End Property

' IOcrResult
Public Property Get Lines() As ReadOnlyList_1 'ReadOnlyList_OcrLine
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_OcrLine
    If m_pIOcrResult <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIOcrResult, _
                  IOcrResult_GetLines, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_OcrLine, pIVectorView) 'ReadOnlyList_OcrLine
            End If
        End If
    End If
    Set Lines = Ret
End Property

Public Property Get TextAngle() As Double
    Dim Ret As Double
    If m_pIOcrResult <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIOcrResult, _
                  IOcrResult_GetTextAngle, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Double
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    TextAngle = Ret
End Property

Public Property Get Text() As String
    Dim Ret As String
    If m_pIOcrResult <> 0& Then
        Dim value As Long
        If Invoke(m_pIOcrResult, _
                  IOcrResult_GetText, _
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


