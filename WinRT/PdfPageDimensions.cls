VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPageDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Pdf_PdfPageDimensions As String = "Windows.Data.Pdf.PdfPageDimensions"
Private Const IID_IPdfPageDimensions As String = "{22170471-313e-44e8-835d-63a3e7624a10}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPdfPageDimensions
    IPdfPageDimensions_GetMediaBox = 6
    IPdfPageDimensions_GetCropBox = 7
    IPdfPageDimensions_GetBleedBox = 8
    IPdfPageDimensions_GetTrimBox = 9
    IPdfPageDimensions_GetArtBox = 10

End Enum

' ----==== Types ====----
Private Type RECTF
    X As Single
    Y As Single
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pIPdfPageDimensions As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPdfPageDimensions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPdfPageDimensions
End Property

Public Property Let Ifc(ByVal pIPdfPageDimensions As Long)
    Call ReleaseIfc(m_pIPdfPageDimensions)
    m_pIPdfPageDimensions = pIPdfPageDimensions
End Property

' IPdfPageDimensions
Public Property Get MediaBox() As Rect
    Dim Ret As Rect
    If m_pIPdfPageDimensions <> 0& Then
        Dim Value As RECTF
        If Invoke(m_pIPdfPageDimensions, _
                  IPdfPageDimensions_GetMediaBox, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Rect
            Ret.X = Value.X
            Ret.Y = Value.Y
            Ret.Width = Value.Width
            Ret.Height = Value.Height
        End If
    End If
    Set MediaBox = Ret
End Property

Public Property Get CropBox() As Rect
    Dim Ret As Rect
    If m_pIPdfPageDimensions <> 0& Then
        Dim Value As RECTF
        If Invoke(m_pIPdfPageDimensions, _
                  IPdfPageDimensions_GetCropBox, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Rect
            Ret.X = Value.X
            Ret.Y = Value.Y
            Ret.Width = Value.Width
            Ret.Height = Value.Height
        End If
    End If
    Set CropBox = Ret
End Property

Public Property Get BleedBox() As Rect
    Dim Ret As Rect
    If m_pIPdfPageDimensions <> 0& Then
        Dim Value As RECTF
        If Invoke(m_pIPdfPageDimensions, _
                  IPdfPageDimensions_GetBleedBox, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Rect
            Ret.X = Value.X
            Ret.Y = Value.Y
            Ret.Width = Value.Width
            Ret.Height = Value.Height
        End If
    End If
    Set BleedBox = Ret
End Property

Public Property Get TrimBox() As Rect
    Dim Ret As Rect
    If m_pIPdfPageDimensions <> 0& Then
        Dim Value As RECTF
        If Invoke(m_pIPdfPageDimensions, _
                  IPdfPageDimensions_GetTrimBox, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Rect
            Ret.X = Value.X
            Ret.Y = Value.Y
            Ret.Width = Value.Width
            Ret.Height = Value.Height
        End If
    End If
    Set TrimBox = Ret
End Property

Public Property Get ArtBox() As Rect
    Dim Ret As Rect
    If m_pIPdfPageDimensions <> 0& Then
        Dim Value As RECTF
        If Invoke(m_pIPdfPageDimensions, _
                  IPdfPageDimensions_GetArtBox, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Rect
            Ret.X = Value.X
            Ret.Y = Value.Y
            Ret.Width = Value.Width
            Ret.Height = Value.Height
        End If
    End If
    Set ArtBox = Ret
End Property

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


