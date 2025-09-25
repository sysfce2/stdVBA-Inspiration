VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPageRenderOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Pdf_PdfPageRenderOptions As String = "Windows.Data.Pdf.PdfPageRenderOptions"
Private Const IID_IPdfPageRenderOption As String = "{3c98056f-b7cf-4c29-9a04-52d90267f425}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPdfPageRenderOptions
    IPdfPageRenderOptions_GetSourceRect = 6
    IPdfPageRenderOptions_PutSourceRect = 7
    IPdfPageRenderOptions_GetDestinationWidth = 8
    IPdfPageRenderOptions_PutDestinationWidth = 9
    IPdfPageRenderOptions_GetDestinationHeight = 10
    IPdfPageRenderOptions_PutDestinationHeight = 11
    IPdfPageRenderOptions_GetBackgroundColor = 12
    IPdfPageRenderOptions_PutBackgroundColor = 13
    IPdfPageRenderOptions_GetIsIgnoringHighContrast = 14
    IPdfPageRenderOptions_PutIsIgnoringHighContrast = 15
    IPdfPageRenderOptions_GetBitmapEncoderId = 16
    IPdfPageRenderOptions_PutBitmapEncoderId = 17

End Enum

' ----==== Types ====----
Private Type RECTF
    X As Single
    Y As Single
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pIPdfPageRenderOptions As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_Data_Pdf_PdfPageRenderOptions, _
                           IID_IPdfPageRenderOption, _
                           m_pIPdfPageRenderOptions) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPdfPageRenderOptions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPdfPageRenderOptions
End Property

' IPdfPageRenderOptions
Public Property Get SourceRect() As Rect
    Dim Ret As Rect
    If m_pIPdfPageRenderOptions <> 0& Then
        Dim Value As RECTF
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_GetSourceRect, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Rect
            Ret.X = Value.X
            Ret.Y = Value.Y
            Ret.Width = Value.Width
            Ret.Height = Value.Height
        End If
    End If
    Set SourceRect = Ret
End Property

Public Property Let SourceRect(ByVal Value As Rect)
    If m_pIPdfPageRenderOptions <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_PutSourceRect, _
                  Value.X, _
                  Value.Y, _
                  Value.Width, _
                  Value.Height) = S_OK Then
        End If
    End If
End Property

Public Property Get DestinationWidth() As Long
    Dim Ret As Long
    If m_pIPdfPageRenderOptions <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_GetDestinationWidth, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    DestinationWidth = Ret
End Property

Public Property Let DestinationWidth(ByVal Value As Long)
    If m_pIPdfPageRenderOptions <> 0& Then
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_PutDestinationWidth, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get DestinationHeight() As Long
    Dim Ret As Long
    If m_pIPdfPageRenderOptions <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_GetDestinationHeight, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    DestinationHeight = Ret
End Property

Public Property Let DestinationHeight(ByVal Value As Long)
    If m_pIPdfPageRenderOptions <> 0& Then
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_PutDestinationHeight, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get BackgroundColor() As Color
    Dim Ret As Color
    If m_pIPdfPageRenderOptions <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_GetBackgroundColor, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set BackgroundColor = Ret
End Property

Public Property Let BackgroundColor(ByVal Value As Color)
    If m_pIPdfPageRenderOptions <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_PutBackgroundColor, _
                  Value.ColorValue) = S_OK Then
        End If
    End If
End Property

Public Property Get IsIgnoringHighContrast() As Boolean
    Dim Ret As Boolean
    If m_pIPdfPageRenderOptions <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_GetIsIgnoringHighContrast, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    IsIgnoringHighContrast = Ret
End Property

Public Property Let IsIgnoringHighContrast(ByVal Value As Boolean)
    If m_pIPdfPageRenderOptions <> 0& Then
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_PutIsIgnoringHighContrast, _
                  Value) = S_OK Then
        End If
    End If
End Property

Friend Property Get BitmapEncoderId() As GUID
    Dim Ret As GUID
    If m_pIPdfPageRenderOptions <> 0& Then
        Dim Value As GUID
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_GetBitmapEncoderId, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    BitmapEncoderId = Ret
End Property

Friend Property Let BitmapEncoderId(ByRef Value As GUID)
    If m_pIPdfPageRenderOptions <> 0& Then
        If Invoke(m_pIPdfPageRenderOptions, _
                  IPdfPageRenderOptions_PutBitmapEncoderId, _
                  Value.data1, _
                  Value.data2, _
                  Value.data3, _
                  Value.data4) = S_OK Then
        End If
    End If
End Property

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


