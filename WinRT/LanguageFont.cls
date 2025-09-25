VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LanguageFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_Fonts_LanguageFont As String = "Windows.Globalization.Fonts.LanguageFont"
Private Const IID_ILanguageFont As String = "{b12e5c3a-b76d-459b-beeb-901151cd77d1}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ILanguageFont
    ILanguageFont_GetFontFamily = 6
    ILanguageFont_GetFontWeight = 7
    ILanguageFont_GetFontStretch = 8
    ILanguageFont_GetFontStyle = 9
    ILanguageFont_GetScaleFactor = 10

End Enum

Private Type FontWeight_
    Weight As Integer
End Type

' ----==== Variablen ====----
Private m_pILanguageFont As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILanguageFont)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pILanguageFont
End Property

Public Property Let Ifc(ByVal pILanguageFont As Long)
    Call ReleaseIfc(m_pILanguageFont)
    m_pILanguageFont = pILanguageFont
End Property

' ILanguageFont
Public Property Get FontFamily() As String
    Dim Ret As String
    If m_pILanguageFont <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguageFont, _
            ILanguageFont_GetFontFamily, _
            VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    FontFamily = Ret
End Property

Public Property Get FontWeight() As FontWeight
    Dim Ret As FontWeight
    If m_pILanguageFont <> 0& Then
        Dim value As FontWeight_
        If Invoke(m_pILanguageFont, _
            ILanguageFont_GetFontWeight, _
            VarPtr(value)) = S_OK Then
            Set Ret = New FontWeight
            Ret.Weight = value.Weight
        End If
    End If
    Set FontWeight = Ret
End Property

Public Property Get FontStretch() As FontStretch
    Dim Ret As FontStretch
    If m_pILanguageFont <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguageFont, _
            ILanguageFont_GetFontStretch, _
            VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FontStretch = Ret
End Property

Public Property Get FontStyle() As FontStyle
    Dim Ret As FontStyle
    If m_pILanguageFont <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguageFont, _
            ILanguageFont_GetFontStyle, _
            VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FontStyle = Ret
End Property

Public Property Get ScaleFactor() As Double
    Dim Ret As Double
    If m_pILanguageFont <> 0& Then
        Dim value As Double
        If Invoke(m_pILanguageFont, _
            ILanguageFont_GetScaleFactor, _
            VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ScaleFactor = Ret
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


