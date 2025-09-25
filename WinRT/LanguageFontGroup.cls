VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LanguageFontGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_Fonts_LanguageFontGroup As String = "Windows.Globalization.Fonts.LanguageFontGroup"
Private Const IID_ILanguageFontGroup As String = "{f33a7fc3-3a5c-4aea-b9ff-b39fb242f7f6}"
Private Const IID_ILanguageFontGroupFactory As String = "{fcaeac67-4e77-49c7-b856-dde934fc735b}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ILanguageFontGroupFactory
    ILanguageFontGroupFactory_CreateLanguageFontGroup = 6
    
    ' ILanguageFontGroup
    ILanguageFontGroup_GetUITextFont = 6
    ILanguageFontGroup_GetUIHeadingFont = 7
    ILanguageFontGroup_GetUITitleFont = 8
    ILanguageFontGroup_GetUICaptionFont = 9
    ILanguageFontGroup_GetUINotificationHeadingFont = 10
    ILanguageFontGroup_GetTraditionalDocumentFont = 11
    ILanguageFontGroup_GetModernDocumentFont = 12
    ILanguageFontGroup_GetDocumentHeadingFont = 13
    ILanguageFontGroup_GetFixedWidthTextFont = 14
    ILanguageFontGroup_GetDocumentAlternate1Font = 15
    ILanguageFontGroup_GetDocumentAlternate2Font = 16
                        
End Enum

' ----==== Variablen ====----
Private m_pILanguageFontGroup As Long
Private m_pILanguageFontGroupFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_Fonts_LanguageFontGroup, _
                            IID_ILanguageFontGroupFactory, _
                            m_pILanguageFontGroupFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILanguageFontGroup)
    Call ReleaseIfc(m_pILanguageFontGroupFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pILanguageFontGroup
End Property

Public Property Let Ifc(ByVal pILanguageFontGroup As Long)
    Call ReleaseIfc(m_pILanguageFontGroup)
    m_pILanguageFontGroup = pILanguageFontGroup
End Property

' ILanguageFontGroup
Public Property Get UITextFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetUITextFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set UITextFont = Ret
End Property

Public Property Get UIHeadingFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetUIHeadingFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set UIHeadingFont = Ret
End Property

Public Property Get UITitleFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetUITitleFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set UITitleFont = Ret
End Property

Public Property Get UICaptionFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetUICaptionFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set UICaptionFont = Ret
End Property

Public Property Get UINotificationHeadingFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetUINotificationHeadingFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set UINotificationHeadingFont = Ret
End Property

Public Property Get TraditionalDocumentFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetTraditionalDocumentFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set TraditionalDocumentFont = Ret
End Property

Public Property Get ModernDocumentFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetModernDocumentFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set ModernDocumentFont = Ret
End Property

Public Property Get DocumentHeadingFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetDocumentHeadingFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set DocumentHeadingFont = Ret
End Property

Public Property Get FixedWidthTextFont() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetFixedWidthTextFont, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set FixedWidthTextFont = Ret
End Property

Public Property Get DocumentAlternate1Font() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetDocumentAlternate1Font, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set DocumentAlternate1Font = Ret
End Property

Public Property Get DocumentAlternate2Font() As LanguageFont
    Dim Ret As LanguageFont
    If m_pILanguageFontGroup <> 0& Then
        Dim pILanguageFont As Long
        If Invoke(m_pILanguageFontGroup, _
                  ILanguageFontGroup_GetDocumentAlternate2Font, _
                  VarPtr(pILanguageFont)) = S_OK Then
            If pILanguageFont <> 0& Then
                Set Ret = New LanguageFont
                Ret.Ifc = pILanguageFont
            End If
        End If
    End If
    Set DocumentAlternate2Font = Ret
End Property

' ----==== Functions ====----
' ILanguageFontGroupFactory
Public Function CreateLanguageFontGroup(ByVal LanguageTag As String) As LanguageFontGroup
    Dim Ret As LanguageFontGroup
    If m_pILanguageFontGroupFactory <> 0& Then
        Dim pILanguageFontGroup As Long
        Dim hLanguageTag As Long
        hLanguageTag = CreateWindowsString(LanguageTag)
        If Invoke(m_pILanguageFontGroupFactory, _
                  ILanguageFontGroupFactory_CreateLanguageFontGroup, _
                  hLanguageTag, _
                  VarPtr(pILanguageFontGroup)) = S_OK Then
            If pILanguageFontGroup <> 0& Then
                Set Ret = New LanguageFontGroup
                Ret.Ifc = pILanguageFontGroup
            End If
        End If
        Call DeleteWindowsString(hLanguageTag)
    End If
    Set CreateLanguageFontGroup = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



