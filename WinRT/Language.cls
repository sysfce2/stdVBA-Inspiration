VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Language"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_Language As String = "Windows.Globalization.Language"
Private Const IID_ILanguage As String = "{ea79a752-f7c2-4265-b1bd-c4dec4e4f080}"
Private Const IID_ILanguage2 As String = "{6a47e5b5-d94d-4886-a404-a5a5b9d5b494}"
Private Const IID_ILanguage3 As String = "{c6af3d10-641a-5ba4-bb43-5e12aed75954}"
Private Const IID_ILanguageFactory As String = "{9b0252ac-0c27-44f8-b792-9793fb66c63e}"
Private Const IID_ILanguageStatics As String = "{b23cd557-0865-46d4-89b8-d59be8990f0d}"
Private Const IID_ILanguageStatics2 As String = "{30199f6e-914b-4b2a-9d6e-e3b0e27dbe4f}"
Private Const IID_ILanguageStatics3 As String = "{d15ecb5a-71de-5752-9542-fac5b4f27261}"
Private Const IID_ILanguageExtensionSubtags As String = "{7d7daf45-368d-4364-852b-dec927037b85}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ILanguageFactory
    ILanguageFactory_CreateLanguage = 6
    
    ' ILanguageStatics
    ILanguageStatics_IsWellFormed = 6
    ILanguageStatics_GetCurrentInputMethodLanguageTag = 7
    
    ' ILanguageStatics2
    ILanguageStatics2_TrySetInputMethodLanguageTag = 6
    
    ' ILanguageStatics3
    ILanguageStatics3_GetMuiCompatibleLanguageListFromLanguageTags = 6

    ' ILanguage
    ILanguage_GetLanguageTag = 6
    ILanguage_GetDisplayName = 7
    ILanguage_GetNativeName = 8
    ILanguage_GetScript = 9
    
    ' ILanguage2
    ILanguage2_GetLayoutDirection = 6
    
    ' ILanguage3
    ILanguage3_GetAbbreviatedName = 6
    
    ' ILanguageExtensionSubtags
    ILanguageExtensionSubtags_GetExtensionSubtags = 6
    
End Enum

' ----==== Variablen ====----
Private m_pILanguage As Long
Private m_pILanguage2 As Long
Private m_pILanguage3 As Long
Private m_pILanguageFactory As Long
Private m_pILanguageStatics As Long
Private m_pILanguageStatics2 As Long
Private m_pILanguageStatics3 As Long
Private m_pILanguageExtensionSubtags As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_Language, _
                            IID_ILanguageFactory, _
                            m_pILanguageFactory) Then
    End If
    If GetActivationFactory(Windows_Globalization_Language, _
                            IID_ILanguageStatics, _
                            m_pILanguageStatics) Then
        Call QueryIfc(m_pILanguageStatics, IID_ILanguageStatics2, m_pILanguageStatics2)
        Call QueryIfc(m_pILanguageStatics, IID_ILanguageStatics3, m_pILanguageStatics3)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILanguageExtensionSubtags)
    Call ReleaseIfc(m_pILanguage3)
    Call ReleaseIfc(m_pILanguage2)
    Call ReleaseIfc(m_pILanguage)
    Call ReleaseIfc(m_pILanguageStatics3)
    Call ReleaseIfc(m_pILanguageStatics2)
    Call ReleaseIfc(m_pILanguageStatics)
    Call ReleaseIfc(m_pILanguageFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pILanguage
End Property

Public Property Let Ifc(ByVal pILanguage As Long)
    Call ReleaseIfc(m_pILanguage3)
    Call ReleaseIfc(m_pILanguage2)
    Call ReleaseIfc(m_pILanguage)
    m_pILanguage = pILanguage
    Call QueryIfc(m_pILanguage, IID_ILanguage2, m_pILanguage2)
    Call QueryIfc(m_pILanguage, IID_ILanguage3, m_pILanguage3)
    Call QueryIfc(m_pILanguage, IID_ILanguageExtensionSubtags, m_pILanguageExtensionSubtags)
End Property

' ILanguageStatics
Public Property Get CurrentInputMethodLanguageTag() As String
    Dim Ret As String
    If m_pILanguageStatics <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguageStatics, _
                  ILanguageStatics_GetCurrentInputMethodLanguageTag, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CurrentInputMethodLanguageTag = Ret
End Property

' ILanguage
Public Property Get LanguageTag() As String
    Dim Ret As String
    If m_pILanguage <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguage, _
                  ILanguage_GetLanguageTag, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    LanguageTag = Ret
End Property

Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pILanguage <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguage, _
                  ILanguage_GetDisplayName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get NativeName() As String
    Dim Ret As String
    If m_pILanguage <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguage, _
                  ILanguage_GetNativeName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    NativeName = Ret
End Property

Public Property Get Script() As String
    Dim Ret As String
    If m_pILanguage <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguage, _
                  ILanguage_GetScript, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Script = Ret
End Property

' ILanguage2
Public Property Get LayoutDirection() As LanguageLayoutDirection
    Dim Ret As LanguageLayoutDirection
    If m_pILanguage2 <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguage2, _
                  ILanguage2_GetLayoutDirection, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LayoutDirection = Ret
End Property

' ILanguage3
Public Property Get AbbreviatedName() As String
    Dim Ret As String
    If m_pILanguage3 <> 0& Then
        Dim value As Long
        If Invoke(m_pILanguage3, _
                  ILanguage3_GetAbbreviatedName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AbbreviatedName = Ret
End Property

' ----==== Functions ====----
' ILanguageFactory
Public Function CreateLanguage(ByVal LanguageTag As String) As Language
    Dim Ret As Language
    If m_pILanguageFactory <> 0& Then
        Dim hString As Long
        Dim pILanguage As Long
        hString = CreateWindowsString(LanguageTag)
        If Invoke(m_pILanguageFactory, _
                  ILanguageFactory_CreateLanguage, _
                  hString, _
                  VarPtr(pILanguage)) = S_OK Then
            If pILanguage <> 0& Then
                Set Ret = New Language
                Ret.Ifc = pILanguage
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateLanguage = Ret
End Function

' ILanguageStatics
Public Function IsWellFormed(ByVal LanguageTag As String) As Boolean
    Dim Ret As Boolean
    If m_pILanguageStatics <> 0& Then
        Dim hString As Long
        Dim Result As Long
        hString = CreateWindowsString(LanguageTag)
        If Invoke(m_pILanguageStatics, _
                  ILanguageStatics_IsWellFormed, _
                  hString, _
                  VarPtr(Result)) = S_OK Then
            Ret = CBool(Result)
        End If
        Call DeleteWindowsString(hString)
    End If
    IsWellFormed = Ret
End Function

' ILanguageStatics2
Public Function TrySetInputMethodLanguageTag(ByVal LanguageTag As String) As Boolean
    Dim Ret As Boolean
    If m_pILanguageStatics2 <> 0& Then
        Dim hString As Long
        Dim Result As Long
        hString = CreateWindowsString(LanguageTag)
        If Invoke(m_pILanguageStatics2, _
                  ILanguageStatics2_TrySetInputMethodLanguageTag, _
                  hString, _
                  VarPtr(Result)) = S_OK Then
            Ret = CBool(Result)
        End If
        Call DeleteWindowsString(hString)
    End If
    TrySetInputMethodLanguageTag = Ret
End Function

' ILanguageStatics3
Public Function GetMuiCompatibleLanguageListFromLanguageTags(ByVal LanguageTags As List_String) As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pILanguageStatics3 <> 0& And _
       IsNotNothing(LanguageTags) Then
        Dim pIVectorView As Long
        If Invoke(m_pILanguageStatics3, _
                  ILanguageStatics3_GetMuiCompatibleLanguageListFromLanguageTags, _
                  LanguageTags.Ifc, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set GetMuiCompatibleLanguageListFromLanguageTags = Ret
End Function

' ILanguageExtensionSubtags
Public Function GetExtensionSubtags(ByVal singleton As String) As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pILanguageExtensionSubtags <> 0& Then
        Dim hString As Long
        Dim pIVectorView As Long
        hString = CreateWindowsString(singleton)
        If Invoke(m_pILanguageExtensionSubtags, _
                  ILanguageExtensionSubtags_GetExtensionSubtags, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetExtensionSubtags = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


