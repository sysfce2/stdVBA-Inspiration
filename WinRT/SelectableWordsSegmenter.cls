VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectableWordsSegmenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Text_SelectableWordsSegmenter As String = "Windows.Data.Text.SelectableWordsSegmenter"
Private Const IID_ISelectableWordsSegmenter As String = "{f6dc31e7-4b13-45c5-8897-7d71269e085d}"
Private Const IID_ISelectableWordsSegmenterFactory As String = "{8c7a7648-6057-4339-bc70-f210010a4150}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISelectableWordsSegmenterFactory
    ISelectableWordsSegmenterFactory_CreateWithLanguage = 6

    ' ISelectableWordsSegmenter
    ISelectableWordsSegmenter_GetResolvedLanguage = 6
    ISelectableWordsSegmenter_GetTokenAt = 7
    ISelectableWordsSegmenter_GetTokens = 8
'    ISelectableWordsSegmenter_Tokenize = 9
                        
End Enum

' ----==== Vars ====----
Private m_pISelectableWordsSegmenter As Long
Private m_pISelectableWordsSegmenterFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Text_SelectableWordsSegmenter, _
                            IID_ISelectableWordsSegmenterFactory, _
                            m_pISelectableWordsSegmenterFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISelectableWordsSegmenter)
    Call ReleaseIfc(m_pISelectableWordsSegmenterFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISelectableWordsSegmenter
End Property

Public Property Let Ifc(ByVal pISelectableWordsSegmenter As Long)
    Call ReleaseIfc(m_pISelectableWordsSegmenter)
    m_pISelectableWordsSegmenter = pISelectableWordsSegmenter
End Property

' ISelectableWordsSegmenter
Public Property Get ResolvedLanguage() As String
    Dim Ret As String
    If m_pISelectableWordsSegmenter <> 0& Then
        Dim value As Long
        If Invoke(m_pISelectableWordsSegmenter, _
                  ISelectableWordsSegmenter_GetResolvedLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ResolvedLanguage = Ret
End Property

' ISelectableWordsSegmenter
Public Property Get TokenAt(ByVal Text As String, _
                            ByVal startIndex As Long) As SelectableWordSegment
    Dim Ret As SelectableWordSegment
    If m_pISelectableWordsSegmenter <> 0& Then
        Dim Result As Long
        Dim hString As Long
        hString = CreateWindowsString(Text)
        If Invoke(m_pISelectableWordsSegmenter, _
                  ISelectableWordsSegmenter_GetTokenAt, _
                  hString, _
                  startIndex, _
                  VarPtr(Result)) = S_OK Then
            If Result <> 0& Then
                Set Ret = New SelectableWordSegment
                Ret.Ifc = Result
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set TokenAt = Ret
End Property

Public Property Get Tokens(ByVal Text As String) As ReadOnlyList_1 'ReadOnlyList_SelectableWordSegment
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_SelectableWordSegment
    If m_pISelectableWordsSegmenter <> 0& Then
        Dim pIVectorView As Long
        Dim hString As Long
        hString = CreateWindowsString(Text)
        If Invoke(m_pISelectableWordsSegmenter, _
                  ISelectableWordsSegmenter_GetTokens, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_SelectableWordSegment, pIVectorView) 'ReadOnlyList_SelectableWordSegment
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Tokens = Ret
End Property

' ----==== Functions ====----
' ISelectableWordsSegmenterFactory
Public Function CreateWithLanguage(ByVal Language As String) As SelectableWordsSegmenter
    Dim Ret As SelectableWordsSegmenter
    If m_pISelectableWordsSegmenterFactory <> 0& Then
        Dim Result As Long
        Dim hString As Long
        hString = CreateWindowsString(Language)
        If Invoke(m_pISelectableWordsSegmenterFactory, _
                  ISelectableWordsSegmenterFactory_CreateWithLanguage, _
                  hString, _
                  VarPtr(Result)) = S_OK Then
            If Result <> 0& Then
                Set Ret = New SelectableWordsSegmenter
                Ret.Ifc = Result
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateWithLanguage = Ret
End Function

' ISelectableWordsSegmenter
'    ISelectableWordsSegmenter_Tokenize = 9

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function




