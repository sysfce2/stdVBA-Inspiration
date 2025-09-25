VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WordsSegmenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Text_WordsSegmenter As String = "Windows.Data.Text.WordsSegmenter"
Private Const IID_IWordsSegmenter As String = "{86b4d4d1-b2fe-4e34-a81d-66640300454f}"
Private Const IID_IWordsSegmenterFactory As String = "{e6977274-fc35-455c-8bfb-6d7f4653ca97}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IWordsSegmenterFactory
    IWordsSegmenterFactory_CreateWithLanguage = 6

    ' IWordsSegmenter
    IWordsSegmenter_GetResolvedLanguage = 6
    IWordsSegmenter_GetTokenAt = 7
    IWordsSegmenter_GetTokens = 8
'    IWordsSegmenter_Tokenize = 9
                        
End Enum

' ----==== Vars ====----
Private m_pIWordsSegmenter As Long
Private m_pIWordsSegmenterFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Text_WordsSegmenter, _
                            IID_IWordsSegmenterFactory, _
                            m_pIWordsSegmenterFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIWordsSegmenter)
    Call ReleaseIfc(m_pIWordsSegmenterFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIWordsSegmenter
End Property

Public Property Let Ifc(ByVal pIWordsSegmenter As Long)
    Call ReleaseIfc(m_pIWordsSegmenter)
    m_pIWordsSegmenter = pIWordsSegmenter
End Property

' IWordsSegmenter
Public Property Get ResolvedLanguage() As String
    Dim Ret As String
    If m_pIWordsSegmenter <> 0& Then
        Dim value As Long
        If Invoke(m_pIWordsSegmenter, _
                  IWordsSegmenter_GetResolvedLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ResolvedLanguage = Ret
End Property

' IWordsSegmenter
Public Property Get TokenAt(ByVal Text As String, _
                            ByVal startIndex As Long) As WordSegment
    Dim Ret As WordSegment
    If m_pIWordsSegmenter <> 0& Then
        Dim Result As Long
        Dim hString As Long
        hString = CreateWindowsString(Text)
        If Invoke(m_pIWordsSegmenter, _
                  IWordsSegmenter_GetTokenAt, _
                  hString, _
                  startIndex, _
                  VarPtr(Result)) = S_OK Then
            If Result <> 0& Then
                Set Ret = New WordSegment
                Ret.Ifc = Result
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set TokenAt = Ret
End Property

Public Property Get Tokens(ByVal Text As String) As ReadOnlyList_1 'ReadOnlyList_WordSegment
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_WordSegment
    If m_pIWordsSegmenter <> 0& Then
        Dim pIVectorView As Long
        Dim hString As Long
        hString = CreateWindowsString(Text)
        If Invoke(m_pIWordsSegmenter, _
                  IWordsSegmenter_GetTokens, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_WordSegment, pIVectorView) 'ReadOnlyList_WordSegment
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Tokens = Ret
End Property

' ----==== Functions ====----
' IWordsSegmenterFactory
Public Function CreateWithLanguage(ByVal Language As String) As WordsSegmenter
    Dim Ret As WordsSegmenter
    If m_pIWordsSegmenterFactory <> 0& Then
        Dim Result As Long
        Dim hString As Long
        hString = CreateWindowsString(Language)
        If Invoke(m_pIWordsSegmenterFactory, _
                  IWordsSegmenterFactory_CreateWithLanguage, _
                  hString, _
                  VarPtr(Result)) = S_OK Then
            If Result <> 0& Then
                Set Ret = New WordsSegmenter
                Ret.Ifc = Result
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateWithLanguage = Ret
End Function

' IWordsSegmenter
'    IWordsSegmenter_Tokenize = 9

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


