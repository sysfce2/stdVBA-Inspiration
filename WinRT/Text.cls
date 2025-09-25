VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

' Namespace Text

Option Explicit

' ----==== Enum Windows.Data.Text.xxx ====----
Public Enum AlternateNormalizationFormat
    AlternateNormalizationFormat_NotNormalized = 0
    AlternateNormalizationFormat_Number = 1
    AlternateNormalizationFormat_Currency = 3
    AlternateNormalizationFormat_Date = 4
    AlternateNormalizationFormat_Time = 5
End Enum

Public Enum FontStretch
    FontStretch_Undefined = 0
    FontStretch_UltraCondensed = 1
    FontStretch_ExtraCondensed = 2
    FontStretch_Condensed = 3
    FontStretch_SemiCondensed = 4
    FontStretch_Normal = 5
    FontStretch_SemiExpanded = 6
    FontStretch_Expanded = 7
    FontStretch_ExtraExpanded = 8
    FontStretch_UltraExpanded = 9
End Enum

Public Enum FontStyle
    FontStyle_Normal = 0
    FontStyle_Oblique = 1
    FontStyle_Italic = 2
End Enum

' ----==== Namespace Text.WordsSegmenter ====----
Public Property Get WordsSegmenter() As WordsSegmenter
    Set WordsSegmenter = New WordsSegmenter
End Property

' ----==== Namespace Text.WordSegment ====----
Public Property Get WordSegment() As WordSegment
    Set WordSegment = New WordSegment
End Property

' ----==== Namespace Text.AlternateWordForm ====----
Public Property Get AlternateWordForm() As AlternateWordForm
    Set AlternateWordForm = New AlternateWordForm
End Property

' ----==== Namespace Text.SelectableWordsSegmenter ====----
Public Property Get SelectableWordsSegmenter() As SelectableWordsSegmenter
    Set SelectableWordsSegmenter = New SelectableWordsSegmenter
End Property

' ----==== Namespace Text.SelectableWordSegment ====----
Public Property Get SelectableWordSegment() As SelectableWordSegment
    Set SelectableWordSegment = New SelectableWordSegment
End Property

' ----==== Namespace Text.TextSegment ====----
Public Property Get TextSegment(Optional ByVal StartPosition = 0, _
                                Optional ByVal Length As Long = 0) As TextSegment
    Set TextSegment = New TextSegment
    TextSegment.StartPosition = StartPosition
    TextSegment.Length = Length
End Property

