VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ocr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

' Namespace Ocr

Option Explicit

' ----==== Enums Windows.Media.Ocr.xxx ====----

' ----==== Namespace Ocr.OcrLine ====----
Public Property Get OcrLine() As OcrLine
    Set OcrLine = New OcrLine
End Property

' ----==== Namespace Ocr.OcrWord ====----
Public Property Get OcrWord() As OcrWord
    Set OcrWord = New OcrWord
End Property

' ----==== Namespace Ocr.OcrResult ====----
Public Property Get OcrResult() As OcrResult
    Set OcrResult = New OcrResult
End Property

' ----==== Namespace Ocr.OcrEngine ====----
Public Property Get OcrEngine() As OcrEngine
    Set OcrEngine = New OcrEngine
End Property

