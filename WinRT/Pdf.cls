VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Pdf

Option Explicit

' ----==== Enums Windows.Data.Pdf.xxx ====----
Public Enum PdfPageRotation
    PdfPageRotation_Normal = 0
    PdfPageRotation_Rotate90 = 1
    PdfPageRotation_Rotate180 = 2
    PdfPageRotation_Rotate270 = 3
End Enum

' ----==== Namespace Pdf.PdfPage ====----
Public Property Get PdfPage() As PdfPage
    Set PdfPage = New PdfPage
End Property

' ----==== Namespace Pdf.PdfDocument ====----
Public Property Get PdfDocument() As PdfDocument
    Set PdfDocument = New PdfDocument
End Property

' ----==== Namespace Pdf.PdfPageDimensions ====----
Public Property Get PdfPageDimensions() As PdfPageDimensions
    Set PdfPageDimensions = New PdfPageDimensions
End Property

' ----==== Namespace Pdf.PdfPageRenderOptions ====----
Public Property Get PdfPageRenderOptions() As PdfPageRenderOptions
    Set PdfPageRenderOptions = New PdfPageRenderOptions
End Property

