VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Streams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Streams

Option Explicit

' ----==== Enums Windows_Storage_Streams.xxx ====----
Public Enum InputStreamOptions
    InputStreamOptions_None = 0
    InputStreamOptions_Partial = 1
    InputStreamOptions_ReadAhead = 2
End Enum

' ----==== Namespace Streams.Buffer ====----
Public Property Get Buffer() As Buffer
    Set Buffer = New Buffer
End Property

' ----==== Namespace Streams.InputStream ====----
Public Property Get InputStream() As InputStream
    Set InputStream = New InputStream
End Property

' ----==== Namespace Streams.OutputStream ====----
Public Property Get OutputStream() As OutputStream
    Set OutputStream = New OutputStream
End Property

' ----==== Namespace Streams.RandomAccessStream ====----
Public Property Get RandomAccessStream() As RandomAccessStream
    Set RandomAccessStream = New RandomAccessStream
End Property

' ----==== Namespace Streams.RandomAccessStreamReference ====----
Public Property Get RandomAccessStreamReference() As RandomAccessStreamReference
    Set RandomAccessStreamReference = New RandomAccessStreamReference
End Property

' ----==== Namespace Streams.RandomAccessStreamWithContentType ====----
Public Property Get RandomAccessStreamWithContentType() As RandomAccessStreamWithContentType
    Set RandomAccessStreamWithContentType = New RandomAccessStreamWithContentType
End Property


