VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cryptography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Cryptography

Option Explicit

' ----==== Enums Windows.Security.Cryptography.xxx ====----
Public Enum BinaryStringEncoding
    BinaryStringEncoding_Utf8 = 0
    BinaryStringEncoding_Utf16LE = 1
    BinaryStringEncoding_Utf16BE = 2
End Enum

' ----==== Namespace Cryptography.CryptographicBuffer ====----
Public Property Get CryptographicBuffer() As CryptographicBuffer
    Set CryptographicBuffer = New CryptographicBuffer
End Property

