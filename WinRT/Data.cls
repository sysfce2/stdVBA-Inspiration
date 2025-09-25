VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Data

Option Explicit

' ----==== Namespace Data.Pdf ====----
Public Property Get Pdf() As Pdf
    Set Pdf = New Pdf
End Property

' ----==== Namespace Data.Html ====----
Public Property Get Html() As Html
    Set Html = New Html
End Property

' ----==== Namespace Data.Json ====----
Public Property Get Json() As Json
    Set Json = New Json
End Property

' ----==== Namespace Data.Text ====----
Public Property Get Text() As Text
    Set Text = New Text
End Property

' ----==== Namespace Data.Xml ====----
Public Property Get Xml() As Xml
    Set Xml = New Xml
End Property

