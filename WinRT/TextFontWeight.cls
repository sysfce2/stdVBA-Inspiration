VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextFontWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_Weight As Integer

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get Weight() As Integer
    Weight = m_Weight
End Property

Public Property Let Weight(ByVal value As Integer)
    m_Weight = value
End Property

Public Property Get ToString() As String
     ToString = "{Weight=" & CStr(m_Weight) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----


