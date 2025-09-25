VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_StartPosition As Long
Private m_Length As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get StartPosition() As Long
    StartPosition = m_StartPosition
End Property

Public Property Let StartPosition(ByVal value As Long)
    m_StartPosition = value
End Property

Public Property Get Length() As Long
    Length = m_Length
End Property

Public Property Let Length(ByVal value As Long)
    m_Length = value
End Property

Public Property Get ToString() As String
     ToString = "{StartPosition=" & CStr(m_StartPosition) & ",Length=" & CStr(m_Length) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----

