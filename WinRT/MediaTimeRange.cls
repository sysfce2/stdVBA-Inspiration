VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaTimeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

'Private Type MediaTimeRange
'    Start As TimeSpan
'    End As TimeSpan
'End Type

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_Start As TimeSpan
Private m_End As TimeSpan

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get StartTime() As TimeSpan
    Set StartTime = m_Start
End Property

Public Property Let StartTime(ByVal value As TimeSpan)
    Set m_Start = value
End Property

Public Property Get EndTime() As TimeSpan
    Set EndTime = m_End
End Property

Public Property Let EndTime(ByVal value As TimeSpan)
    Set m_End = value
End Property

Public Property Get ToString() As String
     ToString = "{Start=" & CStr(m_Start.VbDate) & ",End=" & CStr(m_End.VbDate) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----

