VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_TimeSpanValue As Currency

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get TimeSpanValue() As Currency
    TimeSpanValue = m_TimeSpanValue
End Property

Public Property Let TimeSpanValue(ByVal value As Currency)
    m_TimeSpanValue = value
End Property

Public Property Get VbDate() As Date
    VbDate = TimeSpan2VBDate(m_TimeSpanValue)
End Property

Public Property Get FromTicks(ByVal ticks As Currency) As TimeSpan
    Dim Ret As New TimeSpan
    Ret.TimeSpanValue = ticks
    Set FromTicks = Ret
End Property

Public Property Get FromSeconds(ByVal seconds As Currency) As TimeSpan
    Dim Ret As New TimeSpan
    Ret.TimeSpanValue = seconds * 1000
    Set FromSeconds = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----


