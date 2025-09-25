VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTime"
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
Private m_DateTimeValue As Currency

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get DateTimeValue() As Currency
    DateTimeValue = m_DateTimeValue
End Property

Public Property Let DateTimeValue(ByVal value As Currency)
    m_DateTimeValue = value
End Property

Public Property Get VbDate() As Date
    VbDate = DateTime2VBDate(m_DateTimeValue)
End Property

Public Property Let VbDate(ByVal value As Date)
    m_DateTimeValue = VBDate2DateTime(value)
End Property

Public Property Get Now() As Datetime
    Dim Datetime As New Datetime
    Datetime.DateTimeValue = VBDate2DateTime(VBA.Now)
    Set Now = Datetime
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----

