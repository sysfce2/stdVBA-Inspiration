VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DMultisampleDescription"
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
Private m_Count As Long
Private m_Quality As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Let Count(ByVal value As Long)
    m_Count = value
End Property

Public Property Get Quality() As Long
    Quality = m_Quality
End Property

Public Property Let Quality(ByVal value As Long)
    m_Quality = value
End Property

Public Property Get ToString() As String
     ToString = "{Count=" & CStr(m_Count) & ",Quality=" & CStr(m_Quality) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----


