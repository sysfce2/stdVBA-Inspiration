VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyValuePair_String_Inspectable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

'Public Type Windows.Foundation.Collections.KeyValuePair_HSTRING_IInspectable
'    key As Long 'String
'    value As Long 'Inspectable
'End Type

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_Key As String
Private m_Value As Inspectable

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Set m_Value = Nothing
End Sub

' ----==== Properties ====----
Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal value As String)
    m_Key = value
End Property

Public Property Get value() As Inspectable
    Set value = m_Value
End Property

Public Property Let value(ByVal value As Inspectable)
    Set m_Value = value
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----



