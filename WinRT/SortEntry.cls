VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

'Public Type Windows_Storage_Search_SortEntry
'    PropertyName As String
'    AscendingOrder As Boolean
'End Type

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_PropertyName As String
Private m_AscendingOrder As Boolean

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get PropertyName() As String
    PropertyName = m_PropertyName
End Property

Public Property Let PropertyName(ByVal value As String)
    m_PropertyName = value
End Property

Public Property Get AscendingOrder() As Boolean
    AscendingOrder = m_AscendingOrder
End Property

Public Property Let AscendingOrder(ByVal value As Boolean)
    m_AscendingOrder = value
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----

