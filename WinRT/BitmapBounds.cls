VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapBounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Vars ====----
Private m_X As Long
Private m_Y As Long
Private m_Width As Long
Private m_Height As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get X() As Long
    X = m_X
End Property

Public Property Let X(ByVal value As Long)
    m_X = value
End Property

Public Property Get Y() As Long
    Y = m_Y
End Property

Public Property Let Y(ByVal value As Long)
    m_Y = value
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Width(ByVal value As Long)
    m_Width = value
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Let Height(ByVal value As Long)
    m_Height = value
End Property

Public Property Get ToString() As String
     ToString = "{X=" & CStr(m_X) & ",Y=" & CStr(m_Y) & ",Width=" & CStr(m_Width) & ",Height=" & CStr(m_Height) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----


