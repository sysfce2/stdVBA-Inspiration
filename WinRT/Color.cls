VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----

' ----==== Types ====----
Private Type ARGB
    A As Byte
    R As Byte
    G As Byte
    B As Byte
End Type

Private Type ColorLong
    lColor As Long
End Type

' ----==== Vars ====----
Private m_ARGB As ARGB
Private m_ColorValue As ColorLong

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get ColorValue() As Long
    ColorValue = m_ColorValue.lColor
End Property

Public Property Let ColorValue(ByVal value As Long)
    m_ColorValue.lColor = value
    LSet m_ARGB = m_ColorValue
End Property

Public Property Get A() As Byte
    A = m_ARGB.A
End Property

Public Property Get R() As Byte
    R = m_ARGB.R
End Property

Public Property Get G() As Byte
    G = m_ARGB.G
End Property

Public Property Get B() As Byte
    B = m_ARGB.B
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
