VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapPlaneDescription"
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
Private m_StartIndex As Long
Private m_Stride As Long
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
Public Property Get StartIndex() As Long
    StartIndex = m_StartIndex
End Property

Public Property Let StartIndex(ByVal value As Long)
    m_StartIndex = value
End Property

Public Property Get Stride() As Long
    Stride = m_Stride
End Property

Public Property Let Stride(ByVal value As Long)
    m_Stride = value
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
     ToString = "{StartIndex=" & CStr(m_StartIndex) & ",Width=" & CStr(m_Width) & ",Height=" & CStr(m_Height) & ",Stride=" & CStr(m_Stride) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----


