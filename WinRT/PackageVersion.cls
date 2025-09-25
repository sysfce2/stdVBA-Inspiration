VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackageVersion"
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
Private m_Major As Integer
Private m_Minor As Integer
Private m_Build As Integer
Private m_Revision As Integer

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get Major() As Integer
    Major = m_Major
End Property

Public Property Let Major(ByVal value As Integer)
    m_Major = value
End Property

Public Property Get Minor() As Integer
    Minor = m_Minor
End Property

Public Property Let Minor(ByVal value As Integer)
    m_Minor = value
End Property

Public Property Get Build() As Integer
    Build = m_Build
End Property

Public Property Let Build(ByVal value As Integer)
    m_Build = value
End Property

Public Property Get Revision() As Integer
    Revision = m_Revision
End Property

Public Property Let Revision(ByVal value As Integer)
    m_Revision = value
End Property

Public Property Get ToString() As String
     ToString = "{Major=" & CStr(m_Major) & ",Minor=" & CStr(m_Minor) & ",Build=" & CStr(m_Build) & ",Revision=" & CStr(m_Revision) & "}"
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----


