VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UICommandInvokedHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Variablen ====----
Private m_pIUICommandInvokedHandler As Long

' ----==== Events ====----
Public Event CommandInvoked(ByVal command As UICommand)

' ----==== Class ====----
Private Sub Class_Initialize()
    m_pIUICommandInvokedHandler = UICIH.Create(Me)
End Sub

Private Sub Class_Terminate()
'
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUICommandInvokedHandler
End Property

Public Property Let Ifc(ByVal pIUICommandInvokedHandler As Long)
    m_pIUICommandInvokedHandler = pIUICommandInvokedHandler
End Property

' ----==== Invoke Handler ====----
Public Sub UICommand_Invoke(ByVal command As Long)
    Dim UICommand As UICommand
    If command <> 0& Then
        Set UICommand = New UICommand
        UICommand.Ifc = command
        RaiseEvent CommandInvoked(UICommand)
    End If
End Sub
