VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UICommandSeparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Popups_UICommandSeparator As String = "Windows.UI.Popups.UICommandSeparator"
Private Const IID_IUICommand As String = "{4ff93a75-4145-47ff-ac7f-dff1c1fa5b0f}"

' ----==== Variablen ====----
Private m_pIUICommand As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_UI_Popups_UICommandSeparator, _
                           IID_IUICommand, _
                           m_pIUICommand) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUICommand)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUICommand
End Property
