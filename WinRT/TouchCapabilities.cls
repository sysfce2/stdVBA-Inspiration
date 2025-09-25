VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TouchCapabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Input_TouchCapabilities As String = "Windows.Devices.Input.TouchCapabilities"
Private Const IID_ITouchCapabilities As String = "{20dd55f9-13f1-46c8-9285-2c05fa3eda6f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ITouchCapabilities
    ITouchCapabilities_GetTouchPresent = 6
    ITouchCapabilities_GetContacts = 7
    
End Enum

' ----==== Variablen ====----
Private m_pITouchCapabilities As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pITouchCapabilities As Long
    If GetActivateInstance(Windows_Devices_Input_TouchCapabilities, _
                           IID_ITouchCapabilities, _
                           pITouchCapabilities) Then
        Me.Ifc = pITouchCapabilities
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pITouchCapabilities)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pITouchCapabilities
End Property

Public Property Let Ifc(ByVal pITouchCapabilities As Long)
    Call ReleaseIfc(m_pITouchCapabilities)
    m_pITouchCapabilities = pITouchCapabilities
End Property

' ITouchCapabilities
Public Property Get TouchPresent() As Long
    Dim Ret As Long
    If m_pITouchCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pITouchCapabilities, _
                  ITouchCapabilities_GetTouchPresent, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TouchPresent = Ret
End Property

Public Property Get Contacts() As Long
    Dim Ret As Long
    If m_pITouchCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pITouchCapabilities, _
                  ITouchCapabilities_GetContacts, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Contacts = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


