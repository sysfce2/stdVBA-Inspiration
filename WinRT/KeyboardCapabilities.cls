VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyboardCapabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Input_KeyboardCapabilities As String = "Windows.Devices.Input.KeyboardCapabilities"
Private Const IID_IKeyboardCapabilities As String = "{3a3f9b56-6798-4bbc-833e-0f34b17c65ff}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IKeyboardCapabilities
    IKeyboardCapabilities_GetKeyboardPresent = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIKeyboardCapabilities As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIKeyboardCapabilities As Long
    If GetActivateInstance(Windows_Devices_Input_KeyboardCapabilities, _
                           IID_IKeyboardCapabilities, _
                           pIKeyboardCapabilities) Then
        Me.Ifc = pIKeyboardCapabilities
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIKeyboardCapabilities)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIKeyboardCapabilities
End Property

Public Property Let Ifc(ByVal pIKeyboardCapabilities As Long)
    Call ReleaseIfc(m_pIKeyboardCapabilities)
    m_pIKeyboardCapabilities = pIKeyboardCapabilities
End Property

' IKeyboardCapabilities
Public Property Get KeyboardPresent() As Long
    Dim Ret As Long
    If m_pIKeyboardCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pIKeyboardCapabilities, _
                  IKeyboardCapabilities_GetKeyboardPresent, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    KeyboardPresent = Ret
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

