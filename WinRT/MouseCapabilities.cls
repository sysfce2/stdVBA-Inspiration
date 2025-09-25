VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseCapabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Devices_Input_MouseCapabilities As String = "Windows.Devices.Input.MouseCapabilities"
Private Const IID_IMouseCapabilities As String = "{bca5e023-7dd9-4b6b-9a92-55d43cb38f73}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMouseCapabilities
    IMouseCapabilities_GetMousePresent = 6
    IMouseCapabilities_GetVerticalWheelPresent = 7
    IMouseCapabilities_GetHorizontalWheelPresent = 8
    IMouseCapabilities_GetSwapButtons = 9
    IMouseCapabilities_GetNumberOfButtons = 10
    
End Enum

' ----==== Variablen ====----
Private m_pIMouseCapabilities As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIMouseCapabilities As Long
    If GetActivateInstance(Windows_Devices_Input_MouseCapabilities, _
                           IID_IMouseCapabilities, _
                           pIMouseCapabilities) Then
        Me.Ifc = pIMouseCapabilities
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMouseCapabilities)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMouseCapabilities
End Property

Public Property Let Ifc(ByVal pIMouseCapabilities As Long)
    Call ReleaseIfc(m_pIMouseCapabilities)
    m_pIMouseCapabilities = pIMouseCapabilities
End Property

' IMouseCapabilities
Public Property Get MousePresent() As Long
    Dim Ret As Long
    If m_pIMouseCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pIMouseCapabilities, _
                  IMouseCapabilities_GetMousePresent, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    MousePresent = Ret
End Property

Public Property Get VerticalWheelPresent() As Long
    Dim Ret As Long
    If m_pIMouseCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pIMouseCapabilities, _
                  IMouseCapabilities_GetVerticalWheelPresent, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    VerticalWheelPresent = Ret
End Property

Public Property Get HorizontalWheelPresent() As Long
    Dim Ret As Long
    If m_pIMouseCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pIMouseCapabilities, _
                  IMouseCapabilities_GetHorizontalWheelPresent, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    HorizontalWheelPresent = Ret
End Property

Public Property Get SwapButtons() As Long
    Dim Ret As Long
    If m_pIMouseCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pIMouseCapabilities, _
                  IMouseCapabilities_GetSwapButtons, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SwapButtons = Ret
End Property

Public Property Get NumberOfButtons() As Long
    Dim Ret As Long
    If m_pIMouseCapabilities <> 0& Then
        Dim value As Long
        If Invoke(m_pIMouseCapabilities, _
                  IMouseCapabilities_GetNumberOfButtons, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfButtons = Ret
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


