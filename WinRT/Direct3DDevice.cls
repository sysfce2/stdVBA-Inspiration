VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_DirectX_Direct3D11_Direct3DDevice As String = "Windows.Graphics.DirectX.Direct3D11.Direct3DDevice"
Private Const IID_IDirect3DDevice As String = "{a37624ab-8d5f-4650-9d3e-9eae3d9bc670}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDirect3DDevice
    IDirect3DDevice_Trim = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIDirect3DDevice As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIDirect3DDevice) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDirect3DDevice
End Property

Public Property Let Ifc(ByVal pIDirect3DDevice As Long)
    Call DisposeIfc(m_pIDirect3DDevice) ' <- IClosable
    m_pIDirect3DDevice = pIDirect3DDevice
End Property

' ----==== Function ====----
' IDirect3DDevice
Public Function Trim() As Boolean
    Dim Ret As Boolean
    If m_pIDirect3DDevice <> 0& Then
        If Invoke(m_pIDirect3DDevice, _
                  IDirect3DDevice_Trim) = S_OK Then
            Ret = True
        End If
    End If
    Trim = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


