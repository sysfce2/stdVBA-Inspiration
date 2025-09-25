VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3DSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_DirectX_Direct3D11_Direct3DSurface As String = "Windows.Graphics.DirectX.Direct3D11.Direct3DSurface"
Private Const IID_IDirect3DSurface As String = "{0bf4a146-13c1-4694-bee3-7abf15eaf586}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDirect3DSurface
    IDirect3DSurface_GetDescription = 6
    
End Enum

' ----==== Types ====----
Private Type Direct3DSurfaceDescription_
    Width As Long
    Height As Long
    Format As DirectXPixelFormat
    MultisampleDescriptionCount As Long
    MultisampleDescriptionQuality As Long
End Type

Private Type Direct3DMultisampleDescription_
    count As Long
    Quality As Long
End Type

' ----==== Variablen ====----
Private m_pIDirect3DSurface As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIDirect3DSurface) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDirect3DSurface
End Property

Public Property Let Ifc(ByVal pIDirect3DSurface As Long)
    Call DisposeIfc(m_pIDirect3DSurface) ' <- IClosable
    m_pIDirect3DSurface = pIDirect3DSurface
End Property

' IDirect3DSurface
Public Property Get Description() As Direct3DSurfaceDescription
    Dim Ret As Direct3DSurfaceDescription
    If m_pIDirect3DSurface <> 0& Then
        Dim value As Direct3DSurfaceDescription_
        If Invoke(m_pIDirect3DSurface, _
                  IDirect3DSurface_GetDescription, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New Direct3DSurfaceDescription
            Ret.Height = value.Height
            Ret.Width = value.Width
            Ret.Format = value.Format
            Dim value2 As Direct3DMultisampleDescription
            Set value2 = New Direct3DMultisampleDescription
            value2.count = value.MultisampleDescriptionCount
            value2.Quality = value.MultisampleDescriptionQuality
            Ret.MultisampleDescription = value2
        End If
    End If
    Set Description = Ret
End Property

' ----==== Function ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


