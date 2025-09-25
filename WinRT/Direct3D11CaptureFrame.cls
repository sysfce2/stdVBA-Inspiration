VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3D11CaptureFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Capture_Direct3D11CaptureFrame As String = "Windows.Graphics.Capture.Direct3D11CaptureFrame"
Private Const IID_IDirect3D11CaptureFrame As String = "{fa50c623-38da-4b32-acf3-fa9734ad800e}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDirect3D11CaptureFrame
    IDirect3D11CaptureFrame_GetSurface = 6
    IDirect3D11CaptureFrame_GetSystemRelativeTime = 7
    IDirect3D11CaptureFrame_GetContentSize = 8
    
End Enum

' ----==== Types ====----
Private Type SizeInt32_
    Width As Long
    Height As Long
End Type

' ----==== Variablen ====----
Private m_pIDirect3D11CaptureFrame As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIDirect3D11CaptureFrame) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDirect3D11CaptureFrame
End Property

Public Property Let Ifc(ByVal pIDirect3D11CaptureFrame As Long)
    Call DisposeIfc(m_pIDirect3D11CaptureFrame) ' <- IClosable
    m_pIDirect3D11CaptureFrame = pIDirect3D11CaptureFrame
End Property

' IDirect3D11CaptureFrame
Public Property Get Surface() As Direct3DSurface
    Dim Ret As Direct3DSurface
    If m_pIDirect3D11CaptureFrame <> 0& Then
        Dim pIDirect3DSurface As Long
        If Invoke(m_pIDirect3D11CaptureFrame, _
                  IDirect3D11CaptureFrame_GetSurface, _
                  VarPtr(pIDirect3DSurface)) = S_OK Then
            If pIDirect3DSurface <> 0& Then
                Set Ret = New Direct3DSurface
                Ret.Ifc = pIDirect3DSurface
            End If
        End If
    End If
    Set Surface = Ret
End Property

Public Property Get SystemRelativeTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIDirect3D11CaptureFrame <> 0& Then
        Dim value As Currency
        If Invoke(m_pIDirect3D11CaptureFrame, _
                  IDirect3D11CaptureFrame_GetSystemRelativeTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set SystemRelativeTime = Ret
End Property

Public Property Get ContentSize() As SizeInt32
    Dim Ret As SizeInt32
    If m_pIDirect3D11CaptureFrame <> 0& Then
        Dim value As SizeInt32_
        If Invoke(m_pIDirect3D11CaptureFrame, _
                  IDirect3D11CaptureFrame_GetContentSize, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New SizeInt32
            Ret.Width = value.Width
            Ret.Height = value.Height
        End If
    End If
    Set ContentSize = Ret
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

