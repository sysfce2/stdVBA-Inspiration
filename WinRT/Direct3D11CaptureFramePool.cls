VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3D11CaptureFramePool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Capture_Direct3D11CaptureFramePool As String = "Windows.Graphics.Capture.Direct3D11CaptureFramePool"
Private Const IID_IDirect3D11CaptureFramePool As String = "{24eb6d22-1975-422e-82e7-780dbd8ddf24}"
Private Const IID_IDirect3D11CaptureFramePoolStatics As String = "{7784056a-67aa-4d53-ae54-1088d5a8ca21}"
Private Const IID_IDirect3D11CaptureFramePoolStatics2 As String = "{589b103f-6bbc-5df5-a991-02e28b3b66d5}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDirect3D11CaptureFramePoolStatics
    IDirect3D11CaptureFramePoolStatics_Create = 6
    
    ' IDirect3D11CaptureFramePoolStatics
    IDirect3D11CaptureFramePoolStatics2_CreateFreeThreaded = 6
    
    ' IDirect3D11CaptureFramePool
    IDirect3D11CaptureFramePool_Recreate = 6
    IDirect3D11CaptureFramePool_TryGetNextFrame = 7
    IDirect3D11CaptureFramePool_AddFrameArrived = 8
    IDirect3D11CaptureFramePool_RemoveFrameArrived = 9
    IDirect3D11CaptureFramePool_CreateCaptureSession = 10
'    IDirect3D11CaptureFramePool_GetDispatcherQueue = 11

End Enum

' ----==== Types ====----
Private Type SizeInt32_
    Width As Long
    Height As Long
End Type

' ----==== Variablen ====----
Private m_pIDirect3D11CaptureFramePool As Long
Private m_pIDirect3D11CaptureFramePoolStatics As Long
Private m_pIDirect3D11CaptureFramePoolStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Graphics_Capture_Direct3D11CaptureFramePool, _
                            IID_IDirect3D11CaptureFramePoolStatics, _
                            m_pIDirect3D11CaptureFramePoolStatics) Then
         Call QueryIfc(m_pIDirect3D11CaptureFramePoolStatics, _
                       IID_IDirect3D11CaptureFramePoolStatics2, _
                       m_pIDirect3D11CaptureFramePoolStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDirect3D11CaptureFramePoolStatics2)
    Call ReleaseIfc(m_pIDirect3D11CaptureFramePoolStatics)
    Call DisposeIfc(m_pIDirect3D11CaptureFramePool) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDirect3D11CaptureFramePool
End Property

Public Property Let Ifc(ByVal pIDirect3D11CaptureFramePool As Long)
    Call DisposeIfc(m_pIDirect3D11CaptureFramePool) ' <- IClosable
    m_pIDirect3D11CaptureFramePool = pIDirect3D11CaptureFramePool
End Property

' IDirect3D11CaptureFramePool
'    IDirect3D11CaptureFramePool_GetDispatcherQueue = 11

' ----==== Function ====----
' IDirect3D11CaptureFramePoolStatics
Public Function Create(ByVal Device As Direct3DDevice, _
                       ByVal PixelFormat As DirectXPixelFormat, _
                       ByVal NumberOfBuffers As Long, _
                       ByVal Size As SizeInt32) As Direct3D11CaptureFramePool
    Dim Ret As Direct3D11CaptureFramePool
    If m_pIDirect3D11CaptureFramePoolStatics <> 0& And _
       IsNotNothing(Device) Then
        Dim pIDirect3D11CaptureFramePool As Long
        If Invoke(m_pIDirect3D11CaptureFramePoolStatics, _
                  IDirect3D11CaptureFramePoolStatics_Create, _
                  Device.Ifc, _
                  PixelFormat, _
                  NumberOfBuffers, _
                  Size.Width, _
                  Size.Height, _
                  VarPtr(pIDirect3D11CaptureFramePool)) = S_OK Then
            If pIDirect3D11CaptureFramePool <> 0& Then
                Set Ret = New Direct3D11CaptureFramePool
                Ret.Ifc = pIDirect3D11CaptureFramePool
            End If
        End If
    End If
    Set Create = Ret
End Function
    
' IDirect3D11CaptureFramePoolStatics
Public Function CreateFreeThreaded(ByVal Device As Direct3DDevice, _
                                   ByVal PixelFormat As DirectXPixelFormat, _
                                   ByVal NumberOfBuffers As Long, _
                                   ByVal Size As SizeInt32) As Direct3D11CaptureFramePool
    Dim Ret As Direct3D11CaptureFramePool
    If m_pIDirect3D11CaptureFramePoolStatics2 <> 0& And _
       IsNotNothing(Device) Then
        Dim pIDirect3D11CaptureFramePool As Long
        If Invoke(m_pIDirect3D11CaptureFramePoolStatics2, _
                  IDirect3D11CaptureFramePoolStatics2_CreateFreeThreaded, _
                  Device.Ifc, _
                  PixelFormat, _
                  NumberOfBuffers, _
                  Size.Width, _
                  Size.Height, _
                  VarPtr(pIDirect3D11CaptureFramePool)) = S_OK Then
            If pIDirect3D11CaptureFramePool <> 0& Then
                Set Ret = New Direct3D11CaptureFramePool
                Ret.Ifc = pIDirect3D11CaptureFramePool
            End If
        End If
    End If
    Set CreateFreeThreaded = Ret
End Function

' IDirect3D11CaptureFramePool
Public Function Recreate(ByVal Device As Direct3DDevice, _
                         ByVal PixelFormat As DirectXPixelFormat, _
                         ByVal NumberOfBuffers As Long, _
                         ByVal Size As SizeInt32) As Boolean
    Dim Ret As Boolean
    If m_pIDirect3D11CaptureFramePool <> 0& And _
       IsNotNothing(Device) Then
        If Invoke(m_pIDirect3D11CaptureFramePool, _
                  IDirect3D11CaptureFramePool_Recreate, _
                  Device.Ifc, _
                  PixelFormat, _
                  NumberOfBuffers, _
                  Size.Width, _
                  Size.Height) = S_OK Then
            Ret = True
        End If
    End If
    Recreate = Ret
End Function

Public Function TryGetNextFrame() As Direct3D11CaptureFrame
    Dim Ret As Direct3D11CaptureFrame
    If m_pIDirect3D11CaptureFramePool <> 0& Then
        Dim pIDirect3D11CaptureFrame As Long
        If Invoke(m_pIDirect3D11CaptureFramePool, _
                  IDirect3D11CaptureFramePool_TryGetNextFrame, _
                  VarPtr(pIDirect3D11CaptureFrame)) = S_OK Then
            If pIDirect3D11CaptureFrame <> 0& Then
                Set Ret = New Direct3D11CaptureFrame
                Ret.Ifc = pIDirect3D11CaptureFrame
            End If
        End If
    End If
    Set TryGetNextFrame = Ret
End Function

Public Function AddFrameArrived(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIDirect3D11CaptureFramePool <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pIDirect3D11CaptureFramePool, _
                  IDirect3D11CaptureFramePool_AddFrameArrived, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddFrameArrived = Ret
End Function

Public Function RemoveFrameArrived(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIDirect3D11CaptureFramePool <> 0& Then
        If Invoke(m_pIDirect3D11CaptureFramePool, _
                  IDirect3D11CaptureFramePool_RemoveFrameArrived, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveFrameArrived = Ret
End Function

Public Function CreateCaptureSession() As GraphicsCaptureSession
    Dim Ret As GraphicsCaptureSession
    If m_pIDirect3D11CaptureFramePool <> 0& Then
        Dim pIGraphicsCaptureSession As Long
        If Invoke(m_pIDirect3D11CaptureFramePool, _
                  IDirect3D11CaptureFramePool_CreateCaptureSession, _
                  VarPtr(pIGraphicsCaptureSession)) = S_OK Then
            If pIGraphicsCaptureSession <> 0& Then
                Set Ret = New GraphicsCaptureSession
                Ret.Ifc = pIGraphicsCaptureSession
            End If
        End If
    End If
    Set CreateCaptureSession = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

