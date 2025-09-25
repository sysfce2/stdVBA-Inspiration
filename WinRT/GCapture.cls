VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

' Namespace Capture

Option Explicit

' ----==== Enums Windows.Graphics.Capture.xxx ====----
Public Enum GraphicsCaptureAccessKind
    GraphicsCaptureAccessKind_Borderless = 0
    GraphicsCaptureAccessKind_Programmatic = 1
End Enum

' ----==== Namespace Capture.GraphicsCaptureItem ====----
Public Property Get GraphicsCaptureItem() As GraphicsCaptureItem
    Set GraphicsCaptureItem = New GraphicsCaptureItem
End Property

' ----==== Namespace Capture.GraphicsCaptureSession ====----
Public Property Get GraphicsCaptureSession() As GraphicsCaptureSession
    Set GraphicsCaptureSession = New GraphicsCaptureSession
End Property

' ----==== Namespace Capture.GraphicsCapturePicker ====----
Public Property Get GraphicsCapturePicker() As GraphicsCapturePicker
    Set GraphicsCapturePicker = New GraphicsCapturePicker
End Property

' ----==== Namespace Capture.Direct3D11CaptureFrame ====----
Public Property Get Direct3D11CaptureFrame() As Direct3D11CaptureFrame
    Set Direct3D11CaptureFrame = New Direct3D11CaptureFrame
End Property

' ----==== Namespace Capture.Direct3D11CaptureFramePool ====----
Public Property Get Direct3D11CaptureFramePool() As Direct3D11CaptureFramePool
    Set Direct3D11CaptureFramePool = New Direct3D11CaptureFramePool
End Property

