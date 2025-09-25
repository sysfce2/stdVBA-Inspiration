VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

' Namespace Graphics

Option Explicit

' ----==== Namespace Graphics.PointInt32 ====----
Public Property Get PointInt32(Optional ByVal X As Long = 0, _
                               Optional ByVal Y As Long = 0) As PointInt32
    Set PointInt32 = New PointInt32
    PointInt32.X = X
    PointInt32.Y = Y
End Property

' ----==== Namespace Graphics.SizeInt32 ====----
Public Property Get SizeInt32(Optional ByVal Width As Long = 0, _
                              Optional ByVal Height As Long = 0) As SizeInt32
    Set SizeInt32 = New SizeInt32
    SizeInt32.Width = Width
    SizeInt32.Height = Height
End Property

' ----==== Namespace Graphics.RectInt32 ====----
Public Property Get RectInt32(Optional ByVal X As Long = 0, _
                              Optional ByVal Y As Long = 0, _
                              Optional ByVal Width As Long = 0, _
                              Optional ByVal Height As Long = 0) As RectInt32
    Set RectInt32 = New RectInt32
    RectInt32.X = X
    RectInt32.Y = Y
    RectInt32.Width = Width
    RectInt32.Height = Height
End Property

' ----==== Namespace Graphics.Imaging ====----
Public Property Get Imaging() As Imaging
    Set Imaging = New Imaging
End Property

' ----==== Namespace Graphics.Capture ====----
Public Property Get Capture() As GCapture
    Set Capture = New GCapture
End Property

' ----==== Namespace Graphics.DirectX ====----
Public Property Get DirectX() As DirectX
    Set DirectX = New DirectX
End Property

