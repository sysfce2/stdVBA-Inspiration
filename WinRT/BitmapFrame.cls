VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_BitmapFrame As String = "Windows.Graphics.Imaging.BitmapFrame"
Private Const IID_IBitmapFrame As String = "{72a49a1c-8081-438d-91bc-94ecfc8185c6}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IBitmapFrame
    IBitmapFrame_GetThumbnailAsync = 6
    IBitmapFrame_GetBitmapProperties = 7
    IBitmapFrame_GetBitmapPixelFormat = 8
    IBitmapFrame_GetBitmapAlphaMode = 9
    IBitmapFrame_GetDpiX = 10
    IBitmapFrame_GetDpiY = 11
    IBitmapFrame_GetPixelWidth = 12
    IBitmapFrame_GetPixelHeight = 13
    IBitmapFrame_GetOrientedPixelWidth = 14
    IBitmapFrame_GetOrientedPixelHeight = 15
    IBitmapFrame_GetPixelDataAsync = 16
    IBitmapFrame_GetPixelDataTransformedAsync = 17
                        
End Enum

' ----==== Variablen ====----
Private m_pIBitmapFrame As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBitmapFrame)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBitmapFrame
End Property

Public Property Let Ifc(ByVal pIBitmapFrame As Long)
    Call ReleaseIfc(m_pIBitmapFrame)
    m_pIBitmapFrame = pIBitmapFrame
End Property

' IBitmapFrame
Public Property Get BitmapProperties() As BitmapPropertiesView
    Dim Ret As BitmapPropertiesView
    If m_pIBitmapFrame <> 0& Then
        Dim pBitmapPropertiesView As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetBitmapProperties, _
                  VarPtr(pBitmapPropertiesView)) = S_OK Then
            If pBitmapPropertiesView <> 0& Then
                Set Ret = New BitmapPropertiesView
                Ret.Ifc = pBitmapPropertiesView
            End If
        End If
    End If
    Set BitmapProperties = Ret
End Property

Public Property Get BitmapPixelFormat() As BitmapPixelFormat
    Dim Ret As BitmapPixelFormat
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetBitmapPixelFormat, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    BitmapPixelFormat = Ret
End Property

Public Property Get BitmapAlphaMode() As BitmapAlphaMode
    Dim Ret As BitmapAlphaMode
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetBitmapAlphaMode, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    BitmapAlphaMode = Ret
End Property

Public Property Get dpiX() As Double
    Dim Ret As Double
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Double
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetDpiX, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    dpiX = Ret
End Property

Public Property Get dpiY() As Double
    Dim Ret As Double
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Double
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetDpiY, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    dpiY = Ret
End Property

Public Property Get PixelWidth() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelWidth, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    PixelWidth = Ret
End Property

Public Property Get PixelHeight() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelHeight, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    PixelHeight = Ret
End Property

Public Property Get OrientedPixelWidth() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetOrientedPixelWidth, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    OrientedPixelWidth = Ret
End Property

Public Property Get OrientedPixelHeight() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetOrientedPixelHeight, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    OrientedPixelHeight = Ret
End Property

' ----==== Functions ====----
' IBitmapFrame
Public Function GetThumbnailAsync() As ImageStream
    Dim Ret As ImageStream
    If m_pIBitmapFrame <> 0& Then
        Dim pIImageStream As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetThumbnailAsync, _
                  VarPtr(pIImageStream)) = S_OK Then
            If pIImageStream <> 0& Then
                If Await(pIImageStream) Then
                    Set Ret = New ImageStream
                    Ret.Ifc = pIImageStream
                End If
            End If
        End If
    End If
    Set GetThumbnailAsync = Ret
End Function

Public Function GetPixelDataAsync() As PixelDataProvider
    Dim Ret As PixelDataProvider
    If m_pIBitmapFrame <> 0& Then
        Dim pIPixelDataProvider As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelDataAsync, _
                  VarPtr(pIPixelDataProvider)) = S_OK Then
            If pIPixelDataProvider <> 0& Then
                If Await(pIPixelDataProvider) Then
                    Set Ret = New PixelDataProvider
                    Ret.Ifc = pIPixelDataProvider
                End If
            End If
        End If
    End If
    Set GetPixelDataAsync = Ret
End Function

Public Function GetPixelDataTransformedAsync(ByVal pixelFormat As BitmapPixelFormat, _
                                             ByVal alphaMode As BitmapAlphaMode, _
                                             ByVal transform As BitmapTransform, _
                                             ByVal exOrientationMode As ExifOrientationMode, _
                                             ByVal colManagementMode As ColorManagementMode) As PixelDataProvider
    Dim Ret As PixelDataProvider
    If m_pIBitmapFrame <> 0& And _
       IsNotNothing(transform) Then
        Dim pIPixelDataProvider As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelDataTransformedAsync, _
                  pixelFormat, _
                  alphaMode, _
                  transform.Ifc, _
                  exOrientationMode, _
                  colManagementMode, _
                  VarPtr(pIPixelDataProvider)) = S_OK Then
            If pIPixelDataProvider <> 0& Then
                If Await(pIPixelDataProvider) Then
                    Set Ret = New PixelDataProvider
                    Ret.Ifc = pIPixelDataProvider
                End If
            End If
        End If
    End If
    Set GetPixelDataTransformedAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function




