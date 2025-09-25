VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Imaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

' Namespace Imaging

Option Explicit

' ----==== Enums Windows.Graphics.Imaging.xxx ====----
Public Enum BitmapPixelFormat
    BitmapPixelFormat_Unknown = 0
    BitmapPixelFormat_Rgba16 = 12
    BitmapPixelFormat_Rgba8 = 30
    BitmapPixelFormat_Gray16 = 57
    BitmapPixelFormat_Gray8 = 62
    BitmapPixelFormat_Bgra8 = 87
    BitmapPixelFormat_Nv12 = 103
    BitmapPixelFormat_P010 = 104
    BitmapPixelFormat_Yuy2 = 107
End Enum

Public Enum BitmapAlphaMode
    BitmapAlphaMode_Premultiplied = 0
    BitmapAlphaMode_Straight = 1
    BitmapAlphaMode_Ignore = 2
End Enum

Public Enum BitmapBufferAccessMode
    BitmapBufferAccessMode_Read = 0
    BitmapBufferAccessMode_ReadWrite = 1
    BitmapBufferAccessMode_Write = 2
End Enum

Public Enum BitmapInterpolationMode
    BitmapInterpolationMode_NearestNeighbor = 0
    BitmapInterpolationMode_Linear = 1
    BitmapInterpolationMode_Cubic = 2
    BitmapInterpolationMode_Fant = 3
End Enum

Public Enum BitmapFlip
    BitmapFlip_None = 0
    BitmapFlip_Horizontal = 1
    BitmapFlip_Vertical = 2
End Enum

Public Enum BitmapRotation
    BitmapRotation_None = 0
    BitmapRotation_Clockwise90Degrees = 1
    BitmapRotation_Clockwise180Degrees = 2
    BitmapRotation_Clockwise270Degrees = 3
End Enum

Public Enum ExifOrientationMode
    ExifOrientationMode_IgnoreExifOrientation = 0
    ExifOrientationMode_RespectExifOrientation = 1
End Enum

Public Enum ColorManagementMode
    ColorManagementMode_DoNotColorManage = 0
    ColorManagementMode_ColorManageToSRgb = 1
End Enum

' ----==== Namespace Imaging.BitmapBuffer ====----
Public Property Get BitmapBuffer() As BitmapBuffer
    Set BitmapBuffer = New BitmapBuffer
End Property

' ----==== Namespace Imaging.ImageStream ====----
Public Property Get ImageStream() As ImageStream
    Set ImageStream = New ImageStream
End Property

' ----==== Namespace Imaging.BitmapDecoder ====----
Public Property Get BitmapDecoder() As BitmapDecoder
    Set BitmapDecoder = New BitmapDecoder
End Property

' ----==== Namespace Imaging.BitmapEncoder ====----
Public Property Get BitmapEncoder() As BitmapEncoder
    Set BitmapEncoder = New BitmapEncoder
End Property

' ----==== Namespace Imaging.BitmapFrame ====----
Public Property Get BitmapFrame() As BitmapFrame
    Set BitmapFrame = New BitmapFrame
End Property

' ----==== Namespace Imaging.SoftwareBitmap ====----
Public Property Get SoftwareBitmap() As SoftwareBitmap
    Set SoftwareBitmap = New SoftwareBitmap
End Property

' ----==== Namespace Imaging.BitmapProperties ====----
Public Property Get BitmapProperties() As BitmapProperties
    Set BitmapProperties = New BitmapProperties
End Property

' ----==== Namespace Imaging.BitmapTransform ====----
Public Property Get BitmapTransform() As BitmapTransform
    Set BitmapTransform = New BitmapTransform
End Property

' ----==== Namespace Imaging.BitmapPropertiesView ====----
Public Property Get BitmapPropertiesView() As BitmapPropertiesView
    Set BitmapPropertiesView = New BitmapPropertiesView
End Property

' ----==== Namespace Imaging.PixelDataProvider ====----
Public Property Get PixelDataProvider() As PixelDataProvider
    Set PixelDataProvider = New PixelDataProvider
End Property

' ----==== Namespace Imaging.BitmapCodecInformation ====----
Public Property Get BitmapCodecInformation() As BitmapCodecInformation
    Set BitmapCodecInformation = New BitmapCodecInformation
End Property

' ----==== Namespace Imaging.BitmapSize ====----
Public Property Get BitmapSize(Optional ByVal Width As Long = 0, _
                               Optional ByVal Height As Long = 0) As BitmapSize
    Set BitmapSize = New BitmapSize
    BitmapSize.Width = Width
    BitmapSize.Height = Height
End Property

' ----==== Namespace Imaging.BitmapBounds ====----
Public Property Get BitmapBounds(Optional ByVal X As Long = 0, _
                                 Optional ByVal Y As Long = 0, _
                                 Optional ByVal Width As Long = 0, _
                                 Optional ByVal Height As Long = 0) As BitmapBounds
    Set BitmapBounds = New BitmapBounds
    BitmapBounds.X = X
    BitmapBounds.Y = Y
    BitmapBounds.Width = Width
    BitmapBounds.Height = Height
End Property

' ----==== Namespace Imaging.BitmapPlaneDescription ====----
Public Property Get BitmapPlaneDescription(Optional ByVal StartIndex As Long = 0, _
                                           Optional ByVal Width As Long = 0, _
                                           Optional ByVal Height As Long = 0, _
                                           Optional ByVal Stride As Long = 0) As BitmapPlaneDescription
    Set BitmapPlaneDescription = New BitmapPlaneDescription
    BitmapPlaneDescription.StartIndex = StartIndex
    BitmapPlaneDescription.Width = Width
    BitmapPlaneDescription.Height = Height
    BitmapPlaneDescription.Stride = Stride
End Property


