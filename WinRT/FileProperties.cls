VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace FileProperties

Option Explicit

' ----==== Enums Windows.Storage.FileProperties.xxx ====----
Public Enum VideoOrientation
    VideoOrientation_Normal = 0
    VideoOrientation_Rotate90 = 90
    VideoOrientation_Rotate180 = 180
    VideoOrientation_Rotate270 = 270
End Enum

Public Enum PhotoOrientation
    PhotoOrientation_Unspecified = 0
    PhotoOrientation_Normal = 1
    PhotoOrientation_FlipHorizontal = 2
    PhotoOrientation_Rotate180 = 3
    PhotoOrientation_FlipVertical = 4
    PhotoOrientation_Transpose = 5
    PhotoOrientation_Rotate270 = 6
    PhotoOrientation_Transverse = 7
    PhotoOrientation_Rotate90 = 8
End Enum

Public Enum ThumbnailMode
    ThumbnailMode_PicturesView = 0
    ThumbnailMode_VideosView = 1
    ThumbnailMode_MusicView = 2
    ThumbnailMode_DocumentsView = 3
    ThumbnailMode_ListView = 4
    ThumbnailMode_SingleItem = 5
End Enum

Public Enum ThumbnailOptions
    ThumbnailOptions_None = &H0&
    ThumbnailOptions_ReturnOnlyIfCached = &H1&
    ThumbnailOptions_ResizeThumbnail = &H2&
    ThumbnailOptions_UseCurrentScale = &H4&
End Enum

Public Enum ThumbnailType
    ThumbnailType_Image = 0
    ThumbnailType_Icon = 1
End Enum

' ----==== Namespace FileProperties.BasicProperties ====----
Public Property Get BasicProperties() As BasicProperties
    Set BasicProperties = New BasicProperties
End Property

' ----==== Namespace FileProperties.MusicProperties ====----
Public Property Get MusicProperties() As MusicProperties
    Set MusicProperties = New MusicProperties
End Property

' ----==== Namespace FileProperties.VideoProperties ====----
Public Property Get VideoProperties() As VideoProperties
    Set VideoProperties = New VideoProperties
End Property

' ----==== Namespace FileProperties.ImageProperties ====----
Public Property Get ImageProperties() As ImageProperties
    Set ImageProperties = New ImageProperties
End Property

' ----==== Namespace FileProperties.ImageProperties ====----
Public Property Get DocumentProperties() As DocumentProperties
    Set DocumentProperties = New DocumentProperties
End Property

' ----==== Namespace FileProperties.StorageItemThumbnail ====----
Public Property Get StorageItemThumbnail() As StorageItemThumbnail
    Set StorageItemThumbnail = New StorageItemThumbnail
End Property

' ----==== Namespace FileProperties.StorageItemContentProperties ====----
Public Property Get StorageItemContentProperties() As StorageItemContentProperties
    Set StorageItemContentProperties = New StorageItemContentProperties
End Property

