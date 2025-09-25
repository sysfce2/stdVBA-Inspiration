VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pickers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Pickers

Option Explicit

' ----==== Enums Windows.Storage.Pickers.xxx ====----
Public Enum PickerViewMode
    PickerViewMode_List = 0
    PickerViewMode_Thumbnail = 1
End Enum

Public Enum PickerLocationId
    PickerLocationId_DocumentsLibrary = 0
    PickerLocationId_ComputerFolder = 1
    PickerLocationId_Desktop = 2
    PickerLocationId_Downloads = 3
    PickerLocationId_HomeGroup = 4
    PickerLocationId_MusicLibrary = 5
    PickerLocationId_PicturesLibrary = 6
    PickerLocationId_VideosLibrary = 7
    PickerLocationId_Objects3D = 8
    PickerLocationId_Unspecified = 9
End Enum

' ----==== Namespace Pickers.FolderPicker ====----
Public Property Get FolderPicker() As FolderPicker
    Set FolderPicker = New FolderPicker
End Property

' ----==== Namespace Pickers.FileOpenPicker ====----
Public Property Get FileOpenPicker() As FileOpenPicker
    Set FileOpenPicker = New FileOpenPicker
End Property

' ----==== Namespace Pickers.FileSavePicker ====----
Public Property Get FileSavePicker() As FileSavePicker
    Set FileSavePicker = New FileSavePicker
End Property

' ----==== Namespace Pickers.FilePickerFileTypesOrderedMap ====----
Public Property Get FilePickerFileTypesOrderedMap() As FilePickerFileTypesOrderedMap
    Set FilePickerFileTypesOrderedMap = New FilePickerFileTypesOrderedMap
End Property

