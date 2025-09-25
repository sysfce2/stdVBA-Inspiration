VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Foundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

' Namespace Foundation

Option Explicit

' ----==== Enums Windows.Foundation.xxx ====----
Public Enum PropertyType
    PropertyType_Empty = 0
    PropertyType_UInt8 = 1
    PropertyType_Int16 = 2
    PropertyType_UInt16 = 3
    PropertyType_Int32 = 4
    PropertyType_UInt32 = 5
    PropertyType_Int64 = 6
    PropertyType_UInt64 = 7
    PropertyType_Single = 8
    PropertyType_Double = 9
    PropertyType_Char16 = 10
    PropertyType_Boolean = 11
    PropertyType_String = 12
    PropertyType_Inspectable = 13
    PropertyType_DateTime = 14
    PropertyType_TimeSpan = 15
    PropertyType_Guid = 16
    PropertyType_Point = 17
    PropertyType_Size = 18
    PropertyType_Rect = 19
    PropertyType_OtherType = 20
    PropertyType_UInt8Array = 1025
    PropertyType_Int16Array = 1026
    PropertyType_UInt16Array = 1027
    PropertyType_Int32Array = 1028
    PropertyType_UInt32Array = 1029
    PropertyType_Int64Array = 1030
    PropertyType_UInt64Array = 1031
    PropertyType_SingleArray = 1032
    PropertyType_DoubleArray = 1033
    PropertyType_Char16Array = 1034
    PropertyType_BooleanArray = 1035
    PropertyType_StringArray = 1036
    PropertyType_InspectableArray = 1037
    PropertyType_DateTimeArray = 1038
    PropertyType_TimeSpanArray = 1039
    PropertyType_GuidArray = 1040
    PropertyType_PointArray = 1041
    PropertyType_SizeArray = 1042
    PropertyType_RectArray = 1043
    PropertyType_OtherTypeArray = 1044
End Enum

' ----==== Namespace Foundation.Size ====----
Public Property Get Size(Optional ByVal Width As Single = 0, _
                         Optional ByVal Height As Single = 0) As Size
    Set Size = New Size
    Size.Width = Width
    Size.Height = Height
End Property

' ----==== Namespace Foundation.Point ====----
Public Property Get Point(Optional ByVal X As Single = 0, _
                          Optional ByVal Y As Single = 0) As Point
    Set Point = New Point
    Point.X = X
    Point.Y = Y
End Property

' ----==== Namespace Foundation.Rect ====----
Public Property Get Rect(Optional ByVal X As Single = 0, _
                         Optional ByVal Y As Single = 0, _
                         Optional ByVal Width As Single = 0, _
                         Optional ByVal Height As Single = 0) As Rect
    Set Rect = New Rect
    Rect.X = X
    Rect.Y = Y
    Rect.Width = Width
    Rect.Height = Height
End Property

' ----==== Namespace Foundation.Uri ====----
Public Property Get Uri() As Uri
    Set Uri = New Uri
End Property

' ----==== Namespace Foundation.DateTime ====----
Public Property Get DateTime() As DateTime
    Set DateTime = New DateTime
End Property

' ----==== Namespace Foundation.TimeSpan ====----
Public Property Get TimeSpan() As TimeSpan
    Set TimeSpan = New TimeSpan
End Property

' ----==== Namespace Foundation.Metadata ====----
Public Property Get Metadata() As Metadata
    Set Metadata = New Metadata
End Property

' ----==== Namespace Foundation.MemoryBuffer ====----
Public Property Get MemoryBuffer() As MemoryBuffer
    Set MemoryBuffer = New MemoryBuffer
End Property

' ----==== Namespace Foundation.PropertyValue ====----
Public Property Get PropertyValue() As PropertyValue
    Set PropertyValue = New PropertyValue
End Property

' ----==== Namespace Foundation.MemoryBufferReference ====----
Public Property Get MemoryBufferReference() As MemoryBufferReference
    Set MemoryBufferReference = New MemoryBufferReference
End Property

' ----==== Namespace Foundation.Collections ====----
Public Property Get Collections() As Collections
    Set Collections = New Collections
End Property


