VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropertyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Foundation_PropertyValue As String = "Windows.Foundation.PropertyValue"
Private Const IID_IPropertyValue As String = "{4bd682dd-7554-40e9-9a9b-82654ede7e62}"
Private Const IID_IPropertyValueStatics As String = "{629bdbc8-d932-4ff4-96b9-8d96c5c1e858}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPropertyValueStatics
    IPropertyValueStatics_CreateEmpty = 6
    IPropertyValueStatics_CreateUInt8 = 7
    IPropertyValueStatics_CreateInt16 = 8
    IPropertyValueStatics_CreateUInt16 = 9
    IPropertyValueStatics_CreateInt32 = 10
    IPropertyValueStatics_CreateUInt32 = 11
    IPropertyValueStatics_CreateInt64 = 12
    IPropertyValueStatics_CreateUInt64 = 13
    IPropertyValueStatics_CreateSingle = 14
    IPropertyValueStatics_CreateDouble = 15
    IPropertyValueStatics_CreateChar16 = 16
    IPropertyValueStatics_CreateBoolean = 17
    IPropertyValueStatics_CreateString = 18
    IPropertyValueStatics_CreateInspectable = 19
    IPropertyValueStatics_CreateGuid = 20
    IPropertyValueStatics_CreateDateTime = 21
    IPropertyValueStatics_CreateTimeSpan = 22
    IPropertyValueStatics_CreatePoint = 23
    IPropertyValueStatics_CreateSize = 24
    IPropertyValueStatics_CreateRect = 25
    IPropertyValueStatics_CreateUInt8Array = 26
    IPropertyValueStatics_CreateInt16Array = 27
    IPropertyValueStatics_CreateUInt16Array = 28
    IPropertyValueStatics_CreateInt32Array = 29
    IPropertyValueStatics_CreateUInt32Array = 30
    IPropertyValueStatics_CreateInt64Array = 31
    IPropertyValueStatics_CreateUInt64Array = 32
    IPropertyValueStatics_CreateSingleArray = 33
    IPropertyValueStatics_CreateDoubleArray = 34
    IPropertyValueStatics_CreateChar16Array = 35
    IPropertyValueStatics_CreateBooleanArray = 36
    IPropertyValueStatics_CreateStringArray = 37
    IPropertyValueStatics_CreateInspectableArray = 38
    IPropertyValueStatics_CreateGuidArray = 39
'    IPropertyValueStatics_CreateDateTimeArray = 40
'    IPropertyValueStatics_CreateTimeSpanArray = 41
'    IPropertyValueStatics_CreatePointArray = 42
'    IPropertyValueStatics_CreateSizeArray = 43
'    IPropertyValueStatics_CreateRectArray = 44
                    
    ' IPropertyValue
    IPropertyValue_GetType = 6
    IPropertyValue_GetIsNumericScalar = 7
    IPropertyValue_GetUInt8 = 8
    IPropertyValue_GetInt16 = 9
    IPropertyValue_GetUInt16 = 10
    IPropertyValue_GetInt32 = 11
    IPropertyValue_GetUInt32 = 12
    IPropertyValue_GetInt64 = 13
    IPropertyValue_GetUInt64 = 14
    IPropertyValue_GetSingle = 15
    IPropertyValue_GetDouble = 16
    IPropertyValue_GetChar16 = 17
    IPropertyValue_GetBoolean = 18
    IPropertyValue_GetString = 19
    IPropertyValue_GetGuid = 20
    IPropertyValue_GetDateTime = 21
    IPropertyValue_GetTimeSpan = 22
'    IPropertyValue_GetPoint = 23
'    IPropertyValue_GetSize = 24
'    IPropertyValue_GetRect = 25
'    IPropertyValue_GetUInt8Array = 26
'    IPropertyValue_GetInt16Array = 27
'    IPropertyValue_GetUInt16Array = 28
'    IPropertyValue_GetInt32Array = 29
'    IPropertyValue_GetUInt32Array = 30
'    IPropertyValue_GetInt64Array = 31
'    IPropertyValue_GetUInt64Array = 32
'    IPropertyValue_GetSingleArray = 33
'    IPropertyValue_GetDoubleArray = 34
'    IPropertyValue_GetChar16Array = 35
'    IPropertyValue_GetBooleanArray = 36
'    IPropertyValue_GetStringArray = 37
'    IPropertyValue_GetInspectableArray = 38
'    IPropertyValue_GetGuidArray = 39
'    IPropertyValue_GetDateTimeArray = 40
'    IPropertyValue_GetTimeSpanArray = 41
'    IPropertyValue_GetPointArray = 42
'    IPropertyValue_GetSizeArray = 43
'    IPropertyValue_GetRectArray = 44

End Enum

' ----==== Variablen ====----
Private m_pIPropertyValue As Long
Private m_pIPropertyValueStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Foundation_PropertyValue, _
                            IID_IPropertyValueStatics, _
                            m_pIPropertyValueStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPropertyValue)
    Call ReleaseIfc(m_pIPropertyValueStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPropertyValue
End Property

Public Property Let Ifc(ByVal pIPropertyValue As Long)
    m_pIPropertyValue = pIPropertyValue
End Property

' IPropertyValue
Public Property Get PropType() As PropertyType
    Dim Ret As PropertyType
    If m_pIPropertyValue <> 0& Then
        Dim value As Long
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetType, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PropType = Ret
End Property

Public Property Get IsNumericScalar() As Boolean
    Dim Ret As Boolean
    If m_pIPropertyValue <> 0& Then
        Dim value As Long
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetIsNumericScalar, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsNumericScalar = Ret
End Property

' ----==== Functions ====----
Public Function CreateEmpty() As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateEmpty, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateEmpty = Ret
End Function

Public Function CreateUInt8(ByVal value As Byte) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt8, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt8 = Ret
End Function

Public Function CreateInt16(ByVal value As Integer) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInt16, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInt16 = Ret
End Function

Public Function CreateUInt16(ByVal value As Integer) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt16, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt16 = Ret
End Function

Public Function CreateInt32(ByVal value As Long) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInt32, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInt32 = Ret
End Function

Public Function CreateUInt32(ByVal value As Long) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt32, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt32 = Ret
End Function

Public Function CreateInt64(ByVal value As Currency) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInt64, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInt64 = Ret
End Function

Public Function CreateUInt64(ByVal value As Currency) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt64, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt64 = Ret
End Function

Public Function CreateSingle(ByVal value As Single) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateSingle, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateSingle = Ret
End Function

Public Function CreateDouble(ByVal value As Double) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateDouble, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateDouble = Ret
End Function

Public Function CreateChar16(ByVal value As Integer) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateChar16, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateChar16 = Ret
End Function

Public Function CreateBoolean(ByVal value As Boolean) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateBoolean, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateBoolean = Ret
End Function

Public Function CreateString(ByVal value As String) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateString, _
                  hString, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
        Call DeleteWindowsString(hString)
    End If
    CreateString = Ret
End Function

Public Function CreateInspectable(ByVal value As Long) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInspectable, _
                  value, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInspectable = Ret
End Function

Friend Function CreateGuid(ByRef value As GUID) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateGuid, _
                  value.data1, _
                  value.data2, _
                  value.data3, _
                  value.data4, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateGuid = Ret
End Function

Public Function CreateDateTime(ByVal value As DateTime) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And _
       IsNotNothing(value) Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateDateTime, _
                  value.DateTimeValue, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateDateTime = Ret
End Function

Public Function CreateTimeSpan(ByVal value As TimeSpan) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And _
       IsNotNothing(value) Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateTimeSpan, _
                  value.TimeSpanValue, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateTimeSpan = Ret
End Function

Public Function CreatePoint(ByVal value As Point) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And _
       IsNotNothing(value) Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreatePoint, _
                  value.X, _
                  value.Y, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreatePoint = Ret
End Function

Public Function CreateSize(ByVal value As Size) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And _
       IsNotNothing(value) Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateSize, _
                  value.Width, _
                  value.Height, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateSize = Ret
End Function

Public Function CreateRect(ByVal value As Rect) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And _
       IsNotNothing(value) Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateRect, _
                  value.X, _
                  value.Y, _
                  value.Width, _
                  value.Height, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateRect = Ret
End Function

Public Function CreateUInt8Array(ByVal valueLength As Long, ByRef value() As Byte) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt8Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt8Array = Ret
End Function

Public Function CreateInt16Array(ByVal valueLength As Long, ByRef value() As Integer) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInt16Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInt16Array = Ret
End Function

Public Function CreateUInt16Array(ByVal valueLength As Long, ByRef value() As Integer) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt16Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt16Array = Ret
End Function

Public Function CreateInt32Array(ByVal valueLength As Long, ByRef value() As Long) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInt32Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInt32Array = Ret
End Function

Public Function CreateUInt32Array(ByVal valueLength As Long, ByRef value() As Long) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt32Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt32Array = Ret
End Function

Public Function CreateInt64Array(ByVal valueLength As Long, ByRef value() As Currency) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInt64Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInt64Array = Ret
End Function

Public Function CreateUInt64Array(ByVal valueLength As Long, ByRef value() As Currency) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateUInt64Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateUInt64Array = Ret
End Function

Public Function CreateSingleArray(ByVal valueLength As Long, ByRef value() As Single) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateSingleArray, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateSingleArray = Ret
End Function

Public Function CreateDoubleArray(ByVal valueLength As Long, ByRef value() As Double) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateDoubleArray, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateDoubleArray = Ret
End Function

Public Function CreateChar16Array(ByVal valueLength As Long, ByRef value() As Long) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateChar16Array, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateChar16Array = Ret
End Function

Public Function CreateBooleanArray(ByVal valueLength As Long, ByRef value() As Boolean) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateBooleanArray, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateBooleanArray = Ret
End Function

Public Function CreateStringArray(ByVal valueLength As Long, ByRef value() As String) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim Item As Long
        Dim hString() As Long
        ReDim hString(valueLength - 1)
        For Item = 0 To valueLength - 1
            hString(Item) = CreateWindowsString(value(Item))
        Next
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateStringArray, _
                  valueLength, _
                  VarPtr(hString(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
        For Item = 0 To valueLength - 1
            Call DeleteWindowsString(hString(Item))
        Next
    End If
    CreateStringArray = Ret
End Function

Public Function CreateInspectableArray(ByVal valueLength As Long, ByRef value() As Long) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateInspectableArray, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateInspectableArray = Ret
End Function

Friend Function CreateGuidArray(ByVal valueLength As Long, ByRef value() As GUID) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIPropertyValueStatics, _
                  IPropertyValueStatics_CreateGuidArray, _
                  valueLength, _
                  VarPtr(value(0)), _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = pIInspectable
        End If
    End If
    CreateGuidArray = Ret
End Function

'Friend Function CreateDateTimeArray(ByVal valueLength As Long, ByRef value() As DateTime) As Long
'    Dim Ret As Long
'    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
'        Dim pIInspectable As Long
'        If Invoke(m_pIPropertyValueStatics, _
'                  IPropertyValueStatics_CreateDateTimeArray, _
'                  valueLength, _
'                  VarPtr(value(0)), _
'                  VarPtr(pIInspectable)) = S_OK Then
'            Ret = pIInspectable
'        End If
'    End If
'    CreateDateTimeArray = Ret
'End Function

'Friend Function CreateTimeSpanArray(ByVal valueLength As Long, ByRef value() As TimeSpan) As Long
'    Dim Ret As Long
'    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
'        Dim pIInspectable As Long
'        If Invoke(m_pIPropertyValueStatics, _
'                  IPropertyValueStatics_CreateTimeSpanArray, _
'                  valueLength, _
'                  VarPtr(value(0)), _
'                  VarPtr(pIInspectable)) = S_OK Then
'            Ret = pIInspectable
'        End If
'    End If
'    CreateTimeSpanArray = Ret
'End Function

'Friend Function CreatePointArray(ByVal valueLength As Long, ByRef value() As Point) As Long
'    Dim Ret As Long
'    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
'        Dim pIInspectable As Long
'        If Invoke(m_pIPropertyValueStatics, _
'                  IPropertyValueStatics_CreatePointArray, _
'                  valueLength, _
'                  VarPtr(value(0)), _
'                  VarPtr(pIInspectable)) = S_OK Then
'            Ret = pIInspectable
'        End If
'    End If
'    CreatePointArray = Ret
'End Function
'
'Friend Function CreateSizeArray(ByVal valueLength As Long, ByRef value() As Size) As Long
'    Dim Ret As Long
'    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
'        Dim pIInspectable As Long
'        If Invoke(m_pIPropertyValueStatics, _
'                  IPropertyValueStatics_CreateSizeArray, _
'                  valueLength, _
'                  VarPtr(value(0)), _
'                  VarPtr(pIInspectable)) = S_OK Then
'            Ret = pIInspectable
'        End If
'    End If
'    CreateSizeArray = Ret
'End Function
'
'Friend Function CreateRectArray(ByVal valueLength As Long, ByRef value() As Rect) As Long
'    Dim Ret As Long
'    If m_pIPropertyValueStatics <> 0& And valueLength > 0& Then
'        Dim pIInspectable As Long
'        If Invoke(m_pIPropertyValueStatics, _
'                  IPropertyValueStatics_CreateRectArray, _
'                  valueLength, _
'                  VarPtr(value(0)), _
'                  VarPtr(pIInspectable)) = S_OK Then
'            Ret = pIInspectable
'        End If
'    End If
'    CreateRectArray = Ret
'End Function

' IPropertyValue
Public Function GetUInt8() As Byte
    Dim Ret As Byte
    If m_pIPropertyValue <> 0& Then
        Dim value As Byte
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetUInt8, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetUInt8 = Ret
End Function

Public Function GetInt16() As Integer
    Dim Ret As Integer
    If m_pIPropertyValue <> 0& Then
        Dim value As Integer
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetInt16, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetInt16 = Ret
End Function

Public Function GetUInt16() As Integer
    Dim Ret As Integer
    If m_pIPropertyValue <> 0& Then
        Dim value As Integer
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetUInt16, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetUInt16 = Ret
End Function

Public Function GetInt32() As Long
    Dim Ret As Long
    If m_pIPropertyValue <> 0& Then
        Dim value As Long
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetInt32, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetInt32 = Ret
End Function

Public Function GetUInt32() As Long
    Dim Ret As Long
    If m_pIPropertyValue <> 0& Then
        Dim value As Long
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetUInt32, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetUInt32 = Ret
End Function

Public Function GetInt64() As Currency
    Dim Ret As Currency
    If m_pIPropertyValue <> 0& Then
        Dim value As Currency
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetInt64, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetInt64 = Ret
End Function

Public Function GetUInt64() As Currency
    Dim Ret As Currency
    If m_pIPropertyValue <> 0& Then
        Dim value As Currency
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetUInt64, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetUInt64 = Ret
End Function

Public Function GetSingle() As Single
    Dim Ret As Single
    If m_pIPropertyValue <> 0& Then
        Dim value As Single
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetSingle, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetSingle = Ret
End Function

Public Function GetDouble() As Double
    Dim Ret As Double
    If m_pIPropertyValue <> 0& Then
        Dim value As Single
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetDouble, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetDouble = Ret
End Function

Public Function GetChar16() As Integer
    Dim Ret As Integer
    If m_pIPropertyValue <> 0& Then
        Dim value As Integer
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetChar16, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetChar16 = Ret
End Function

Public Function GetBoolean() As Boolean
    Dim Ret As Boolean
    If m_pIPropertyValue <> 0& Then
        Dim value As Long
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetBoolean, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    GetBoolean = Ret
End Function

Public Function GetString() As String
    Dim Ret As String
    If m_pIPropertyValue <> 0& Then
        Dim value As Long
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetString, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GetString = Ret
End Function

Friend Function GetGuid() As GUID
    Dim Ret As GUID
    If m_pIPropertyValue <> 0& Then
        Dim value As GUID
        If Invoke(m_pIPropertyValue, _
                  IPropertyValue_GetGuid, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetGuid = Ret
End Function

'Friend Function GetDateTime() As DateTime
'    Dim Ret As DateTime
'    If m_pIPropertyValue <> 0& Then
'        Dim value As Windows_Foundation_DateTime
'        If Invoke(m_pIPropertyValue, _
'                  IPropertyValue_GetDateTime, _
'                  VarPtr(value)) = S_OK Then
'            Ret = value
'        End If
'    End If
'    GetDateTime = Ret
'End Function
'
'Friend Function GetTimeSpan() As TimeSpan
'    Dim Ret As TimeSpan
'    If m_pIPropertyValue <> 0& Then
'        Dim value As Windows_Foundation_TimeSpan
'        If Invoke(m_pIPropertyValue, _
'                  IPropertyValue_GetTimeSpan, _
'                  VarPtr(value)) = S_OK Then
'            Ret = value
'        End If
'    End If
'    GetTimeSpan = Ret
'End Function

'    IPropertyValue_GetPoint = 23
'    IPropertyValue_GetSize = 24
'    IPropertyValue_GetRect = 25
'    IPropertyValue_GetUInt8Array = 26
'    IPropertyValue_GetInt16Array = 27
'    IPropertyValue_GetUInt16Array = 28
'    IPropertyValue_GetInt32Array = 29
'    IPropertyValue_GetUInt32Array = 30
'    IPropertyValue_GetInt64Array = 31
'    IPropertyValue_GetUInt64Array = 32
'    IPropertyValue_GetSingleArray = 33
'    IPropertyValue_GetDoubleArray = 34
'    IPropertyValue_GetChar16Array = 35
'    IPropertyValue_GetBooleanArray = 36
'    IPropertyValue_GetStringArray = 37
'    IPropertyValue_GetInspectableArray = 38
'    IPropertyValue_GetGuidArray = 39
'    IPropertyValue_GetDateTimeArray = 40
'    IPropertyValue_GetTimeSpanArray = 41
'    IPropertyValue_GetPointArray = 42
'    IPropertyValue_GetSizeArray = 43
'    IPropertyValue_GetRectArray = 44

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


