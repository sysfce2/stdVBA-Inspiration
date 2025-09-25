VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Json_JsonArray As String = "Windows.Data.Json.JsonArray"
Private Const IID_IJsonArray As String = "{08c1ddb6-0cbd-4a9a-b5d3-2f852dc37e81}"
Private Const IID_IJsonValue As String = "{a3219ecb-f0b3-4dcd-beee-19d48cd3ed1e}"
Private Const IID_IStringable As String = "{96369f54-8eb6-48f0-abce-c1b211e627c3}"
Private Const IID_IJsonArrayStatics As String = "{db1434a9-e164-499f-93e2-8a8f49bb90ba}"
Private Const IID_IVector_IJsonValue As String = "{d44662bc-dce3-59a8-9272-4b210f33908b}"
Private Const IID_IIterable_IJsonValue As String = "{cb0492b6-4113-55cf-b2c5-99eb428ba493}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IJsonArrayStatics
    IJsonArrayStatics_Parse = 6
    IJsonArrayStatics_TryParse = 7
    
    'IJsonArray
    IJsonArray_GetObjectAt = 6
    IJsonArray_GetArrayAt = 7
    IJsonArray_GetStringAt = 8
    IJsonArray_GetNumberAt = 9
    IJsonArray_GetBooleanAt = 10
    
    'IJsonValue
    IJsonValue_GetValueType = 6
    IJsonValue_Stringify = 7
    IJsonValue_GetString = 8
    IJsonValue_GetNumber = 9
    IJsonValue_GetBoolean = 10
    IJsonValue_GetArray = 11
    IJsonValue_GetObject = 12

    ' IStringable
    IStringable_ToString = 6

    ' IVector_IJsonValue
    IVector_IJsonValue_GetAt = 6
    IVector_IJsonValue_GetSize = 7
    IVector_IJsonValue_GetView = 8
    IVector_IJsonValue_IndexOf = 9
    IVector_IJsonValue_SetAt = 10
    IVector_IJsonValue_InsertAt = 11
    IVector_IJsonValue_RemoveAt = 12
    IVector_IJsonValue_Append = 13
    IVector_IJsonValue_RemoveAtEnd = 14
    IVector_IJsonValue_Clear = 15
    IVector_IJsonValue_GetMany = 16
    IVector_IJsonValue_ReplaceAll = 17

    ' IIterable_IJsonValue
    IIterable_IJsonValue_First = 6

    ' IIterator_IJsonValue
    IIterator_IJsonValue_GetCurrent = 6
    IIterator_IJsonValue_GetHasCurrent = 7
    IIterator_IJsonValue_MoveNext = 8
    IIterator_IJsonValue_GetMany = 9

End Enum

' ----==== Variablen ====----
Private cJsonError As JsonError
Private m_ErrorStatus As JsonErrorStatus
Private m_pIJsonArray As Long
Private m_pIJsonValue As Long
Private m_pIStringable As Long
Private m_pIJsonArrayStatics As Long
Private m_pIVector_IJsonValue As Long
Private m_pIIterable_IJsonValue As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Json_JsonArray, _
                            IID_IJsonArrayStatics, _
                            m_pIJsonArrayStatics) Then
    End If
    Dim pIJsonArray As Long
    If GetActivateInstance(Windows_Data_Json_JsonArray, _
                           IID_IJsonArray, _
                           pIJsonArray) Then
        Me.Ifc = pIJsonArray
    End If
    Set cJsonError = New JsonError
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIJsonValue)
    Call ReleaseIfc(m_pIJsonArray)
    Call ReleaseIfc(m_pIJsonArrayStatics)
    Call ReleaseIfc(m_pIIterable_IJsonValue)
    Call ReleaseIfc(m_pIVector_IJsonValue)
    Set cJsonError = Nothing
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIJsonArray
End Property

Public Property Let Ifc(ByVal pIJsonArray As Long)
    Call ReleaseIfc(m_pIIterable_IJsonValue)
    Call ReleaseIfc(m_pIVector_IJsonValue)
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIJsonValue)
    Call ReleaseIfc(m_pIJsonArray)
    m_pIJsonArray = pIJsonArray
    Call QueryIfc(m_pIJsonArray, IID_IJsonValue, m_pIJsonValue)
    Call QueryIfc(m_pIJsonArray, IID_IStringable, m_pIStringable)
    Call QueryIfc(m_pIJsonArray, IID_IVector_IJsonValue, m_pIVector_IJsonValue)
    Call QueryIfc(m_pIJsonArray, IID_IIterable_IJsonValue, m_pIIterable_IJsonValue)
End Property

Public Property Get JsonStatus() As JsonErrorStatus
    JsonStatus = m_ErrorStatus
End Property

'IJsonValue
Public Property Get ValueType() As JsonValueType
    Dim Ret As JsonValueType
    If m_pIJsonValue <> 0& Then
        Dim Value As Long
        If Invoke(m_pIJsonValue, _
                  IJsonValue_GetValueType, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    ValueType = Ret
End Property

' IVector_IJsonValue
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVector_IJsonValue <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
' IJsonArrayStatics
Public Function Parse(ByVal inputStr As String) As JsonArray
    Dim Ret As JsonArray
    If m_pIJsonArrayStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(inputStr)
        Dim pIJsonArray As Long
        If Invoke(m_pIJsonArrayStatics, _
                  IJsonArrayStatics_Parse, _
                  hString, _
                  VarPtr(pIJsonArray)) = S_OK Then
            If pIJsonArray <> 0& Then
                Set Ret = New JsonArray
                Ret.Ifc = pIJsonArray
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Parse = Ret
End Function

Public Function TryParse(ByVal inputStr As String, _
                         ByRef result As JsonArray) As Boolean
    Dim bolRet As Boolean
    If m_pIJsonArrayStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(inputStr)
        Dim pIJsonArray As Long
        Dim succeeded As Boolean
        If Invoke(m_pIJsonArrayStatics, _
                  IJsonArrayStatics_TryParse, _
                  hString, _
                  VarPtr(pIJsonArray), _
                  VarPtr(succeeded)) = S_OK Then
            bolRet = succeeded
            If pIJsonArray <> 0& Then
                Set result = New JsonArray
                result.Ifc = pIJsonArray
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    TryParse = bolRet
End Function

'IJsonArray
Public Function GetObjectAt(ByVal index As Long) As JsonObject
    Dim Ret As JsonObject
    If m_pIJsonArray <> 0& Then
        Dim pIJsonObject As Long
        If Invoke(m_pIJsonArray, _
                  IJsonArray_GetObjectAt, _
                  index, _
                  VarPtr(pIJsonObject)) = S_OK Then
            If pIJsonObject <> 0& Then
                Set Ret = New JsonObject
                Ret.Ifc = pIJsonObject
            End If
        End If
    End If
    Set GetObjectAt = Ret
End Function

Public Function GetArrayAt(ByVal index As Long) As JsonArray
    Dim Ret As JsonArray
    If m_pIJsonArray <> 0& Then
        Dim pIJsonArray As Long
        If Invoke(m_pIJsonArray, _
                  IJsonArray_GetArrayAt, _
                  index, _
                  VarPtr(pIJsonArray)) = S_OK Then
            If pIJsonArray <> 0& Then
                Set Ret = New JsonArray
                Ret.Ifc = pIJsonArray
            End If
        End If
    End If
    Set GetArrayAt = Ret
End Function

Public Function GetStringAt(ByVal index As Long) As String
    Dim Ret As String
    If m_pIJsonArray <> 0& Then
        Dim Value As Long
        If Invoke(m_pIJsonArray, _
                  IJsonArray_GetStringAt, _
                  index, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    GetStringAt = Ret
End Function

Public Function GetNumberAt(ByVal index As Long) As Double
    Dim Ret As Double
    If m_pIJsonArray <> 0& Then
        Dim Value As Double
        If Invoke(m_pIJsonArray, _
                  IJsonArray_GetNumberAt, _
                  index, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GetNumberAt = Ret
End Function

Public Function GetBooleanAt(ByVal index As Long) As Boolean
    Dim Ret As Boolean
    If m_pIJsonArray <> 0& Then
        Dim Value As Long
        If Invoke(m_pIJsonArray, _
                  IJsonArray_GetBooleanAt, _
                  index, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    GetBooleanAt = Ret
End Function

'IJsonValue
Public Function Stringify() As String
    Dim Ret As String
    If m_pIJsonValue <> 0& Then
        Dim Value As Long
        If Invoke(m_pIJsonValue, _
                  IJsonValue_Stringify, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Stringify = Ret
End Function

Public Function GetString() As String
    Dim Ret As String
    If m_pIJsonValue <> 0& Then
        Dim Value As Long
        If Invoke(m_pIJsonValue, _
                  IJsonValue_GetString, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    GetString = Ret
End Function

Public Function GetNumber() As Double
    Dim Ret As Double
    If m_pIJsonValue <> 0& Then
        Dim Value As Double
        If Invoke(m_pIJsonValue, _
                  IJsonValue_GetNumber, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GetNumber = Ret
End Function

Public Function GetBoolean() As Boolean
    Dim Ret As Boolean
    If m_pIJsonValue <> 0& Then
        Dim Value As Boolean
        If Invoke(m_pIJsonValue, _
                  IJsonValue_GetBoolean, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GetBoolean = Ret
End Function

Public Function GetArray() As JsonArray
    Dim Ret As JsonArray
    If m_pIJsonValue <> 0& Then
        Dim pIJsonArray As Long
        If Invoke(m_pIJsonValue, _
                  IJsonValue_GetArray, _
                  VarPtr(pIJsonArray)) = S_OK Then
            If pIJsonArray <> 0& Then
                Set Ret = New JsonArray
                Ret.Ifc = pIJsonArray
            End If
        End If
    End If
    Set GetArray = Ret
End Function

Public Function GetObject() As JsonObject
    Dim Ret As JsonObject
    If m_pIJsonValue <> 0& Then
        Dim pIJsonObject As Long
        If Invoke(m_pIJsonValue, _
                  IJsonValue_GetObject, _
                  VarPtr(pIJsonObject)) = S_OK Then
            If pIJsonObject <> 0& Then
                Set Ret = New JsonObject
                Ret.Ifc = pIJsonObject
            End If
        End If
    End If
    Set GetObject = Ret
End Function

' IStringable
Public Function ToString() As String
    Dim Ret As String
    If m_pIStringable <> 0& Then
        Dim Value As Long
        If Invoke(m_pIStringable, _
                  IStringable_ToString, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    ToString = Ret
End Function

' IVector_IJsonValue
Public Function GetAt(ByVal index As Long) As JsonValue
    Dim Ret As JsonValue
    If m_pIVector_IJsonValue <> 0& Then
        Dim pIJsonValue As Long
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_GetAt, _
                  index, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue <> 0& Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
    End If
    Set GetAt = Ret
End Function

Public Function IndexOf(ByVal Value As JsonValue) As Long
    Dim Ret As Long
    If m_pIVector_IJsonValue <> 0& And _
       IsNotNothing(Value) Then
        Dim succeeded As Boolean
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_IndexOf, _
                  Value.Ifc, _
                  VarPtr(Ret), _
                  VarPtr(succeeded)) = S_OK Then
            If Not succeeded Then
                Ret = -1
            End If
        End If
    End If
    IndexOf = Ret
End Function

Public Function SetAt(ByVal index As Long, _
                      ByVal Value As JsonValue) As Boolean
    Dim Ret As Boolean
    If m_pIVector_IJsonValue <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_SetAt, _
                  index, _
                  Value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    SetAt = Ret
End Function

Public Function InsertAt(ByVal index As Long, _
                         ByVal Value As JsonValue) As Boolean
    Dim Ret As Boolean
    If m_pIVector_IJsonValue <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_InsertAt, _
                  index, _
                  Value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    InsertAt = Ret
End Function

Public Function RemoveAt(ByVal index As Long) As Boolean
    Dim Ret As Boolean
    If m_pIVector_IJsonValue <> 0& Then
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_RemoveAt, _
                  index) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAt = Ret
End Function

Public Function Append(ByVal Value As JsonValue) As Boolean
    Dim Ret As Boolean
    If m_pIVector_IJsonValue <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_Append, _
                  Value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    Append = Ret
End Function

Public Function RemoveAtEnd() As Boolean
    Dim Ret As Boolean
    If m_pIVector_IJsonValue <> 0& Then
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_RemoveAtEnd) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAtEnd = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIVector_IJsonValue <> 0& Then
        If Invoke(m_pIVector_IJsonValue, _
                  IVector_IJsonValue_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

' IIterable_IJsonValue
' IIterator_IJsonValue
Public Function GetArrayValues() As JsonValue()
    Dim Ret() As JsonValue
    If m_pIJsonArray <> 0 Then
        Dim count As Long
        count = Me.Size
        If count > 0& Then
            ReDim Ret(count - 1)
            Dim pIIterator_IJsonValue As Long
            If Invoke(m_pIIterable_IJsonValue, _
                      IIterable_IJsonValue_First, _
                      VarPtr(pIIterator_IJsonValue)) = S_OK Then
                Dim Value As Long
                Dim Item As Long
                Dim JsonValue As JsonValue
                For Item = 0 To count - 1
                    If Item = 0 Then
                        If Invoke(pIIterator_IJsonValue, _
                                  IIterator_IJsonValue_GetCurrent, _
                                  VarPtr(Value)) = S_OK Then
                            If Value <> 0& Then
                                Set JsonValue = New JsonValue
                                JsonValue.Ifc = Value
                                Set Ret(Item) = JsonValue
                            End If
                        End If
                    Else
                        Dim moveNext As Long
                        If Invoke(pIIterator_IJsonValue, _
                                  IIterator_IJsonValue_MoveNext, _
                                  VarPtr(moveNext)) = S_OK Then
                            If CBool(moveNext) Then
                                If Invoke(pIIterator_IJsonValue, _
                                          IIterator_IJsonValue_GetCurrent, _
                                          VarPtr(Value)) = S_OK Then
                                    If Value <> 0& Then
                                        Set JsonValue = New JsonValue
                                        JsonValue.Ifc = Value
                                        Set Ret(Item) = JsonValue
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                Call ReleaseIfc(pIIterator_IJsonValue)
            End If
        End If
    End If
    GetArrayValues = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Dim hResult As Long
        hResult = OleInvoke(pInterface, vtb, var)
        m_ErrorStatus = cJsonError.GetJsonStatus(hResult)
        Invoke = hResult
    End If
End Function
