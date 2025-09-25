VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Json_JsonObject As String = "Windows.Data.Json.JsonObject"
Private Const IID_IJsonValue As String = "{a3219ecb-f0b3-4dcd-beee-19d48cd3ed1e}"
Private Const IID_IJsonObject As String = "{064e24dd-29c2-4f83-9ac1-9ee11578beb3}"
Private Const IID_IStringable As String = "{96369f54-8eb6-48f0-abce-c1b211e627c3}"
Private Const IID_IJsonObjectStatics As String = "{2289f159-54de-45d8-abcc-22603fa066a0}"
Private Const IID_IJsonObjectWithDefaultValues As String = "{d960d2a2-b7f0-4f00-8e44-d82cf415ea13}"
Private Const IID_IMap_hString_IJsonValue As String = "{c9d9a725-786b-5113-b4b7-9b61764c220b}"
Private Const IID_IIterable_IKeyValuePair_hString_IJsonValue As String = "{dfabb6e1-0411-5a8f-aa87-354e7110f099}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IJsonObjectStatics
    IJsonObjectStatics_Parse = 6
    IJsonObjectStatics_TryParse = 7
    
    ' IJsonObject
    IJsonObject_GetNamedValue = 6
    IJsonObject_SetNamedValue = 7
    IJsonObject_GetNamedObject = 8
    IJsonObject_GetNamedArray = 9
    IJsonObject_GetNamedString = 10
    IJsonObject_GetNamedNumber = 11
    IJsonObject_GetNamedBoolean = 12
    
    ' IJsonObjectWithDefaultValues
    IJsonObjectWithDefaultValues_GetNamedValueOrDefault = 6
    IJsonObjectWithDefaultValues_GetNamedObjectOrDefault = 7
    IJsonObjectWithDefaultValues_GetNamedStringOrDefault = 8
    IJsonObjectWithDefaultValues_GetNamedArrayOrDefault = 9
    IJsonObjectWithDefaultValues_GetNamedNumberOrDefault = 10
    IJsonObjectWithDefaultValues_GetNamedBooleanOrDefault = 11

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
    
    ' IMap_hString_IJsonValue
    IMap_hString_IJsonValue_Lookup = 6
    IMap_hString_IJsonValue_GetSize = 7
    IMap_hString_IJsonValue_HasKey = 8
    IMap_hString_IJsonValue_GetView = 9
    IMap_hString_IJsonValue_Insert = 10
    IMap_hString_IJsonValue_Remove = 11
    IMap_hString_IJsonValue_Clear = 12
    
    ' IIterable_IKeyValuePair_hString_IJsonValue
    IIterable_IKeyValuePair_hString_IJsonValue_First = 6
    
    ' IIterator_IKeyValuePair_hString_IJsonValue
    IIterator_IKeyValuePair_hString_IJsonValue_GetCurrent = 6
    IIterator_IKeyValuePair_hString_IJsonValue_GetHasCurrent = 7
    IIterator_IKeyValuePair_hString_IJsonValue_MoveNext = 8
    IIterator_IKeyValuePair_hString_IJsonValue_GetMany = 9

    ' IKeyValuePair_hString_IJsonValue
    IKeyValuePair_hString_IJsonValue_GetKey = 6
    IKeyValuePair_hString_IJsonValue_GetValue = 7
    
End Enum

' ----==== Variablen ====----
Private cJsonError As JsonError
Private m_ErrorStatus As JsonErrorStatus
Private m_pIJsonValue As Long
Private m_pIStringable As Long
Private m_pIJsonObject As Long
Private m_pIJsonObjectStatics As Long
Private m_pIJsonObjectWithDefaultValues As Long
Private m_pIMap_hString_IJsonValue As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Json_JsonObject, _
                            IID_IJsonObjectStatics, _
                            m_pIJsonObjectStatics) Then
    End If
    Dim pIJsonObject As Long
    If GetActivateInstance(Windows_Data_Json_JsonObject, _
                           IID_IJsonObject, _
                           pIJsonObject) Then
        Me.Ifc = pIJsonObject
    End If
    Set cJsonError = New JsonError
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIJsonValue)
    Call ReleaseIfc(m_pIJsonObject)
    Call ReleaseIfc(m_pIJsonObjectStatics)
    Call ReleaseIfc(m_pIJsonObjectWithDefaultValues)
    Call ReleaseIfc(m_pIMap_hString_IJsonValue)
    Set cJsonError = Nothing
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIJsonObject
End Property

Public Property Let Ifc(ByVal pIJsonObject As Long)
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIJsonValue)
    Call ReleaseIfc(m_pIJsonObject)
    Call ReleaseIfc(m_pIMap_hString_IJsonValue)
    m_pIJsonObject = pIJsonObject
    Call QueryIfc(m_pIJsonObject, IID_IJsonValue, m_pIJsonValue)
    Call QueryIfc(m_pIJsonObject, IID_IStringable, m_pIStringable)
    Call QueryIfc(m_pIJsonObject, IID_IJsonObjectWithDefaultValues, m_pIJsonObjectWithDefaultValues)
    Call QueryIfc(m_pIJsonObject, IID_IMap_hString_IJsonValue, m_pIMap_hString_IJsonValue)
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

' IMap_hString_IJsonValue
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIMap_hString_IJsonValue <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMap_hString_IJsonValue, _
                  IMap_hString_IJsonValue_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
' IJsonObjectStatics
Public Function Parse(ByVal inputStr As String) As JsonObject
    Dim Ret As JsonObject
    If m_pIJsonObjectStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(inputStr)
        Dim pIJsonObject As Long
        If Invoke(m_pIJsonObjectStatics, _
                  IJsonObjectStatics_Parse, _
                  hString, _
                  VarPtr(pIJsonObject)) = S_OK Then
            If pIJsonObject <> 0& Then
                Set Ret = New JsonObject
                Ret.Ifc = pIJsonObject
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Parse = Ret
End Function

Public Function TryParse(ByVal inputStr As String, _
                         ByRef result As JsonObject) As Boolean
    Dim Ret As Boolean
    If m_pIJsonObjectStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(inputStr)
        Dim pIJsonObject As Long
        Dim succeeded As Boolean
        If Invoke(m_pIJsonObjectStatics, _
                  IJsonObjectStatics_TryParse, _
                  hString, _
                  VarPtr(pIJsonObject), _
                  VarPtr(succeeded)) = S_OK Then
            Ret = succeeded
            If pIJsonObject <> 0& Then
                Set result = New JsonObject
                result.Ifc = pIJsonObject
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    TryParse = Ret
End Function

' IJsonObject
Public Function GetNamedValue(ByVal Name As String) As JsonValue
    Dim Ret As JsonValue
    If m_pIJsonObject <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIJsonValue As Long
        If Invoke(m_pIJsonObject, _
                  IJsonObject_GetNamedValue, _
                  hString, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue <> 0& Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetNamedValue = Ret
End Function

Public Function SetNamedValue(ByVal Name As String, _
                              ByVal Value As Object) As Boolean
    Dim Ret As Boolean
    If m_pIJsonObject <> 0& And _
       IsNotNothing(Value) Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        If hString <> 0& Then
            If Invoke(m_pIJsonObject, _
                      IJsonObject_SetNamedValue, _
                      hString, _
                      Value.Ifc) = S_OK Then
                Ret = True
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
    SetNamedValue = Ret
End Function

Public Function GetNamedObject(ByVal Name As String) As JsonObject
    Dim Ret As JsonObject
    If m_pIJsonObject <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIJsonObject As Long
        If Invoke(m_pIJsonObject, _
                  IJsonObject_GetNamedObject, _
                  hString, _
                  VarPtr(pIJsonObject)) = S_OK Then
            If pIJsonObject <> 0& Then
                Set Ret = New JsonObject
                Ret.Ifc = pIJsonObject
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetNamedObject = Ret
End Function

Public Function GetNamedArray(ByVal Name As String) As JsonArray
    Dim Ret As JsonArray
    If m_pIJsonObject <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIJsonArray As Long
        If Invoke(m_pIJsonObject, _
                  IJsonObject_GetNamedArray, _
                  hString, _
                  VarPtr(pIJsonArray)) = S_OK Then
            If pIJsonArray <> 0& Then
                Set Ret = New JsonArray
                Ret.Ifc = pIJsonArray
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetNamedArray = Ret
End Function

Public Function GetNamedString(ByVal Name As String) As String
    Dim Ret As String
    If m_pIJsonObject <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        If hString <> 0& Then
            Dim Value As Long
            If Invoke(m_pIJsonObject, _
                      IJsonObject_GetNamedString, _
                      hString, _
                      VarPtr(Value)) = S_OK Then
                Ret = GetWindowsString(Value)
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
    GetNamedString = Ret
End Function

Public Function GetNamedNumber(ByVal Name As String) As Double
    Dim Ret As Double
    If m_pIJsonObject <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        If hString <> 0& Then
            Dim Value As Double
            If Invoke(m_pIJsonObject, _
                      IJsonObject_GetNamedNumber, _
                      hString, _
                      VarPtr(Value)) = S_OK Then
                Ret = Value
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
    GetNamedNumber = Ret
End Function

Public Function GetNamedBoolean(ByVal Name As String) As Boolean
    Dim Ret As Boolean
    If m_pIJsonObject <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        If hString <> 0& Then
            Dim Value As Boolean
            If Invoke(m_pIJsonObject, _
                      IJsonObject_GetNamedBoolean, _
                      hString, _
                      VarPtr(Value)) = S_OK Then
                Ret = Value
            End If
            Call DeleteWindowsString(hString)
        End If
    End If
    GetNamedBoolean = Ret
End Function

' IJsonObjectWithDefaultValues
Public Function GetNamedValueOrDefault(ByVal Name As String, _
                                       ByVal defaultValue As JsonValue) As JsonValue
    Dim Ret As JsonValue
    If m_pIJsonObjectWithDefaultValues <> 0& And _
       IsNotNothing(defaultValue) Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIJsonValue As Long
        If Invoke(m_pIJsonObjectWithDefaultValues, _
                  IJsonObjectWithDefaultValues_GetNamedValueOrDefault, _
                  hString, _
                  defaultValue.Ifc, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue <> 0& Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetNamedValueOrDefault = Ret
End Function

Public Function GetNamedObjectOrDefault(ByVal Name As String, _
                                        ByVal defaultValue As JsonObject) As JsonObject
    Dim Ret As JsonObject
    If m_pIJsonObjectWithDefaultValues <> 0& And _
       IsNotNothing(defaultValue) Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIJsonObject As Long
        If Invoke(m_pIJsonObjectWithDefaultValues, _
                  IJsonObjectWithDefaultValues_GetNamedObjectOrDefault, _
                  hString, _
                  defaultValue.Ifc, _
                  VarPtr(pIJsonObject)) = S_OK Then
            If pIJsonObject <> 0& Then
                Set Ret = New JsonObject
                Ret.Ifc = pIJsonObject
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetNamedObjectOrDefault = Ret
End Function

Public Function GetNamedStringOrDefault(ByVal Name As String, _
                                        ByVal defaultValue As String) As String
    Dim Ret As String
    If m_pIJsonObjectWithDefaultValues <> 0& Then
        Dim hName As Long
        Dim hValue As Long
        hName = CreateWindowsString(Name)
        hValue = CreateWindowsString(defaultValue)
        Dim Value As Long
        If Invoke(m_pIJsonObjectWithDefaultValues, _
                  IJsonObjectWithDefaultValues_GetNamedStringOrDefault, _
                  hName, _
                  hValue, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
        Call DeleteWindowsString(hValue)
        Call DeleteWindowsString(hName)
    End If
    GetNamedStringOrDefault = Ret
End Function

Public Function GetNamedArrayOrDefault(ByVal Name As String, _
                                       ByVal defaultValue As JsonArray) As JsonArray
    Dim Ret As JsonArray
    If m_pIJsonObjectWithDefaultValues <> 0& And _
       IsNotNothing(defaultValue) Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIJsonArray As Long
        If Invoke(m_pIJsonObjectWithDefaultValues, _
                  IJsonObjectWithDefaultValues_GetNamedArrayOrDefault, _
                  hString, _
                  defaultValue.Ifc, _
                  VarPtr(pIJsonArray)) = S_OK Then
            If pIJsonArray <> 0& Then
                Set Ret = New JsonArray
                Ret.Ifc = pIJsonArray
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetNamedArrayOrDefault = Ret
End Function

Public Function GetNamedNumberOrDefault(ByVal Name As String, _
                                        ByVal defaultValue As Double) As Double
    Dim Ret As Double
    If m_pIJsonObjectWithDefaultValues <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim Value As Double
        If Invoke(m_pIJsonObjectWithDefaultValues, _
                  IJsonObjectWithDefaultValues_GetNamedNumberOrDefault, _
                  hString, _
                  defaultValue, _
                  VarPtr(Value)) = S_OK Then
                Ret = Value
        End If
        Call DeleteWindowsString(hString)
    End If
    GetNamedNumberOrDefault = Ret
End Function

Public Function GetNamedBooleanOrDefault(ByVal Name As String, _
                                         ByVal defaultValue As Boolean) As Boolean
    Dim Ret As Boolean
    If m_pIJsonObjectWithDefaultValues <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim Value As Long
        If Invoke(m_pIJsonObjectWithDefaultValues, _
                  IJsonObjectWithDefaultValues_GetNamedBooleanOrDefault, _
                  hString, _
                  defaultValue, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
        Call DeleteWindowsString(hString)
    End If
    GetNamedBooleanOrDefault = Ret
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

' IMap_hString_IJsonValue
Public Function Lookup(ByVal Key As String) As JsonValue
    Dim Ret As JsonValue
    If m_pIMap_hString_IJsonValue <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim pIJsonValue As Long
        If Invoke(m_pIMap_hString_IJsonValue, _
                  IMap_hString_IJsonValue_Lookup, _
                  hString, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue <> 0& Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Lookup = Ret
End Function

Public Function HasKey(ByVal Key As String) As Boolean
    Dim Ret As Boolean
    If m_pIMap_hString_IJsonValue <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        Dim Value As Long
        If Invoke(m_pIMap_hString_IJsonValue, _
                  IMap_hString_IJsonValue_HasKey, _
                  hString, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
        Call DeleteWindowsString(hString)
    End If
    HasKey = Ret
End Function

Public Function Insert(ByVal Key As String, _
                       ByVal Value As JsonValue, _
                       Optional ByRef replaced As Boolean) As Boolean
    Dim Ret As Boolean
    If m_pIMap_hString_IJsonValue <> 0& And _
       IsNotNothing(Value) Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        If Invoke(m_pIMap_hString_IJsonValue, _
                  IMap_hString_IJsonValue_Insert, _
                  hString, _
                  Value.Ifc, _
                  VarPtr(replaced)) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    Insert = Ret
End Function

Public Function Remove(ByVal Key As String) As Boolean
    Dim Ret As Boolean
    If m_pIMap_hString_IJsonValue <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Key)
        If Invoke(m_pIMap_hString_IJsonValue, _
                  IMap_hString_IJsonValue_Remove, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    Remove = Ret
End Function

Public Function Clear() As Boolean
    Dim Ret As Boolean
    If m_pIMap_hString_IJsonValue <> 0& Then
        If Invoke(m_pIMap_hString_IJsonValue, _
                  IMap_hString_IJsonValue_Clear) = S_OK Then
            Ret = True
        End If
    End If
    Clear = Ret
End Function

' IIterable_IKeyValuePair_hString_IJsonValue
' IIterator_IKeyValuePair_hString_IJsonValue
' IKeyValuePair_hString_IJsonValue
Public Function GetKeyValuePairs() As KeyValuePair_String_JsonValue()
    Dim Ret() As KeyValuePair_String_JsonValue
    If m_pIJsonObject <> 0& Then
        Dim count As Long
        count = Me.Size
        If count > 0& Then
            ReDim Ret(count - 1)
            Dim pIIterable_IKeyValuePair_hString_IJsonValue As Long
            If QueryIfc(m_pIJsonObject, _
                        IID_IIterable_IKeyValuePair_hString_IJsonValue, _
                        pIIterable_IKeyValuePair_hString_IJsonValue) Then
                Dim pIIterator_IKeyValuePair_hString_IJsonValue As Long
                If Invoke(pIIterable_IKeyValuePair_hString_IJsonValue, _
                          IIterable_IKeyValuePair_hString_IJsonValue_First, _
                          VarPtr(pIIterator_IKeyValuePair_hString_IJsonValue)) = S_OK Then
                    Dim Item As Long
                    For Item = 0 To count - 1
                        Dim Value As Long
                        Dim pIKeyValuePair_hString_IJsonValue As Long
                        Dim JsonValue As JsonValue
                        Set Ret(Item) = New KeyValuePair_String_JsonValue
                        If Item = 0 Then
                            If Invoke(pIIterator_IKeyValuePair_hString_IJsonValue, _
                                      IIterator_IKeyValuePair_hString_IJsonValue_GetCurrent, _
                                      VarPtr(pIKeyValuePair_hString_IJsonValue)) = S_OK Then
                                If Invoke(pIKeyValuePair_hString_IJsonValue, _
                                          IKeyValuePair_hString_IJsonValue_GetKey, _
                                          VarPtr(Value)) = S_OK Then
                                    If Value <> 0& Then
                                        Ret(Item).Key = GetWindowsString(Value)
                                    End If
                                End If
                                If Invoke(pIKeyValuePair_hString_IJsonValue, _
                                          IKeyValuePair_hString_IJsonValue_GetValue, _
                                          VarPtr(Value)) = S_OK Then
                                    If Value <> 0& Then
                                        Set JsonValue = New JsonValue
                                        JsonValue.Ifc = Value
                                        Ret(Item).Value = JsonValue
                                    End If
                                End If
                                Call ReleaseIfc(pIKeyValuePair_hString_IJsonValue)
                            End If
                        Else
                            Dim moveNext As Long
                            If Invoke(pIIterator_IKeyValuePair_hString_IJsonValue, _
                                      IIterator_IKeyValuePair_hString_IJsonValue_MoveNext, _
                                      VarPtr(moveNext)) = S_OK Then
                                If CBool(moveNext) Then
                                    If Invoke(pIIterator_IKeyValuePair_hString_IJsonValue, _
                                              IIterator_IKeyValuePair_hString_IJsonValue_GetCurrent, _
                                              VarPtr(pIKeyValuePair_hString_IJsonValue)) = S_OK Then
                                        If Invoke(pIKeyValuePair_hString_IJsonValue, _
                                                  IKeyValuePair_hString_IJsonValue_GetKey, _
                                                  VarPtr(Value)) = S_OK Then
                                            If Value <> 0& Then
                                                Ret(Item).Key = GetWindowsString(Value)
                                            End If
                                        End If
                                        If Invoke(pIKeyValuePair_hString_IJsonValue, _
                                                  IKeyValuePair_hString_IJsonValue_GetValue, _
                                                  VarPtr(Value)) = S_OK Then
                                            If Value <> 0& Then
                                                Set JsonValue = New JsonValue
                                                JsonValue.Ifc = Value
                                                Ret(Item).Value = JsonValue
                                            End If
                                        End If
                                        Call ReleaseIfc(pIKeyValuePair_hString_IJsonValue)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    Call ReleaseIfc(pIIterator_IKeyValuePair_hString_IJsonValue)
                End If
                Call ReleaseIfc(pIIterable_IKeyValuePair_hString_IJsonValue)
            End If
        End If
    End If
    GetKeyValuePairs = Ret
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

