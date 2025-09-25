VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Json_JsonValue As String = "Windows.Data.Json.JsonValue"
Private Const IID_IJsonValue As String = "{a3219ecb-f0b3-4dcd-beee-19d48cd3ed1e}"
Private Const IID_IStringable As String = "{96369f54-8eb6-48f0-abce-c1b211e627c3}"
Private Const IID_IJsonValueStatics As String = "{5f6b544a-2f53-48e1-91a3-f78b50a6345c}"
Private Const IID_IJsonValueStatics2 As String = "{1d9ecbe4-3fe8-4335-8392-93d8e36865f0}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IJsonValueStatics
    IJsonValueStatics_Parse = 6
    IJsonValueStatics_TryParse = 7
    IJsonValueStatics_CreateBooleanValue = 8
    IJsonValueStatics_CreateNumberValue = 9
    IJsonValueStatics_CreateStringValue = 10

    ' IJsonValueStatics2
    IJsonValueStatics2_CreateNullValue = 6

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

End Enum

' ----==== Variablen ====----
Private cJsonError As JsonError
Private m_ErrorStatus As JsonErrorStatus
Private m_pIJsonValue As Long
Private m_pIStringable As Long
Private m_pIJsonValueStatics As Long
Private m_pIJsonValueStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Json_JsonValue, _
                            IID_IJsonValueStatics, _
                            m_pIJsonValueStatics) Then
        Call QueryIfc(m_pIJsonValueStatics, _
                      IID_IJsonValueStatics2, _
                      m_pIJsonValueStatics2)
    End If
    Dim pIJsonValue As Long
    If GetActivateInstance(Windows_Data_Json_JsonValue, _
                           IID_IJsonValue, _
                           pIJsonValue) Then
        Me.Ifc = pIJsonValue
    End If
    Set cJsonError = New JsonError
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIJsonValue)
    Call ReleaseIfc(m_pIJsonValueStatics)
    Call ReleaseIfc(m_pIJsonValueStatics2)
    Set cJsonError = Nothing
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIJsonValue
End Property

Public Property Let Ifc(ByVal pIJsonValue As Long)
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIJsonValue)
    m_pIJsonValue = pIJsonValue
    Call QueryIfc(m_pIJsonValue, IID_IStringable, m_pIStringable)
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

' ----==== Functions ====----
' IJsonValueStatics
Public Function Parse(ByVal inputStr As String) As JsonValue
    Dim Ret As JsonValue
    If m_pIJsonValueStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(inputStr)
        Dim pIJsonValue As Long
        If Invoke(m_pIJsonValueStatics, _
                  IJsonValueStatics_Parse, _
                  hString, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue <> 0& Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Parse = Ret
End Function

Public Function TryParse(ByVal inputStr As String, _
                         ByRef result As JsonValue) As Boolean
    Dim Ret As Boolean
    If m_pIJsonValueStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(inputStr)
        Dim pIJsonValue As Long
        Dim succeeded As Long
        If Invoke(m_pIJsonValueStatics, _
                  IJsonValueStatics_TryParse, _
                  hString, _
                  VarPtr(pIJsonValue), _
                  VarPtr(succeeded)) = S_OK Then
            Ret = CBool(succeeded)
            If pIJsonValue <> 0& Then
                Set result = New JsonValue
                result.Ifc = pIJsonValue
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    TryParse = Ret
End Function

Public Function CreateBooleanValue(inputBool As Boolean) As JsonValue
    Dim Ret As JsonValue
    If m_pIJsonValueStatics <> 0& Then
        Dim pIJsonValue As Long
        If Invoke(m_pIJsonValueStatics, _
                  IJsonValueStatics_CreateBooleanValue, _
                  inputBool, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue <> 0& Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
    End If
    Set CreateBooleanValue = Ret
End Function

Public Function CreateNumberValue(inputDbl As Double) As JsonValue
    Dim Ret As JsonValue
    If m_pIJsonValueStatics <> 0& Then
        Dim pIJsonValue As Long
        If Invoke(m_pIJsonValueStatics, _
                  IJsonValueStatics_CreateNumberValue, _
                  inputDbl, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
    End If
    Set CreateNumberValue = Ret
End Function

Public Function CreateStringValue(inputStr As String) As JsonValue
    Dim Ret As JsonValue
    If m_pIJsonValueStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(inputStr)
        Dim pIJsonValue As Long
        If Invoke(m_pIJsonValueStatics, _
                  IJsonValueStatics_CreateStringValue, _
                  hString, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateStringValue = Ret
End Function

' IJsonValueStatics2
Public Function CreateNullValue() As JsonValue
    Dim Ret As JsonValue
    If m_pIJsonValueStatics2 <> 0& Then
        Dim pIJsonValue As Long
        If Invoke(m_pIJsonValueStatics2, _
                  IJsonValueStatics2_CreateNullValue, _
                  VarPtr(pIJsonValue)) = S_OK Then
            If pIJsonValue Then
                Set Ret = New JsonValue
                Ret.Ifc = pIJsonValue
            End If
        End If
    End If
    Set CreateNullValue = Ret
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
        Dim Value As Long
        If Invoke(m_pIJsonValue, _
                  IJsonValue_GetBoolean, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
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
