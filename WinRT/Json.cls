VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Json"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Json

Option Explicit

' ----==== Enums Windows.Data.Json.xxx ====----
Public Enum JsonValueType
    JsonValueType_Null = 0
    JsonValueType_Boolean = 1
    JsonValueType_Number = 2
    JsonValueType_String = 3
    JsonValueType_Array = 4
    JsonValueType_Object = 5
End Enum

Public Enum JsonErrorStatus
    JsonErrorStatus_Unknown = 0
    JsonErrorStatus_InvalidJsonString = 1
    JsonErrorStatus_InvalidJsonNumber = 2
    JsonErrorStatus_JsonValueNotFound = 3
    JsonErrorStatus_ImplementationLimit = 4
End Enum

' ----==== Namespace Json.JsonError ====----
Public Property Get JsonError() As JsonError
    Set JsonError = New JsonError
End Property

' ----==== Namespace Json.JsonArray ====----
Public Property Get JsonArray() As JsonArray
    Set JsonArray = New JsonArray
End Property

' ----==== Namespace Json.JsonValue ====----
Public Property Get JsonValue() As JsonValue
    Set JsonValue = New JsonValue
End Property

' ----==== Namespace Json.JsonObject ====----
Public Property Get JsonObject() As JsonObject
    Set JsonObject = New JsonObject
End Property

