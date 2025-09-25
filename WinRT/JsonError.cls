VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Json_JsonError As String = "Windows.Data.Json.JsonError"
Private Const IID_IJsonErrorStatics2 As String = "{404030da-87d0-436c-83ab-fc7b12c0cc26}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IJsonErrorStatics2
    IJsonErrorStatics2_GetJsonStatus = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIJsonErrorStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Json_JsonError, _
                            IID_IJsonErrorStatics2, _
                            m_pIJsonErrorStatics2) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIJsonErrorStatics2)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIJsonErrorStatics2
End Property

Public Property Let Ifc(ByVal pIJsonError As Long)
    m_pIJsonErrorStatics2 = pIJsonError
End Property

' ----==== Functions ====----
Public Function GetJsonStatus(ByVal hResult As Long) As JsonErrorStatus
    Dim Ret As JsonErrorStatus
    If m_pIJsonErrorStatics2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIJsonErrorStatics2, _
                  IJsonErrorStatics2_GetJsonStatus, _
                  hResult, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GetJsonStatus = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


