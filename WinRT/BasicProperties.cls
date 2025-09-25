VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasicProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 04/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_FileProperties_BasicProperties As String = "Windows.Storage.FileProperties.BasicProperties"
Private Const IID_IBasicProperties As String = "{d05d55db-785e-4a66-be02-9beec58aea81}"
Private Const IID_IStorageItemExtraProperties As String = "{c54361b2-54cd-432b-bdbc-4b19c4b470d7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IBasicProperties
    IBasicProperties_GetSize = 6
    IBasicProperties_GetDateModified = 7
    IBasicProperties_GetItemDate = 8
    
    ' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7
    IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault = 8

End Enum

' ----==== Vars ====----
Private m_pIBasicProperties As Long
Private m_pIStorageItemExtraProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBasicProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBasicProperties
End Property

Public Property Let Ifc(ByVal pIBasicProperties As Long)
    Call ReleaseIfc(m_pIBasicProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
    m_pIBasicProperties = pIBasicProperties
    Call QueryIfc(m_pIBasicProperties, _
                  IID_IStorageItemExtraProperties, _
                  m_pIStorageItemExtraProperties)
End Property

' IBasicProperties
Public Property Get Size() As Currency
    Dim Ret As Currency
    If m_pIBasicProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIBasicProperties, _
                  IBasicProperties_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value * 10000
        End If
    End If
    Size = Ret
End Property

Public Property Get DateModified() As DateTime
    Dim Ret As DateTime
    If m_pIBasicProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIBasicProperties, _
                  IBasicProperties_GetDateModified, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set DateModified = Ret
End Property

Public Property Get ItemDate() As DateTime
    Dim Ret As DateTime
    If m_pIBasicProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIBasicProperties, _
                  IBasicProperties_GetItemDate, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set ItemDate = Ret
End Property

' ----==== Functions ====----
' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7

Public Function SavePropertiesAsyncOverloadDefault() As Boolean
    Dim Ret As Boolean
    If m_pIStorageItemExtraProperties <> 0& Then
        Dim operation As Long
        If Invoke(m_pIStorageItemExtraProperties, _
                  IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    SavePropertiesAsyncOverloadDefault = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


