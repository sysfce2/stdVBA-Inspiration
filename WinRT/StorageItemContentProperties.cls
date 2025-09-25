VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageItemContentProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 04/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_File_Properties_StorageItemContentProperties As String = "Windows.Storage.FileProperties.StorageItemContentProperties"
Private Const IID_IStorageItemExtraProperties As String = "{c54361b2-54cd-432b-bdbc-4b19c4b470d7}"
Private Const IID_IStorageItemContentProperties As String = "{05294bad-bc38-48bf-85d7-770e0e2ae0ba}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IStorageItemContentProperties
    IStorageItemContentProperties_GetMusicPropertiesAsync = 6
    IStorageItemContentProperties_GetVideoPropertiesAsync = 7
    IStorageItemContentProperties_GetImagePropertiesAsync = 8
    IStorageItemContentProperties_GetDocumentPropertiesAsync = 9
                        
    ' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7
    IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault = 8

End Enum

' ----==== Vars ====----
Private m_pIStorageItemExtraProperties As Long
Private m_pIStorageItemContentProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
    Call ReleaseIfc(m_pIStorageItemContentProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIStorageItemContentProperties
End Property

Public Property Let Ifc(ByVal pIStorageItemContentProperties As Long)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
    Call ReleaseIfc(m_pIStorageItemContentProperties)
    m_pIStorageItemContentProperties = pIStorageItemContentProperties
    Call QueryIfc(m_pIStorageItemContentProperties, _
                  IID_IStorageItemExtraProperties, _
                  m_pIStorageItemExtraProperties)
End Property

' ----==== Functions ====----
' IStorageItemContentProperties
Public Function GetMusicPropertiesAsync() As MusicProperties
    Dim Ret As MusicProperties
    If m_pIStorageItemContentProperties <> 0& Then
        Dim pIMusicProperties As Long
        If Invoke(m_pIStorageItemContentProperties, _
                  IStorageItemContentProperties_GetMusicPropertiesAsync, _
                  VarPtr(pIMusicProperties)) = S_OK Then
            If pIMusicProperties <> 0& Then
                If Await(pIMusicProperties) Then
                    Set Ret = New MusicProperties
                    Ret.Ifc = pIMusicProperties
                End If
            End If
        End If
    End If
    Set GetMusicPropertiesAsync = Ret
End Function

Public Function GetVideoPropertiesAsync() As VideoProperties
    Dim Ret As VideoProperties
    If m_pIStorageItemContentProperties <> 0& Then
        Dim pIVideoProperties As Long
        If Invoke(m_pIStorageItemContentProperties, _
                  IStorageItemContentProperties_GetVideoPropertiesAsync, _
                  VarPtr(pIVideoProperties)) = S_OK Then
            If pIVideoProperties <> 0& Then
                If Await(pIVideoProperties) Then
                    Set Ret = New VideoProperties
                    Ret.Ifc = pIVideoProperties
                End If
            End If
        End If
    End If
    Set GetVideoPropertiesAsync = Ret
End Function

Public Function GetImagePropertiesAsync() As ImageProperties
    Dim Ret As ImageProperties
    If m_pIStorageItemContentProperties <> 0& Then
        Dim pIImageProperties As Long
        If Invoke(m_pIStorageItemContentProperties, _
                  IStorageItemContentProperties_GetImagePropertiesAsync, _
                  VarPtr(pIImageProperties)) = S_OK Then
            If pIImageProperties <> 0& Then
                If Await(pIImageProperties) Then
                    Set Ret = New ImageProperties
                    Ret.Ifc = pIImageProperties
                End If
            End If
        End If
    End If
    Set GetImagePropertiesAsync = Ret
End Function

Public Function GetDocumentPropertiesAsync() As DocumentProperties
    Dim Ret As DocumentProperties
    If m_pIStorageItemContentProperties <> 0& Then
        Dim pIDocumentProperties As Long
        If Invoke(m_pIStorageItemContentProperties, _
                  IStorageItemContentProperties_GetDocumentPropertiesAsync, _
                  VarPtr(pIDocumentProperties)) = S_OK Then
            If pIDocumentProperties <> 0& Then
                If Await(pIDocumentProperties) Then
                    Set Ret = New DocumentProperties
                    Ret.Ifc = pIDocumentProperties
                End If
            End If
        End If
    End If
    Set GetDocumentPropertiesAsync = Ret
End Function
                        
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


