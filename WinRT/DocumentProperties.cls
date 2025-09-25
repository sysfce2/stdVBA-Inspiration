VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_FileProperties_DocumentProperties As String = "Windows.Storage.FileProperties.DocumentProperties"
Private Const IID_IDocumentProperties As String = "{7eab19bc-1821-4923-b4a9-0aea404d0070}"
Private Const IID_IStorageItemExtraProperties As String = "{c54361b2-54cd-432b-bdbc-4b19c4b470d7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IDocumentProperties
    IDocumentProperties_GetAuthor = 6
    IDocumentProperties_GetTitle = 7
    IDocumentProperties_PutTitle = 8
    IDocumentProperties_GetKeywords = 9
    IDocumentProperties_GetComment = 10
    IDocumentProperties_PutComment = 11

    ' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7
    IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault = 8

End Enum

' ----==== Vars ====----
Private m_pIDocumentProperties As Long
Private m_pIStorageItemExtraProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDocumentProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDocumentProperties
End Property

Public Property Let Ifc(ByVal pIDocumentProperties As Long)
    Call ReleaseIfc(m_pIDocumentProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
    m_pIDocumentProperties = pIDocumentProperties
    Call QueryIfc(m_pIDocumentProperties, _
                  IID_IStorageItemExtraProperties, _
                  m_pIStorageItemExtraProperties)
End Property

' IDocumentProperties
Public Property Get Author() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIDocumentProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIDocumentProperties, _
                  IDocumentProperties_GetAuthor, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Author = Ret
End Property

Public Property Get Title() As String
    Dim Ret As String
    If m_pIDocumentProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIDocumentProperties, _
                  IDocumentProperties_GetTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    Dim Ret As String
    If m_pIDocumentProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIDocumentProperties, _
                  IDocumentProperties_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Keywords() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIDocumentProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIDocumentProperties, _
                  IDocumentProperties_GetKeywords, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Keywords = Ret
End Property

Public Property Get Comment() As String
    Dim Ret As String
    If m_pIDocumentProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIDocumentProperties, _
                  IDocumentProperties_GetComment, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Comment = Ret
End Property

Public Property Let Comment(ByVal value As String)
    Dim Ret As String
    If m_pIDocumentProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIDocumentProperties, _
                  IDocumentProperties_PutComment, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
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


