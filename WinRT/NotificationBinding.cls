VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotificationBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_NotificationBinding As String = "Windows.UI.Notifications.NotificationBinding"
Private Const IID_INotificationBinding As String = "{f29e4b85-0370-4ad3-b4ea-da9e35e7eabf}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' INotificationBinding
    INotificationBinding_GetTemplate = 6
    INotificationBinding_PutTemplate = 7
    INotificationBinding_GetLanguage = 8
    INotificationBinding_PutLanguage = 9
    INotificationBinding_GetHints = 10
    INotificationBinding_GetTextElements = 11
                        
End Enum

' ----==== Variablen ====----
Private m_pINotificationBinding As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pINotificationBinding)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pINotificationBinding
End Property

Public Property Let Ifc(ByVal pINotificationBinding As Long)
    Call ReleaseIfc(m_pINotificationBinding)
    m_pINotificationBinding = pINotificationBinding
End Property

' INotificationBinding
Public Property Get Template() As String
    Dim Ret As String
    If m_pINotificationBinding <> 0& Then
        Dim hString As Long
        If Invoke(m_pINotificationBinding, _
                  INotificationBinding_GetTemplate, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Template = Ret
End Property

Public Property Let Template(ByVal value As String)
    If m_pINotificationBinding <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pINotificationBinding, _
                  INotificationBinding_PutTemplate, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Language() As String
    Dim Ret As String
    If m_pINotificationBinding <> 0& Then
        Dim hString As Long
        If Invoke(m_pINotificationBinding, _
                  INotificationBinding_GetLanguage, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Language = Ret
End Property

Public Property Let Language(ByVal value As String)
    If m_pINotificationBinding <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pINotificationBinding, _
                  INotificationBinding_PutLanguage, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Hints() As List_String_String
    Dim Ret As List_String_String
    If m_pINotificationBinding <> 0& Then
        Dim pIMap_HSTRING_HSTRING As Long
        If Invoke(m_pINotificationBinding, _
                  INotificationBinding_GetHints, _
                  VarPtr(pIMap_HSTRING_HSTRING)) = S_OK Then
            If pIMap_HSTRING_HSTRING <> 0& Then
                Set Ret = New List_String_String
                Ret.Ifc = pIMap_HSTRING_HSTRING
            End If
        End If
    End If
    Set Hints = Ret
End Property

' ----==== Functions ====----
' INotificationBinding
Public Function GetTextElements() As ReadOnlyList_1 'ReadOnlyList_AdaptiveNotificationText
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AdaptiveNotificationText
    If m_pINotificationBinding <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pINotificationBinding, _
                  INotificationBinding_GetTextElements, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_AdaptiveNotificationText, pIVectorView) 'ReadOnlyList_AdaptiveNotificationText
            End If
        End If
    End If
    Set GetTextElements = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


