VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdaptiveNotificationText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 03/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_AdaptiveNotificationText As String = "Windows.UI.Notifications.AdaptiveNotificationText"
Private Const IID_IAdaptiveNotificationText As String = "{46d4a3be-609a-4326-a40b-bfde872034a3}"
Private Const IID_IAdaptiveNotificationContent As String = "{eb0dbe66-7448-448d-9db8-d78acd2abba9}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IAdaptiveNotificationText
    IAdaptiveNotificationText_GetText = 6
    IAdaptiveNotificationText_PutText = 7
    IAdaptiveNotificationText_GetLanguage = 8
    IAdaptiveNotificationText_PutLanguage = 9
                        
    ' IAdaptiveNotificationContent
    IAdaptiveNotificationContent_GetKind = 6
    IAdaptiveNotificationContent_GetHints = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIAdaptiveNotificationText As Long
Private m_pIAdaptiveNotificationContent As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAdaptiveNotificationText)
    Call ReleaseIfc(m_pIAdaptiveNotificationContent)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIAdaptiveNotificationText
End Property

Public Property Let Ifc(ByVal pIAdaptiveNotificationText As Long)
    Call ReleaseIfc(m_pIAdaptiveNotificationText)
    Call ReleaseIfc(m_pIAdaptiveNotificationContent)
    m_pIAdaptiveNotificationText = pIAdaptiveNotificationText
    Call QueryIfc(m_pIAdaptiveNotificationText, _
                  IID_IAdaptiveNotificationContent, _
                  m_pIAdaptiveNotificationContent)
End Property

' IAdaptiveNotificationText
Public Property Get Text() As String
    Dim Ret As String
    If m_pIAdaptiveNotificationText <> 0& Then
        Dim hString As Long
        If Invoke(m_pIAdaptiveNotificationText, _
                  IAdaptiveNotificationText_GetText, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Text = Ret
End Property

Public Property Let Text(ByVal Value As String)
    If m_pIAdaptiveNotificationText <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIAdaptiveNotificationText, _
                  IAdaptiveNotificationText_PutText, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get language() As String
    Dim Ret As String
    If m_pIAdaptiveNotificationText <> 0& Then
        Dim hString As Long
        If Invoke(m_pIAdaptiveNotificationText, _
                  IAdaptiveNotificationText_GetLanguage, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    language = Ret
End Property

Public Property Let language(ByVal Value As String)
    If m_pIAdaptiveNotificationText <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIAdaptiveNotificationText, _
                  IAdaptiveNotificationText_PutLanguage, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property
                        
' IAdaptiveNotificationContent
Public Property Get Kind() As AdaptiveNotificationContentKind
    Dim Ret As AdaptiveNotificationContentKind
    If m_pIAdaptiveNotificationContent <> 0& Then
        Dim Value As Long
        If Invoke(m_pIAdaptiveNotificationContent, _
                  IAdaptiveNotificationContent_GetKind, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Kind = Ret
End Property

Public Property Get Hints() As List_String_String
    Dim Ret As List_String_String
    If m_pIAdaptiveNotificationContent <> 0& Then
        Dim pIMap_HSTRING_HSTRING As Long
        If Invoke(m_pIAdaptiveNotificationContent, _
                  IAdaptiveNotificationContent_GetHints, _
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

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


