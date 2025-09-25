VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotificationVisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_NotificationVisual As String = "Windows.UI.Notifications.NotificationVisual"
Private Const IID_INotificationVisual As String = "{68835b8e-aa56-4e11-86d3-5f9a6957bc5b}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' INotificationVisual
    INotificationVisual_GetLanguage = 6
    INotificationVisual_PutLanguage = 7
    INotificationVisual_GetBindings = 8
    INotificationVisual_GetBinding = 9
                        
End Enum

' ----==== Variablen ====----
Private m_pINotificationVisual As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pINotificationVisual)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pINotificationVisual
End Property

Public Property Let Ifc(ByVal pINotificationVisual As Long)
    Call ReleaseIfc(m_pINotificationVisual)
    m_pINotificationVisual = pINotificationVisual
End Property

' INotificationVisual
Public Property Get Language() As String
    Dim Ret As String
    If m_pINotificationVisual <> 0& Then
        Dim hString As Long
        If Invoke(m_pINotificationVisual, _
                  INotificationVisual_GetLanguage, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Language = Ret
End Property

Public Property Let Language(ByVal value As String)
    Dim Ret As String
    If m_pINotificationVisual <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pINotificationVisual, _
                  INotificationVisual_PutLanguage, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Bindings() As ReadOnlyList_1 'ReadOnlyList_NotificationBinding
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_NotificationBinding
    If m_pINotificationVisual <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pINotificationVisual, _
                  INotificationVisual_GetBindings, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_NotificationBinding, pIVectorView) 'ReadOnlyList_NotificationBinding
            End If
        End If
    End If
    Set Bindings = Ret
End Property

' ----==== Functions ====----
' INotificationVisual
Public Function GetBinding(Optional ByVal templateName As String = "ToastGeneric") As NotificationBinding
    Dim Ret As NotificationBinding
    If m_pINotificationVisual <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(templateName)
        Dim pINotificationBinding As Long
        If Invoke(m_pINotificationVisual, _
                  INotificationVisual_GetBinding, _
                  hString, _
                  VarPtr(pINotificationBinding)) = S_OK Then
            If pINotificationBinding <> 0& Then
                Set Ret = New NotificationBinding
                Ret.Ifc = pINotificationBinding
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetBinding = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

