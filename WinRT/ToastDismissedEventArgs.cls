VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastDismissedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastDismissedEventArgs As String = "Windows.UI.Notifications.ToastDismissedEventArgs"
Private Const IID_IToastDismissedEventArgs As String = "{3f89d935-d9cb-4538-a0f0-ffe7659938f8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastDismissedEventArgs
    IToastDismissedEventArgs_GetReason = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIToastDismissedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastDismissedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastDismissedEventArgs
End Property

Public Property Let Ifc(ByVal pIToastDismissedEventArgs As Long)
    Call ReleaseIfc(m_pIToastDismissedEventArgs)
    m_pIToastDismissedEventArgs = pIToastDismissedEventArgs
End Property

' IToastDismissedEventArgs
Public Property Get Reason() As ToastDismissalReason
    Dim Ret As ToastDismissalReason
    If m_pIToastDismissedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastDismissedEventArgs, _
                  IToastDismissedEventArgs_GetReason, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Reason = Ret
End Property
    
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


