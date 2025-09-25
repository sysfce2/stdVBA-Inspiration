VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastFailedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastFailedEventArgs As String = "Windows.UI.Notifications.ToastFailedEventArgs"
Private Const IID_IToastFailedEventArgs As String = "{35176862-cfd4-44f8-ad64-f500fd896c3b}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastFailedEventArgs
    IToastFailedEventArgs_GetErrorCode = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIToastFailedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastFailedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastFailedEventArgs
End Property

Public Property Let Ifc(ByVal pIToastFailedEventArgs As Long)
    Call ReleaseIfc(m_pIToastFailedEventArgs)
    m_pIToastFailedEventArgs = pIToastFailedEventArgs
End Property

' IToastFailedEventArgs
Public Property Get ErrorCode() As Long
    Dim Ret As Long
    If m_pIToastFailedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastFailedEventArgs, _
                  IToastFailedEventArgs_GetErrorCode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ErrorCode = Ret
End Property
    
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

