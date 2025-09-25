VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppDisplayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_ApplicationModel_AppDisplayInfo As String = "Windows.ApplicationModel.AppDisplayInfo"
Private Const IID_IAppDisplayInfo As String = "{1aeb1103-e4d4-41aa-a4f6-c4a276e79eac}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAppDisplayInfo
    IAppDisplayInfo_GetDisplayName = 6
    IAppDisplayInfo_GetDescription = 7
    IAppDisplayInfo_GetLogo = 8
    
End Enum

' ----==== Variablen ====----
Private m_pIAppDisplayInfo As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAppDisplayInfo)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIAppDisplayInfo As Long)
    Call ReleaseIfc(m_pIAppDisplayInfo)
    m_pIAppDisplayInfo = pIAppDisplayInfo
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIAppDisplayInfo
End Property

' IAppDisplayInfo
Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIAppDisplayInfo <> 0& Then
        Dim Value As Long
        If Invoke(m_pIAppDisplayInfo, _
                  IAppDisplayInfo_GetDisplayName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get Description() As String
    Dim Ret As String
    If m_pIAppDisplayInfo <> 0& Then
        Dim Value As Long
        If Invoke(m_pIAppDisplayInfo, _
                  IAppDisplayInfo_GetDescription, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Description = Ret
End Property

' ----==== Functions ====----
' IAppDisplayInfo
Public Function GetLogo(ByVal logoSize As Size) As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIAppDisplayInfo <> 0& And _
       IsNotNothing(logoSize) Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIAppDisplayInfo, _
                  IAppDisplayInfo_GetLogo, _
                  logoSize.Width, _
                  logoSize.Height, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                Set Ret = New RandomAccessStreamReference
                Ret.Ifc = pIRandomAccessStreamReference
            End If
        End If
    End If
    Set GetLogo = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


