VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoreWindowPopupShowingEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Core_CoreWindowPopupShowingEventArgs As String = "Windows.UI.Core.CoreWindowPopupShowingEventArgs"
Private Const IID_ICoreWindowPopupShowingEventArgs As String = "{26155fa2-5ba5-4ea4-a3b4-2dc7d63c8e26}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICoreWindowPopupShowingEventArgs
    ICoreWindowPopupShowingEventArgs_SetDesiredSize = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pICoreWindowPopupShowingEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICoreWindowPopupShowingEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICoreWindowPopupShowingEventArgs
End Property

Public Property Let Ifc(ByVal pICoreWindowPopupShowingEventArgs As Long)
    Call ReleaseIfc(m_pICoreWindowPopupShowingEventArgs)
    m_pICoreWindowPopupShowingEventArgs = pICoreWindowPopupShowingEventArgs
End Property

' ----==== Functions ====----
Public Function SetDesiredSize(ByVal value As Size) As Boolean
    Dim Ret As Boolean
    If m_pICoreWindowPopupShowingEventArgs <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pICoreWindowPopupShowingEventArgs, _
                  ICoreWindowPopupShowingEventArgs_SetDesiredSize, _
                  value.Width, _
                  value.Height) = S_OK Then
            Ret = True
        End If
    End If
    SetDesiredSize = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


