VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LauncherUIOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_LauncherUIOptions As String = "Windows.System.LauncherUIOptions"
Private Const IID_ILauncherUIOptions As String = "{1b25da6e-8aa6-41e9-8251-4165f5985f49}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ILauncherUIOptions
    ILauncherUIOptions_GetInvocationPoint = 6
'    ILauncherUIOptions_PutInvocationPoint = 7
    ILauncherUIOptions_GetSelectionRect = 8
'    ILauncherUIOptions_PutSelectionRect = 9
    ILauncherUIOptions_GetPreferredPlacement = 10
    ILauncherUIOptions_PutPreferredPlacement = 11
                    
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Types ====----
Private Type Point_
    X As Single
    Y As Single
End Type

Private Type Rect_
    X As Single
    Y As Single
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pILauncherUIOptions As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILauncherUIOptions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pILauncherUIOptions
End Property

Public Property Let Ifc(ByVal pILauncherUIOptions As Long)
    m_pILauncherUIOptions = pILauncherUIOptions
End Property

' ILauncherUIOptions
Public Property Get InvocationPoint() As Point
    Dim Ret As Point
    If m_pILauncherUIOptions <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pILauncherUIOptions, _
                  ILauncherUIOptions_GetInvocationPoint, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Point_
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Set Ret = New Point
                    Ret.X = value.X
                    Ret.Y = value.Y
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Set InvocationPoint = Ret
End Property

'    ILauncherUIOptions_PutInvocationPoint = 7

Public Property Get SelectionRect() As Rect
    Dim Ret As Rect
    If m_pILauncherUIOptions <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pILauncherUIOptions, _
                  ILauncherUIOptions_GetSelectionRect, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Rect_
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Set Ret = New Point
                    Ret.X = value.X
                    Ret.Y = value.Y
                    Ret.Width = value.Width
                    Ret.Height = value.Height
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Set SelectionRect = Ret
End Property

'    ILauncherUIOptions_PutSelectionRect = 9

Public Property Get PreferredPlacement() As Placement
    Dim Ret As Placement
    If m_pILauncherUIOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pILauncherUIOptions, _
                  ILauncherUIOptions_GetPreferredPlacement, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PreferredPlacement = Ret
End Property

Public Property Let PreferredPlacement(ByVal value As Placement)
    If m_pILauncherUIOptions <> 0& Then
        If Invoke(m_pILauncherUIOptions, _
                  ILauncherUIOptions_PutPreferredPlacement, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Function ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


