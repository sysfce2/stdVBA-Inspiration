VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Popups_PopupMenu As String = "Windows.UI.Popups.PopupMenu"
Private Const IID_IPopupMenu As String = "{4e9bc6dc-880d-47fc-a0a1-72b639e62559}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPopupMenu
    IPopupMenu_GetCommands = 6
    IPopupMenu_ShowAsync = 7
    IPopupMenu_ShowAsyncWithRect = 8
    IPopupMenu_ShowAsyncWithRectAndPlacement = 9
    
End Enum

' ----==== Variablen ====----
Private m_pIPopupMenu As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_UI_Popups_PopupMenu, _
                           IID_IPopupMenu, _
                           m_pIPopupMenu) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPopupMenu)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPopupMenu
End Property

Public Property Let Ifc(ByVal pIPopupMenu As Long)
    Call ReleaseIfc(m_pIPopupMenu)
    m_pIPopupMenu = pIPopupMenu
End Property

' IPopupMenu
Public Property Get Commands() As List_UICommand
    Dim Ret As List_UICommand
    If m_pIPopupMenu <> 0& Then
        Dim pIUICommandList As Long
        If Invoke(m_pIPopupMenu, _
                  IPopupMenu_GetCommands, _
                  VarPtr(pIUICommandList)) = S_OK Then
            If pIUICommandList <> 0& Then
                Set Ret = New List_UICommand
                Ret.Ifc = pIUICommandList
            End If
        End If
    End If
    Set Commands = Ret
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pIPopupMenu <> 0& Then
        If InitWithWindow(m_pIPopupMenu, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' IPopupMenu
Public Function ShowAsync(ByVal invocationPoint As Point) As UICommand
    Dim Ret As UICommand
    If m_pIPopupMenu <> 0& And _
       IsNotNothing(invocationPoint) Then
        Dim pIUICommand As Long
        If Invoke(m_pIPopupMenu, _
                  IPopupMenu_ShowAsync, _
                  invocationPoint.X, _
                  invocationPoint.Y, _
                  VarPtr(pIUICommand)) = S_OK Then
            If pIUICommand <> 0& Then
                If Await(pIUICommand) Then
                    Set Ret = New UICommand
                    Ret.Ifc = pIUICommand
                End If
            End If
        End If
    End If
    Set ShowAsync = Ret
End Function

Public Function ShowAsyncWithRect(ByVal selection As Rect) As UICommand
    Dim Ret As UICommand
    If m_pIPopupMenu <> 0& And _
       IsNotNothing(selection) Then
        Dim pIUICommand As Long
        If Invoke(m_pIPopupMenu, _
                  IPopupMenu_ShowAsyncWithRect, _
                  selection.X, _
                  selection.Y, _
                  selection.Width, _
                  selection.Height, _
                  VarPtr(pIUICommand)) = S_OK Then
            If pIUICommand <> 0& Then
                If Await(pIUICommand) Then
                    Set Ret = New UICommand
                    Ret.Ifc = pIUICommand
                End If
            End If
        End If
    End If
    Set ShowAsyncWithRect = Ret
End Function

Public Function ShowAsyncWithRectAndPlacement(ByVal selection As Rect, _
                                              ByVal preferredPlacement As Placement) As UICommand
    Dim Ret As UICommand
    If m_pIPopupMenu <> 0& And _
       IsNotNothing(selection) Then
        Dim pIUICommand As Long
        If Invoke(m_pIPopupMenu, _
                  IPopupMenu_ShowAsyncWithRectAndPlacement, _
                  selection.X, _
                  selection.Y, _
                  selection.Width, _
                  selection.Height, _
                  preferredPlacement, _
                  VarPtr(pIUICommand)) = S_OK Then
            If pIUICommand <> 0& Then
                If Await(pIUICommand) Then
                    Set Ret = New UICommand
                    Ret.Ifc = pIUICommand
                End If
            End If
        End If
    End If
    Set ShowAsyncWithRectAndPlacement = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


