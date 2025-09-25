VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoreWindowFlyout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Core_CoreWindowFlyout As String = "Windows.UI.Core.CoreWindowFlyout"
Private Const IID_ICoreWindowFlyout As String = "{e89d854d-2050-40bb-b344-f6f355eeb314}"
Private Const IID_ICoreWindowFlyoutFactory As String = "{dec4c6c4-93e8-4f7c-be27-cefaa1af68a7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ICoreWindowFlyoutFactory
    ICoreWindowFlyoutFactory_Create = 6
    ICoreWindowFlyoutFactory_CreateWithTitle = 7
    
    ' ICoreWindowFlyout
    ICoreWindowFlyout_AddShowing = 6
    ICoreWindowFlyout_RemoveShowing = 7
    ICoreWindowFlyout_GetMaxSize = 8
    ICoreWindowFlyout_GetMinSize = 9
    ICoreWindowFlyout_GetTitle = 10
    ICoreWindowFlyout_PutTitle = 11
    ICoreWindowFlyout_GetIsInteractionDelayed = 12
    ICoreWindowFlyout_PutIsInteractionDelayed = 13
    ICoreWindowFlyout_GetCommands = 14
    ICoreWindowFlyout_GetDefaultCommandIndex = 15
    ICoreWindowFlyout_PutDefaultCommandIndex = 16
    ICoreWindowFlyout_GetBackButtonCommand = 17
    ICoreWindowFlyout_PutBackButtonCommand = 18
    ICoreWindowFlyout_ShowAsync = 19

End Enum

' ----==== Types ====----
Private Type Windows_Foundation_Size
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pICoreWindowFlyout As Long
Private m_pICoreWindowFlyoutFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Core_CoreWindowFlyout, _
                            IID_ICoreWindowFlyoutFactory, _
                            m_pICoreWindowFlyoutFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICoreWindowFlyout)
    Call ReleaseIfc(m_pICoreWindowFlyoutFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICoreWindowFlyout
End Property

Public Property Let Ifc(ByVal pICoreWindowFlyout As Long)
    Call ReleaseIfc(m_pICoreWindowFlyout)
    m_pICoreWindowFlyout = pICoreWindowFlyout
End Property

Public Property Get MaxSize() As Size
    Dim Ret As Size
    If m_pICoreWindowFlyout <> 0& Then
        Dim value As Windows_Foundation_Size
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_GetMaxSize, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New Size
            Ret.Width = value.Width
            Ret.Height = value.Height
        End If
    End If
    Set MaxSize = Ret
End Property

Public Property Get MinSize() As Size
    Dim Ret As Size
    If m_pICoreWindowFlyout <> 0& Then
        Dim value As Windows_Foundation_Size
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_GetMinSize, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New Size
            Ret.Width = value.Width
            Ret.Height = value.Height
        End If
    End If
    Set MinSize = Ret
End Property

Public Property Get Title() As String
    Dim Ret As String
    If m_pICoreWindowFlyout <> 0& Then
        Dim hString As Long
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_GetTitle, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pICoreWindowFlyout <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get IsInteractionDelayed() As Boolean
    Dim Ret As Boolean
    If m_pICoreWindowFlyout <> 0& Then
        Dim value As Long
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_GetIsInteractionDelayed, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsInteractionDelayed = Ret
End Property

Public Property Let IsInteractionDelayed(ByVal value As Boolean)
    If m_pICoreWindowFlyout <> 0& Then
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_PutIsInteractionDelayed, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Commands() As List_UICommand
    Dim Ret As List_UICommand
    If m_pICoreWindowFlyout <> 0& Then
        Dim pIUICommandList As Long
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_GetCommands, _
                  VarPtr(pIUICommandList)) = S_OK Then
            If pIUICommandList <> 0& Then
                Set Ret = New List_UICommand
                Ret.Ifc = pIUICommandList
            End If
        End If
    End If
    Set Commands = Ret
End Property

Public Property Get DefaultCommandIndex() As Long
    Dim Ret As Long
    If m_pICoreWindowFlyout <> 0& Then
        Dim value As Long
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_GetDefaultCommandIndex, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DefaultCommandIndex = Ret
End Property

Public Property Let DefaultCommandIndex(ByVal value As Long)
    If m_pICoreWindowFlyout <> 0& Then
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_PutDefaultCommandIndex, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get BackButtonCommand() As UICommandInvokedHandler
    Dim Ret As UICommandInvokedHandler
    If m_pICoreWindowFlyout <> 0& Then
        Dim pIUICommandInvokedHandler As Long
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_GetBackButtonCommand, _
                  VarPtr(pIUICommandInvokedHandler)) = S_OK Then
            If pIUICommandInvokedHandler <> 0& Then
                Set Ret = New UICommandInvokedHandler
                Ret.Ifc = pIUICommandInvokedHandler
            End If
        End If
    End If
    Set BackButtonCommand = Ret
End Property

Public Property Let BackButtonCommand(ByVal value As UICommandInvokedHandler)
    If m_pICoreWindowFlyout <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_PutBackButtonCommand, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pICoreWindowFlyout <> 0& Then
        If InitWithWindow(m_pICoreWindowFlyout, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' ICoreWindowFlyoutFactory
Public Function Create(ByVal Position As Point) As CoreWindowFlyout
    Dim Ret As CoreWindowFlyout
    If m_pICoreWindowFlyoutFactory <> 0& And _
       IsNotNothing(Position) Then
        Dim pICoreWindowFlyout As Long
        If Invoke(m_pICoreWindowFlyoutFactory, _
                  ICoreWindowFlyoutFactory_Create, _
                  Position.X, _
                  Position.Y, _
                  VarPtr(pICoreWindowFlyout)) = S_OK Then
            If pICoreWindowFlyout <> 0& Then
                Set Ret = New CoreWindowFlyout
                Ret.Ifc = pICoreWindowFlyout
            End If
        End If
    End If
    Set Create = Ret
End Function

Public Function CreateWithTitle(ByVal Position As Point, _
                                ByVal Title As String) As CoreWindowFlyout
    Dim Ret As CoreWindowFlyout
    If m_pICoreWindowFlyoutFactory <> 0& And _
       IsNotNothing(Position) Then
        Dim hTitle As Long
        hTitle = CreateWindowsString(Title)
        Dim pICoreWindowFlyout As Long
        If Invoke(m_pICoreWindowFlyoutFactory, _
                  ICoreWindowFlyoutFactory_CreateWithTitle, _
                  Position.X, _
                  Position.Y, _
                  hTitle, _
                  VarPtr(pICoreWindowFlyout)) = S_OK Then
            If pICoreWindowFlyout <> 0& Then
                Set Ret = New CoreWindowFlyout
                Ret.Ifc = pICoreWindowFlyout
            End If
        End If
        Call DeleteWindowsString(hTitle)
    End If
    Set CreateWithTitle = Ret
End Function

' ICoreWindowFlyout
Public Function AddShowing(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pICoreWindowFlyout <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_AddShowing, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddShowing = Ret
End Function

Public Function RemoveShowing(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pICoreWindowFlyout <> 0& Then
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_RemoveShowing, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveShowing = Ret
End Function

Public Function ShowAsync() As UICommand
    Dim Ret As UICommand
    If m_pICoreWindowFlyout <> 0& Then
        Dim pIUICommand As Long
        If Invoke(m_pICoreWindowFlyout, _
                  ICoreWindowFlyout_ShowAsync, _
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

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


