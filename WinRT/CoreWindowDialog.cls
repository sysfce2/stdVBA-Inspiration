VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoreWindowDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 03/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Core_CoreWindowDialog As String = "Windows.UI.Core.CoreWindowDialog"
Private Const IID_ICoreWindowDialog As String = "{e7392ce0-c78d-427e-8b2c-01ff420c69d5}"
Private Const IID_ICoreWindowDialogFactory As String = "{cfb2a855-1c59-4b13-b1e5-16e29805f7c4}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ICoreWindowDialogFactory
    ICoreWindowDialogFactory_CreateWithTitle = 6
    
    ' ICoreWindowDialog
    ICoreWindowDialog_AddShowing = 6
    ICoreWindowDialog_RemoveShowing = 7
    ICoreWindowDialog_GetMaxSize = 8
    ICoreWindowDialog_GetMinSize = 9
    ICoreWindowDialog_GetTitle = 10
    ICoreWindowDialog_PutTitle = 11
    ICoreWindowDialog_GetIsInteractionDelayed = 12
    ICoreWindowDialog_PutIsInteractionDelayed = 13
    ICoreWindowDialog_GetCommands = 14
    ICoreWindowDialog_GetDefaultCommandIndex = 15
    ICoreWindowDialog_PutDefaultCommandIndex = 16
    ICoreWindowDialog_GetCancelCommandIndex = 17
    ICoreWindowDialog_PutCancelCommandIndex = 18
    ICoreWindowDialog_GetBackButtonCommand = 19
    ICoreWindowDialog_PutBackButtonCommand = 20
    ICoreWindowDialog_ShowAsync = 21

End Enum

' ----==== Types ====----
Private Type Windows_Foundation_Size
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pICoreWindowDialog As Long
Private m_pICoreWindowDialogFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Core_CoreWindowDialog, _
                            IID_ICoreWindowDialogFactory, _
                            m_pICoreWindowDialogFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICoreWindowDialog)
    Call ReleaseIfc(m_pICoreWindowDialogFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICoreWindowDialog
End Property

Public Property Let Ifc(ByVal pICoreWindowDialog As Long)
    Call ReleaseIfc(m_pICoreWindowDialog)
    m_pICoreWindowDialog = pICoreWindowDialog
End Property

' ICoreWindowDialog
Public Property Get MaxSize() As Size
    Dim Ret As Size
    If m_pICoreWindowDialog <> 0& Then
        Dim value As Windows_Foundation_Size
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetMaxSize, _
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
    If m_pICoreWindowDialog <> 0& Then
        Dim value As Windows_Foundation_Size
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetMinSize, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New Size
            Ret.Width = value.Width
            Ret.Height = value.Height
        End If
    End If
    MinSize = Ret
End Property

Public Property Get Title() As String
    Dim Ret As String
    If m_pICoreWindowDialog <> 0& Then
        Dim hString As Long
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetTitle, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pICoreWindowDialog <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get IsInteractionDelayed() As Boolean
    Dim Ret As Boolean
    If m_pICoreWindowDialog <> 0& Then
        Dim value As Long
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetIsInteractionDelayed, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsInteractionDelayed = Ret
End Property

Public Property Let IsInteractionDelayed(ByVal value As Boolean)
    If m_pICoreWindowDialog <> 0& Then
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_PutIsInteractionDelayed, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Commands() As List_UICommand
    Dim Ret As List_UICommand
    If m_pICoreWindowDialog <> 0& Then
        Dim pIUICommandList As Long
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetCommands, _
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
    If m_pICoreWindowDialog <> 0& Then
        Dim value As Long
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetDefaultCommandIndex, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DefaultCommandIndex = Ret
End Property

Public Property Let DefaultCommandIndex(ByVal value As Long)
    If m_pICoreWindowDialog <> 0& Then
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_PutDefaultCommandIndex, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get CancelCommandIndex() As Long
    Dim Ret As Long
    If m_pICoreWindowDialog <> 0& Then
        Dim value As Long
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetCancelCommandIndex, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    CancelCommandIndex = Ret
End Property

Public Property Let CancelCommandIndex(ByVal value As Long)
    If m_pICoreWindowDialog <> 0& Then
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_PutCancelCommandIndex, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get BackButtonCommand() As UICommandInvokedHandler
    Dim Ret As UICommandInvokedHandler
    If m_pICoreWindowDialog <> 0& Then
        Dim pIUICommandInvokedHandler As Long
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_GetBackButtonCommand, _
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
    If m_pICoreWindowDialog <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_PutBackButtonCommand, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pICoreWindowDialog <> 0& Then
        If InitWithWindow(m_pICoreWindowDialog, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' ICoreWindowDialogFactory
Public Function CreateWithTitle(ByVal Title As String) As CoreWindowDialog
    Dim Ret As CoreWindowDialog
    If m_pICoreWindowDialogFactory <> 0& Then
        Dim hTitle As Long
        hTitle = CreateWindowsString(Title)
        Dim pICoreWindowDialog As Long
        If Invoke(m_pICoreWindowDialogFactory, _
                  ICoreWindowDialogFactory_CreateWithTitle, _
                  hTitle, _
                  VarPtr(pICoreWindowDialog)) = S_OK Then
            If pICoreWindowDialog <> 0& Then
                Set Ret = New CoreWindowDialog
                Ret.Ifc = pICoreWindowDialog
            End If
        End If
        Call DeleteWindowsString(hTitle)
    End If
    Set CreateWithTitle = Ret
End Function

' ICoreWindowDialog
Public Function AddShowing(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pICoreWindowDialog <> 0& And handler <> 0& Then
        Dim cookie As Currency
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_AddShowing, _
                  handler, _
                  VarPtr(cookie)) = S_OK Then
            Ret = cookie
        End If
    End If
    AddShowing = Ret
End Function

Public Function RemoveShowing(ByVal cookie As Currency) As Boolean
    Dim Ret As Boolean
    If m_pICoreWindowDialog <> 0& Then
        If Invoke(m_pICoreWindowDialog, _
                  ICoreWindowDialog_RemoveShowing, _
                  cookie) = S_OK Then
            Ret = True
        End If
    End If
    RemoveShowing = Ret
End Function

Public Function ShowAsync() As UICommand
    Dim Ret As UICommand
    If m_pICoreWindowDialog <> 0& Then
        Dim pIUICommand As Long
        If Invoke(m_pICoreWindowDialog, _
                    ICoreWindowDialog_ShowAsync, _
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


