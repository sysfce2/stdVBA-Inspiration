VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Popups_MessageDialog As String = "Windows.UI.Popups.MessageDialog"
Private Const IID_IMessageDialog As String = "{33f59b01-5325-43ab-9ab3-bdae440e4121}"
Private Const IID_IMessageDialogFactory As String = "{2d161777-a66f-4ea5-bb87-793ffa4941f2}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMessageDialogFactory
    IMessageDialogFactory_Create = 6
    IMessageDialogFactory_CreateWithTitle = 7
    
    ' IMessageDialog
    IMessageDialog_GetTitle = 6
    IMessageDialog_PutTitle = 7
    IMessageDialog_GetCommands = 8
    IMessageDialog_GetDefaultCommandIndex = 9
    IMessageDialog_PutDefaultCommandIndex = 10
    IMessageDialog_GetCancelCommandIndex = 11
    IMessageDialog_PutCancelCommandIndex = 12
    IMessageDialog_GetContent = 13
    IMessageDialog_PutContent = 14
    IMessageDialog_ShowAsync = 15
    IMessageDialog_GetOptions = 16
    IMessageDialog_PutOptions = 17

End Enum

' ----==== Variablen ====----
Private m_pIMessageDialog As Long
Private m_pIMessageDialogFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Popups_MessageDialog, _
                            IID_IMessageDialogFactory, _
                            m_pIMessageDialogFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMessageDialog)
    Call ReleaseIfc(m_pIMessageDialogFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMessageDialog
End Property

Public Property Let Ifc(ByVal pIMessageDialog As Long)
    Call ReleaseIfc(m_pIMessageDialog)
    m_pIMessageDialog = pIMessageDialog
End Property

' IMessageDialog
Public Property Get Title() As String
    Dim Ret As String
    If m_pIMessageDialog <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_GetTitle, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal Value As String)
    If m_pIMessageDialog <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Commands() As List_UICommand
    Dim Ret As List_UICommand
    If m_pIMessageDialog <> 0& Then
        Dim pIUICommandList As Long
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_GetCommands, _
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
    If m_pIMessageDialog <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_GetDefaultCommandIndex, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    DefaultCommandIndex = Ret
End Property

Public Property Let DefaultCommandIndex(ByVal Value As Long)
    If m_pIMessageDialog <> 0& Then
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_PutDefaultCommandIndex, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get CancelCommandIndex() As Long
    Dim Ret As Long
    If m_pIMessageDialog <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_GetCancelCommandIndex, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    CancelCommandIndex = Ret
End Property

Public Property Let CancelCommandIndex(ByVal Value As Long)
    If m_pIMessageDialog <> 0& Then
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_PutCancelCommandIndex, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Content() As String
    Dim Ret As String
    If m_pIMessageDialog <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_GetContent, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Content = Ret
End Property

Public Property Let Content(ByVal Value As String)
    If m_pIMessageDialog <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_PutContent, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get options() As MessageDialogOptions
    Dim Ret As MessageDialogOptions
    If m_pIMessageDialog <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_GetOptions, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    options = Ret
End Property

Public Property Let options(ByVal Value As MessageDialogOptions)
    If m_pIMessageDialog <> 0& Then
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_PutOptions, _
                  Value) = S_OK Then
        End If
    End If
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pIMessageDialog <> 0& Then
        If InitWithWindow(m_pIMessageDialog, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Functions ====----
' IMessageDialogFactory
Public Function Create(ByVal Content As String) As MessageDialog
    Dim Ret As MessageDialog
    If m_pIMessageDialogFactory <> 0& Then
        Dim hContent As Long
        hContent = CreateWindowsString(Content)
        Dim pIMessageDialog As Long
        If Invoke(m_pIMessageDialogFactory, _
                  IMessageDialogFactory_Create, _
                  hContent, _
                  VarPtr(pIMessageDialog)) = S_OK Then
            If pIMessageDialog <> 0& Then
                Set Ret = New MessageDialog
                Ret.Ifc = pIMessageDialog
            End If
        End If
        Call DeleteWindowsString(hContent)
    End If
    Set Create = Ret
End Function

Public Function CreateWithTitle(ByVal Content As String, _
                                ByVal Title As String) As MessageDialog
    Dim Ret As MessageDialog
    If m_pIMessageDialogFactory <> 0& Then
        Dim hContent As Long
        Dim hTitle As Long
        hContent = CreateWindowsString(Content)
        hTitle = CreateWindowsString(Title)
        Dim pIMessageDialog As Long
        If Invoke(m_pIMessageDialogFactory, _
                  IMessageDialogFactory_CreateWithTitle, _
                  hContent, _
                  hTitle, _
                  VarPtr(pIMessageDialog)) = S_OK Then
            If pIMessageDialog <> 0& Then
                Set Ret = New MessageDialog
                Ret.Ifc = pIMessageDialog
            End If
        End If
        Call DeleteWindowsString(hContent)
        Call DeleteWindowsString(hTitle)
    End If
    Set CreateWithTitle = Ret
End Function

' IMessageDialog
Public Function ShowAsync() As UICommand
    Dim Ret As UICommand
    If m_pIMessageDialog <> 0& Then
        Dim pIUICommand As Long
        If Invoke(m_pIMessageDialog, _
                  IMessageDialog_ShowAsync, _
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


