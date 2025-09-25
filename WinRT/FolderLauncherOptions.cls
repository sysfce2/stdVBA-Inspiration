VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderLauncherOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_FolderLauncherOptions As String = "Windows.System.FolderLauncherOptions"
Private Const IID_ILauncherViewOptions As String = "{8a9b29f1-7ca7-49de-9bd3-3c5b7184f616}"
Private Const IID_IFolderLauncherOptions As String = "{bb91c27d-6b87-432a-bd04-776c6f5fb2ab}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IFolderLauncherOptions
    IFolderLauncherOptions_GetItemsToSelect = 6
    
    ' ILauncherViewOptions
    ILauncherViewOptions_GetDesiredRemainingView = 6
    ILauncherViewOptions_PutDesiredRemainingView = 7
    
End Enum

' ----==== Variablen ====----
Private m_pILauncherViewOptions As Long
Private m_pIFolderLauncherOptions As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_System_FolderLauncherOptions, _
                           IID_IFolderLauncherOptions, _
                           m_pIFolderLauncherOptions) Then
        Call QueryIfc(m_pIFolderLauncherOptions, _
                      IID_ILauncherViewOptions, _
                      m_pILauncherViewOptions)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILauncherViewOptions)
    Call ReleaseIfc(m_pIFolderLauncherOptions)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIFolderLauncherOptions
End Property

' IFolderLauncherOptions
Public Property Get ItemsToSelect() As List_Insprectable
    Dim Ret As List_Insprectable
    If m_pIFolderLauncherOptions <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pIFolderLauncherOptions, _
                  IFolderLauncherOptions_GetItemsToSelect, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_Insprectable
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set ItemsToSelect = Ret
End Property

' ILauncherViewOptions
Public Property Get DesiredRemainingView() As ViewSizePreference
    Dim Ret As ViewSizePreference
    If m_pILauncherViewOptions <> 0& Then
        Dim Value As Long
        If Invoke(m_pILauncherViewOptions, _
                  ILauncherViewOptions_GetDesiredRemainingView, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    DesiredRemainingView = Ret
End Property
    
Public Property Let DesiredRemainingView(ByVal Value As ViewSizePreference)
    If m_pILauncherViewOptions <> 0& Then
        If Invoke(m_pILauncherViewOptions, _
                  ILauncherViewOptions_PutDesiredRemainingView, _
                  Value) = S_OK Then
        End If
    End If
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pIFolderLauncherOptions <> 0& Then
        If InitWithWindow(m_pIFolderLauncherOptions, ownerHwnd) Then
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

