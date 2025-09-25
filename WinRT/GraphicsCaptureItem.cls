VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicsCaptureItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const GWL_STYLE As Long = -16
Private Const WS_MINIMIZE As Long = &H20000000

Private Const Windows_Graphics_Capture_GraphicsCaptureItem As String = "Windows.Graphics.Capture.GraphicsCaptureItem"
Private Const IID_IGraphicsCaptureItem As String = "{79c3f95b-31f7-4ec2-a464-632ef5d30760}"
Private Const IID_IGraphicsCaptureItemInterop As String = "{3628e81b-3cac-4c60-b7f4-23ce0e0c3356}"
Private Const IID_IGraphicsCaptureItemStatics As String = "{a87ebea5-457c-5788-ab47-0cf1d3637e74}"
Private Const IID_IGraphicsCaptureItemStatics2 As String = "{3b92acc9-e584-5862-bf5c-9c316c6d2dbb}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGraphicsCaptureItemStatics
'    IGraphicsCaptureItemStatics_CreateFromVisual = 6
    
    ' IGraphicsCaptureItemStatics2
'    IGraphicsCaptureItemStatics2_TryCreateFromWindowId = 6
'    IGraphicsCaptureItemStatics2_TryCreateFromDisplayId = 7
    
    ' IGraphicsCaptureItemInterop
    IGraphicsCaptureItemInterop_CreateForWindow = 3
    IGraphicsCaptureItemInterop_CreateForMonitor = 4
    
    ' IGraphicsCaptureItem
    IGraphicsCaptureItem_GetDisplayName = 6
    IGraphicsCaptureItem_GetSize = 7
    IGraphicsCaptureItem_AddClosed = 8
    IGraphicsCaptureItem_RemoveClosed = 9
    
End Enum

' ----==== Types ====----
Private Type SizeInt32_
    Width As Long
    Height As Long
End Type

' ----==== User32.dll Declarations ====----
Private Declare Function FindWindow Lib "User32.dll" _
                         Alias "FindWindowW" ( _
                         ByVal lpClassName As Long, _
                         ByVal lpWindowName As Long) As Long
                         
Private Declare Function GetWindowLong Lib "User32.dll" _
                         Alias "GetWindowLongA" ( _
                         ByVal hwnd As Long, _
                         ByVal nIndex As Long) As Long

' ----==== Variablen ====----
Private m_pIGraphicsCaptureItem As Long
Private m_pIGraphicsCaptureItemInterop As Long
Private m_pIGraphicsCaptureItemStatics As Long
Private m_pIGraphicsCaptureItemStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Graphics_Capture_GraphicsCaptureItem, _
                            IID_IGraphicsCaptureItemStatics, _
                            m_pIGraphicsCaptureItemStatics) Then
        Call QueryIfc(m_pIGraphicsCaptureItemStatics, _
                      IID_IGraphicsCaptureItemStatics2, _
                      m_pIGraphicsCaptureItemStatics2)
        Call QueryIfc(m_pIGraphicsCaptureItemStatics, _
                      IID_IGraphicsCaptureItemInterop, _
                      m_pIGraphicsCaptureItemInterop)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGraphicsCaptureItemInterop)
    Call ReleaseIfc(m_pIGraphicsCaptureItemStatics2)
    Call ReleaseIfc(m_pIGraphicsCaptureItemStatics)
    Call ReleaseIfc(m_pIGraphicsCaptureItem)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGraphicsCaptureItem
End Property

Public Property Let Ifc(ByVal pIGraphicsCaptureItem As Long)
    Call ReleaseIfc(m_pIGraphicsCaptureItem)
    m_pIGraphicsCaptureItem = pIGraphicsCaptureItem
End Property

' IGraphicsCaptureItem
Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIGraphicsCaptureItem <> 0& Then
        Dim value As Long
        If Invoke(m_pIGraphicsCaptureItem, _
                  IGraphicsCaptureItem_GetDisplayName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get Size() As SizeInt32
    Dim Ret As SizeInt32
    If m_pIGraphicsCaptureItem <> 0& Then
        Dim value As SizeInt32_
        If Invoke(m_pIGraphicsCaptureItem, _
                  IGraphicsCaptureItem_GetSize, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New SizeInt32
            Ret.Width = value.Width
            Ret.Height = value.Height
        End If
    End If
    Set Size = Ret
End Property

' extra Property
Public Property Get IsMinimized() As Boolean
    Dim Ret As Boolean
    If m_pIGraphicsCaptureItem <> 0& Then
        Dim hwnd As Long
        hwnd = FindWindow(0&, StrPtr(Me.DisplayName))
        If hwnd <> 0& Then
            If CBool(GetWindowLong(hwnd, GWL_STYLE) And WS_MINIMIZE) Then
                Ret = True
            End If
        End If
    End If
    IsMinimized = Ret
End Property

' ----==== Function ====----
' IGraphicsCaptureItemStatics
'    IGraphicsCaptureItemStatics_CreateFromVisual = 6
    
' IGraphicsCaptureItemStatics2
'    IGraphicsCaptureItemStatics2_TryCreateFromWindowId = 6
'    IGraphicsCaptureItemStatics2_TryCreateFromDisplayId = 7

' IGraphicsCaptureItemInterop
Public Function CreateForWindow(ByVal hwnd As Long) As GraphicsCaptureItem
    Dim Ret As GraphicsCaptureItem
    If m_pIGraphicsCaptureItemInterop <> 0& Then
        Dim pIGraphicsCaptureItem As Long
        If Invoke(m_pIGraphicsCaptureItemInterop, _
                  IGraphicsCaptureItemInterop_CreateForWindow, _
                  hwnd, _
                  VarPtr(Str2Guid(IID_IGraphicsCaptureItem)), _
                  VarPtr(pIGraphicsCaptureItem)) = S_OK Then
            If pIGraphicsCaptureItem <> 0& Then
                Set Ret = New GraphicsCaptureItem
                Ret.Ifc = pIGraphicsCaptureItem
            End If
        End If
    End If
    Set CreateForWindow = Ret
End Function

Public Function CreateForMonitor(ByVal hMonitor As Long) As GraphicsCaptureItem
    Dim Ret As GraphicsCaptureItem
    If m_pIGraphicsCaptureItemInterop <> 0& Then
        Dim pIGraphicsCaptureItem As Long
        If Invoke(m_pIGraphicsCaptureItemInterop, _
                  IGraphicsCaptureItemInterop_CreateForMonitor, _
                  hMonitor, _
                  VarPtr(Str2Guid(IID_IGraphicsCaptureItem)), _
                  VarPtr(pIGraphicsCaptureItem)) = S_OK Then
            If pIGraphicsCaptureItem <> 0& Then
                Set Ret = New GraphicsCaptureItem
                Ret.Ifc = pIGraphicsCaptureItem
            End If
        End If
    End If
    Set CreateForMonitor = Ret
End Function

' IGraphicsCaptureItem
Public Function AddClosed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIGraphicsCaptureItem <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIGraphicsCaptureItem, _
                  IGraphicsCaptureItem_AddClosed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddClosed = Ret
End Function

Public Function RemoveClosed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIGraphicsCaptureItem <> 0& And token <> 0@ Then
        If Invoke(m_pIGraphicsCaptureItem, _
                  IGraphicsCaptureItem_RemoveClosed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveClosed = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


