VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicsCapturePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Capture_GraphicsCapturePicker As String = "Windows.Graphics.Capture.GraphicsCapturePicker"
Private Const IID_IGraphicsCapturePicker As String = "{5a1711b3-ad79-4b4a-9336-1318fdde3539}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGraphicsCapturePicker
    IGraphicsCapturePicker_PickSingleItemAsync = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIGraphicsCapturePicker As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIGraphicsCapturePicker As Long
    If GetActivateInstance(Windows_Graphics_Capture_GraphicsCapturePicker, _
                           IID_IGraphicsCapturePicker, _
                           pIGraphicsCapturePicker) Then
        Me.Ifc = pIGraphicsCapturePicker
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGraphicsCapturePicker)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGraphicsCapturePicker
End Property

Public Property Let Ifc(ByVal pIGraphicsCapturePicker As Long)
    Call ReleaseIfc(m_pIGraphicsCapturePicker)
    m_pIGraphicsCapturePicker = pIGraphicsCapturePicker
End Property

' extra Property
Public Property Let ParentHwnd(ByVal ownerHwnd As Long)
    If m_pIGraphicsCapturePicker <> 0& Then
        If InitWithWindow(m_pIGraphicsCapturePicker, ownerHwnd) Then
        End If
    End If
End Property

' ----==== Function ====----
' IGraphicsCapturePicker
Public Function PickSingleItemAsync() As GraphicsCaptureItem
    Dim Ret As GraphicsCaptureItem
    If m_pIGraphicsCapturePicker <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGraphicsCapturePicker, _
                  IGraphicsCapturePicker_PickSingleItemAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Set Ret = New GraphicsCaptureItem
                    Ret.Ifc = operation
                End If
            End If
        End If
    End If
    Set PickSingleItemAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


