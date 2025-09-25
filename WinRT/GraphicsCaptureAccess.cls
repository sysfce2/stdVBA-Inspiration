VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicsCaptureAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Capture_GraphicsCaptureAccess As String = "Windows.Graphics.Capture.GraphicsCaptureAccess"
Private Const IID_IGraphicsCaptureAccessStatics As String = "{743ed370-06ec-5040-a58a-901f0f757095}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGraphicsCaptureAccessStatics
    IGraphicsCaptureAccessStatics_RequestAccessAsync = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIGraphicsCaptureAccessStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Graphics_Capture_GraphicsCaptureAccess, _
                            IID_IGraphicsCaptureAccessStatics, _
                            m_pIGraphicsCaptureAccessStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGraphicsCaptureAccessStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGraphicsCaptureAccessStatics
End Property

Public Property Let Ifc(ByVal pIGraphicsCaptureAccessStatics As Long)
    Call ReleaseIfc(m_pIGraphicsCaptureAccessStatics)
    m_pIGraphicsCaptureAccessStatics = pIGraphicsCaptureAccessStatics
End Property

' ----==== Function ====----
' IGraphicsCaptureAccessStatics
Public Function RequestAccessAsync(ByVal request As GraphicsCaptureAccessKind) As AppCapabilityAccessStatus
    Dim Ret As AppCapabilityAccessStatus
    If m_pIGraphicsCaptureAccessStatics <> 0& Then
        Dim operation As Long
        If Invoke(m_pIGraphicsCaptureAccessStatics, _
                  IGraphicsCaptureAccessStatics_RequestAccessAsync, _
                  request, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = operation
                End If
            End If
        End If
    End If
    RequestAccessAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function




