VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicsCaptureSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Capture_GraphicsCaptureSession As String = "Windows.Graphics.Capture.GraphicsCaptureSession"
Private Const IID_IGraphicsCaptureSession As String = "{814e42a9-f70f-4ad7-939b-fddcc6eb880d}"
Private Const IID_IGraphicsCaptureSession2 As String = "{2c39ae40-7d2e-5044-804e-8b6799d4cf9e}"
Private Const IID_IGraphicsCaptureSession3 As String = "{f2cdd966-22ae-5ea1-9596-3a289344c3be}"
Private Const IID_IGraphicsCaptureSessionStatics As String = "{2224a540-5974-49aa-b232-0882536f4cb5}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IGraphicsCaptureSessionStatics
    IGraphicsCaptureSessionStatics_IsSupported = 6
    
    ' IGraphicsCaptureSession
    IGraphicsCaptureSession_StartCapture = 6
    
    ' IGraphicsCaptureSession2
    IGraphicsCaptureSession2_GetIsCursorCaptureEnabled = 6
    IGraphicsCaptureSession2_PutIsCursorCaptureEnabled = 7
    
    ' IGraphicsCaptureSession3
    IGraphicsCaptureSession3_GetIsBorderRequired = 6
    IGraphicsCaptureSession3_PutIsBorderRequired = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIGraphicsCaptureSession As Long
Private m_pIGraphicsCaptureSession2 As Long
Private m_pIGraphicsCaptureSession3 As Long
Private m_pIGraphicsCaptureSessionStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Graphics_Capture_GraphicsCaptureSession, _
                            IID_IGraphicsCaptureSessionStatics, _
                            m_pIGraphicsCaptureSessionStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGraphicsCaptureSessionStatics)
    Call ReleaseIfc(m_pIGraphicsCaptureSession3)
    Call ReleaseIfc(m_pIGraphicsCaptureSession2)
    Call DisposeIfc(m_pIGraphicsCaptureSession) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGraphicsCaptureSession
End Property

Public Property Let Ifc(ByVal pIGraphicsCaptureSession As Long)
    Call ReleaseIfc(m_pIGraphicsCaptureSession3)
    Call ReleaseIfc(m_pIGraphicsCaptureSession2)
    Call DisposeIfc(m_pIGraphicsCaptureSession) ' <- IClosable
    m_pIGraphicsCaptureSession = pIGraphicsCaptureSession
    Call QueryIfc(m_pIGraphicsCaptureSession, IID_IGraphicsCaptureSession2, m_pIGraphicsCaptureSession2)
    Call QueryIfc(m_pIGraphicsCaptureSession, IID_IGraphicsCaptureSession3, m_pIGraphicsCaptureSession3)
End Property

' IGraphicsCaptureSession2
Public Property Get IsCursorCaptureEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIGraphicsCaptureSession2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIGraphicsCaptureSession2, _
                  IGraphicsCaptureSession2_GetIsCursorCaptureEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsCursorCaptureEnabled = Ret
End Property

Public Property Let IsCursorCaptureEnabled(ByVal value As Boolean)
    Dim Ret As Boolean
    If m_pIGraphicsCaptureSession2 <> 0& Then
        If Invoke(m_pIGraphicsCaptureSession2, _
                  IGraphicsCaptureSession2_PutIsCursorCaptureEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

' IGraphicsCaptureSession3
Public Property Get IsBorderRequired() As Boolean
    Dim Ret As Boolean
    If m_pIGraphicsCaptureSession3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIGraphicsCaptureSession3, _
                  IGraphicsCaptureSession3_GetIsBorderRequired, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsBorderRequired = Ret
End Property

Public Property Let IsBorderRequired(ByVal value As Boolean)
    Dim Ret As Boolean
    If m_pIGraphicsCaptureSession3 <> 0& Then
        If Invoke(m_pIGraphicsCaptureSession3, _
                  IGraphicsCaptureSession3_PutIsBorderRequired, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Function ====----
' IGraphicsCaptureSessionStatics
Public Function IsSupported() As Boolean
    Dim Ret As Boolean
    If m_pIGraphicsCaptureSessionStatics <> 0& Then
        Dim result As Long
        If Invoke(m_pIGraphicsCaptureSessionStatics, _
                  IGraphicsCaptureSessionStatics_IsSupported, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
    End If
    IsSupported = Ret
End Function

' IGraphicsCaptureSession
Public Function StartCapture() As Boolean
    Dim Ret As Boolean
    If m_pIGraphicsCaptureSession <> 0& Then
        If Invoke(m_pIGraphicsCaptureSession, _
                  IGraphicsCaptureSession_StartCapture) = S_OK Then
            Ret = True
        End If
    End If
    StartCapture = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
