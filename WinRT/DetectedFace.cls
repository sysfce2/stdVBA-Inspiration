VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetectedFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_FaceAnalysis_DetectedFace As String = "Windows.Media.FaceAnalysis.DetectedFace"
Private Const IID_IDetectedFace As String = "{8200d454-66bc-34df-9410-e89400195414}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IDetectedFace
    IDetectedFace_GetFaceBox = 6
    
End Enum

' ----==== Types ====----
Private Type BitmapBounds_
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type

' ----==== Variablen ====----
Private m_pIDetectedFace As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDetectedFace)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDetectedFace
End Property

Public Property Let Ifc(ByVal pIDetectedFace As Long)
    Call ReleaseIfc(m_pIDetectedFace)
    m_pIDetectedFace = pIDetectedFace
End Property

' IDetectedFace
Public Property Get FaceBox() As BitmapBounds
    Dim Ret As BitmapBounds
    If m_pIDetectedFace <> 0& Then
        Dim tBitmapBounds As BitmapBounds_
        If Invoke(m_pIDetectedFace, _
                  IDetectedFace_GetFaceBox, _
                  VarPtr(tBitmapBounds)) = S_OK Then
            Set Ret = New BitmapBounds
            Ret.X = tBitmapBounds.X
            Ret.Y = tBitmapBounds.Y
            Ret.Width = tBitmapBounds.Width
            Ret.Height = tBitmapBounds.Height
        End If
    End If
    Set FaceBox = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


