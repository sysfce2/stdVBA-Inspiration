VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_BitmapTransform As String = "Windows.Graphics.Imaging.BitmapTransform"
Private Const IID_IBitmapTransform As String = "{ae755344-e268-4d35-adcf-e995d31a8d34}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IBitmapTransform
    IBitmapTransform_GetScaledWidth = 6
    IBitmapTransform_PutScaledWidth = 7
    IBitmapTransform_GetScaledHeight = 8
    IBitmapTransform_PutScaledHeight = 9
    IBitmapTransform_GetInterpolationMode = 10
    IBitmapTransform_PutInterpolationMode = 11
    IBitmapTransform_GetFlip = 12
    IBitmapTransform_PutFlip = 13
    IBitmapTransform_GetRotation = 14
    IBitmapTransform_PutRotation = 15
    IBitmapTransform_GetBounds = 16
    IBitmapTransform_PutBounds = 17
                        
End Enum

Private Type BitmapBounds_
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type
                
' ----==== Vars ====----
Private m_pIBitmapTransform As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIBitmapTransform As Long
    If GetActivateInstance(Windows_Graphics_Imaging_BitmapTransform, _
                           IID_IBitmapTransform, _
                           pIBitmapTransform) Then
        Me.Ifc = pIBitmapTransform
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBitmapTransform)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBitmapTransform
End Property

Public Property Let Ifc(ByVal pIBitmapTransform As Long)
    Call ReleaseIfc(m_pIBitmapTransform)
    m_pIBitmapTransform = pIBitmapTransform
End Property

' IBitmapTransform
Public Property Get ScaledWidth() As Long
    Dim Ret As Long
    If m_pIBitmapTransform <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_GetScaledWidth, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    ScaledWidth = Ret
End Property

Public Property Let ScaledWidth(ByVal Value As Long)
    If m_pIBitmapTransform <> 0& Then
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_PutScaledWidth, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get ScaledHeight() As Long
    Dim Ret As Long
    If m_pIBitmapTransform <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_GetScaledHeight, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    ScaledHeight = Ret
End Property

Public Property Let ScaledHeight(ByVal Value As Long)
    If m_pIBitmapTransform <> 0& Then
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_PutScaledHeight, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get InterpolationMode() As BitmapInterpolationMode
    Dim Ret As BitmapInterpolationMode
    If m_pIBitmapTransform <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_GetInterpolationMode, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    InterpolationMode = Ret
End Property

Public Property Let InterpolationMode(ByVal Value As BitmapInterpolationMode)
    If m_pIBitmapTransform <> 0& Then
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_PutInterpolationMode, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Flip() As BitmapFlip
    Dim Ret As BitmapFlip
    If m_pIBitmapTransform <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_GetFlip, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Flip = Ret
End Property

Public Property Let Flip(ByVal Value As BitmapFlip)
    If m_pIBitmapTransform <> 0& Then
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_PutFlip, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Rotation() As BitmapRotation
    Dim Ret As BitmapRotation
    If m_pIBitmapTransform <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_GetRotation, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Rotation = Ret
End Property

Public Property Let Rotation(ByVal Value As BitmapRotation)
    If m_pIBitmapTransform <> 0& Then
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_PutRotation, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Bounds() As BitmapBounds
    Dim Ret As BitmapBounds
    If m_pIBitmapTransform <> 0& Then
        Dim Value As BitmapBounds_
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_GetBounds, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New BitmapBounds
            Ret.X = Value.X
            Ret.Y = Value.Y
            Ret.Width = Value.Width
            Ret.Height = Value.Height
        End If
    End If
    Set Bounds = Ret
End Property

Public Property Let Bounds(ByVal Value As BitmapBounds)
    If m_pIBitmapTransform <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIBitmapTransform, _
                  IBitmapTransform_PutBounds, _
                  Value.X, _
                  Value.Y, _
                  Value.Width, _
                  Value.Height) = S_OK Then
        End If
    End If
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

