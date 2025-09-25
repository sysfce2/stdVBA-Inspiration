VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageEncodingProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_MediaProperties_ImageEncodingProperties As String = "Windows.Media.MediaProperties.ImageEncodingProperties"
Private Const IID_IMediaEncodingProperties As String = "{b4002af6-acd4-4e5a-a24b-5d7498a8b8c4}"
Private Const IID_IImageEncodingProperties As String = "{78625635-f331-4189-b1c3-b48d5ae034f1}"
Private Const IID_IImageEncodingProperties2 As String = "{c854a2df-c923-469b-ac8e-6a9f3c1cd9e3}"
Private Const IID_IImageEncodingPropertiesStatics As String = "{257c68dc-8b99-439e-aa59-913a36161297}"
Private Const IID_IImageEncodingPropertiesStatics2 As String = "{f6c25b29-3824-46b0-956e-501329e1be3c}"
Private Const IID_IImageEncodingPropertiesStatics3 As String = "{48f4814d-a2ff-48dc-8ea0-e90680663656}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IImageEncodingPropertiesStatics
    IImageEncodingPropertiesStatics_CreateJpeg = 6
    IImageEncodingPropertiesStatics_CreatePng = 7
    IImageEncodingPropertiesStatics_CreateJpegXR = 8
    
    ' IImageEncodingPropertiesStatics2
    IImageEncodingPropertiesStatics2_CreateUncompressed = 6
    IImageEncodingPropertiesStatics2_CreateBmp = 7
    
    ' IImageEncodingPropertiesStatics3
    IImageEncodingPropertiesStatics3_CreateHeif = 6
    
    ' IMediaEncodingProperties
'    IMediaEncodingProperties_GetProperties = 6
    IMediaEncodingProperties_GetType = 7
    IMediaEncodingProperties_PutSubtype = 8
    IMediaEncodingProperties_GetSubtype = 9
                    
    ' IImageEncodingProperties
    IImageEncodingProperties_PutWidth = 6
    IImageEncodingProperties_GetWidth = 7
    IImageEncodingProperties_PutHeight = 8
    IImageEncodingProperties_GetHeight = 9
    
    ' IImageEncodingProperties2
    IImageEncodingProperties2_Copy = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaEncodingProperties As Long
Private m_pIImageEncodingProperties As Long
Private m_pIImageEncodingProperties2 As Long
Private m_pIImageEncodingPropertiesStatics As Long
Private m_pIImageEncodingPropertiesStatics2 As Long
Private m_pIImageEncodingPropertiesStatics3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_MediaProperties_ImageEncodingProperties, _
                            IID_IImageEncodingPropertiesStatics, _
                            m_pIImageEncodingPropertiesStatics) Then
        Call QueryIfc(m_pIImageEncodingPropertiesStatics, IID_IImageEncodingPropertiesStatics2, m_pIImageEncodingPropertiesStatics2)
        Call QueryIfc(m_pIImageEncodingPropertiesStatics, IID_IImageEncodingPropertiesStatics3, m_pIImageEncodingPropertiesStatics3)
    End If
    Dim pIMediaEncodingProperties As Long
    If GetActivateInstance(Windows_Media_MediaProperties_ImageEncodingProperties, _
                           IID_IMediaEncodingProperties, _
                           pIMediaEncodingProperties) Then
        Me.Ifc = pIMediaEncodingProperties
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaEncodingProperties)
    Call ReleaseIfc(m_pIImageEncodingProperties2)
    Call ReleaseIfc(m_pIImageEncodingProperties)
    Call ReleaseIfc(m_pIImageEncodingPropertiesStatics3)
    Call ReleaseIfc(m_pIImageEncodingPropertiesStatics2)
    Call ReleaseIfc(m_pIImageEncodingPropertiesStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIImageEncodingProperties
End Property

Public Property Let Ifc(ByVal pIImageEncodingProperties As Long)
    Call ReleaseIfc(m_pIMediaEncodingProperties)
    Call ReleaseIfc(m_pIImageEncodingProperties2)
    Call ReleaseIfc(m_pIImageEncodingProperties)
    m_pIImageEncodingProperties = pIImageEncodingProperties
    Call QueryIfc(m_pIImageEncodingProperties, IID_IMediaEncodingProperties, m_pIMediaEncodingProperties)
    Call QueryIfc(m_pIImageEncodingProperties, IID_IImageEncodingProperties2, m_pIImageEncodingProperties2)
End Property

' IMediaEncodingProperties
'    IMediaEncodingProperties_GetProperties = 6

Public Property Get MediaType() As String
    Dim Ret As String
    If m_pIMediaEncodingProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMediaEncodingProperties, _
                  IMediaEncodingProperties_GetType, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MediaType = Ret
End Property

Public Property Let MediaSubtype(ByVal Value As String)
    If m_pIMediaEncodingProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIMediaEncodingProperties, _
                  IMediaEncodingProperties_PutSubtype, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get MediaSubtype() As String
    Dim Ret As String
    If m_pIMediaEncodingProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMediaEncodingProperties, _
                  IMediaEncodingProperties_GetSubtype, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MediaSubtype = Ret
End Property

' IImageEncodingProperties
Public Property Let Width(ByVal Value As Long)
    If m_pIImageEncodingProperties <> 0& Then
        If Invoke(m_pIImageEncodingProperties, _
                  IImageEncodingProperties_PutWidth, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Width() As Long
    Dim Ret As Long
    If m_pIImageEncodingProperties <> 0& Then
        Dim Value As Long
        If Invoke(m_pIImageEncodingProperties, _
                  IImageEncodingProperties_GetWidth, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Width = Ret
End Property

Public Property Let Height(ByVal Value As Long)
    If m_pIImageEncodingProperties <> 0& Then
        If Invoke(m_pIImageEncodingProperties, _
                  IImageEncodingProperties_PutHeight, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Height() As Long
    Dim Ret As Long
    If m_pIImageEncodingProperties <> 0& Then
        Dim Value As Long
        If Invoke(m_pIImageEncodingProperties, _
                  IImageEncodingProperties_GetHeight, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Height = Ret
End Property

' ----==== Functions ====----
' IImageEncodingPropertiesStatics
Public Function CreateJpeg() As ImageEncodingProperties
    Dim Ret As ImageEncodingProperties
    If m_pIImageEncodingPropertiesStatics <> 0& Then
        Dim pIImageEncodingProperties As Long
        If Invoke(m_pIImageEncodingPropertiesStatics, _
                  IImageEncodingPropertiesStatics_CreateJpeg, _
                  VarPtr(pIImageEncodingProperties)) = S_OK Then
            If pIImageEncodingProperties <> 0& Then
                Set Ret = New ImageEncodingProperties
                Ret.Ifc = pIImageEncodingProperties
            End If
        End If
    End If
    Set CreateJpeg = Ret
End Function

Public Function CreatePng() As ImageEncodingProperties
    Dim Ret As ImageEncodingProperties
    If m_pIImageEncodingPropertiesStatics <> 0& Then
        Dim pIImageEncodingProperties As Long
        If Invoke(m_pIImageEncodingPropertiesStatics, _
                  IImageEncodingPropertiesStatics_CreatePng, _
                  VarPtr(pIImageEncodingProperties)) = S_OK Then
            If pIImageEncodingProperties <> 0& Then
                Set Ret = New ImageEncodingProperties
                Ret.Ifc = pIImageEncodingProperties
            End If
        End If
    End If
    Set CreatePng = Ret
End Function

Public Function CreateJpegXR() As ImageEncodingProperties
    Dim Ret As New ImageEncodingProperties
    If m_pIImageEncodingPropertiesStatics <> 0& Then
        Dim pIImageEncodingProperties As Long
        If Invoke(m_pIImageEncodingPropertiesStatics, _
                  IImageEncodingPropertiesStatics_CreateJpegXR, _
                  VarPtr(pIImageEncodingProperties)) = S_OK Then
            If pIImageEncodingProperties <> 0& Then
                Set Ret = New ImageEncodingProperties
                Ret.Ifc = pIImageEncodingProperties
            End If
        End If
    End If
    Set CreateJpegXR = Ret
End Function

' IImageEncodingPropertiesStatics2
Public Function CreateUncompressed(ByVal pixelFormat As MediaPixelFormat) As ImageEncodingProperties
    Dim Ret As ImageEncodingProperties
    If m_pIImageEncodingPropertiesStatics2 <> 0& Then
        Dim pIImageEncodingProperties As Long
        If Invoke(m_pIImageEncodingPropertiesStatics2, _
                  IImageEncodingPropertiesStatics2_CreateUncompressed, _
                  pixelFormat, _
                  VarPtr(pIImageEncodingProperties)) = S_OK Then
            If pIImageEncodingProperties <> 0& Then
                Set Ret = New ImageEncodingProperties
                Ret.Ifc = pIImageEncodingProperties
            End If
        End If
    End If
    Set CreateUncompressed = Ret
End Function

Public Function CreateBmp() As ImageEncodingProperties
    Dim Ret As ImageEncodingProperties
    If m_pIImageEncodingPropertiesStatics2 <> 0& Then
        Dim pIImageEncodingProperties As Long
        If Invoke(m_pIImageEncodingPropertiesStatics2, _
                  IImageEncodingPropertiesStatics2_CreateBmp, _
                  VarPtr(pIImageEncodingProperties)) = S_OK Then
            If pIImageEncodingProperties <> 0& Then
                Set Ret = New ImageEncodingProperties
                Ret.Ifc = pIImageEncodingProperties
            End If
        End If
    End If
    Set CreateBmp = Ret
End Function
    
' IImageEncodingPropertiesStatics3
Public Function CreateHeif() As ImageEncodingProperties
    Dim Ret As ImageEncodingProperties
    If m_pIImageEncodingPropertiesStatics3 <> 0& Then
        Dim pIImageEncodingProperties As Long
        If Invoke(m_pIImageEncodingPropertiesStatics3, _
                  IImageEncodingPropertiesStatics3_CreateHeif, _
                  VarPtr(pIImageEncodingProperties)) = S_OK Then
            If pIImageEncodingProperties <> 0& Then
                Set Ret = New ImageEncodingProperties
                Ret.Ifc = pIImageEncodingProperties
            End If
        End If
    End If
    Set CreateHeif = Ret
End Function

' IImageEncodingProperties2
Public Function Copy() As ImageEncodingProperties
    Dim Ret As ImageEncodingProperties
    If m_pIImageEncodingProperties2 <> 0& Then
        Dim pIImageEncodingProperties As Long
        If Invoke(m_pIImageEncodingProperties2, _
                  IImageEncodingProperties2_Copy, _
                  VarPtr(pIImageEncodingProperties)) = S_OK Then
            If pIImageEncodingProperties <> 0& Then
                Set Ret = New ImageEncodingProperties
                Ret.Ifc = pIImageEncodingProperties
            End If
        End If
    End If
    Set Copy = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

