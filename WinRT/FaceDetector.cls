VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaceDetector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_FaceAnalysis_FaceDetector As String = "Windows.Media.FaceAnalysis.FaceDetector"
Private Const IID_IFaceDetector As String = "{16b672dc-fe6f-3117-8d95-c3f04d51630c}"
Private Const IID_IFaceDetectorStatics As String = "{bc042d67-9047-33f6-881b-6746c1b218b8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IFaceDetectorStatics
    IFaceDetectorStatics_CreateAsync = 6
    IFaceDetectorStatics_GetSupportedBitmapPixelFormats = 7
    IFaceDetectorStatics_IsBitmapPixelFormatSupported = 8
    IFaceDetectorStatics_GetIsSupported = 9
    
    ' IFaceDetector
    IFaceDetector_DetectFacesAsync = 6
    IFaceDetector_DetectFacesWithSearchAreaAsync = 7
    IFaceDetector_GetMinDetectableFaceSize = 8
    IFaceDetector_PutMinDetectableFaceSize = 9
    IFaceDetector_GetMaxDetectableFaceSize = 10
    IFaceDetector_PutMaxDetectableFaceSize = 11
    
End Enum

' ----==== Types ====----
Private Type BitmapSize_
    Width As Long
    Height As Long
End Type

' ----==== Variablen ====----
Private m_pIFaceDetector As Long
Private m_pIFaceDetectorStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_FaceAnalysis_FaceDetector, _
                            IID_IFaceDetectorStatics, _
                            m_pIFaceDetectorStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIFaceDetector)
    Call ReleaseIfc(m_pIFaceDetectorStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIFaceDetector
End Property

Public Property Let Ifc(ByVal pIFaceDetector As Long)
    Call ReleaseIfc(m_pIFaceDetector)
    m_pIFaceDetector = pIFaceDetector
End Property

' IFaceDetectorStatics
Public Property Get IsSupported() As Boolean
    Dim Ret As Boolean
    If m_pIFaceDetectorStatics <> 0& Then
        Dim value As Long
        If Invoke(m_pIFaceDetectorStatics, _
                  IFaceDetectorStatics_GetIsSupported, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsSupported = Ret
End Property

' IFaceDetector
Public Property Get MinDetectableFaceSize() As BitmapSize
    Dim Ret As BitmapSize
    If m_pIFaceDetector <> 0& Then
        Dim tBitmapSize As BitmapSize_
        If Invoke(m_pIFaceDetector, _
                  IFaceDetector_GetMinDetectableFaceSize, _
                  VarPtr(tBitmapSize)) = S_OK Then
            Set Ret = New BitmapSize
            Ret.Width = tBitmapSize.Width
            Ret.Height = tBitmapSize.Height
        End If
    End If
    Set MinDetectableFaceSize = Ret
End Property

Public Property Let MinDetectableFaceSize(ByVal value As BitmapSize)
    If m_pIFaceDetector <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIFaceDetector, _
                  IFaceDetector_PutMinDetectableFaceSize, _
                  value.Width, _
                  value.Height) = S_OK Then
        End If
    End If
End Property

Public Property Get MaxDetectableFaceSize() As BitmapSize
    Dim Ret As BitmapSize
    If m_pIFaceDetector <> 0& Then
        Dim tBitmapSize As BitmapSize_
        If Invoke(m_pIFaceDetector, _
                  IFaceDetector_GetMaxDetectableFaceSize, _
                  VarPtr(tBitmapSize)) = S_OK Then
            Set Ret = New BitmapSize
            Ret.Width = tBitmapSize.Width
            Ret.Height = tBitmapSize.Height
        End If
    End If
    Set MaxDetectableFaceSize = Ret
End Property

Public Property Let MaxDetectableFaceSize(ByVal value As BitmapSize)
    If m_pIFaceDetector <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIFaceDetector, _
                  IFaceDetector_PutMaxDetectableFaceSize, _
                  value.Width, _
                  value.Height) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IFaceDetectorStatics
Public Function CreateAsync() As FaceDetector
    Dim Ret As FaceDetector
    If m_pIFaceDetectorStatics <> 0& Then
        Dim pIFaceDetector As Long
        If Invoke(m_pIFaceDetectorStatics, _
                  IFaceDetectorStatics_CreateAsync, _
                  VarPtr(pIFaceDetector)) = S_OK Then
            If pIFaceDetector <> 0& Then
                If Await(pIFaceDetector) Then
                    Set Ret = New FaceDetector
                    Ret.Ifc = pIFaceDetector
                End If
            End If
        End If
    End If
    Set CreateAsync = Ret
End Function

Public Function GetSupportedBitmapPixelFormats() As ReadOnlyList_1 'ReadOnlyList_BitmapPixelFormat
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_BitmapPixelFormat
    If m_pIFaceDetectorStatics <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIFaceDetectorStatics, _
                  IFaceDetectorStatics_GetSupportedBitmapPixelFormats, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_BitmapPixelFormat, pIVectorView) 'ReadOnlyList_BitmapPixelFormat
            End If
        End If
    End If
    Set GetSupportedBitmapPixelFormats = Ret
End Function

Public Function IsBitmapPixelFormatSupported(ByVal PixelFormat As BitmapPixelFormat) As Boolean
    Dim Ret As Boolean
    If m_pIFaceDetectorStatics <> 0& Then
        Dim value As Long
        If Invoke(m_pIFaceDetectorStatics, _
                  IFaceDetectorStatics_IsBitmapPixelFormatSupported, _
                  PixelFormat, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsBitmapPixelFormatSupported = Ret
End Function

' IFaceDetector
Public Function DetectFacesAsync(ByVal image As SoftwareBitmap) As ReadOnlyList_1 'ReadOnlyList_DetectedFace
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_DetectedFace
    If m_pIFaceDetector <> 0& And _
       IsNotNothing(image) Then
        Dim pIVectorView As Long
        If Invoke(m_pIFaceDetector, _
                  IFaceDetector_DetectFacesAsync, _
                  image.Ifc, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_DetectedFace, pIVectorView) 'ReadOnlyList_DetectedFace
                End If
            End If
        End If
    End If
    Set DetectFacesAsync = Ret
End Function

Public Function DetectFacesWithSearchAreaAsync(ByVal image As SoftwareBitmap, _
                                               ByVal searchArea As BitmapBounds) As ReadOnlyList_1 'ReadOnlyList_DetectedFace
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_DetectedFace
    If m_pIFaceDetector <> 0& And _
       IsNotNothing(image) And _
       IsNotNothing(searchArea) Then
        Dim pIVectorView As Long
        If Invoke(m_pIFaceDetector, _
                  IFaceDetector_DetectFacesWithSearchAreaAsync, _
                  image.Ifc, _
                  searchArea.X, _
                  searchArea.Y, _
                  searchArea.Width, _
                  searchArea.Height, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_DetectedFace, pIVectorView) 'ReadOnlyList_DetectedFace
                End If
            End If
        End If
    End If
    Set DetectFacesWithSearchAreaAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


