VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoftwareBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_SoftwareBitmap As String = "Windows.Graphics.Imaging.SoftwareBitmap"
Private Const IID_ISoftwareBitmap As String = "{689e0708-7eef-483f-963f-da938818e073}"
Private Const IID_ISoftwareBitmapFactory As String = "{c99feb69-2d62-4d47-a6b3-4fdb6a07fdf8}"
Private Const IID_ISoftwareBitmapStatics As String = "{df0385db-672f-4a9d-806e-c2442f343e86}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ISoftwareBitmapFactory
    ISoftwareBitmapFactory_Create = 6
    ISoftwareBitmapFactory_CreateWithAlpha = 7
                        
    ' ISoftwareBitmapStatics
    ISoftwareBitmapStatics_Copy = 6
    ISoftwareBitmapStatics_Convert = 7
    ISoftwareBitmapStatics_ConvertWithAlpha = 8
    ISoftwareBitmapStatics_CreateCopyFromBuffer = 9
    ISoftwareBitmapStatics_CreateCopyWithAlphaFromBuffer = 10
'    ISoftwareBitmapStatics_CreateCopyFromSurfaceAsync = 11
'    ISoftwareBitmapStatics_CreateCopyWithAlphaFromSurfaceAsync = 12

    ' ISoftwareBitmap
    ISoftwareBitmap_GetBitmapPixelFormat = 6
    ISoftwareBitmap_GetBitmapAlphaMode = 7
    ISoftwareBitmap_GetPixelWidth = 8
    ISoftwareBitmap_GetPixelHeight = 9
    ISoftwareBitmap_GetIsReadOnly = 10
    ISoftwareBitmap_PutDpiX = 11
    ISoftwareBitmap_GetDpiX = 12
    ISoftwareBitmap_PutDpiY = 13
    ISoftwareBitmap_GetDpiY = 14
    ISoftwareBitmap_LockBuffer = 15
    ISoftwareBitmap_CopyTo = 16
    ISoftwareBitmap_CopyFromBuffer = 17
    ISoftwareBitmap_CopyToBuffer = 18
    ISoftwareBitmap_GetReadOnlyView = 19

End Enum

' ----==== Variablen ====----
Private m_pISoftwareBitmap As Long
Private m_pISoftwareBitmapStatics As Long
Private m_pISoftwareBitmapFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Graphics_Imaging_SoftwareBitmap, _
                            IID_ISoftwareBitmapFactory, _
                            m_pISoftwareBitmapFactory) Then
    End If
    If GetActivationFactory(Windows_Graphics_Imaging_SoftwareBitmap, _
                            IID_ISoftwareBitmapStatics, _
                            m_pISoftwareBitmapStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pISoftwareBitmap) ' <- IClosable
    Call ReleaseIfc(m_pISoftwareBitmapStatics)
    Call ReleaseIfc(m_pISoftwareBitmapFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISoftwareBitmap
End Property

Public Property Let Ifc(ByVal pISoftwareBitmap As Long)
    Call DisposeIfc(m_pISoftwareBitmap) ' <- IClosable
    m_pISoftwareBitmap = pISoftwareBitmap
End Property

' ISoftwareBitmap
Public Property Get BitmapPixelFormat() As BitmapPixelFormat
    Dim Ret As BitmapPixelFormat
    If m_pISoftwareBitmap <> 0& Then
        Dim value As Long
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetBitmapPixelFormat, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BitmapPixelFormat = Ret
End Property

Public Property Get BitmapAlphaMode() As BitmapAlphaMode
    Dim Ret As BitmapAlphaMode
    If m_pISoftwareBitmap <> 0& Then
        Dim value As Long
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetBitmapAlphaMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BitmapAlphaMode = Ret
End Property

Public Property Get PixelWidth() As Long
    Dim Ret As Long
    If m_pISoftwareBitmap <> 0& Then
        Dim value As Long
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetPixelWidth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PixelWidth = Ret
End Property

Public Property Get PixelHeight() As Long
    Dim Ret As Long
    If m_pISoftwareBitmap <> 0& Then
        Dim value As Long
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetPixelHeight, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PixelHeight = Ret
End Property

Public Property Get IsReadOnly() As Boolean
    Dim Ret As Boolean
    If m_pISoftwareBitmap <> 0& Then
        Dim value As Long
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetIsReadOnly, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsReadOnly = Ret
End Property

Public Property Let dpiX(ByVal value As Double)
    If m_pISoftwareBitmap <> 0& Then
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_PutDpiX, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get dpiX() As Double
    Dim Ret As Double
    If m_pISoftwareBitmap <> 0& Then
        Dim value As Double
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetDpiX, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    dpiX = Ret
End Property

Public Property Let dpiY(ByVal value As Double)
    If m_pISoftwareBitmap <> 0& Then
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_PutDpiY, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get dpiY() As Double
    Dim Ret As Double
    If m_pISoftwareBitmap <> 0& Then
        Dim value As Double
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetDpiY, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    dpiY = Ret
End Property

' ----==== Functions ====----
' ISoftwareBitmapFactory
Public Function Create(ByVal format As BitmapPixelFormat, _
                       ByVal picWidth As Long, _
                       ByVal picHeight As Long) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmapFactory <> 0& Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmapFactory, _
                  ISoftwareBitmapFactory_Create, _
                  format, _
                  picWidth, _
                  picHeight, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set Create = Ret
End Function

Public Function CreateWithAlpha(ByVal format As BitmapPixelFormat, _
                                ByVal picWidth As Long, _
                                ByVal picHeight As Long, _
                                ByVal alphaMode As BitmapAlphaMode) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmapFactory <> 0& Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmapFactory, _
                  ISoftwareBitmapFactory_CreateWithAlpha, _
                  format, _
                  picWidth, _
                  picHeight, _
                  alphaMode, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set CreateWithAlpha = Ret
End Function
                        
' ISoftwareBitmapStatics
Public Function Copy(ByVal Source As SoftwareBitmap) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmapStatics <> 0& And _
       IsNotNothing(Source) Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmapStatics, _
                  ISoftwareBitmapStatics_Copy, _
                  Source.Ifc, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set Copy = Ret
End Function

Public Function Convert(ByVal Source As SoftwareBitmap, _
                        ByVal format As BitmapPixelFormat) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmapStatics <> 0& And _
       IsNotNothing(Source) Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmapStatics, _
                  ISoftwareBitmapStatics_Convert, _
                  Source.Ifc, _
                  format, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set Convert = Ret
End Function

Public Function ConvertWithAlpha(ByVal Source As SoftwareBitmap, _
                                 ByVal format As BitmapPixelFormat, _
                                 ByVal alphaMode As BitmapAlphaMode) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmapStatics <> 0& And _
       IsNotNothing(Source) Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmapStatics, _
                  ISoftwareBitmapStatics_ConvertWithAlpha, _
                  Source.Ifc, _
                  format, _
                  alphaMode, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set ConvertWithAlpha = Ret
End Function

Public Function CreateCopyFromBuffer(ByVal Source As Buffer, _
                                     ByVal format As BitmapPixelFormat, _
                                     ByVal picWidth As Long, _
                                     ByVal picHeight As Long) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmapStatics <> 0& And _
       IsNotNothing(Source) Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmapStatics, _
                  ISoftwareBitmapStatics_CreateCopyFromBuffer, _
                  Source.Ifc, _
                  format, _
                  picWidth, _
                  picHeight, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set CreateCopyFromBuffer = Ret
End Function

Public Function CreateCopyWithAlphaFromBuffer(ByVal Source As Buffer, _
                                              ByVal format As BitmapPixelFormat, _
                                              ByVal picWidth As Long, _
                                              ByVal picHeight As Long, _
                                              ByVal alphaMode As BitmapAlphaMode) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmapStatics <> 0& And _
       IsNotNothing(Source) Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmapStatics, _
                  ISoftwareBitmapStatics_CreateCopyWithAlphaFromBuffer, _
                  Source.Ifc, _
                  format, _
                  picWidth, _
                  picHeight, _
                  alphaMode, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set CreateCopyWithAlphaFromBuffer = Ret
End Function

'    ISoftwareBitmapStatics_CreateCopyFromSurfaceAsync = 11
'    ISoftwareBitmapStatics_CreateCopyWithAlphaFromSurfaceAsync = 12

' ISoftwareBitmap
Public Function LockBuffer(ByVal mode As BitmapBufferAccessMode) As BitmapBuffer
    Dim Ret As BitmapBuffer
    If m_pISoftwareBitmap <> 0& Then
        Dim pIBitmapBuffer As Long
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_LockBuffer, _
                  mode, _
                  VarPtr(pIBitmapBuffer)) = S_OK Then
            If pIBitmapBuffer <> 0& Then
                Set Ret = New BitmapBuffer
                Ret.Ifc = pIBitmapBuffer
            End If
        End If
    End If
    Set LockBuffer = Ret
End Function

Public Function CopyTo(ByVal Source As SoftwareBitmap) As Boolean
    Dim Ret As Boolean
    If m_pISoftwareBitmap <> 0& And _
       IsNotNothing(Source) Then
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_CopyTo, _
                  Source.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    CopyTo = Ret
End Function

Public Function CopyFromBuffer(ByVal Source As Buffer) As Boolean
    Dim Ret As Boolean
    If m_pISoftwareBitmap <> 0& And _
       IsNotNothing(Source) Then
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_CopyFromBuffer, _
                  Source.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    CopyFromBuffer = Ret
End Function

Public Function CopyToBuffer(ByVal Source As Buffer) As Boolean
    Dim Ret As Boolean
    If m_pISoftwareBitmap <> 0& And _
       IsNotNothing(Source) Then
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_CopyToBuffer, _
                  Source.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    CopyToBuffer = Ret
End Function

Public Function GetReadOnlyView() As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pISoftwareBitmap <> 0& Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pISoftwareBitmap, _
                  ISoftwareBitmap_GetReadOnlyView, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                Set Ret = New SoftwareBitmap
                Ret.Ifc = pISoftwareBitmap
            End If
        End If
    End If
    Set GetReadOnlyView = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


