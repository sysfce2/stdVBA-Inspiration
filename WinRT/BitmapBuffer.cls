VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_BitmapBuffer As String = "Windows.Graphics.Imaging.BitmapBuffer"
Private Const IID_IBitmapBuffer As String = "{a53e04c4-399c-438c-b28f-a63a6b83d1a1}"
Private Const IID_IMemoryBuffer As String = "{fbc4dd2a-245b-11e4-af98-689423260cf8}"
 
' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IBitmapBuffer
    IBitmapBuffer_GetPlaneCount = 6
    IBitmapBuffer_GetPlaneDescription = 7
    
    ' IMemoryBuffer
    IMemoryBuffer_CreateReference = 6
    
End Enum

Private Type BmpPlaneDescription
    startIndex As Long
    Width As Long
    Height As Long
    Stride As Long
End Type

' ----==== Variablen ====----
Private m_pIBitmapBuffer As Long
Private m_pIMemoryBuffer As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMemoryBuffer)
    Call DisposeIfc(m_pIBitmapBuffer) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBitmapBuffer
End Property

Public Property Let Ifc(ByVal pIBitmapBuffer As Long)
    Call ReleaseIfc(m_pIMemoryBuffer)
    Call DisposeIfc(m_pIBitmapBuffer) ' <- IClosable
    m_pIBitmapBuffer = pIBitmapBuffer
    Call QueryIfc(m_pIBitmapBuffer, IID_IMemoryBuffer, m_pIMemoryBuffer)
End Property

' ----==== Functions ====----
' IBitmapBuffer
Public Function GetPlaneCount() As Long
    Dim Ret As Long
    If m_pIBitmapBuffer <> 0& Then
        Dim Value As Long
        If Invoke(m_pIBitmapBuffer, _
                  IBitmapBuffer_GetPlaneCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GetPlaneCount = Ret
End Function

Public Function GetPlaneDescription(ByVal index As Long) As BitmapPlaneDescription
    Dim Ret As BitmapPlaneDescription
    If m_pIBitmapBuffer <> 0& Then
        Dim tBmpPlaneDescription As BmpPlaneDescription
        If Invoke(m_pIBitmapBuffer, _
                  IBitmapBuffer_GetPlaneDescription, _
                  index, _
                  VarPtr(tBmpPlaneDescription)) = S_OK Then
            Set Ret = New BitmapPlaneDescription
            Ret.startIndex = tBmpPlaneDescription.startIndex
            Ret.Width = tBmpPlaneDescription.Width
            Ret.Height = tBmpPlaneDescription.Height
            Ret.Stride = tBmpPlaneDescription.Stride
        End If
    End If
    Set GetPlaneDescription = Ret
End Function

' IMemoryBuffer
Public Function CreateReference() As MemoryBufferReference
    Dim Ret As MemoryBufferReference
    If m_pIMemoryBuffer <> 0& Then
        Dim pIMemoryBufferReference As Long
        If Invoke(m_pIMemoryBuffer, _
                  IMemoryBuffer_CreateReference, _
                  VarPtr(pIMemoryBufferReference)) = S_OK Then
            If pIMemoryBufferReference <> 0& Then
                Set Ret = New MemoryBufferReference
                Ret.Ifc = pIMemoryBufferReference
            End If
        End If
    End If
    Set CreateReference = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


