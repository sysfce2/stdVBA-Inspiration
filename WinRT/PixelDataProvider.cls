VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PixelDataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_PixelDataProvider As String = "Windows.Graphics.Imaging.PixelDataProvider"
Private Const IID_IPixelDataProvider As String = "{dd831f25-185c-4595-9fb9-ccbe6ec18a6f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPixelDataProvider
    IPixelDataProvider_DetachPixelData = 6

End Enum

' ----==== Variablen ====----
Private m_pIPixelDataProvider As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPixelDataProvider)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPixelDataProvider
End Property

Public Property Let Ifc(ByVal pIPixelDataProvider As Long)
    Call ReleaseIfc(m_pIPixelDataProvider)
    m_pIPixelDataProvider = pIPixelDataProvider
End Property

' ----==== Functions ====----
' IPixelDataProvider
Public Function DetachPixelData(ByRef pixelDataLength As Long) As Byte()
    Dim Ret() As Byte
    If m_pIPixelDataProvider <> 0& Then
        Dim pixelData As Long
        If Invoke(m_pIPixelDataProvider, _
                  IPixelDataProvider_DetachPixelData, _
                  VarPtr(pixelDataLength), _
                  VarPtr(pixelData)) = S_OK Then
            If pixelDataLength > 0& And pixelData <> 0& Then
                ReDim Ret(pixelDataLength - 1)
                Call CopyMemory(Ret(0), ByVal pixelData, pixelDataLength)
                Call CoTaskMemFree(pixelData)
            End If
        End If
    End If
    DetachPixelData = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


