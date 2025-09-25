VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Pdf_PdfPdfPage As String = "Windows.Data.Pdf.PdfPage"
Private Const IID_IPdfPage As String = "{9db4b0c8-5320-4cfc-ad76-493fdad0e594}"

' * Class implements the interface:
' *    Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPdfPage
    IPdfPage_RenderToStreamAsync = 6
    IPdfPage_RenderWithOptionsToStreamAsync = 7
    IPdfPage_PreparePageAsync = 8
    IPdfPage_GetIndex = 9
    IPdfPage_GetSize = 10
    IPdfPage_GetDimensions = 11
    IPdfPage_GetRotation = 12
    IPdfPage_GetPreferredZoom = 13
                        
End Enum

' ----==== Types ====----
Private Type SIZEF
    Width As Single
    Height As Single
End Type

' ----==== Variablen ====----
Private m_pIPdfPage As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pIPdfPage) '<- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPdfPage
End Property

Public Property Let Ifc(ByVal pIPdfPage As Long)
    Call DisposeIfc(m_pIPdfPage) '<- IClosable
    m_pIPdfPage = pIPdfPage
End Property

' IPdfPage
Public Property Get index() As Long
    Dim Ret As Long
    If m_pIPdfPage <> 0& Then
        Dim value As Long
        If Invoke(m_pIPdfPage, _
                  IPdfPage_GetIndex, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    index = Ret
End Property

Public Property Get Size() As Size
    Dim Ret As Size
        If m_pIPdfPage <> 0& Then
            Dim value As SIZEF
            If Invoke(m_pIPdfPage, _
                      IPdfPage_GetSize, _
                      VarPtr(value)) = S_OK Then
                Set Ret = New Size
                Ret.Width = value.Width
                Ret.Height = value.Height
            End If
        End If
    Set Size = Ret
End Property

Public Property Get Dimensions() As PdfPageDimensions
    Dim Ret As PdfPageDimensions
        If m_pIPdfPage <> 0& Then
            Dim pIPdfPageDimensions As Long
            If Invoke(m_pIPdfPage, _
                      IPdfPage_GetDimensions, _
                      VarPtr(pIPdfPageDimensions)) = S_OK Then
                If pIPdfPageDimensions <> 0& Then
                    Set Ret = New PdfPageDimensions
                    Ret.Ifc = pIPdfPageDimensions
                End If
            End If
        End If
    Set Dimensions = Ret
End Property

Public Property Get Rotation() As PdfPageRotation
    Dim Ret As PdfPageRotation
        If m_pIPdfPage <> 0& Then
            Dim value As Long
            If Invoke(m_pIPdfPage, _
                      IPdfPage_GetRotation, _
                      VarPtr(value)) = S_OK Then
                Ret = value
            End If
        End If
    Rotation = Ret
End Property

Public Property Get PreferredZoom() As Single
    Dim Ret As Single
        If m_pIPdfPage <> 0& Then
            Dim value As Single
            If Invoke(m_pIPdfPage, _
                      IPdfPage_GetPreferredZoom, _
                      VarPtr(value)) = S_OK Then
                Ret = value
            End If
        End If
    PreferredZoom = Ret
End Property

' ----==== Functions ====----
' IPdfPage
Public Function RenderToStreamAsync(ByVal OutputStream As RandomAccessStream) As Boolean
    Dim Ret As Boolean
        If m_pIPdfPage <> 0& And _
           IsNotNothing(OutputStream) Then
            Dim pIAsyncAction As Long
            If Invoke(m_pIPdfPage, _
                      IPdfPage_RenderToStreamAsync, _
                      OutputStream.Ifc, _
                      VarPtr(pIAsyncAction)) = S_OK Then
                If pIAsyncAction <> 0& Then
                    If Await(pIAsyncAction, False) Then
                        Ret = True
                    End If
                    Call ReleaseIfc(pIAsyncAction)
                End If
            End If
        End If
    RenderToStreamAsync = Ret
End Function

Public Function RenderWithOptionsToStreamAsync(ByVal OutputStream As RandomAccessStream, _
                                               ByVal options As PdfPageRenderOptions) As Boolean
    Dim Ret As Boolean
        If m_pIPdfPage <> 0& And _
           IsNotNothing(options) And _
           IsNotNothing(OutputStream) Then
            Dim pIAsyncAction As Long
            If Invoke(m_pIPdfPage, _
                      IPdfPage_RenderWithOptionsToStreamAsync, _
                      OutputStream.Ifc, _
                      options.Ifc, _
                      VarPtr(pIAsyncAction)) = S_OK Then
                If pIAsyncAction <> 0& Then
                    If Await(pIAsyncAction, False) Then
                        Ret = True
                    End If
                    Call ReleaseIfc(pIAsyncAction)
                End If
            End If
        End If
    RenderWithOptionsToStreamAsync = Ret
End Function

Public Function PreparePageAsync() As Boolean
    Dim Ret As Boolean
        If m_pIPdfPage <> 0& Then
            Dim pIAsyncAction As Long
            If Invoke(m_pIPdfPage, _
                      IPdfPage_PreparePageAsync, _
                      VarPtr(pIAsyncAction)) = S_OK Then
                If pIAsyncAction <> 0& Then
                    If Await(pIAsyncAction, False) Then
                        Ret = True
                    End If
                    Call ReleaseIfc(pIAsyncAction)
                End If
            End If
        End If
    PreparePageAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


