VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PdfDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Pdf_PdfDocument As String = "Windows.Data.Pdf.PdfDocument"
Private Const IID_IPdfDocument As String = "{ac7ebedd-80fa-4089-846e-81b77ff5a86c}"
Private Const IID_IPdfDocumentStatics As String = "{433a0b5f-c007-4788-90f2-08143d922599}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    'IPdfDocumentStatics
    IPdfDocumentStatics_LoadFromFileAsync = 6
    IPdfDocumentStatics_LoadFromFileWithPasswordAsync = 7
    IPdfDocumentStatics_LoadFromStreamAsync = 8
    IPdfDocumentStatics_LoadFromStreamWithPasswordAsync = 9

    ' IPdfDocument
    IPdfDocument_GetPage = 6
    IPdfDocument_GetPageCount = 7
    IPdfDocument_GetIsPasswordProtected = 8

End Enum

' ----==== Variablen ====----
Private m_pIPdfDocument As Long
Private m_pIPdfDocumentStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Pdf_PdfDocument, _
                            IID_IPdfDocumentStatics, _
                            m_pIPdfDocumentStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPdfDocument)
    Call ReleaseIfc(m_pIPdfDocumentStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPdfDocument
End Property

Public Property Let Ifc(ByVal pIPdfDocument As Long)
    Call ReleaseIfc(m_pIPdfDocument)
    m_pIPdfDocument = pIPdfDocument
End Property

' IPdfDocument
Public Property Get PageCount() As Long
    Dim Ret As Long
    If m_pIPdfDocument <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPdfDocument, _
                  IPdfDocument_GetPageCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    PageCount = Ret
End Property

Public Property Get IsPasswordProtected() As Boolean
    Dim Ret As Boolean
    If m_pIPdfDocument <> 0& Then
        Dim Value As Long
        If Invoke(m_pIPdfDocument, _
                  IPdfDocument_GetIsPasswordProtected, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    IsPasswordProtected = Ret
End Property

' ----==== Functions ====----
'IPdfDocumentStatics
Public Function LoadFromFileAsync(ByVal file As StorageFile) As PdfDocument
    Dim Ret As PdfDocument
    If m_pIPdfDocumentStatics <> 0& And _
       IsNotNothing(file) Then
        Dim pIPdfDocument As Long
        If Invoke(m_pIPdfDocumentStatics, _
                  IPdfDocumentStatics_LoadFromFileAsync, _
                  file.Ifc, _
                  VarPtr(pIPdfDocument)) = S_OK Then
            If pIPdfDocument <> 0& Then
                If Await(pIPdfDocument) Then
                    Set Ret = New PdfDocument
                    Ret.Ifc = pIPdfDocument
                End If
            End If
        End If
    End If
    Set LoadFromFileAsync = Ret
End Function

Public Function LoadFromFileWithPasswordAsync(ByVal file As StorageFile, _
                                              ByVal password As String) As PdfDocument
    Dim Ret As PdfDocument
    If m_pIPdfDocumentStatics <> 0& And _
       IsNotNothing(file) Then
        Dim hString As Long
        hString = CreateWindowsString(password)
        Dim pIPdfDocument As Long
        If Invoke(m_pIPdfDocumentStatics, _
                  IPdfDocumentStatics_LoadFromFileWithPasswordAsync, _
                  file.Ifc, _
                  hString, _
                  VarPtr(pIPdfDocument)) = S_OK Then
            If pIPdfDocument <> 0& Then
                If Await(pIPdfDocument) Then
                    Set Ret = New PdfDocument
                    Ret.Ifc = pIPdfDocument
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set LoadFromFileWithPasswordAsync = Ret
End Function

Public Function LoadFromStreamAsync(ByVal stream As RandomAccessStream) As PdfDocument
    Dim Ret As PdfDocument
    If m_pIPdfDocumentStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim pIPdfDocument As Long
        If Invoke(m_pIPdfDocumentStatics, _
                  IPdfDocumentStatics_LoadFromStreamAsync, _
                  stream.Ifc, _
                  VarPtr(pIPdfDocument)) = S_OK Then
            If pIPdfDocument <> 0& Then
                If Await(pIPdfDocument) Then
                    Set Ret = New PdfDocument
                    Ret.Ifc = pIPdfDocument
                End If
            End If
        End If
    End If
    Set LoadFromStreamAsync = Ret
End Function

Public Function LoadFromStreamWithPasswordAsync(ByVal stream As RandomAccessStream, _
                                                ByVal password As String) As PdfDocument
    Dim Ret As PdfDocument
    If m_pIPdfDocumentStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim hString As Long
        hString = CreateWindowsString(password)
        Dim pIPdfDocument As Long
        If Invoke(m_pIPdfDocumentStatics, _
                  IPdfDocumentStatics_LoadFromStreamWithPasswordAsync, _
                  stream.Ifc, _
                  hString, _
                  VarPtr(pIPdfDocument)) = S_OK Then
            If pIPdfDocument <> 0& Then
                If Await(pIPdfDocument) Then
                    Set Ret = New PdfDocument
                    Ret.Ifc = pIPdfDocument
                End If
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set LoadFromStreamWithPasswordAsync = Ret
End Function

' IPdfDocument
Public Function GetPage(ByVal pageIndex As Long) As PdfPage
    Dim Ret As New PdfPage
        If m_pIPdfDocument <> 0& Then
            Dim pIPdfPage As Long
            If Invoke(m_pIPdfDocument, _
                      IPdfDocument_GetPage, _
                      pageIndex, _
                      VarPtr(pIPdfPage)) = S_OK Then
                If pIPdfPage <> 0& Then
                    Set Ret = New PdfPage
                    Ret.Ifc = pIPdfPage
                End If
            End If
        End If
    Set GetPage = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


