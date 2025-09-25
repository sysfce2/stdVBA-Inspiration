VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpeechSynthesisStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_SpeechSynthesis_SpeechSynthesisStream As String = "Windows.Media.SpeechSynthesis.SpeechSynthesisStream"
Private Const IID_ISpeechSynthesisStream As String = "{83e46e93-244c-4622-ba0b-6229c4d0d65d}"
Private Const IID_IStream As String = "{0000000c-0000-0000-c000-000000000046}"
Private Const IID_IInputStream As String = "{905a0fe2-bc53-11df-8c49-001e4fc686da}"
Private Const IID_IOutputStream As String = "{905a0fe6-bc53-11df-8c49-001e4fc686da}"
Private Const IID_IRandomAccessStream As String = "{905a0fe1-bc53-11df-8c49-001e4fc686da}"
Private Const IID_IContentTypeProvider As String = "{97d098a5-3b99-4de9-88a5-e11d2f50c795}"
Private Const IID_IRandomAccessStreamWithContentType As String = "{cc254827-4b3d-438f-9232-10c76bc7e038}"
Private Const IID_ITimedMetadataTrackProvider As String = "{3b7f2024-f74e-4ade-93c5-219da05b6856}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISpeechSynthesisStream
'    ISpeechSynthesisStream_GetMarkers = 6
    
    ' ITimedMetadataTrackProvider
'    ITimedMetadataTrackProvider_GetTimedMetadataTracks = 6
    
    ' IRandomAccessStreamWithContentType
    ' has no member
    
    ' IRandomAccessStream
    IRandomAccessStream_GetSize = 6
    IRandomAccessStream_PutSize = 7
    IRandomAccessStream_GetInputStreamAt = 8
    IRandomAccessStream_GetOutputStreamAt = 9
    IRandomAccessStream_GetPosition = 10
    IRandomAccessStream_Seek = 11
    IRandomAccessStream_CloneStream = 12
    IRandomAccessStream_GetCanRead = 13
    IRandomAccessStream_GetCanWrite = 14

    ' IInputStream
'    IInputStream_ReadAsync = 6
      
    ' IOutputStream
'    IOutputStream_WriteAsync = 6
    IOutputStream_FlushAsync = 7
    
    ' IContentTypeProvider
    IContentTypeProvider_GetContentType = 6

End Enum

' ----==== Variablen ====----
Private m_pISpeechSynthesisStream As Long
Private m_pIInputStream As Long
Private m_pIOutputStream As Long
Private m_pIRandomAccessStream As Long
Private m_pIContentTypeProvider As Long
Private m_pITimedMetadataTrackProvider As Long
Private m_pIRandomAccessStreamWithContentType As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIInputStream)
    Call ReleaseIfc(m_pIOutputStream)
    Call ReleaseIfc(m_pIRandomAccessStream)
    Call ReleaseIfc(m_pIContentTypeProvider)
    Call ReleaseIfc(m_pITimedMetadataTrackProvider)
    Call ReleaseIfc(m_pIRandomAccessStreamWithContentType)
    Call DisposeIfc(m_pISpeechSynthesisStream) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pISpeechSynthesisStream
End Property

Public Property Let Ifc(ByVal pISpeechSynthesisStream As Long)
    Call ReleaseIfc(m_pIInputStream)
    Call ReleaseIfc(m_pIOutputStream)
    Call ReleaseIfc(m_pIRandomAccessStream)
    Call ReleaseIfc(m_pIContentTypeProvider)
    Call ReleaseIfc(m_pITimedMetadataTrackProvider)
    Call ReleaseIfc(m_pIRandomAccessStreamWithContentType)
    Call DisposeIfc(m_pISpeechSynthesisStream) ' <- IClosable
    m_pISpeechSynthesisStream = pISpeechSynthesisStream
    Call QueryIfc(m_pISpeechSynthesisStream, IID_IOutputStream, m_pIOutputStream)
    Call QueryIfc(m_pISpeechSynthesisStream, IID_IInputStream, m_pIInputStream)
    Call QueryIfc(m_pISpeechSynthesisStream, IID_IRandomAccessStream, m_pIRandomAccessStream)
    Call QueryIfc(m_pISpeechSynthesisStream, IID_IContentTypeProvider, m_pIContentTypeProvider)
    Call QueryIfc(m_pISpeechSynthesisStream, IID_ITimedMetadataTrackProvider, m_pITimedMetadataTrackProvider)
    Call QueryIfc(m_pISpeechSynthesisStream, IID_IRandomAccessStreamWithContentType, m_pIRandomAccessStreamWithContentType)
End Property

' Extra Property
Public Property Get RandomAccessStream() As RandomAccessStream
    Dim Ret As RandomAccessStream
    If m_pISpeechSynthesisStream <> 0& Then
        Dim pIRandomAccessStream As Long
        If QueryIfc(m_pISpeechSynthesisStream, IID_IRandomAccessStream, pIRandomAccessStream) Then
            If pIRandomAccessStream <> 0& Then
                Set Ret = New RandomAccessStream
                Ret.Ifc = m_pIRandomAccessStream
            End If
        End If
    End If
    Set RandomAccessStream = Ret
End Property

' ISpeechSynthesisStream
'    ISpeechSynthesisStream_GetMarkers = 6

' ITimedMetadataTrackProvider
'    ITimedMetadataTrackProvider_GetTimedMetadataTracks = 6

' IRandomAccessStream
Public Property Get Size() As Currency
    Dim Ret As Currency
    If m_pIRandomAccessStream <> 0& Then
        Dim value As Currency
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

Public Property Let Size(ByVal value As Currency)
    If m_pIRandomAccessStream <> 0& Then
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_PutSize, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Position() As Currency
    Dim Ret As Currency
    If m_pIRandomAccessStream <> 0& Then
        Dim value As Currency
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_GetPosition, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Position = Ret
End Property

Public Property Get CanRead() As Boolean
    Dim Ret As Boolean
    If m_pIRandomAccessStream <> 0& Then
        Dim value As Long
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_GetCanRead, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CanRead = Ret
End Property

Public Property Get CanWrite() As Boolean
    Dim Ret As Boolean
    If m_pIRandomAccessStream <> 0& Then
        Dim value As Long
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_GetCanWrite, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CanWrite = Ret
End Property

' IContentTypeProvider
Public Property Get ContentType() As String
    Dim Ret As String
    If m_pIContentTypeProvider <> 0& Then
        Dim value As Long
        If Invoke(m_pIContentTypeProvider, _
                  IContentTypeProvider_GetContentType, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ContentType = Ret
End Property

' ----==== Functions ====----
' IRandomAccessStream
Public Function GetInputStreamAt(ByVal Position As Currency) As InputStream
    Dim Ret As InputStream
    If m_pIRandomAccessStream <> 0& Then
        Dim pIInputStream As Long
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_GetInputStreamAt, _
                  Position, _
                  VarPtr(pIInputStream)) = S_OK Then
            If pIInputStream <> 0& Then
                Set Ret = New InputStream
                Ret.Ifc = pIInputStream
            End If
        End If
    End If
    Set GetInputStreamAt = Ret
End Function

Public Function GetOutputStreamAt(ByVal Position As Currency) As OutputStream
    Dim Ret As OutputStream
    If m_pIRandomAccessStream <> 0& Then
        Dim pIOutputStream As Long
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_GetOutputStreamAt, _
                  Position, _
                  VarPtr(pIOutputStream)) = S_OK Then
            If pIOutputStream <> 0& Then
                Set Ret = New OutputStream
                Ret.Ifc = pIOutputStream
            End If
        End If
    End If
    Set GetOutputStreamAt = Ret
End Function

Public Function SeekStream(ByVal Position As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIRandomAccessStream <> 0& Then
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_Seek, _
                  Position) = S_OK Then
            Ret = True
        End If
    End If
    SeekStream = Ret
End Function

Public Function CloneStream() As RandomAccessStream
    Dim Ret As RandomAccessStream
    If m_pIRandomAccessStream <> 0& Then
        Dim pIRandomAccessStream As Long
        If Invoke(m_pIRandomAccessStream, _
                  IRandomAccessStream_CloneStream, _
                  VarPtr(pIRandomAccessStream)) = S_OK Then
            If pIRandomAccessStream <> 0& Then
                Set Ret = New RandomAccessStream
                Ret.Ifc = pIRandomAccessStream
            End If
        End If
    End If
    Set CloneStream = Ret
End Function

' IInputStream
'    IInputStream_ReadAsync = 6

' IOutputStream
'    IOutputStream_WriteAsync = 6

Public Function FlushAsync() As Boolean
    Dim Ret As Boolean
    If m_pIOutputStream <> 0& Then
        Dim operation As Long
        If Invoke(m_pIOutputStream, _
                  IOutputStream_FlushAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = CBool(operation)
                End If
            End If
        End If
    End If
    FlushAsync = Ret
End Function

' ----==== extra Functions ====----
Public Function ToIStream() As Long
    Dim Ret As Long
    If m_pIRandomAccessStream <> 0& Then
        Dim pIStream As Long
        If CreateStreamOverRandomAccessStream(m_pIRandomAccessStream, _
                                              Str2Guid(IID_IStream), _
                                              pIStream) = S_OK Then
            Ret = pIStream
        End If
    End If
    ToIStream = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

