VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomAccessStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Streams_RandomAccessStream As String = "Windows.Storage.Streams.RandomAccessStream"
Private Const IID_IStream As String = "{0000000c-0000-0000-c000-000000000046}"
Private Const IID_IInputStream As String = "{905a0fe2-bc53-11df-8c49-001e4fc686da}"
Private Const IID_IOutputStream As String = "{905a0fe6-bc53-11df-8c49-001e4fc686da}"
Private Const IID_IRandomAccessStream As String = "{905a0fe1-bc53-11df-8c49-001e4fc686da}"
Private Const IID_IRandomAccessStreamStatics As String = "{524cedcf-6e29-4ce5-9573-6b753db66c3a}"

' * Any object which implements this interface must also implement the following interfaces:
' *     Windows.Foundation.IClosable

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IRandomAccessStreamStatics
'    IRandomAccessStreamStatics_CopyAsync = 6
'    IRandomAccessStreamStatics_CopySizeAsync = 7
'    IRandomAccessStreamStatics_CopyAndCloseAsync = 8

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
    
End Enum

' ----==== Variablen ====----
Private m_pIInputStream As Long
Private m_pIOutputStream As Long
Private m_pIRandomAccessStream As Long
Private m_pIRandomAccessStreamStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_Streams_RandomAccessStream, _
                            IID_IRandomAccessStreamStatics, _
                            m_pIRandomAccessStreamStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIInputStream)
    Call ReleaseIfc(m_pIOutputStream)
    Call ReleaseIfc(m_pIRandomAccessStreamStatics)
    Call DisposeIfc(m_pIRandomAccessStream) '<-IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIRandomAccessStream
End Property

Public Property Let Ifc(ByVal pIRandomAccessStream As Long)
    Call ReleaseIfc(m_pIInputStream)
    Call ReleaseIfc(m_pIOutputStream)
    Call DisposeIfc(m_pIRandomAccessStream) '<-IClosable
    m_pIRandomAccessStream = pIRandomAccessStream
    Call QueryIfc(m_pIRandomAccessStream, _
                  IID_IOutputStream, _
                  m_pIOutputStream)
    Call QueryIfc(m_pIRandomAccessStream, _
                  IID_IInputStream, _
                  m_pIInputStream)
End Property

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

' ----==== Functions ====----
' IRandomAccessStreamStatics
'    IRandomAccessStreamStatics_CopyAsync = 6
'    IRandomAccessStreamStatics_CopySizeAsync = 7
'    IRandomAccessStreamStatics_CopyAndCloseAsync = 8

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

Public Function FromIStream(ByVal pIStream As Long, _
                            ByVal Options As BSOS_OPTIONS) As RandomAccessStream
    Dim Ret As RandomAccessStream
    If pIStream <> 0& Then
        Dim pIRandomAccessStream As Long
        If CreateRandomAccessStreamOverStream(pIStream, _
                                              Options, _
                                              Str2Guid(IID_IRandomAccessStream), _
                                              pIRandomAccessStream) = S_OK Then
            If pIRandomAccessStream <> 0& Then
                Set Ret = New RandomAccessStream
                Ret.Ifc = pIRandomAccessStream
            End If
        End If
    End If
    Set FromIStream = Ret
End Function

Public Function FromFile(ByVal fileName As String, _
                         ByVal accessMode As FileAccessMode) As RandomAccessStream
    Dim Ret As RandomAccessStream
    Dim pIRandomAccessStream As Long
    If CreateRandomAccessStreamOnFile(StrPtr(fileName), _
                                      accessMode, _
                                      Str2Guid(IID_IRandomAccessStream), _
                                      pIRandomAccessStream) = S_OK Then
        If pIRandomAccessStream <> 0& Then
            Set Ret = New RandomAccessStream
            Ret.Ifc = pIRandomAccessStream
        End If
    End If
    Set FromFile = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

