VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomAccessStreamReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_Streams_RandomAccessStreamReference As String = "Windows.Storage.Streams.RandomAccessStreamReference"
Private Const IID_IRandomAccessStreamReference As String = "{33ee3134-1dd6-4e3a-8067-d1c162e8642b}"
Private Const IID_IRandomAccessStreamReferenceStatics As String = "{857309dc-3fbf-4e7d-986f-ef3b1a07a964}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IRandomAccessStreamReferenceStatics
    IRandomAccessStreamReferenceStatics_CreateFromFile = 6
    IRandomAccessStreamReferenceStatics_CreateFromUri = 7
    IRandomAccessStreamReferenceStatics_CreateFromStream = 8
    
    ' IRandomAccessStreamReference
    IRandomAccessStreamReference_OpenReadAsync = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIRandomAccessStreamReference As Long
Private m_pIRandomAccessStreamReferenceStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Storage_Streams_RandomAccessStreamReference, _
                            IID_IRandomAccessStreamReferenceStatics, _
                            m_pIRandomAccessStreamReferenceStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIRandomAccessStreamReference)
    Call ReleaseIfc(m_pIRandomAccessStreamReferenceStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIRandomAccessStreamReference
End Property

Public Property Let Ifc(ByVal pIRandomAccessStreamReference As Long)
    Call ReleaseIfc(m_pIRandomAccessStreamReference)
    m_pIRandomAccessStreamReference = pIRandomAccessStreamReference
End Property

' ----==== Functions ====----
' IRandomAccessStreamReferenceStatics
Public Function CreateFromFile(ByVal file As StorageFile) As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIRandomAccessStreamReferenceStatics <> 0& And _
       IsNotNothing(file) Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIRandomAccessStreamReferenceStatics, _
                  IRandomAccessStreamReferenceStatics_CreateFromFile, _
                  file.Ifc, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                Set Ret = New RandomAccessStreamReference
                Ret.Ifc = pIRandomAccessStreamReference
            End If
        End If
    End If
    Set CreateFromFile = Ret
End Function

Public Function CreateFromUri(ByVal Uri As Uri) As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIRandomAccessStreamReferenceStatics <> 0& And _
       IsNotNothing(Uri) Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIRandomAccessStreamReferenceStatics, _
                  IRandomAccessStreamReferenceStatics_CreateFromUri, _
                  Uri.Ifc, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                Set Ret = New RandomAccessStreamReference
                Ret.Ifc = pIRandomAccessStreamReference
            End If
        End If
    End If
    Set CreateFromUri = Ret
End Function

Public Function CreateFromStream(ByVal stream As RandomAccessStream) As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIRandomAccessStreamReferenceStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIRandomAccessStreamReferenceStatics, _
                  IRandomAccessStreamReferenceStatics_CreateFromStream, _
                  stream.Ifc, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                Set Ret = New RandomAccessStreamReference
                Ret.Ifc = pIRandomAccessStreamReference
            End If
        End If
    End If
    Set CreateFromStream = Ret
End Function

' IRandomAccessStreamReference
Public Function OpenReadAsync() As RandomAccessStreamWithContentType
    Dim Ret As RandomAccessStreamWithContentType
    If m_pIRandomAccessStreamReference <> 0& Then
        Dim pIRandomAccessStreamWithContentType As Long
        If Invoke(m_pIRandomAccessStreamReference, _
                  IRandomAccessStreamReference_OpenReadAsync, _
                  VarPtr(pIRandomAccessStreamWithContentType)) = S_OK Then
            If pIRandomAccessStreamWithContentType <> 0& Then
                If Await(pIRandomAccessStreamWithContentType) Then
                    Set Ret = New RandomAccessStreamWithContentType
                    Ret.Ifc = pIRandomAccessStreamWithContentType
                End If
            End If
        End If
    End If
    Set OpenReadAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


