VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptographicBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Security_Cryptography_CryptographicBuffer As String = "Windows.Security.Cryptography.CryptographicBuffer"
Private Const IID_ICryptographicBufferStatics As String = "{320b7e22-3cb0-4cdf-8663-1d28910065eb}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ICryptographicBufferStatics
    ICryptographicBufferStatics_Compare = 6
    ICryptographicBufferStatics_GenerateRandom = 7
    ICryptographicBufferStatics_GenerateRandomNumber = 8
    ICryptographicBufferStatics_CreateFromByteArray = 9
    ICryptographicBufferStatics_CopyToByteArray = 10
    ICryptographicBufferStatics_DecodeFromHexString = 11
    ICryptographicBufferStatics_EncodeToHexString = 12
    ICryptographicBufferStatics_DecodeFromBase64String = 13
    ICryptographicBufferStatics_EncodeToBase64String = 14
    ICryptographicBufferStatics_ConvertStringToBinary = 15
    ICryptographicBufferStatics_ConvertBinaryToString = 16
    
End Enum

' ----==== Variablen ====----
Private m_pICryptographicBufferStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Security_Cryptography_CryptographicBuffer, _
                            IID_ICryptographicBufferStatics, _
                            m_pICryptographicBufferStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICryptographicBufferStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICryptographicBufferStatics
End Property

' ----==== Functions ====----
' ICryptographicBufferStatics
Public Function Compare(ByVal object1 As Buffer, _
                        ByVal object2 As Buffer) As Boolean
    Dim Ret As Boolean
    If m_pICryptographicBufferStatics <> 0& And _
       IsNotNothing(object1) And _
       IsNotNothing(object2) Then
        Dim Value As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_Compare, _
                  object1.Ifc, _
                  object2.Ifc, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    Compare = Ret
End Function

Public Function GenerateRandom(ByVal Length As Long) As Buffer
    Dim Ret As Buffer
    If m_pICryptographicBufferStatics <> 0& Then
        Dim pIBuffer As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_GenerateRandom, _
                  Length, _
                  VarPtr(pIBuffer)) = S_OK Then
            If pIBuffer <> 0& Then
                Set Ret = New Buffer
                Ret.Ifc = pIBuffer
            End If
        End If
    End If
    Set GenerateRandom = Ret
End Function

Public Function GenerateRandomNumber() As Long
    Dim Ret As Long
    If m_pICryptographicBufferStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_GenerateRandomNumber, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    GenerateRandomNumber = Ret
End Function

Public Function CreateFromByteArray(ByVal valueLength As Long, _
                                    ByRef Value() As Byte) As Buffer
    Dim Ret As Buffer
    If m_pICryptographicBufferStatics <> 0& Then
        Dim pIBuffer As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_CreateFromByteArray, _
                  valueLength, _
                  VarPtr(Value(0)), _
                  VarPtr(pIBuffer)) = S_OK Then
            If pIBuffer <> 0& Then
                Set Ret = New Buffer
                Ret.Ifc = pIBuffer
            End If
        End If
    End If
    Set CreateFromByteArray = Ret
End Function

Public Function CopyToByteArray(ByVal Buffer As Buffer) As Byte()
    Dim Ret() As Byte
    If m_pICryptographicBufferStatics <> 0& And _
       IsNotNothing(Buffer) Then
        Dim pRet As Long
        Dim valueLength As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_CopyToByteArray, _
                  Buffer.Ifc, _
                  VarPtr(valueLength), _
                  VarPtr(pRet)) = S_OK Then
            If pRet <> 0& Then
                ReDim Ret(valueLength - 1)
                Call CopyMemory(Ret(0), ByVal pRet, valueLength)
            End If
        End If
    End If
    CopyToByteArray = Ret
End Function

Public Function DecodeFromHexString(ByVal Value As String) As Buffer
    Dim Ret As Buffer
    If m_pICryptographicBufferStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        Dim pIBuffer As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_DecodeFromHexString, _
                  hString, _
                  VarPtr(pIBuffer)) = S_OK Then
            If pIBuffer <> 0& Then
                Set Ret = New Buffer
                Ret.Ifc = pIBuffer
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set DecodeFromHexString = Ret
End Function

Public Function EncodeToHexString(ByVal Buffer As Buffer) As String
    Dim Ret As String
    If m_pICryptographicBufferStatics <> 0& And _
       IsNotNothing(Buffer) Then
        Dim hString As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_EncodeToHexString, _
                  Buffer.Ifc, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    EncodeToHexString = Ret
End Function

Public Function DecodeFromBase64String(ByVal Value As String) As Buffer
    Dim Ret As Buffer
    If m_pICryptographicBufferStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        Dim pIBuffer As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_DecodeFromBase64String, _
                  hString, _
                  VarPtr(pIBuffer)) = S_OK Then
            If pIBuffer <> 0& Then
                Set Ret = New Buffer
                Ret.Ifc = pIBuffer
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set DecodeFromBase64String = Ret
End Function

Public Function EncodeToBase64String(ByVal Buffer As Buffer) As String
    Dim Ret As String
    If m_pICryptographicBufferStatics <> 0& And _
       IsNotNothing(Buffer) Then
        Dim hString As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_EncodeToBase64String, _
                  Buffer.Ifc, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    EncodeToBase64String = Ret
End Function

Public Function ConvertStringToBinary(ByVal Value As String, _
                                      ByVal encoding As BinaryStringEncoding) As Buffer
    Dim Ret As Buffer
    If m_pICryptographicBufferStatics <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        Dim pIBuffer As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_ConvertStringToBinary, _
                  hString, _
                  encoding, _
                  VarPtr(pIBuffer)) = S_OK Then
            If pIBuffer <> 0& Then
                Set Ret = New Buffer
                Ret.Ifc = pIBuffer
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set ConvertStringToBinary = Ret
End Function

Public Function ConvertBinaryToString(ByVal encoding As BinaryStringEncoding, _
                                      ByVal Buffer As Buffer) As String
    Dim Ret As String
    If m_pICryptographicBufferStatics <> 0& And _
       IsNotNothing(Buffer) Then
        Dim hString As Long
        If Invoke(m_pICryptographicBufferStatics, _
                  ICryptographicBufferStatics_ConvertBinaryToString, _
                  encoding, _
                  Buffer.Ifc, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    ConvertBinaryToString = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


