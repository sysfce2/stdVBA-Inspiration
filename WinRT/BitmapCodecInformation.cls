VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapCodecInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_BitmapCodecInformation As String = "Windows.Graphics.Imaging.BitmapCodecInformation"
Private Const IID_IBitmapCodecInformation As String = "{400caaf2-c4b0-4392-a3b0-6f6f9ba95cb4}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IBitmapCodecInformation
    IBitmapCodecInformation_GetCodecId = 6
    IBitmapCodecInformation_GetFileExtensions = 7
    IBitmapCodecInformation_GetFriendlyName = 8
    IBitmapCodecInformation_GetMimeTypes = 9
    
End Enum

' ----==== Variablen ====----
Private m_pIBitmapCodecInformation As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBitmapCodecInformation)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBitmapCodecInformation
End Property

Public Property Let Ifc(ByVal pIBitmapCodecInformation As Long)
    Call ReleaseIfc(m_pIBitmapCodecInformation)
    m_pIBitmapCodecInformation = pIBitmapCodecInformation
End Property

' IBitmapCodecInformation
Friend Property Get CodecID() As GUID
    Dim Ret As GUID
    If m_pIBitmapCodecInformation <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapCodecInformation, _
                  IBitmapCodecInformation_GetCodecId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    CodecID = Ret
End Property

Public Property Get FileExtensions() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIBitmapCodecInformation <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIBitmapCodecInformation, _
                  IBitmapCodecInformation_GetFileExtensions, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set FileExtensions = Ret
End Property

Public Property Get FriendlyName() As String
    Dim Ret As String
    If m_pIBitmapCodecInformation <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapCodecInformation, _
                  IBitmapCodecInformation_GetFriendlyName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    FriendlyName = Ret
End Property

Public Property Get MimeTypes() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIBitmapCodecInformation <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIBitmapCodecInformation, _
                  IBitmapCodecInformation_GetMimeTypes, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set MimeTypes = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


