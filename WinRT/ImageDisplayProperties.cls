VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageDisplayProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_ImageDisplayProperties As String = "Windows.Media.ImageDisplayProperties"
Private Const IID_IImageDisplayProperties As String = "{cd0bc7ef-54e7-411f-9933-f0e98b0a96d2}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IImageDisplayProperties
    IImageDisplayProperties_GetTitle = 6
    IImageDisplayProperties_PutTitle = 7
    IImageDisplayProperties_GetSubtitle = 8
    IImageDisplayProperties_PutSubtitle = 9

End Enum

' ----==== Variablen ====----
Private m_pIImageDisplayProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIImageDisplayProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIImageDisplayProperties
End Property

Public Property Let Ifc(ByVal pIImageDisplayProperties As Long)
    Call ReleaseIfc(m_pIImageDisplayProperties)
    m_pIImageDisplayProperties = pIImageDisplayProperties
End Property

' IImageDisplayProperties
Public Property Get Title() As String
    Dim Ret As String
    If m_pIImageDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageDisplayProperties, _
                  IImageDisplayProperties_GetTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pIImageDisplayProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIImageDisplayProperties, _
                  IImageDisplayProperties_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Subtitle() As String
    Dim Ret As String
    If m_pIImageDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIImageDisplayProperties, _
                  IImageDisplayProperties_GetSubtitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Subtitle = Ret
End Property

Public Property Let Subtitle(ByVal value As String)
    If m_pIImageDisplayProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIImageDisplayProperties, _
                  IImageDisplayProperties_PutSubtitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
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


