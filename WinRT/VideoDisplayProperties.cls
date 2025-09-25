VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoDisplayProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_VideoDisplayProperties As String = "Windows.Media.VideoDisplayProperties"
Private Const IID_IVideoDisplayProperties As String = "{5609fdb1-5d2d-4872-8170-45dee5bc2f5c}"
Private Const IID_IVideoDisplayProperties2 As String = "{b410e1ce-ab52-41ab-a486-cc10fab152f9}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVideoDisplayProperties
    IVideoDisplayProperties_GetTitle = 6
    IVideoDisplayProperties_PutTitle = 7
    IVideoDisplayProperties_GetSubtitle = 8
    IVideoDisplayProperties_PutSubtitle = 9
                    
    ' IVideoDisplayProperties2
    IVideoDisplayProperties2_GetGenres = 6

End Enum

' ----==== Variablen ====----
Private m_pIVideoDisplayProperties As Long
Private m_pIVideoDisplayProperties2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVideoDisplayProperties2)
    Call ReleaseIfc(m_pIVideoDisplayProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVideoDisplayProperties
End Property

Public Property Let Ifc(ByVal pIVideoDisplayProperties As Long)
    Call ReleaseIfc(m_pIVideoDisplayProperties2)
    Call ReleaseIfc(m_pIVideoDisplayProperties)
    m_pIVideoDisplayProperties = pIVideoDisplayProperties
    Call QueryIfc(m_pIVideoDisplayProperties, IID_IVideoDisplayProperties2, m_pIVideoDisplayProperties2)
End Property

' IVideoDisplayProperties
Public Property Get Title() As String
    Dim Ret As String
    If m_pIVideoDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoDisplayProperties, _
                  IVideoDisplayProperties_GetTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pIVideoDisplayProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIVideoDisplayProperties, _
                  IVideoDisplayProperties_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Subtitle() As String
    Dim Ret As String
    If m_pIVideoDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoDisplayProperties, _
                  IVideoDisplayProperties_GetSubtitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Subtitle = Ret
End Property

Public Property Let Subtitle(ByVal value As String)
    If m_pIVideoDisplayProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIVideoDisplayProperties, _
                  IVideoDisplayProperties_PutSubtitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property
                    
' IVideoDisplayProperties2
Public Property Get Genres() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIVideoDisplayProperties2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIVideoDisplayProperties2, _
                  IVideoDisplayProperties2_GetGenres, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Genres = Ret
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


