VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapPropertiesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_BitmapProperties As String = "Windows.Graphics.Imaging.BitmapProperties"
Private Const IID_IBitmapPropertiesView As String = "{7e0fe87a-3a70-48f8-9c55-196cf5a545f5}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IBitmapPropertiesView
    IBitmapPropertiesView_GetPropertiesAsync = 6
    
End Enum

' ----==== Vars ====----
Private m_pIBitmapPropertiesView As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBitmapPropertiesView)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBitmapPropertiesView
End Property

Public Property Let Ifc(ByVal pIBitmapPropertiesView As Long)
    Call ReleaseIfc(m_pIBitmapPropertiesView)
    m_pIBitmapPropertiesView = pIBitmapPropertiesView
End Property

' ----==== Functions ====----
' IBitmapPropertiesView
'    IBitmapPropertiesView_GetPropertiesAsync = 6

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
