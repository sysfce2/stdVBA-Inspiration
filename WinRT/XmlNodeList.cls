VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlNodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Dom_XmlNodeList As String = "Windows.Data.Xml.Dom.XmlNodeList"
Private Const IID_IXmlNodeList As String = "{8c60ad77-83a4-4ec1-9c54-7ba429e13da6}"

'Private Const IID_IVectorView_IXmlNode As String = "{139d959e-e7b5-5cb6-a596-4b544478da9b}"
'Private Const IID_IIterable_IXmlNode As String = "{f1146ffc-8c92-56e8-93f1-711f86722633}"

' * Class implements the following interfaces:
' *    Windows.Foundation.Collections.IVectorView`1<Windows.Data.Xml.Dom.IXmlNode>
' *    Windows.Foundation.Collections.IIterable`1<Windows.Data.Xml.Dom.IXmlNode>

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IXmlNodeList
    IXmlNodeList_GetLength = 6
    IXmlNodeList_Item = 7
    
End Enum
' ----==== Variablen ====----
Private m_pIXmlNodeList As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXmlNodeList)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXmlNodeList
End Property

Public Property Let Ifc(ByVal pIXmlNodeList As Long)
    Call ReleaseIfc(m_pIXmlNodeList)
    m_pIXmlNodeList = pIXmlNodeList
End Property

' IXmlNodeList
Public Property Get Length() As Long
    Dim Ret As Long
    If m_pIXmlNodeList <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNodeList, _
                  IXmlNodeList_GetLength, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Length = Ret
End Property

' ----==== Functions ====----
' IXmlNodeList
Public Function Item(ByVal index As Long) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNodeList <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNodeList, _
                  IXmlNodeList_Item, _
                  index, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set Item = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

