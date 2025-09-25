VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlNamedNodeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Dom_XmlNamedNodeMap As String = "Windows.Data.Xml.Dom.XmlNamedNodeMap"
Private Const IID_IXmlNamedNodeMap As String = "{b3a69eb0-aab0-4b82-a6fa-b1453f7c021b}"

'Private Const IID_IVectorView_IXmlNode As String = "{139d959e-e7b5-5cb6-a596-4b544478da9b}"
'Private Const IID_IIterable_IXmlNode As String = "{f1146ffc-8c92-56e8-93f1-711f86722633}"

' * Class implements the following interfaces:
' *    Windows.Foundation.Collections.IVectorView`1<Windows.Data.Xml.Dom.IXmlNode>
' *    Windows.Foundation.Collections.IIterable`1<Windows.Data.Xml.Dom.IXmlNode>

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IXmlNamedNodeMap
    IXmlNamedNodeMap_GetLength = 6
    IXmlNamedNodeMap_Item = 7
    IXmlNamedNodeMap_GetNamedItem = 8
    IXmlNamedNodeMap_SetNamedItem = 9
    IXmlNamedNodeMap_RemoveNamedItem = 10
    IXmlNamedNodeMap_GetNamedItemNS = 11
    IXmlNamedNodeMap_RemoveNamedItemNS = 12
    IXmlNamedNodeMap_SetNamedItemNS = 13

End Enum
' ----==== Variablen ====----
Private m_pIXmlNamedNodeMap As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXmlNamedNodeMap)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXmlNamedNodeMap
End Property

Public Property Let Ifc(ByVal pIXmlNamedNodeMap As Long)
    Call ReleaseIfc(m_pIXmlNamedNodeMap)
    m_pIXmlNamedNodeMap = pIXmlNamedNodeMap
End Property

' IXmlNamedNodeMap
Public Property Get Length() As Long
    Dim Ret As Long
    If m_pIXmlNamedNodeMap <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_GetLength, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Length = Ret
End Property

' ----==== Functions ====----
' IXmlNamedNodeMap
Public Function Item(ByVal index As Long) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNamedNodeMap <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_Item, _
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

Public Function GetNamedItem(ByVal name As String) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNamedNodeMap <> 0& Then
        Dim hName As Long
        Dim pIXmlNode As Long
        hName = CreateWindowsString(name)
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_GetNamedItem, _
                  hName, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
        Call DeleteWindowsString(hName)
    End If
    Set GetNamedItem = Ret
End Function

Public Function SetNamedItem(ByVal node As XmlNode) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNamedNodeMap <> 0& And _
       IsNotNothing(node) Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_SetNamedItem, _
                  node.Ifc, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set SetNamedItem = Ret
End Function

Public Function RemoveNamedItem(ByVal name As String) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNamedNodeMap <> 0& Then
        Dim hName As Long
        Dim pIXmlNode As Long
        hName = CreateWindowsString(name)
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_RemoveNamedItem, _
                  hName, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
        Call DeleteWindowsString(hName)
    End If
    Set RemoveNamedItem = Ret
End Function

Public Function GetNamedItemNS(ByVal namespaceUri As Inspectable, _
                               ByVal name As String) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNamedNodeMap <> 0& And _
       IsNotNothing(namespaceUri) Then
        Dim hName As Long
        Dim pIXmlNode As Long
        hName = CreateWindowsString(name)
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_GetNamedItemNS, _
                  namespaceUri.Ifc, _
                  hName, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
        Call DeleteWindowsString(hName)
    End If
    Set GetNamedItemNS = Ret
End Function

Public Function RemoveNamedItemNS(ByVal namespaceUri As Inspectable, _
                                  ByVal name As String) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNamedNodeMap <> 0& And _
       IsNotNothing(namespaceUri) Then
        Dim hName As Long
        Dim pIXmlNode As Long
        hName = CreateWindowsString(name)
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_RemoveNamedItemNS, _
                  namespaceUri.Ifc, _
                  hName, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
        Call DeleteWindowsString(hName)
    End If
    Set RemoveNamedItemNS = Ret
End Function

Public Function SetNamedItemNS(ByVal node As XmlNode) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNamedNodeMap <> 0& And _
       IsNotNothing(node) Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNamedNodeMap, _
                  IXmlNamedNodeMap_SetNamedItemNS, _
                  node.Ifc, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set SetNamedItemNS = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

