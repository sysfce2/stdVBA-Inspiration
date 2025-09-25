VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

' Namespace Dom

Option Explicit

' ----==== Enums Windows.Data.Xml.Dom.xxx ====----
Public Enum NodeType
    NodeType_Invalid = 0
    NodeType_ElementNode = 1
    NodeType_AttributeNode = 2
    NodeType_TextNode = 3
    NodeType_DataSectionNode = 4
    NodeType_EntityReferenceNode = 5
    NodeType_EntityNode = 6
    NodeType_ProcessingInstructionNode = 7
    NodeType_CommentNode = 8
    NodeType_DocumentNode = 9
    NodeType_DocumentTypeNode = 10
    NodeType_DocumentFragmentNode = 11
    NodeType_NotationNode = 12
End Enum

' ----==== Namespace Dom.XmlDocument ====----
Public Property Get XmlDocument() As XmlDocument
    Set XmlDocument = New XmlDocument
End Property

' ----==== Namespace Dom.XmlLoadSettings ====----
Public Property Get XmlLoadSettings() As XmlLoadSettings
    Set XmlLoadSettings = New XmlLoadSettings
End Property

' ----==== Namespace Dom.XmlNode ====----
Public Property Get XmlNode() As XmlNode
    Set XmlNode = New XmlNode
End Property

' ----==== Namespace Dom.XmlElement ====----
Public Property Get XmlElement() As XmlElement
    Set XmlElement = New XmlElement
End Property

' ----==== Namespace Dom.XmlNodeList ====----
Public Property Get XmlNodeList() As XmlNodeList
    Set XmlNodeList = New XmlNodeList
End Property

' ----==== Namespace Dom.XmlText ====----
Public Property Get XmlText() As XmlText
    Set XmlText = New XmlText
End Property

' ----==== Namespace Dom.XmlNamedNodeMap ====----
Public Property Get XmlNamedNodeMap() As XmlNamedNodeMap
    Set XmlNamedNodeMap = New XmlNamedNodeMap
End Property

' ----==== Namespace Dom.XmlAttribute ====----
Public Property Get XmlAttribute() As XmlAttribute
    Set XmlAttribute = New XmlAttribute
End Property

' ----==== Namespace Dom.XmlDocumentType ====----
Public Property Get XmlDocumentType() As XmlDocumentType
    Set XmlDocumentType = New XmlDocumentType
End Property

' ----==== Namespace Dom.XmlDomImplementation ====----
Public Property Get XmlDomImplementation() As XmlDomImplementation
    Set XmlDomImplementation = New XmlDomImplementation
End Property

' ----==== Namespace Dom.XmlDocumentFragment ====----
Public Property Get XmlDocumentFragment() As XmlDocumentFragment
    Set XmlDocumentFragment = New XmlDocumentFragment
End Property

' ----==== Namespace Dom.XmlComment ====----
Public Property Get XmlComment() As XmlComment
    Set XmlComment = New XmlComment
End Property

' ----==== Namespace Dom.XmlEntityReference ====----
Public Property Get XmlEntityReference() As XmlEntityReference
    Set XmlEntityReference = New XmlEntityReference
End Property

' ----==== Namespace Dom.XmlCDataSection ====----
Public Property Get XmlCDataSection() As XmlCDataSection
    Set XmlCDataSection = New XmlCDataSection
End Property

