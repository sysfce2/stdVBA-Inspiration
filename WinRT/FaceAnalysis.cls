VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaceAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

' Namespace FaceAnalysis

Option Explicit

' ----==== Enums Windows.Media.FaceAnalysis.xxx ====----

' ----==== Namespace FaceAnalysis.FaceDetector ====----
Public Property Get FaceDetector() As FaceDetector
    Set FaceDetector = New FaceDetector
End Property

' ----==== Namespace FaceAnalysis.DetectedFace ====----
Public Property Get DetectedFace() As DetectedFace
    Set DetectedFace = New DetectedFace
End Property

