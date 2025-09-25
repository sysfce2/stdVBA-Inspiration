VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVectorView_xxx
    IVectorView_xxx_GetAt = 6
    IVectorView_xxx_GetSize = 7
    
End Enum

Public Enum OfType_xxx
    Of_AdaptiveNotificationText
    Of_AlternateWordForm
    Of_AppDiagnosticInfo
    Of_AppResourceGroupInfo
    Of_AudioTrack
    Of_BitmapCodecInformation
    Of_BitmapPixelFormat
    Of_DetectedFace
    Of_Geoposition
    Of_GlobalSystemMediaTransportControlsSession
    Of_Language
    Of_MediaBreak
    Of_MediaCaptureVideoProfile
    Of_MediaPlaybackItem
    Of_MediaTimeRange
    Of_NotificationBinding
    Of_OcrLine
    Of_OcrWord
    Of_Package
    Of_ProcessDiagnosticInfo
    Of_ScheduledToastNotification
    Of_SelectableWordSegment
    Of_SortEntry
    Of_SpeechRecognitionResult
    Of_StorageFile
    Of_StorageFolder
    Of_StorageItem
    Of_StorageLibraryChange
    Of_String
    Of_ToastNotification
    Of_User
    Of_UserNotification
    Of_VideoTrack
    Of_VoiceInformation
    Of_WordSegment
End Enum

' ----==== Types ====----
Private Type SortEntry_
    PropertyName As Long
    AscendingOrder As Long
End Type

Private Type MediaTimeRange_
    Start As Currency
    End As Currency
End Type

' ----==== Variablen ====----
Private m_OfType_xxx As OfType_xxx
Private m_pIVectorView_xxx As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    m_OfType_xxx = -1
    Call ReleaseIfc(m_pIVectorView_xxx)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVectorView_xxx
End Property

Public Property Let Ifc(ByVal pIVectorView_xxx As Long)
    Call ReleaseIfc(m_pIVectorView_xxx)
    m_pIVectorView_xxx = pIVectorView_xxx
End Property

Public Property Let Of(ByVal ReturnAs As OfType_xxx)
    m_OfType_xxx = ReturnAs
End Property

'IVectorView_xxx
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVectorView_xxx <> 0& Then
        Dim value As Long
        If Invoke(m_pIVectorView_xxx, _
                  IVectorView_xxx_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVectorView_xxx
Public Function GetAt(ByVal index As Long) As Variant
    Dim Ret As Variant
    If m_pIVectorView_xxx <> 0& Then
        Dim value As Long
        If Invoke(m_pIVectorView_xxx, _
                  IVectorView_xxx_GetAt, _
                  index, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Select Case m_OfType_xxx
                    Case OfType_xxx.Of_AdaptiveNotificationText
                        Set Ret = New AdaptiveNotificationText
                        Ret.Ifc = value
                    Case OfType_xxx.Of_AlternateWordForm
                        Set Ret = New AlternateWordForm
                        Ret.Ifc = value
                    Case OfType_xxx.Of_AppDiagnosticInfo
                        Set Ret = New AppDiagnosticInfo
                        Ret.Ifc = value
                    Case OfType_xxx.Of_AppResourceGroupInfo
                        Set Ret = New AppResourceGroupInfo
                        Ret.Ifc = value
                    Case OfType_xxx.Of_AudioTrack
                        Set Ret = New AudioTrack
                        Ret.Ifc = value
                    Case OfType_xxx.Of_BitmapCodecInformation
                        Set Ret = New BitmapCodecInformation
                        Ret.Ifc = value
                    Case OfType_xxx.Of_BitmapPixelFormat
                        Ret = value
                    Case OfType_xxx.Of_DetectedFace
                        Set Ret = New DetectedFace
                        Ret.Ifc = value
                    Case OfType_xxx.Of_Geoposition
                        Set Ret = New Geoposition
                        Ret.Ifc = value
                    Case OfType_xxx.Of_GlobalSystemMediaTransportControlsSession
                        Set Ret = New GlobalSystemMediaTransportControlsSession
                        Ret.Ifc = value
                    Case OfType_xxx.Of_Language
                        Set Ret = New Language
                        Ret.Ifc = value
                    Case OfType_xxx.Of_MediaBreak
                        Set Ret = New MediaBreak
                        Ret.Ifc = value
                    Case OfType_xxx.Of_MediaCaptureVideoProfile
                        Set Ret = New MediaCaptureVideoProfile
                        Ret.Ifc = value
                    Case OfType_xxx.Of_MediaPlaybackItem
                        Set Ret = New MediaPlaybackItem
                        Ret.Ifc = value
                    Case OfType_xxx.Of_MediaTimeRange
                        Set Ret = New MediaTimeRange
                        Dim tMediaTimeRange As MediaTimeRange_
                        Call CopyMemory(tMediaTimeRange, value, LenB(tMediaTimeRange))
                        Dim TimeRangeStart As New TimeSpan
                        Dim TimeRangeEnd As New TimeSpan
                        TimeRangeStart.TimeSpanValue = tMediaTimeRange.Start
                        TimeRangeEnd.TimeSpanValue = tMediaTimeRange.End
                        Ret.StartTime = TimeRangeStart
                        Ret.EndTime = TimeRangeEnd
                    Case OfType_xxx.Of_NotificationBinding
                        Set Ret = New NotificationBinding
                        Ret.Ifc = value
                    Case OfType_xxx.Of_OcrLine
                        Set Ret = New OcrLine
                        Ret.Ifc = value
                    Case OfType_xxx.Of_OcrWord
                        Set Ret = New OcrWord
                        Ret.Ifc = value
                    Case OfType_xxx.Of_Package
                        Set Ret = New Package
                        Ret.Ifc = value
                    Case OfType_xxx.Of_ProcessDiagnosticInfo
                        Set Ret = New ProcessDiagnosticInfo
                        Ret.Ifc = value
                    Case OfType_xxx.Of_ScheduledToastNotification
                        Set Ret = New ScheduledToastNotification
                        Ret.Ifc = value
                    Case OfType_xxx.Of_SelectableWordSegment
                        Set Ret = New SelectableWordSegment
                        Ret.Ifc = value
                    Case OfType_xxx.Of_SortEntry
                        Set Ret = New SortEntry
                        Dim tSortEntry As SortEntry_
                        Call CopyMemory(tSortEntry, value, LenB(tSortEntry))
                        Ret.PropertyName = GetWindowsString(tSortEntry.PropertyName)
                        Ret.AscendingOrder = CBool(tSortEntry.AscendingOrder)
                    Case OfType_xxx.Of_SpeechRecognitionResult
                        Set Ret = New SpeechRecognitionResult
                        Ret.Ifc = value
                    Case OfType_xxx.Of_StorageFile
                        Set Ret = New StorageFile
                        Ret.Ifc = value
                    Case OfType_xxx.Of_StorageFolder
                        Set Ret = New StorageFolder
                        Ret.Ifc = value
                    Case OfType_xxx.Of_StorageItem
                        Set Ret = New StorageItem
                        Ret.Ifc = value
                    Case OfType_xxx.Of_StorageLibraryChange
                        Set Ret = New StorageLibraryChange
                        Ret.Ifc = value
                    Case OfType_xxx.Of_String
                        Ret = GetWindowsString(value)
                    Case OfType_xxx.Of_ToastNotification
                        Set Ret = New ToastNotification
                        Ret.Ifc = value
                    Case OfType_xxx.Of_User
                        Set Ret = New User
                        Ret.Ifc = value
                    Case OfType_xxx.Of_UserNotification
                        Set Ret = New UserNotification
                        Ret.Ifc = value
                    Case OfType_xxx.Of_VideoTrack
                        Set Ret = New VideoTrack
                        Ret.Ifc = value
                    Case OfType_xxx.Of_VoiceInformation
                        Set Ret = New VoiceInformation
                        Ret.Ifc = value
                    Case OfType_xxx.Of_WordSegment
                        Set Ret = New WordSegment
                        Ret.Ifc = value
                End Select
            End If
        End If
    End If
    If VarType(Ret) = vbObject Then
        Set GetAt = Ret
    Else
        GetAt = Ret
    End If
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As Variant()
    Dim Ret() As Variant
    Dim count As Long
    count = Me.Size
    If count > 0& Then
        ReDim Ret(count - 1)
        Dim Item As Long
        For Item = 0 To count - 1
            Select Case m_OfType_xxx
                Case OfType_xxx.Of_String, OfType_xxx.Of_BitmapPixelFormat
                    Ret(Item) = Me.GetAt(Item)
                Case Else
                    Set Ret(Item) = Me.GetAt(Item)
            End Select
        Next
    End If
    GetAll = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function
