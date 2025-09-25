VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinRT-Samples"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1218
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command66 
      Caption         =   "Test"
      Height          =   525
      Left            =   15210
      TabIndex        =   65
      Top             =   5760
      Width           =   3000
   End
   Begin VB.CommandButton Command65 
      Caption         =   "GraphicsCapturePicker"
      Height          =   525
      Left            =   12180
      TabIndex        =   64
      Top             =   5760
      Width           =   3000
   End
   Begin VB.CommandButton Command64 
      Caption         =   "UserDataPaths"
      Height          =   525
      Left            =   9150
      TabIndex        =   63
      Top             =   5760
      Width           =   3000
   End
   Begin VB.CommandButton Command63 
      Caption         =   "KnownFolders"
      Height          =   525
      Left            =   6120
      TabIndex        =   62
      Top             =   5760
      Width           =   3000
   End
   Begin VB.CommandButton Command62 
      Caption         =   "LanguageFontGroup"
      Height          =   525
      Left            =   3090
      TabIndex        =   61
      Top             =   5760
      Width           =   3000
   End
   Begin VB.CommandButton Command61 
      Caption         =   "GeographicRegion"
      Height          =   525
      Left            =   60
      TabIndex        =   60
      Top             =   5760
      Width           =   3000
   End
   Begin VB.CommandButton Command60 
      Caption         =   "PhoneNumberFormatter"
      Height          =   525
      Left            =   15210
      TabIndex        =   59
      Top             =   5190
      Width           =   3000
   End
   Begin VB.CommandButton Command59 
      Caption         =   "PhoneNumberInfo"
      Height          =   525
      Left            =   12180
      TabIndex        =   58
      Top             =   5190
      Width           =   3000
   End
   Begin VB.CommandButton Command58 
      Caption         =   "IncrementNumberRounder"
      Height          =   525
      Left            =   9150
      TabIndex        =   57
      Top             =   5190
      Width           =   3000
   End
   Begin VB.CommandButton Command57 
      Caption         =   "SignificantDigitsNumberRounder"
      Height          =   525
      Left            =   6120
      TabIndex        =   56
      Top             =   5190
      Width           =   3000
   End
   Begin VB.CommandButton Command56 
      Caption         =   "NumeralSystemTranslator"
      Height          =   525
      Left            =   3090
      TabIndex        =   55
      Top             =   5190
      Width           =   3000
   End
   Begin VB.CommandButton Command55 
      Caption         =   "PermilleFormatter"
      Height          =   525
      Left            =   60
      TabIndex        =   54
      Top             =   5190
      Width           =   3000
   End
   Begin VB.CommandButton Command54 
      Caption         =   "PercentFormatter"
      Height          =   525
      Left            =   15210
      TabIndex        =   53
      Top             =   4620
      Width           =   3000
   End
   Begin VB.CommandButton Command53 
      Caption         =   "DecimalFormatter"
      Height          =   525
      Left            =   12180
      TabIndex        =   52
      Top             =   4620
      Width           =   3000
   End
   Begin VB.CommandButton Command52 
      Caption         =   "CurrencyFormatter"
      Height          =   525
      Left            =   9150
      TabIndex        =   51
      Top             =   4620
      Width           =   3000
   End
   Begin VB.CommandButton Command51 
      Caption         =   "DateTimeFormatter"
      Height          =   525
      Left            =   6120
      TabIndex        =   50
      Top             =   4620
      Width           =   3000
   End
   Begin VB.CommandButton Command50 
      Caption         =   "MouseCapabilities"
      Height          =   525
      Left            =   3090
      TabIndex        =   49
      Top             =   4620
      Width           =   3000
   End
   Begin VB.CommandButton Command49 
      Caption         =   "KeyboardCapabilities"
      Height          =   525
      Left            =   60
      TabIndex        =   48
      Top             =   4620
      Width           =   3000
   End
   Begin VB.CommandButton Command48 
      Caption         =   "SpeechRecognition"
      Height          =   525
      Left            =   15210
      TabIndex        =   47
      Top             =   4050
      Width           =   3000
   End
   Begin VB.CommandButton Command47 
      Caption         =   "SpeechSynthesizer"
      Height          =   525
      Left            =   12180
      TabIndex        =   46
      Top             =   4050
      Width           =   3000
   End
   Begin VB.CommandButton Command46 
      Caption         =   "MediaPlayer (Audio)"
      Height          =   525
      Left            =   9150
      TabIndex        =   45
      Top             =   4050
      Width           =   3000
   End
   Begin VB.CommandButton Command45 
      Caption         =   "GlobalSystemMediaTransport"
      Height          =   525
      Left            =   6120
      TabIndex        =   44
      Top             =   4050
      Width           =   3000
   End
   Begin VB.CommandButton Command44 
      Caption         =   "User"
      Height          =   525
      Left            =   3090
      TabIndex        =   43
      Top             =   4050
      Width           =   3000
   End
   Begin VB.CommandButton Command43 
      Caption         =   "ToastNotification"
      Height          =   525
      Left            =   60
      TabIndex        =   42
      Top             =   4050
      Width           =   3000
   End
   Begin VB.CommandButton Command42 
      Caption         =   "XmlDocument"
      Height          =   525
      Left            =   15210
      TabIndex        =   41
      Top             =   3480
      Width           =   3000
   End
   Begin VB.CommandButton Command41 
      Caption         =   "SelectableWordsSegmenter"
      Height          =   525
      Left            =   12180
      TabIndex        =   40
      Top             =   3480
      Width           =   3000
   End
   Begin VB.CommandButton Command40 
      Caption         =   "WordsSegmenter"
      Height          =   525
      Left            =   9150
      TabIndex        =   39
      Top             =   3480
      Width           =   3000
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Geolocator"
      Height          =   525
      Left            =   6120
      TabIndex        =   38
      Top             =   3480
      Width           =   3000
   End
   Begin VB.CommandButton Command38 
      Caption         =   "FileSavePicker"
      Height          =   525
      Left            =   3090
      TabIndex        =   37
      Top             =   3480
      Width           =   3000
   End
   Begin VB.CommandButton Command37 
      Caption         =   "BitmapEncoder Info"
      Height          =   525
      Left            =   60
      TabIndex        =   36
      Top             =   3480
      Width           =   3000
   End
   Begin VB.CommandButton Command36 
      Caption         =   "BitmapDecoder Info"
      Height          =   525
      Left            =   15210
      TabIndex        =   35
      Top             =   2910
      Width           =   3000
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Convert Image"
      Height          =   525
      Left            =   12180
      TabIndex        =   34
      Top             =   2910
      Width           =   3000
   End
   Begin VB.CommandButton Command34 
      Caption         =   "FaceDetector"
      Height          =   525
      Left            =   9150
      TabIndex        =   33
      Top             =   2910
      Width           =   3000
   End
   Begin VB.CommandButton Command33 
      Caption         =   "AppDiagnosticInfo"
      Height          =   525
      Left            =   6120
      TabIndex        =   32
      Top             =   2910
      Width           =   3000
   End
   Begin VB.CommandButton Command32 
      Caption         =   "ProcessDiagnosticInfo"
      Height          =   525
      Left            =   3090
      TabIndex        =   31
      Top             =   2910
      Width           =   3000
   End
   Begin VB.CommandButton Command31 
      Caption         =   "SystemDiagnosticInfo"
      Height          =   525
      Left            =   60
      TabIndex        =   30
      Top             =   2910
      Width           =   3000
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Ocr"
      Height          =   525
      Left            =   15210
      TabIndex        =   29
      Top             =   2340
      Width           =   3000
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Search Files in Folder"
      Height          =   525
      Left            =   12180
      TabIndex        =   28
      Top             =   2340
      Width           =   3000
   End
   Begin VB.CommandButton Command28 
      Caption         =   "CastingDevicePicker"
      Height          =   525
      Left            =   9150
      TabIndex        =   27
      Top             =   2340
      Width           =   3000
   End
   Begin VB.CommandButton Command27 
      Caption         =   "DevicePicker"
      Height          =   525
      Left            =   6120
      TabIndex        =   26
      Top             =   2340
      Width           =   3000
   End
   Begin VB.CommandButton Command26 
      Caption         =   "PdfViewer"
      Height          =   525
      Left            =   3090
      TabIndex        =   25
      Top             =   2340
      Width           =   3000
   End
   Begin VB.CommandButton Command25 
      Caption         =   "HtmlUtilities"
      Height          =   525
      Left            =   60
      TabIndex        =   24
      Top             =   2340
      Width           =   3000
   End
   Begin VB.CommandButton Command24 
      Caption         =   "UserNotificationListener"
      Height          =   525
      Left            =   15210
      TabIndex        =   23
      Top             =   1770
      Width           =   3000
   End
   Begin VB.CommandButton Command23 
      Caption         =   "PopupMenu"
      Height          =   525
      Left            =   12180
      TabIndex        =   22
      Top             =   1770
      Width           =   3000
   End
   Begin VB.CommandButton Command22 
      Caption         =   "CoreWindowFlyout"
      Height          =   525
      Left            =   9150
      TabIndex        =   21
      Top             =   1770
      Width           =   3000
   End
   Begin VB.CommandButton Command21 
      Caption         =   "CoreWindowDialog"
      Height          =   525
      Left            =   6120
      TabIndex        =   20
      Top             =   1770
      Width           =   3000
   End
   Begin VB.CommandButton Command20 
      Caption         =   "MessageDialog"
      Height          =   525
      Left            =   3090
      TabIndex        =   19
      Top             =   1770
      Width           =   3000
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H0000FFFF&
      Caption         =   "Create Json String"
      Height          =   525
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1770
      Width           =   3000
   End
   Begin VB.CommandButton Command18 
      Caption         =   "CryptographicBuffer"
      Height          =   525
      Left            =   15210
      TabIndex        =   17
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Launch Uri with ApplicationPicker"
      Height          =   525
      Left            =   12180
      TabIndex        =   16
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Launch Uri"
      Height          =   525
      Left            =   9150
      TabIndex        =   15
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Launch File with ApplicationPicker"
      Height          =   525
      Left            =   6120
      TabIndex        =   14
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Launch File"
      Height          =   525
      Left            =   3090
      TabIndex        =   13
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Launch Drive/Folder with selection"
      Height          =   525
      Left            =   60
      TabIndex        =   12
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Launch Drive/Folder"
      Height          =   525
      Left            =   15210
      TabIndex        =   11
      Top             =   630
      Width           =   3000
   End
   Begin VB.CommandButton Command11 
      Caption         =   "WinRT ApiInformation"
      Height          =   525
      Left            =   12180
      TabIndex        =   10
      Top             =   630
      Width           =   3000
   End
   Begin VB.CommandButton Command10 
      Caption         =   "FileProperties"
      Height          =   525
      Left            =   9150
      TabIndex        =   9
      Top             =   630
      Width           =   3000
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Calendar"
      Height          =   525
      Left            =   6120
      TabIndex        =   8
      Top             =   630
      Width           =   3000
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "Start ThreadPoolTimer"
      Height          =   525
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   630
      Width           =   3000
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Transcode Audio"
      Height          =   525
      Left            =   60
      TabIndex        =   6
      Top             =   630
      Width           =   3000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Transcode Video"
      Height          =   525
      Left            =   15210
      TabIndex        =   5
      Top             =   60
      Width           =   3000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FolderPicker"
      Height          =   525
      Left            =   12180
      TabIndex        =   4
      Top             =   60
      Width           =   3000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FileOpenPicker (Multi Select)"
      Height          =   525
      Left            =   9150
      TabIndex        =   3
      Top             =   60
      Width           =   3000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FileOpenPicker (Single Select)"
      Height          =   525
      Left            =   6120
      TabIndex        =   2
      Top             =   60
      Width           =   3000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Start VideoCapture from WebCam"
      Height          =   525
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PhotoCapture from WebCam"
      Height          =   525
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' Namespace Windows
Private Windows As New Windows

' Namespace Windows.System.Threading
Private WithEvents ThreadPoolTimer As ThreadPoolTimer
Attribute ThreadPoolTimer.VB_VarHelpID = -1

' Namespace Windows.Media.Transcoding
Private PrepareTranscodeResult As PrepareTranscodeResult
Private WithEvents AsyncActionWithProgress_Double As AsyncActionWithProgress_Double
Attribute AsyncActionWithProgress_Double.VB_VarHelpID = -1

Private WithEvents UICommandInvokedHandler As UICommandInvokedHandler
Attribute UICommandInvokedHandler.VB_VarHelpID = -1

Private CastingDeviceSelectedEventHandler As Long
Private CastingDeviceSelectedEventCookie As Currency
Private CastingDevicePickerDismissedHandler As Long
Private CastingDevicePickerDismissedCookie As Currency

Private Sub Form_Load()
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsNotNothing(ThreadPoolTimer) Then
        Call ThreadPoolTimer.Cancel
        Set ThreadPoolTimer = Nothing
    End If
    Set Windows = Nothing
End Sub

Private Sub Command1_Click()
    Dim MediaCapture As MediaCapture
    Dim FolderPicker As FolderPicker
    Set FolderPicker = Windows.Storage.Pickers.FolderPicker
    If IsNotNothing(FolderPicker) Then
        FolderPicker.ParentHwnd = Me.hwnd
        Dim StorageFolder As StorageFolder
        Set StorageFolder = FolderPicker.PickSingleFolderAsync
        If IsNotNothing(StorageFolder) Then
            Set MediaCapture = Windows.Media.Capture.MediaCapture
            If IsNotNothing(MediaCapture) Then
                If MediaCapture.InitializeAsync Then
                    Dim ImageEncodingProperties As ImageEncodingProperties
                    Set ImageEncodingProperties = Windows.Media.MediaProperties.ImageEncodingProperties.CreatePng
                    If IsNotNothing(ImageEncodingProperties) Then
                        ImageEncodingProperties.Width = 1280
                        ImageEncodingProperties.Height = 960
                        Dim StorageFile As StorageFile
                        Set StorageFile = StorageFolder.CreateFileAsync("Capture.png", CreationCollisionOption_GenerateUniqueName)
                        If IsNotNothing(StorageFile) Then
                            Command1.Enabled = False
                            If MediaCapture.CapturePhotoToStorageFileAsync(ImageEncodingProperties, StorageFile) Then
                                Debug.Print "PhotoCapture from WebCam = OK"
                                Command1.Enabled = True
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Debug.Print "FolderPicker = Cancel"
        End If
    End If
    Set MediaCapture = Nothing
End Sub

Private Sub Command2_Click()
    Static IsRecord As Boolean
    Static MediaCapture As MediaCapture
    If Not IsRecord Then
        Set MediaCapture = Nothing
        Dim FolderPicker As FolderPicker
        Set FolderPicker = Windows.Storage.Pickers.FolderPicker
        If IsNotNothing(FolderPicker) Then
            FolderPicker.ParentHwnd = Me.hwnd
            Dim StorageFolder As StorageFolder
            Set StorageFolder = FolderPicker.PickSingleFolderAsync
            If IsNotNothing(StorageFolder) Then
                Set MediaCapture = Windows.Media.Capture.MediaCapture
                If IsNotNothing(MediaCapture) Then
                    If MediaCapture.InitializeAsync Then
                        Dim MediaEncodingProfile As MediaEncodingProfile
                        Set MediaEncodingProfile = Windows.Media.MediaProperties.MediaEncodingProfile.CreateMp4(VideoEncodingQuality_HD1080p)
                        If IsNotNothing(MediaEncodingProfile) Then
                            Dim StorageFile As StorageFile
                            Set StorageFile = StorageFolder.CreateFileAsync("Capture.mp4", CreationCollisionOption_GenerateUniqueName)
                            If IsNotNothing(StorageFile) Then
                                If MediaCapture.StartRecordToStorageFileAsync(MediaEncodingProfile, StorageFile) Then
                                    IsRecord = True
                                    Command2.BackColor = vbRed
                                    Command2.Caption = "Stop VideoCapture from WebCam"
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                Debug.Print "FolderPicker = Cancel"
            End If
        End If
    Else
        If MediaCapture.StopRecordAsync Then
            IsRecord = False
            Command2.BackColor = vbGreen
            Command2.Caption = "Start VideoCapture from WebCam"
            Set MediaCapture = Nothing
        End If
    End If
End Sub

Private Sub Command3_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = FileOpenPicker.FileTypeFilter
        Call FileTypeFilter.Append(".jpg")
        Call FileTypeFilter.Append(".gif")
        Call FileTypeFilter.Append(".bmp")
        Dim StorageFile As StorageFile
        Set StorageFile = FileOpenPicker.PickSingleFileAsync
        If IsNotNothing(StorageFile) Then
            Debug.Print StorageFile.Path
        Else
            Debug.Print "FileOpenPicker = Cancel"
        End If
    End If
End Sub

Private Sub Command4_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Call FileOpenPicker.FileTypeFilter.Clear
        Dim StorageFileList As ReadOnlyList_1 'ReadOnlyList_StorageFile
        Set StorageFileList = FileOpenPicker.PickMultipleFilesAsync
        If IsNotNothing(StorageFileList) Then
            Dim FileCount As Long
            FileCount = StorageFileList.Size
            If FileCount > 0 Then
                Dim vStorageFile As Variant 'StorageFile
                For Each vStorageFile In StorageFileList.GetAll
                   Debug.Print "StorageFile.Path: " & vStorageFile.Path
                Next
            End If
        Else
            Debug.Print "FileOpenPicker = Cancel"
        End If
    End If
End Sub

Private Sub Command5_Click()
    Dim FolderPicker As FolderPicker
    Set FolderPicker = Windows.Storage.Pickers.FolderPicker
    If IsNotNothing(FolderPicker) Then
        FolderPicker.ParentHwnd = Me.hwnd
        Dim StorageFolder As StorageFolder
        Set StorageFolder = FolderPicker.PickSingleFolderAsync
        If IsNotNothing(StorageFolder) Then
            Debug.Print StorageFolder.Path
        Else
            Debug.Print "FolderPicker = Cancel"
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = FileOpenPicker.FileTypeFilter
        Call FileTypeFilter.Append(".avi")
        Call FileTypeFilter.Append(".mp4")
        Call FileTypeFilter.Append(".wmv")
        Call FileTypeFilter.Append(".wav")
        Call FileTypeFilter.Append(".mp3")
        Call FileTypeFilter.Append(".wma")
        Dim StorageFileIn As StorageFile
        Set StorageFileIn = FileOpenPicker.PickSingleFileAsync
        If IsNotNothing(StorageFileIn) Then
            Dim FolderPicker As FolderPicker
            Set FolderPicker = Windows.Storage.Pickers.FolderPicker
            If IsNotNothing(FolderPicker) Then
                FolderPicker.ParentHwnd = Me.hwnd
                Dim StorageFolder As StorageFolder
                Set StorageFolder = FolderPicker.PickSingleFolderAsync
                If IsNotNothing(StorageFolder) Then
                    Dim StorageFileOut As StorageFile
                    Set StorageFileOut = StorageFolder.CreateFileAsync("Converted.mp4", CreationCollisionOption_GenerateUniqueName)
                    If IsNotNothing(StorageFileOut) Then
                        Dim MediaTranscoder As MediaTranscoder
                        Set MediaTranscoder = Windows.Media.Transcoding.MediaTranscoder
                        If IsNotNothing(MediaTranscoder) Then
                            MediaTranscoder.HardwareAccelerationEnabled = True
                            Dim MediaEncodingProfile As MediaEncodingProfile
                            Set MediaEncodingProfile = Windows.Media.MediaProperties.MediaEncodingProfile.CreateMp4(VideoEncodingQuality_HD1080p)
                            If IsNotNothing(MediaEncodingProfile) Then
                                Set PrepareTranscodeResult = MediaTranscoder.PrepareFileTranscodeAsync(StorageFileIn, StorageFileOut, MediaEncodingProfile)
                                If IsNotNothing(PrepareTranscodeResult) Then
                                    If PrepareTranscodeResult.CanTranscode Then
                                        Set AsyncActionWithProgress_Double = PrepareTranscodeResult.TranscodeAsync
                                        If IsNotNothing(AsyncActionWithProgress_Double) Then
                                            Command6.Enabled = False
                                        End If
                                    Else
                                        Debug.Print "PrepareTranscodeResult.FailureReason: " & CStr(PrepareTranscodeResult.FailureReason)
                                        Set PrepareTranscodeResult = Nothing
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    Debug.Print "FolderPicker = Cancel"
                End If
            End If
        Else
            Debug.Print "FileOpenPicker = Cancel"
        End If
    End If
End Sub

Private Sub Command7_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = FileOpenPicker.FileTypeFilter
        Call FileTypeFilter.Append(".avi")
        Call FileTypeFilter.Append(".mp4")
        Call FileTypeFilter.Append(".wmv")
        Call FileTypeFilter.Append(".wav")
        Call FileTypeFilter.Append(".mp3")
        Call FileTypeFilter.Append(".wma")
        Dim StorageFileIn As StorageFile
        Set StorageFileIn = FileOpenPicker.PickSingleFileAsync
        If IsNotNothing(StorageFileIn) Then
            Dim FolderPicker As FolderPicker
            Set FolderPicker = Windows.Storage.Pickers.FolderPicker
            If IsNotNothing(FolderPicker) Then
                FolderPicker.ParentHwnd = Me.hwnd
                Dim StorageFolder As StorageFolder
                Set StorageFolder = FolderPicker.PickSingleFolderAsync
                If IsNotNothing(StorageFolder) Then
                    Dim StorageFileOut As StorageFile
                    Set StorageFileOut = StorageFolder.CreateFileAsync("Converted.wav", CreationCollisionOption_GenerateUniqueName)
                    'Set StorageFileOut = StorageFolder.CreateFileAsync("Converted.flac", CreationCollisionOption_GenerateUniqueName)
                    If IsNotNothing(StorageFileOut) Then
                        Dim MediaTranscoder As MediaTranscoder
                        Set MediaTranscoder = Windows.Media.Transcoding.MediaTranscoder
                        If IsNotNothing(MediaTranscoder) Then
                            MediaTranscoder.HardwareAccelerationEnabled = True
                            Dim MediaEncodingProfile As MediaEncodingProfile
                            Set MediaEncodingProfile = Windows.Media.MediaProperties.MediaEncodingProfile.CreateWav(AudioEncodingQuality_High)
                            'Set MediaEncodingProfile = Windows.Media.MediaProperties.MediaEncodingProfile.CreateFlac(AudioEncodingQuality_High)
                            If IsNotNothing(MediaEncodingProfile) Then
                                Set PrepareTranscodeResult = MediaTranscoder.PrepareFileTranscodeAsync(StorageFileIn, StorageFileOut, MediaEncodingProfile)
                                If IsNotNothing(PrepareTranscodeResult) Then
                                    If PrepareTranscodeResult.CanTranscode Then
                                        Set AsyncActionWithProgress_Double = PrepareTranscodeResult.TranscodeAsync
                                        If IsNotNothing(AsyncActionWithProgress_Double) Then
                                            Command7.Enabled = False
                                        End If
                                    Else
                                        Debug.Print "PrepareTranscodeResult.FailureReason: " & CStr(PrepareTranscodeResult.FailureReason)
                                        Set PrepareTranscodeResult = Nothing
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    Debug.Print "FolderPicker = Cancel"
                End If
            End If
        Else
            Debug.Print "Cancel"
        End If
    End If
End Sub

Private Sub Command8_Click()
    If IsNothing(ThreadPoolTimer) Then
'        Set ThreadPoolTimer = Windows.System.Threading.ThreadPoolTimer.CreateTimer(2000) ' without Event TimerDestroyed
'        Set ThreadPoolTimer = Windows.System.Threading.ThreadPoolTimer.CreateTimerWithCompletion(2000) ' with Event TimerDestroyed
'        Set ThreadPoolTimer = Windows.System.Threading.ThreadPoolTimer.CreatePeriodicTimer(500) ' without Event TimerDestroyed
        Set ThreadPoolTimer = Windows.System.Threading.ThreadPoolTimer.CreatePeriodicTimerWithCompletion(250) ' with Event TimerDestroyed
        If IsNotNothing(ThreadPoolTimer) Then
            Command8.BackColor = vbRed
            Command8.Caption = "Stop ThreadPoolTimer"
        End If
    Else
        ThreadPoolTimer.Cancel
        Set ThreadPoolTimer = Nothing
        Command8.BackColor = vbGreen
        Command8.Caption = "Start ThreadPoolTimer"
    End If
End Sub

Private Sub Command9_Click()
    Dim Calendar As Calendar
    Set Calendar = Windows.Globalization.Calendar
    If IsNotNothing(Calendar) Then
        Debug.Print Calendar.DayAsPaddedString & "." & _
                    Calendar.MonthAsFullString & " " & _
                    Calendar.YearAsPaddedString
        Debug.Print Calendar.HourAsPaddedString & ":" & _
                    Calendar.MinuteAsPaddedString & ":" & _
                    Calendar.SecondAsPaddedString & "." & _
                    Calendar.NanosecondAsPaddedString
        Debug.Print "Calendar.GetDateTime: " & Calendar.GetDateTime.VbDate
        Debug.Print "Calendar.GetCalendarSystem: " & Calendar.GetCalendarSystem
        Debug.Print "Calendar.GetTimeZone: " & Calendar.GetTimeZone
        Debug.Print "Calendar.GetClock: " & Calendar.GetClock
    End If
End Sub

Private Sub Command10_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Dim PropItem As Long
        Dim StorageFile As StorageFile
        Dim ReadOnlyList As ReadOnlyList_1 'ReadOnlyList_String
        Set StorageFile = FileOpenPicker.PickSingleFileAsync
        If IsNotNothing(StorageFile) Then
            Debug.Print "StorageFile.Attributes: " & StorageFile.Attributes
            Debug.Print "StorageFile.ContentType: " & StorageFile.ContentType
            Debug.Print "StorageFile.DateCreated: " & StorageFile.DateCreated.VbDate
            Debug.Print "StorageFile.DisplayName: " & StorageFile.DisplayName
            Debug.Print "StorageFile.DisplayType: " & StorageFile.DisplayType
            Debug.Print "StorageFile.FileType: " & StorageFile.FileType
            Debug.Print "StorageFile.FolderRelativeId: " & StorageFile.FolderRelativeId
            Debug.Print "StorageFile.IsAvailable: " & StorageFile.IsAvailable
            Debug.Print "StorageFile.Name: " & StorageFile.Name
            Debug.Print "StorageFile.Path: " & StorageFile.Path
            Debug.Print "StorageFile.Provider.DisplayName: " & StorageFile.Provider.DisplayName
            Debug.Print "StorageFile.Provider.Id: " & StorageFile.Provider.Id
            Dim BasicProperties As BasicProperties
            Set BasicProperties = StorageFile.GetBasicPropertiesAsync
            If IsNotNothing(BasicProperties) Then
                Debug.Print "StorageFile.BasicProperties.DateModified: " & BasicProperties.DateModified.VbDate
                Debug.Print "StorageFile.BasicProperties.ItemDate: " & BasicProperties.ItemDate.VbDate
                Debug.Print "StorageFile.BasicProperties.Size: " & BasicProperties.Size
            End If
            If InStr(1, StorageFile.ContentType, "audio") Then
                Dim MusicProperties As MusicProperties
                Set MusicProperties = StorageFile.Properties.GetMusicPropertiesAsync
                If IsNotNothing(MusicProperties) Then
                    Debug.Print "StorageFile.MusicProperties.Album: " & MusicProperties.Album
                    Debug.Print "StorageFile.MusicProperties.AlbumArtist: " & MusicProperties.AlbumArtist
                    Debug.Print "StorageFile.MusicProperties.Artist: " & MusicProperties.Artist
                    Debug.Print "StorageFile.MusicProperties.Bitrate: " & MusicProperties.Bitrate
                    Debug.Print "StorageFile.MusicProperties.Duration: " & MusicProperties.Duration.VbDate
                    Debug.Print "StorageFile.MusicProperties.Publisher: " & MusicProperties.Publisher
                    Debug.Print "StorageFile.MusicProperties.Rating: " & MusicProperties.Rating
                    Debug.Print "StorageFile.MusicProperties.Subtitle: " & MusicProperties.Subtitle
                    Debug.Print "StorageFile.MusicProperties.Title: " & MusicProperties.Title
                    Debug.Print "StorageFile.MusicProperties.TrackNumber: " & MusicProperties.TrackNumber
                    Debug.Print "StorageFile.MusicProperties.Year: " & MusicProperties.Year
                    Set ReadOnlyList = MusicProperties.Composers
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.MusicProperties.Composers: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = MusicProperties.Conductors
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.MusicProperties.Conductors: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = MusicProperties.Genre
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.MusicProperties.Genre: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = MusicProperties.Producers
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.MusicProperties.Producers: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = MusicProperties.Writers
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.MusicProperties.Writers: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                End If
            ElseIf InStr(1, StorageFile.ContentType, "video") Then
                Dim VideoProperties As VideoProperties
                Set VideoProperties = StorageFile.Properties.GetVideoPropertiesAsync
                If IsNotNothing(VideoProperties) Then
                    Debug.Print "StorageFile.VideoProperties.Bitrate: " & VideoProperties.Bitrate
                    Debug.Print "StorageFile.VideoProperties.Duration: " & VideoProperties.Duration.VbDate
                    Debug.Print "StorageFile.VideoProperties.Width: " & VideoProperties.Width
                    Debug.Print "StorageFile.VideoProperties.Height: " & VideoProperties.Height
                    Debug.Print "StorageFile.VideoProperties.Latitude: " & VideoProperties.Latitude
                    Debug.Print "StorageFile.VideoProperties.Longitude: " & VideoProperties.Longitude
                    Debug.Print "StorageFile.VideoProperties.Orientation: " & VideoProperties.Orientation
                    Debug.Print "StorageFile.VideoProperties.Publisher: " & VideoProperties.Publisher
                    Debug.Print "StorageFile.VideoProperties.Rating: " & VideoProperties.Rating
                    Debug.Print "StorageFile.VideoProperties.Subtitle: " & VideoProperties.Subtitle
                    Debug.Print "StorageFile.VideoProperties.Title: " & VideoProperties.Title
                    Debug.Print "StorageFile.VideoProperties.Year: " & VideoProperties.Year
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = VideoProperties.Directors
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.VideoProperties.Directors: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = VideoProperties.Keywords
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.VideoProperties.Keywords: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = VideoProperties.Producers
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.VideoProperties.Producers: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = VideoProperties.Writers
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.VideoProperties.Writers: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                End If
            ElseIf InStr(1, StorageFile.ContentType, "image") Then
                Dim ImageProperties As ImageProperties
                Set ImageProperties = StorageFile.Properties.GetImagePropertiesAsync
                If IsNotNothing(ImageProperties) Then
                    Debug.Print "StorageFile.ImageProperties.CameraManufacturer: " & ImageProperties.CameraManufacturer
                    Debug.Print "StorageFile.ImageProperties.CameraModel: " & ImageProperties.CameraModel
                    Debug.Print "StorageFile.ImageProperties.DateTaken: " & ImageProperties.DateTaken.VbDate
                    Debug.Print "StorageFile.ImageProperties.Width: " & ImageProperties.Width
                    Debug.Print "StorageFile.ImageProperties.Height: " & ImageProperties.Height
                    Debug.Print "StorageFile.ImageProperties.Latitude: " & ImageProperties.Latitude
                    Debug.Print "StorageFile.ImageProperties.Longitude: " & ImageProperties.Longitude
                    Debug.Print "StorageFile.ImageProperties.Orientation: " & ImageProperties.Orientation
                    Debug.Print "StorageFile.ImageProperties.Rating: " & ImageProperties.Rating
                    Debug.Print "StorageFile.ImageProperties.Title: " & ImageProperties.Title
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = ImageProperties.Keywords
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.ImageProperties.Keywords: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = ImageProperties.PeopleNames
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.ImageProperties.PeopleNames: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                End If
            Else
                Dim DocumentProperties As DocumentProperties
                Set DocumentProperties = StorageFile.Properties.GetDocumentPropertiesAsync
                If IsNotNothing(DocumentProperties) Then
                    Debug.Print "StorageFile.DocumentProperties.Title: " & DocumentProperties.Title
                    Debug.Print "StorageFile.DocumentProperties.Comment: " & DocumentProperties.Comment
                    Set ReadOnlyList = Nothing
                    Set ReadOnlyList = DocumentProperties.Author
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.DocumentProperties.Author: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                    Set ReadOnlyList = DocumentProperties.Keywords
                    Set ReadOnlyList = Nothing
                    If IsNotNothing(ReadOnlyList) Then
                        If ReadOnlyList.Size > 0& Then
                            For PropItem = 0 To ReadOnlyList.Size - 1
                                Debug.Print "StorageFile.DocumentProperties.Comment: " & ReadOnlyList.GetAt(PropItem)
                            Next
                        End If
                    End If
                End If
            End If
        Else
            Debug.Print "FileOpenPicker = Cancel"
        End If
    End If
End Sub

Private Sub Command11_Click()
    Dim ApiInformation As ApiInformation
    Set ApiInformation = Windows.Foundation.Metadata.ApiInformation
    If IsNotNothing(ApiInformation) Then
        Debug.Print "ApiInformation.IsApiContractPresentByMajor: " & ApiInformation.IsApiContractPresentByMajor("Windows.ApplicationModel.Calls.CallsVoipContract", 1)
        Debug.Print "ApiInformation.IsApiContractPresentByMajorAndMinor: " & ApiInformation.IsApiContractPresentByMajorAndMinor("Windows.ApplicationModel.Calls.CallsVoipContract", 1, 1)
        Debug.Print "ApiInformation.IsEnumNamedValuePresent: " & ApiInformation.IsEnumNamedValuePresent("Windows.UI.Xaml.Automation.Peers.AutomationControlType", "ComboBox")
        ' usw.
    End If
End Sub

Private Sub Command12_Click()
    Dim Launcher As Launcher
    Set Launcher = Windows.System.Launcher
    If IsNotNothing(Launcher) Then
        If Launcher.LaunchFolderPathAsync("C:\Windows") Then
            Debug.Print "Ok"
        Else
            Debug.Print "Not Ok"
        End If
    End If
End Sub

Private Sub Command13_Click()
    Dim StorageFile As StorageFile
    Set StorageFile = Windows.Storage.StorageFile.GetFileFromPathAsync("C:\Windows\explorer.exe")
    If IsNotNothing(StorageFile) Then
        Dim StorageFolder As StorageFolder
        Set StorageFolder = Windows.Storage.StorageFolder.GetFolderFromPathAsync("C:\Windows\System32")
        If IsNotNothing(StorageFolder) Then
            Dim FolderLauncherOptions As FolderLauncherOptions
            Set FolderLauncherOptions = Windows.System.FolderLauncherOptions
            If IsNotNothing(FolderLauncherOptions) Then
                FolderLauncherOptions.ParentHwnd = Me.hwnd
                Dim List_Insprectable As List_Insprectable
                Set List_Insprectable = FolderLauncherOptions.ItemsToSelect
                If IsNotNothing(List_Insprectable) Then
                    Call List_Insprectable.Append(StorageFile)
                    Call List_Insprectable.Append(StorageFolder)
                    Dim Launcher As Launcher
                    Set Launcher = Windows.System.Launcher
                    If IsNotNothing(Launcher) Then
                        If Launcher.LaunchFolderPathWithOptionsAsync("C:\Windows", FolderLauncherOptions) Then
                            Debug.Print "Ok"
                        Else
                            Debug.Print "Not Ok"
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Command14_Click()
    Dim StorageFile As StorageFile
    Set StorageFile = Windows.Storage.StorageFile.GetFileFromPathAsync("C:\Windows\win.ini")
    If IsNotNothing(StorageFile) Then
        Dim Launcher As Launcher
        Set Launcher = Windows.System.Launcher
        If IsNotNothing(Launcher) Then
            If Launcher.LaunchFileAsync(StorageFile) Then
                Debug.Print "Ok"
            Else
                Debug.Print "Not Ok"
            End If
        End If
    End If
End Sub

Private Sub Command15_Click()
    Dim StorageFile As StorageFile
    Set StorageFile = Windows.Storage.StorageFile.GetFileFromPathAsync("C:\Windows\win.ini")
    If IsNotNothing(StorageFile) Then
        Dim LauncherOptions As LauncherOptions
        Set LauncherOptions = Windows.System.LauncherOptions
        If IsNotNothing(LauncherOptions) Then
            LauncherOptions.DisplayApplicationPicker = True
            LauncherOptions.ParentHwnd = Me.hwnd
            Dim Launcher As Launcher
            Set Launcher = Windows.System.Launcher
            If IsNotNothing(Launcher) Then
                If Launcher.LaunchFileWithOptionsAsync(StorageFile, LauncherOptions) Then
                    Debug.Print "Ok"
                Else
                    Debug.Print "Not Ok"
                End If
            End If
        End If
    End If
End Sub

Private Sub Command16_Click()
    Dim Uri As Uri
    Set Uri = Windows.Foundation.Uri.CreateUri("https://www.google.de")
    If IsNotNothing(Uri) Then
        Dim Launcher As Launcher
        Set Launcher = Windows.System.Launcher
        If IsNotNothing(Launcher) Then
            If Launcher.LaunchUriAsync(Uri) Then
                Debug.Print "Ok"
            Else
                Debug.Print "Not Ok"
            End If
        End If
    End If
End Sub

Private Sub Command17_Click()
    Dim Uri As Uri
    Set Uri = Windows.Foundation.Uri.CreateUri("https://www.google.de")
    If IsNotNothing(Uri) Then
        Dim LauncherOptions As LauncherOptions
        Set LauncherOptions = Windows.System.LauncherOptions
        If IsNotNothing(LauncherOptions) Then
            LauncherOptions.DisplayApplicationPicker = True
            LauncherOptions.ParentHwnd = Me.hwnd
            Dim Launcher As Launcher
            Set Launcher = Windows.System.Launcher
            If IsNotNothing(Launcher) Then
                If Launcher.LaunchUriWithOptionsAsync(Uri, LauncherOptions) Then
                    Debug.Print "Ok"
                Else
                    Debug.Print "Not Ok"
                End If
            End If
        End If
    End If
End Sub

Private Sub Command18_Click()
    Dim TestString As String
    TestString = "MyTestString"
    Dim Buffer As Buffer
    Set Buffer = Windows.Security.Cryptography.CryptographicBuffer.ConvertStringToBinary(TestString, BinaryStringEncoding_Utf16BE)
    If IsNotNothing(Buffer) Then
    
'        Dim StringArray() As Byte
'        StringArray = Windows.Security.Cryptography.CryptographicBuffer.CopyToByteArray(Buffer)
'        Debug.Print StringArray
    
        Debug.Print "Utf16BE-String -> String = " & Windows.Security.Cryptography.CryptographicBuffer.ConvertBinaryToString(BinaryStringEncoding_Utf16BE, Buffer)
        Dim Base64String As String
        Base64String = Windows.Security.Cryptography.CryptographicBuffer.EncodeToBase64String(Buffer)
        Debug.Print "String -> Base64String = " & Base64String
        Set Buffer = Windows.Security.Cryptography.CryptographicBuffer.DecodeFromBase64String(Base64String)
        If IsNotNothing(Buffer) Then
            Debug.Print "Base64String -> String = " & Windows.Security.Cryptography.CryptographicBuffer.ConvertBinaryToString(BinaryStringEncoding_Utf16BE, Buffer)
        End If
        Dim HexString As String
        HexString = Windows.Security.Cryptography.CryptographicBuffer.EncodeToHexString(Buffer)
        Debug.Print "String -> HexString = " & HexString
        Set Buffer = Windows.Security.Cryptography.CryptographicBuffer.DecodeFromHexString(HexString)
        If IsNotNothing(Buffer) Then
            Debug.Print "HexString -> String = " & Windows.Security.Cryptography.CryptographicBuffer.ConvertBinaryToString(BinaryStringEncoding_Utf16BE, Buffer)
        End If
    End If
    Set Buffer = Windows.Security.Cryptography.CryptographicBuffer.GenerateRandom(8)
    If IsNotNothing(Buffer) Then
        Dim ByteArray() As Byte
        ByteArray = Windows.Security.Cryptography.CryptographicBuffer.CopyToByteArray(Buffer)
        Dim ArrayItem As Long
        Debug.Print "---=== Random ByteArray ===---"
        For ArrayItem = 0 To UBound(ByteArray)
            Debug.Print "ByteArray(" & CStr(ArrayItem) & ") = " & CStr(ByteArray(ArrayItem))
        Next
        Debug.Print "---========================---"
        Dim Buffer2 As Buffer
        Set Buffer2 = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(Buffer.Length, ByteArray)
        If IsNotNothing(Buffer2) Then
            Debug.Print "Compare(Buffer, Buffer2) = " & CStr(Windows.Security.Cryptography.CryptographicBuffer.Compare(Buffer, Buffer2))
        End If
    End If
    Debug.Print "Random Number: " & CStr(Windows.Security.Cryptography.CryptographicBuffer.GenerateRandomNumber)
End Sub

Private Sub Command19_Click()
    Static JsonString As String
    Dim JsonValue As JsonValue
    Dim JsonArray As JsonArray
    Dim JsonObject As JsonObject
    Dim JsonObject2 As JsonObject
    Set JsonValue = Windows.Data.Json.JsonValue
    Set JsonArray = Windows.Data.Json.JsonArray
    Set JsonObject = Windows.Data.Json.JsonObject
    Set JsonObject2 = Windows.Data.Json.JsonObject
    If JsonString = vbNullString Then
        If IsNotNothing(JsonObject) And IsNotNothing(JsonObject2) And IsNotNothing(JsonValue) And IsNotNothing(JsonArray) Then
            Call JsonObject.SetNamedValue("Boolean", JsonValue.CreateBooleanValue(True))
            Call JsonObject.SetNamedValue("String", JsonValue.CreateStringValue("String"))
            Call JsonObject.SetNamedValue("Null", JsonValue.CreateNullValue)
            Call JsonObject.SetNamedValue("Number", JsonValue.CreateNumberValue(5.12))
            Call JsonArray.Append(JsonValue.CreateStringValue("ABC"))
            Call JsonArray.Append(JsonValue.CreateStringValue("DEF"))
            Call JsonArray.Append(JsonValue.CreateNumberValue(100))
            Call JsonObject.SetNamedValue("Array", JsonArray)
            Call JsonObject2.SetNamedValue("Name", JsonValue.CreateStringValue("YXZ"))
            Call JsonObject2.SetNamedValue("Price", JsonValue.CreateNumberValue(3.99))
            Call JsonObject.SetNamedValue("Object", JsonObject2)
            JsonString = JsonObject.ToString
            Debug.Print "JsonObject.ToString = " & JsonString
            Command19.BackColor = vbGreen
            Command19.Caption = "Parse Json String"
        End If
    Else
        If Windows.Data.Json.JsonObject.TryParse(JsonString, JsonObject) Then
            Dim JsonObjectItems As Long
            JsonObjectItems = JsonObject.Size
            If JsonObjectItems > 0 Then
                Dim JsonObjectItem As Long
                Dim KeyValuePair As KeyValuePair_String_JsonValue
                Dim KeyValuePairs() As KeyValuePair_String_JsonValue
                KeyValuePairs = JsonObject.GetKeyValuePairs
                Debug.Print "---=== JsonObjectItems ===---"
                For JsonObjectItem = 0 To JsonObjectItems - 1
                    Set KeyValuePair = KeyValuePairs(JsonObjectItem)
                    Debug.Print KeyValuePair.Key & " = " & KeyValuePair.value.ToString
                Next
                Debug.Print "---=======================---"
                
                If JsonObject.HasKey("Array") Then
                    Set JsonArray = JsonObject.GetNamedArray("Array")
                    If IsNotNothing(JsonArray) Then
                        Dim JsonArrayItems As Long
                        JsonArrayItems = JsonArray.Size
                        If JsonArrayItems > 0 Then
                            Dim JsonArrayItem As Long
                            Dim JsonValues() As JsonValue
                            JsonValues = JsonArray.GetArrayValues
                            Debug.Print "---=== JsonArrayItems ===---"
                            For JsonArrayItem = 0 To JsonArrayItems - 1
                                Debug.Print JsonValues(JsonArrayItem).Stringify
                            Next
                            Debug.Print "---======================---"
                        End If
                    End If
                End If
            End If
        Else
            Debug.Print "InvalidJsonString"
        End If
    End If
End Sub

Private Sub Command20_Click()
    Dim MessageDialog As MessageDialog
    Set MessageDialog = Windows.UI.Popups.MessageDialog.CreateWithTitle("Content", "Title")
    If IsNotNothing(MessageDialog) Then
        MessageDialog.ParentHwnd = Me.hwnd
        Dim Commands As List_UICommand
        Set Commands = MessageDialog.Commands
        Set UICommandInvokedHandler = Windows.UI.Popups.UICommandInvokedHandler
        If IsNotNothing(Commands) Then
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Ok"))
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Maybe"))
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Cancel"))
            
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Ok", UICommandInvokedHandler))
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Maybe", UICommandInvokedHandler))
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Cancel", UICommandInvokedHandler))
        
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Ok", UICommandInvokedHandler, 1))
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Maybe", UICommandInvokedHandler, 2))
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Cancel", UICommandInvokedHandler, 3))
        End If
        MessageDialog.CancelCommandIndex = 2 ' Esc
        MessageDialog.DefaultCommandIndex = 2
        'MessageDialog.Options = MessageDialogOptions_AcceptUserInputAfterDelay
        Dim CommandRet As UICommand
        Set CommandRet = MessageDialog.ShowAsync
        If IsNotNothing(CommandRet) Then
            Debug.Print "Return from MessageDialog.ShowAsync"
            Debug.Print "UICommand.Label = " & CommandRet.Label
            If VarType(CommandRet.Id) <> vbEmpty Then Debug.Print "UICommand.Id = " & CommandRet.Id
            Debug.Print
        End If
        Set UICommandInvokedHandler = Nothing
    End If
End Sub

Private Sub Command21_Click()
    Dim CoreWindowPopupShowingEventCookie As Currency
    Dim CoreWindowPopupShowingEventHandler As Long
    Dim CoreWindowDialog As CoreWindowDialog
    Set CoreWindowDialog = Windows.UI.Core.CoreWindowDialog.CreateWithTitle("Title")
    If IsNotNothing(CoreWindowDialog) Then
        CoreWindowDialog.ParentHwnd = Me.hwnd
        CoreWindowPopupShowingEventHandler = ITEH_CoreWindowPopupShowingEvent.Create(Me)
        CoreWindowPopupShowingEventCookie = CoreWindowDialog.AddShowing(CoreWindowPopupShowingEventHandler)
        Dim Commands As List_UICommand
        Set Commands = CoreWindowDialog.Commands
        Set UICommandInvokedHandler = Windows.UI.Popups.UICommandInvokedHandler
        If IsNotNothing(Commands) Then
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Ok"))
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Maybe"))
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Cancel"))
            
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Ok", UICommandInvokedHandler))
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Maybe", UICommandInvokedHandler))
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Cancel", UICommandInvokedHandler))
        
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Ok", UICommandInvokedHandler, 1))
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Maybe", UICommandInvokedHandler, 2))
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Cancel", UICommandInvokedHandler, 3))
        End If
        CoreWindowDialog.CancelCommandIndex = 2 ' Esc
        CoreWindowDialog.DefaultCommandIndex = 2
        CoreWindowDialog.BackButtonCommand = UICommandInvokedHandler
        
        Dim CommandRet As UICommand
        Set CommandRet = CoreWindowDialog.ShowAsync
        If IsNotNothing(CommandRet) Then
            Debug.Print "Return from CoreWindowDialog.ShowAsync"
            Debug.Print "UICommand.Label = " & CommandRet.Label
            If VarType(CommandRet.Id) <> vbEmpty Then Debug.Print "UICommand.Id = " & CommandRet.Id
            Debug.Print
        End If
        Call CoreWindowDialog.RemoveShowing(CoreWindowPopupShowingEventCookie)
        Set UICommandInvokedHandler = Nothing
    End If
End Sub

Private Sub Command22_Click()
    Dim CoreWindowPopupShowingEventCookie As Currency
    Dim CoreWindowPopupShowingEventHandler As Long
    Dim CoreWindowFlyout As CoreWindowFlyout
    Dim Point As New Point
    Point.X = 150
    Point.Y = 150
    Set CoreWindowFlyout = Windows.UI.Core.CoreWindowFlyout.CreateWithTitle(Point, "Title")
    If IsNotNothing(CoreWindowFlyout) Then
        CoreWindowFlyout.ParentHwnd = Me.hwnd
        CoreWindowPopupShowingEventHandler = ITEH_CoreWindowPopupShowingEvent.Create(Me)
        CoreWindowPopupShowingEventCookie = CoreWindowFlyout.AddShowing(CoreWindowPopupShowingEventHandler)
        Dim Commands As List_UICommand
        Set Commands = CoreWindowFlyout.Commands
        Set UICommandInvokedHandler = Windows.UI.Popups.UICommandInvokedHandler
        If IsNotNothing(Commands) Then
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Ok"))
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Cancel"))
            
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Ok", UICommandInvokedHandler))
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Cancel", UICommandInvokedHandler))
        
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Ok", UICommandInvokedHandler, 1))
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Cancel", UICommandInvokedHandler, 2))
        End If
        CoreWindowFlyout.DefaultCommandIndex = 1
        CoreWindowFlyout.BackButtonCommand = UICommandInvokedHandler

        Dim CommandRet As UICommand
        Set CommandRet = CoreWindowFlyout.ShowAsync
        If IsNotNothing(CommandRet) Then
            Debug.Print "Return from CoreWindowFlyout.ShowAsync"
            Debug.Print "UICommand.Label = " & CommandRet.Label
            If VarType(CommandRet.Id) <> vbEmpty Then Debug.Print "UICommand.Id = " & CommandRet.Id
            Debug.Print
        End If
        Call CoreWindowFlyout.RemoveShowing(CoreWindowPopupShowingEventCookie)
        Set UICommandInvokedHandler = Nothing
    End If
End Sub

Private Sub Command23_Click()
    Dim PopupMenu As PopupMenu
    Set PopupMenu = Windows.UI.Popups.PopupMenu
    If IsNotNothing(PopupMenu) Then
        PopupMenu.ParentHwnd = Me.hwnd
        Dim Commands As List_UICommand
        Set Commands = PopupMenu.Commands
        Set UICommandInvokedHandler = Windows.UI.Popups.UICommandInvokedHandler
        If IsNotNothing(Commands) Then
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Ok"))
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Maybe"))
'            Call Commands.Append(Windows.UI.Popups.UICommandSeparator)
'            Call Commands.Append(Windows.UI.Popups.UICommand.Create("&Cancel"))
            
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Ok", UICommandInvokedHandler))
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Maybe", UICommandInvokedHandler))
'            Call Commands.Append(Windows.UI.Popups.UICommandSeparator)
'            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandler("&Cancel", UICommandInvokedHandler))
        
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Ok", UICommandInvokedHandler, 1))
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Maybe", UICommandInvokedHandler, 2))
            Call Commands.Append(Windows.UI.Popups.UICommandSeparator)
            Call Commands.Append(Windows.UI.Popups.UICommand.CreateWithHandlerAndId("&Cancel", UICommandInvokedHandler, 3))
        End If
    
        Dim CommandRet As UICommand
        Dim Point As New Point
        Point.X = 35
        Point.Y = 0
        Set CommandRet = PopupMenu.ShowAsync(Point)
        If IsNotNothing(CommandRet) Then
            Debug.Print "Return from PopupMenu.ShowAsync"
            Debug.Print "UICommand.Label = " & CommandRet.Label
            If VarType(CommandRet.Id) <> vbEmpty Then Debug.Print "UICommand.Id = " & CommandRet.Id
            Debug.Print
        End If
        Set UICommandInvokedHandler = Nothing
    End If
End Sub

Private Sub Command24_Click()
    Dim ToastHeader As String
    Dim ToastMessage As String
    Dim UserNotificationListener As UserNotificationListener
    Set UserNotificationListener = Windows.UI.Notifications.Management.UserNotificationListener.Current
    If IsNotNothing(UserNotificationListener) Then
        If UserNotificationListener.RequestAccessAsync = UserNotificationListenerAccessStatus_Allowed Then
            Dim ListUserNotification  As ReadOnlyList_1 'ReadOnlyList_UserNotification
            Set ListUserNotification = UserNotificationListener.GetNotificationsAsync(NotificationKinds_Toast)
            If IsNotNothing(ListUserNotification) Then
                Dim UserNotificationCount As Long
                UserNotificationCount = ListUserNotification.Size
                If UserNotificationCount > 0& Then
                    Dim UserNotificationItem As Long
                    For UserNotificationItem = 0 To UserNotificationCount - 1
                        Dim UserNotification As UserNotification
                        Set UserNotification = ListUserNotification.GetAt(UserNotificationItem)
                        If IsNotNothing(UserNotification) Then
                            Debug.Print "---=== Start UserNotification: " & CStr(UserNotificationItem + 1) & " ===---"
                            Debug.Print "UserNotification.Id = " & CStr(UserNotification.Id)
                            Debug.Print "UserNotification.CreationTime = " & CStr(UserNotification.CreationTime.VbDate)
                            Debug.Print "UserNotification.ExpirationTime = " & CStr(UserNotification.notification.ExpirationTime.VbDate)
                            Dim NotificationBindingList As ReadOnlyList_1 'ReadOnlyList_NotificationBinding
                            Set NotificationBindingList = UserNotification.notification.Visual.Bindings
                            If IsNotNothing(NotificationBindingList) Then
                                Dim NotificationBindingCount As Long
                                NotificationBindingCount = NotificationBindingList.Size
                                If NotificationBindingCount > 0& Then
                                    Dim NotificationBindingItem As Long
                                    For NotificationBindingItem = 0 To NotificationBindingCount - 1
                                        Dim NotificationBinding As NotificationBinding
                                        Set NotificationBinding = NotificationBindingList.GetAt(NotificationBindingItem)
                                        If IsNotNothing(NotificationBinding) Then
                                            Debug.Print vbTab & "---=== Start NotificationBinding: " & CStr(NotificationBindingItem + 1) & " ===---"
                                            Debug.Print vbTab & "NotificationBinding.Language = " & NotificationBinding.Language
                                            Debug.Print vbTab & "NotificationBinding.Template = " & NotificationBinding.Template
                                            Dim NotificationBindingHintsList As List_String_String
                                            Set NotificationBindingHintsList = NotificationBinding.Hints
                                            If IsNotNothing(NotificationBindingHintsList) Then
                                                Dim NotificationBindingHintsListCount As Long
                                                NotificationBindingHintsListCount = NotificationBindingHintsList.Size
                                                If NotificationBindingHintsListCount > 0& Then
                                                    Dim NotificationBindingHintsListItem As Long
                                                    Dim KeyValuePair0() As KeyValuePair_String_String
                                                    KeyValuePair0 = NotificationBindingHintsList.GetKeyValuePairs
                                                    For NotificationBindingHintsListItem = 0 To NotificationBindingHintsListCount - 1
                                                        Debug.Print vbTab & vbTab & "---=== Start NotificationBinding.Hints: " & CStr(NotificationBindingHintsListItem + 1) & " ===---"
                                                        Debug.Print vbTab & vbTab & "NotificationBinding.Hints.Key = " & KeyValuePair0(NotificationBindingHintsListItem).Key
                                                        Debug.Print vbTab & vbTab & "NotificationBinding.Hints.Value = " & KeyValuePair0(NotificationBindingHintsListItem).value
                                                        Debug.Print vbTab & vbTab & "---=== End NotificationBinding.Hints: " & CStr(NotificationBindingHintsListItem + 1) & " ===---"
                                                    Next
                                                End If
                                            End If
                                            Dim AdaptiveNotificationTextList As ReadOnlyList_1 'ReadOnlyList_AdaptiveNotificationText
                                            Set AdaptiveNotificationTextList = NotificationBinding.GetTextElements
                                            If IsNotNothing(AdaptiveNotificationTextList) Then
                                                Dim AdaptiveNotificationTextListCount As Long
                                                AdaptiveNotificationTextListCount = AdaptiveNotificationTextList.Size
                                                If AdaptiveNotificationTextListCount > 0& Then
                                                    Dim AdaptiveNotificationTextListItem As Long
                                                    For AdaptiveNotificationTextListItem = 0 To AdaptiveNotificationTextListCount - 1
                                                        Dim AdaptiveNotificationText As AdaptiveNotificationText
                                                        Set AdaptiveNotificationText = AdaptiveNotificationTextList.GetAt(AdaptiveNotificationTextListItem)
                                                        If IsNotNothing(AdaptiveNotificationText) Then
                                                            Debug.Print vbTab & vbTab & "---=== Start AdaptiveNotificationText: " & CStr(AdaptiveNotificationTextListItem + 1) & " ===---"
                                                            Debug.Print vbTab & vbTab & "AdaptiveNotificationText.Language = " & AdaptiveNotificationText.Language
                                                            Debug.Print vbTab & vbTab & "AdaptiveNotificationText.Kind = " & AdaptiveNotificationText.Kind
                                                            Debug.Print vbTab & vbTab & "AdaptiveNotificationText.Text = " & AdaptiveNotificationText.Text
                                                            If AdaptiveNotificationTextListItem = 0 Then
                                                                ToastHeader = AdaptiveNotificationText.Text
                                                            Else
                                                                ToastMessage = ToastMessage & AdaptiveNotificationText.Text
                                                                If AdaptiveNotificationTextListItem < AdaptiveNotificationTextListCount - 1 Then
                                                                    ToastMessage = ToastMessage & vbNewLine
                                                                End If
                                                            End If
                                                            Dim AdaptiveNotificationTextHintsList As List_String_String
                                                            Set AdaptiveNotificationTextHintsList = AdaptiveNotificationText.Hints
                                                            If IsNotNothing(AdaptiveNotificationTextHintsList) Then
                                                                Dim AdaptiveNotificationTextHintsListCount As Long
                                                                AdaptiveNotificationTextHintsListCount = AdaptiveNotificationTextHintsList.Size
                                                                If AdaptiveNotificationTextHintsListCount > 0& Then
                                                                    Dim AdaptiveNotificationTextHintsListItem As Long
                                                                    Dim KeyValuePair1() As KeyValuePair_String_String
                                                                    KeyValuePair1 = AdaptiveNotificationTextHintsList.GetKeyValuePairs
                                                                    For AdaptiveNotificationTextHintsListItem = 0 To AdaptiveNotificationTextHintsListCount - 1
                                                                        Debug.Print vbTab & vbTab & vbTab & "---=== Start AdaptiveNotificationText.Hints: " & CStr(AdaptiveNotificationTextHintsListItem + 1) & " ===---"
                                                                        Debug.Print vbTab & vbTab & vbTab & "AdaptiveNotificationText.Hints.Key = " & KeyValuePair1(AdaptiveNotificationTextHintsListItem).Key
                                                                        Debug.Print vbTab & vbTab & vbTab & "AdaptiveNotificationText.Hints.Value = " & KeyValuePair1(AdaptiveNotificationTextHintsListItem).value
                                                                        Debug.Print vbTab & vbTab & vbTab & "---=== End AdaptiveNotificationText.Hints: " & CStr(AdaptiveNotificationTextHintsListItem + 1) & " ===---"
                                                                    Next
                                                                End If
                                                            End If
                                                            Debug.Print vbTab & vbTab & "---=== End AdaptiveNotificationText: " & CStr(AdaptiveNotificationTextListItem + 1) & " ===---"
                                                        End If
                                                    Next
                                                End If
                                            End If
                                            Debug.Print vbTab & "---=== End NotificationBinding: " & CStr(NotificationBindingItem + 1) & " ===---"
                                        End If
                                    Next
                                End If
                            End If
                            Debug.Print "---=== Toast ===---"
                            Debug.Print "ToastHeader = " & ToastHeader
                            Debug.Print "ToastMessage = " & ToastMessage
                            Debug.Print "---=============---"
                            Debug.Print "---=== End UserNotification: " & CStr(UserNotificationItem + 1) & " ===---"
                            Debug.Print
                            ToastHeader = vbNullString
                            ToastMessage = vbNullString
                        End If
                    Next
                Else
                    Debug.Print "Keine aktiven Toast im Benachrichtigungscenter gefunden!"
                End If
            End If
        End If
    End If
End Sub

Private Sub Command25_Click()
    Dim strHtml As String
    strHtml = "<!DOCTYPE html><html><head><title>This is a title</title></head><body><div><p>Hello world!</p></div></body></html>"
    Debug.Print "Html2Text = " & Windows.Data.Html.HtmlUtilities.ConvertToText(strHtml)
End Sub

Private Sub Command26_Click()
    Load frmPdf
    frmPdf.Show vbModal, Me
    Unload frmPdf
End Sub

Private Sub Command27_Click()

    'Debug.Print Windows.Devices.Enumeration.DevicePicker.ShowWithPlacement(Windows.Foundation.Rect(0, 0, 300, 300), Placement_Right)
    
    Dim DevicePicker As DevicePicker
    Set DevicePicker = Windows.Devices.Enumeration.DevicePicker
    If IsNotNothing(DevicePicker) Then
        Dim DeviceSelectedEventCookie As Currency
        Dim DeviceSelectedEventHandler As Long
        Dim DevicePickerDismissedEventCookie As Currency
        Dim DevicePickerDismissedEventHandler As Long
        Dim DeviceDisconnectButtonClickedEventCookie As Currency
        Dim DeviceDisconnectButtonClickedEventHandler As Long
        DeviceSelectedEventHandler = ITEH_DeviceSelectedEvent.Create(Me)
        DevicePickerDismissedEventHandler = ITEH_DevicePickerDismissed.Create(Me)
        DeviceDisconnectButtonClickedEventHandler = ITEH_DeviceDisconnectButtonClickedEvent.Create(Me)
        DeviceSelectedEventCookie = DevicePicker.AddDeviceSelected(DeviceSelectedEventHandler)
        DevicePickerDismissedEventCookie = DevicePicker.AddDevicePickerDismissed(DevicePickerDismissedEventHandler)
        DeviceDisconnectButtonClickedEventCookie = DevicePicker.AddDisconnectButtonClicked(DeviceDisconnectButtonClickedEventHandler)
        Dim DevicePickerAppearance As DevicePickerAppearance
        Set DevicePickerAppearance = DevicePicker.Appearance
        If IsNotNothing(DevicePickerAppearance) Then
            DevicePickerAppearance.Title = "Select Device"
        End If
        Dim DeviceClassList As List_DeviceClass
        Set DeviceClassList = DevicePicker.Filter.SupportedDeviceClasses
        If IsNotNothing(DeviceClassList) Then
            Call DeviceClassList.Append(DeviceClass_AudioCapture)
            Call DeviceClassList.Append(DeviceClass_VideoCapture)
            Call DeviceClassList.Append(DeviceClass_AudioRender)
            Call DeviceClassList.Append(DeviceClass_PortableStorageDevice)
        End If
        Dim DeviceInformation As DeviceInformation
        Set DeviceInformation = DevicePicker.PickSingleDeviceAsync(Windows.Foundation.Rect(0, 0, 300, 300))
        If IsNotNothing(DeviceInformation) Then
        
'            Dim DeviceThumbnail As DeviceThumbnail
'            Set DeviceThumbnail = DeviceInformation.GetThumbnailAsync
'            'Set DeviceThumbnail = DeviceInformation.GetGlyphThumbnailAsync
'            If IsNotNothing(DeviceThumbnail) Then
'                Dim pIStream As Long
'                pIStream = DeviceThumbnail.ToIStream
'                If pIStream <> 0& Then
'                    Me.Picture = GetPictureFromIStream(pIStream)
'                    Call ReleaseIfc(pIStream)
'                End If
'            End If
        
            Debug.Print "Name = " & DeviceInformation.Name
            Debug.Print "Id = " & DeviceInformation.Id
            Debug.Print "IsDefault = " & DeviceInformation.IsDefault
            Debug.Print "IsEnabled = " & DeviceInformation.IsEnabled
            Debug.Print "Kind = " & DeviceInformation.Kind
            Dim PropertyList As ReadOnlyList_2 'ReadOnlyList_String_Inspectable
            Set PropertyList = DeviceInformation.Properties
            If IsNotNothing(PropertyList) Then
                Dim PropertyCount As Long
                PropertyCount = PropertyList.Size
                If PropertyCount > 0& Then
                    Dim vKeyValuePair As Variant ' KeyValuePair_String_Inspectable
                    Dim PropertyValue As New PropertyValue
                    Dim PropertyVal As String
                    For Each vKeyValuePair In PropertyList.GetKeyValuePairs
                        PropertyVal = vbNullString
                        If IsNotNothing(vKeyValuePair.value) Then
                            PropertyValue.Ifc = vKeyValuePair.value.Ifc
                            Select Case PropertyValue.PropType
                                Case PropertyType.PropertyType_String
                                    PropertyVal = PropertyValue.GetString
                                Case PropertyType.PropertyType_Boolean
                                    PropertyVal = CStr(PropertyValue.GetBoolean)
                                Case PropertyType.PropertyType_UInt32
                                    PropertyVal = CStr(PropertyValue.GetUInt32)
                                Case PropertyType.PropertyType_Guid
                                    PropertyVal = Guid2Str(PropertyValue.GetGuid)
                            End Select
                        End If
                        Debug.Print "Property: " & vKeyValuePair.Key & " = " & PropertyVal
                    Next
                End If
                Set PropertyList = Nothing
                Debug.Print
            End If
        Else
            Debug.Print "DevicePicker Cancel"
        End If
        Call DevicePicker.RemoveDeviceSelected(DeviceSelectedEventCookie)
        Call DevicePicker.RemoveDevicePickerDismissed(DevicePickerDismissedEventCookie)
        Call DevicePicker.RemoveDisconnectButtonClicked(DeviceDisconnectButtonClickedEventCookie)
    End If
End Sub

Private Sub Command28_Click()
    Dim CastingDevicePicker As CastingDevicePicker
    Set CastingDevicePicker = Windows.Media.Casting.CastingDevicePicker
    If IsNotNothing(CastingDevicePicker) Then
        CastingDeviceSelectedEventHandler = ITEH_CastingDeviceSelectedEvent.Create(Me)
        CastingDevicePickerDismissedHandler = ITEH_CastingDevicePickerDismissed.Create(Me)
        CastingDeviceSelectedEventCookie = CastingDevicePicker.AddCastingDeviceSelected(CastingDeviceSelectedEventHandler)
        CastingDevicePickerDismissedCookie = CastingDevicePicker.AddCastingDevicePickerDismissed(CastingDevicePickerDismissedHandler)
        Dim DevicePickerAppearance As DevicePickerAppearance
        Set DevicePickerAppearance = CastingDevicePicker.Appearance
        If IsNotNothing(DevicePickerAppearance) Then
            DevicePickerAppearance.Title = "Select Device"
        End If
        Dim CastingDevicePickerFilter As CastingDevicePickerFilter
        Set CastingDevicePickerFilter = CastingDevicePicker.Filter
        If IsNotNothing(CastingDevicePickerFilter) Then
            CastingDevicePickerFilter.SupportsAudio = True
            CastingDevicePickerFilter.SupportsVideo = True
            CastingDevicePickerFilter.SupportsPictures = True
        End If
        If CastingDevicePicker.ShowWithPlacement(Windows.Foundation.Rect(0, 0, 300, 300), Placement_Right) Then
        
        End If
    End If
End Sub

Private Sub Command29_Click()
    Dim QueryOptions As QueryOptions
    Set QueryOptions = Windows.Storage.Search.QueryOptions
    If IsNotNothing(QueryOptions) Then
        Dim FileTypeFilterList As List_String
        Set FileTypeFilterList = QueryOptions.FileTypeFilter
        If IsNotNothing(FileTypeFilterList) Then
            Call FileTypeFilterList.Append(".exe")
            If QueryOptions.CreateCommonFileQuery(CommonFileQuery_DefaultQuery, FileTypeFilterList) Then
                QueryOptions.IndexerOption = IndexerOption_UseIndexerWhenAvailable
                Dim StorageFolder As StorageFolder
                Set StorageFolder = Windows.Storage.StorageFolder.GetFolderFromPathAsync("C:\Windows")
                If IsNotNothing(StorageFolder) Then
                    Dim StorageFileQueryResult As StorageFileQueryResult
                    Set StorageFileQueryResult = StorageFolder.CreateFileQueryWithOptions(QueryOptions)
                    If IsNotNothing(StorageFileQueryResult) Then
                        Debug.Print CStr(StorageFileQueryResult.GetItemCountAsync) & " Dateien in " & StorageFolder.Path & " gefunden."
                        If StorageFileQueryResult.GetItemCountAsync > 0& Then
                            Dim StorageFileList As ReadOnlyList_1 'ReadOnlyList_StorageFile
                            Set StorageFileList = StorageFileQueryResult.GetFilesAsyncDefaultStartAndCount
                            If IsNotNothing(StorageFileList) Then
                                Dim FileCount As Long
                                FileCount = StorageFileList.Size
                                If FileCount > 0 Then
                                    Dim FileItem As Long
                                    Dim StorageFile As StorageFile
                                    For FileItem = 0 To FileCount - 1
                                        Set StorageFile = StorageFileList.GetAt(FileItem)
                                        If IsNotNothing(StorageFile) Then
                                            Debug.Print "StorageFile.Path: " & StorageFile.Path
                                            Set StorageFile = Nothing
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Command30_Click()
    Dim OcrEngine As OcrEngine
    Set OcrEngine = Windows.Media.Ocr.OcrEngine
    If IsNotNothing(OcrEngine) Then
        Debug.Print "OcrEngine.MaxImageDimension = " & CStr(OcrEngine.MaxImageDimension)
        Dim AvailableRecognizerLanguageList As ReadOnlyList_1 'ReadOnlyList_Language
        Set AvailableRecognizerLanguageList = OcrEngine.AvailableRecognizerLanguages
        If IsNotNothing(AvailableRecognizerLanguageList) Then
            Dim LanguageCount As Long
            LanguageCount = AvailableRecognizerLanguageList.Size
            If LanguageCount > 0& Then
                Debug.Print "---=== AvailableRecognizerLanguages ===---"
                Dim LanguageItem As Long
                Dim Language As Language
                For LanguageItem = 0 To LanguageCount - 1
                    Set Language = AvailableRecognizerLanguageList.GetAt(LanguageItem)
                    If IsNotNothing(Language) Then
                        Debug.Print "Language.AbbreviatedName = " & Language.AbbreviatedName
                        Debug.Print "Language.CurrentInputMethodLanguageTag = " & Language.CurrentInputMethodLanguageTag
                        Debug.Print "Language.DisplayName = " & Language.DisplayName
                        Debug.Print "Language.LanguageTag = " & Language.LanguageTag
                        Debug.Print "Language.LayoutDirection = " & Language.LayoutDirection
                        Debug.Print "Language.NativeName = " & Language.NativeName
                        Debug.Print "Language.Script = " & Language.Script
                        Debug.Print "OcrEngine.IsLanguageSupported = " & OcrEngine.IsLanguageSupported(Language)
                        Debug.Print
                    End If
                    Set Language = Nothing
                Next
            End If
        End If
        Dim OcrReader As OcrEngine
        'Set OcrReader = OcrEngine.TryCreateFromLanguage(Language) <- AvailableRecognizerLanguage
        Set OcrReader = OcrEngine.TryCreateFromUserProfileLanguages ' <- DefaultUserLanguage
        If IsNotNothing(OcrReader) Then
            Dim CurrentLanguage As Language
            Set CurrentLanguage = OcrReader.RecognizerLanguage
            If IsNotNothing(CurrentLanguage) Then
                Debug.Print "---=== UsedRecognizerLanguage ===---"
                Debug.Print "CurrentLanguage.AbbreviatedName = " & CurrentLanguage.AbbreviatedName
                Debug.Print "CurrentLanguage.CurrentInputMethodLanguageTag = " & CurrentLanguage.CurrentInputMethodLanguageTag
                Debug.Print "CurrentLanguage.DisplayName = " & CurrentLanguage.DisplayName
                Debug.Print "CurrentLanguage.LanguageTag = " & CurrentLanguage.LanguageTag
                Debug.Print "CurrentLanguage.LayoutDirection = " & CurrentLanguage.LayoutDirection
                Debug.Print "CurrentLanguage.NativeName = " & CurrentLanguage.NativeName
                Debug.Print "CurrentLanguage.Script = " & CurrentLanguage.Script
                Debug.Print
            End If
            Dim FileOpenPicker As FileOpenPicker
            Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
            If IsNotNothing(FileOpenPicker) Then
                FileOpenPicker.ParentHwnd = Me.hwnd
                Call FileOpenPicker.FileTypeFilter.Clear
                Dim FileTypeFilter As List_String
                Set FileTypeFilter = FileOpenPicker.FileTypeFilter
                Call FileTypeFilter.Append(".jpg")
                Call FileTypeFilter.Append(".gif")
                Call FileTypeFilter.Append(".bmp")
                Call FileTypeFilter.Append(".png")
                Dim StorageFile As StorageFile
                Set StorageFile = FileOpenPicker.PickSingleFileAsync
                If IsNotNothing(StorageFile) Then
                    Dim RandomAccessStream As RandomAccessStream
                    Set RandomAccessStream = StorageFile.OpenAsync(FileAccessMode_Read)
                    If IsNotNothing(RandomAccessStream) Then
                        Dim SoftwareBitmap As SoftwareBitmap
                        Set SoftwareBitmap = Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(RandomAccessStream).GetSoftwareBitmapAsync
                        If IsNotNothing(SoftwareBitmap) Then
                            Dim OcrResult As OcrResult
                            Set OcrResult = OcrReader.RecognizeAsync(SoftwareBitmap)
                            If IsNotNothing(OcrResult) Then
                                Debug.Print "OcrResult.TextAngle = " & CStr(OcrResult.TextAngle)
                                Debug.Print "OcrResult.Text = " & OcrResult.Text
'                                Dim OcrLineList As ReadOnlyList_1 'ReadOnlyList_OcrLine
'                                Set OcrLineList = OcrResult.Lines
'                                If IsNotNothing(OcrLineList) Then
'                                    Dim OcrLineCount As Long
'                                    OcrLineCount = OcrLineList.Size
'                                    If OcrLineCount > 0& Then
'                                        Dim OcrLineItem As Long
'                                        For OcrLineItem = 0 To OcrLineCount - 1
'                                            Dim OcrLine As OcrLine
'                                            Set OcrLine = OcrLineList.GetAt(OcrLineItem)
'                                            If IsNotNothing(OcrLine) Then
'                                                Debug.Print "OcrLine" & CStr(OcrLineItem + 1) & ".Text = " & OcrLine.Text
''                                                Dim OcrWordList As ReadOnlyList_1 'ReadOnlyList_OcrWord
''                                                Set OcrWordList = OcrLine.Words
''                                                If IsNotNothing(OcrWordList) Then
''                                                    Dim OcrWordCount As Long
''                                                    OcrWordCount = OcrWordList.Size
''                                                    If OcrWordCount > 0& Then
''                                                        Dim OcrWordItem As Long
''                                                        For OcrWordItem = 0 To OcrWordCount - 1
''                                                            Dim OcrWord As OcrWord
''                                                            Set OcrWord = OcrWordList.GetAt(OcrWordItem)
''                                                            If IsNotNothing(OcrWord) Then
''                                                                Debug.Print "OcrLine" & CStr(OcrLineItem + 1) & "-> OcrWord" & CStr(OcrWordItem + 1) & ".Text = " & _
''                                                                            OcrWord.Text & " -> OcrWord.BoundingRect = " & OcrWord.BoundingRect.ToString
''                                                                Set OcrWord = Nothing
''                                                            End If
''                                                        Next
''                                                    End If
''                                                End If
'                                                Set OcrLine = Nothing
'                                            End If
'                                        Next
'                                    End If
'                                End If
                            End If
                        End If
                    End If
                Else
                    Debug.Print "FileOpenPicker = Cancel"
                End If
            End If
        End If
    End If
End Sub

Private Sub Command31_Click()
    Dim SystemDiagnosticInfo As SystemDiagnosticInfo
    Set SystemDiagnosticInfo = Windows.System.Diagnostics.SystemDiagnosticInfo.GetForCurrentSystem
    If IsNotNothing(SystemDiagnosticInfo) Then
        Dim SystemCpuUsageReport As SystemCpuUsageReport
        Set SystemCpuUsageReport = SystemDiagnosticInfo.CpuUsage.GetReport
        If IsNotNothing(SystemCpuUsageReport) Then
            Debug.Print "SystemCpuUsageReport.IdleTime = " & SystemCpuUsageReport.IdleTime.VbDate
            Debug.Print "SystemCpuUsageReport.UserTime = " & SystemCpuUsageReport.UserTime.VbDate
            Debug.Print "SystemCpuUsageReport.KernelTime = " & SystemCpuUsageReport.KernelTime.VbDate
        End If
    
        Dim SystemMemoryUsageReport As SystemMemoryUsageReport
        Set SystemMemoryUsageReport = SystemDiagnosticInfo.MemoryUsage.GetReport
        If IsNotNothing(SystemMemoryUsageReport) Then
            Debug.Print "SystemMemoryUsageReport.AvailableSizeInBytes = " & SystemMemoryUsageReport.AvailableSizeInBytes
            Debug.Print "SystemMemoryUsageReport.CommittedSizeInBytes = " & SystemMemoryUsageReport.CommittedSizeInBytes
            Debug.Print "SystemMemoryUsageReport.TotalPhysicalSizeInBytes = " & SystemMemoryUsageReport.TotalPhysicalSizeInBytes
        End If
    End If
End Sub

Private Sub Command32_Click()
    Dim ProcessDiagnosticInfo As ProcessDiagnosticInfo
    Set ProcessDiagnosticInfo = Windows.System.Diagnostics.ProcessDiagnosticInfo.GetForCurrentProcess
    If IsNotNothing(ProcessDiagnosticInfo) Then
        Debug.Print "ProcessDiagnosticInfo.ProcessId = " & ProcessDiagnosticInfo.ProcessId
        Debug.Print "ProcessDiagnosticInfo.ExecutableFileName = " & ProcessDiagnosticInfo.ExecutableFileName
        Debug.Print "ProcessDiagnosticInfo.IsPackaged = " & ProcessDiagnosticInfo.IsPackaged
        Debug.Print "ProcessDiagnosticInfo.ProcessStartTime = " & ProcessDiagnosticInfo.ProcessStartTime.VbDate
        Dim ProcessCpuUsageReport As ProcessCpuUsageReport
        Set ProcessCpuUsageReport = ProcessDiagnosticInfo.CpuUsage.GetReport
        If IsNotNothing(ProcessCpuUsageReport) Then
            Debug.Print "ProcessCpuUsageReport.UserTime = " & ProcessCpuUsageReport.UserTime.VbDate
            Debug.Print "ProcessCpuUsageReport.KernelTime = " & ProcessCpuUsageReport.KernelTime.VbDate
        End If
        Dim ProcessMemoryUsageReport As ProcessMemoryUsageReport
        Set ProcessMemoryUsageReport = ProcessDiagnosticInfo.MemoryUsage.GetReport
        If IsNotNothing(ProcessMemoryUsageReport) Then
            Debug.Print "ProcessMemoryUsageReport.NonPagedPoolSizeInBytes = " & ProcessMemoryUsageReport.NonPagedPoolSizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PagedPoolSizeInBytes = " & ProcessMemoryUsageReport.PagedPoolSizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PageFaultCount = " & ProcessMemoryUsageReport.PageFaultCount
            Debug.Print "ProcessMemoryUsageReport.PageFileSizeInBytes = " & ProcessMemoryUsageReport.PageFileSizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PeakNonPagedPoolSizeInBytes = " & ProcessMemoryUsageReport.PeakNonPagedPoolSizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PeakPagedPoolSizeInBytes = " & ProcessMemoryUsageReport.PeakPagedPoolSizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PeakPageFileSizeInBytes = " & ProcessMemoryUsageReport.PeakPageFileSizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PeakVirtualMemorySizeInBytes = " & ProcessMemoryUsageReport.PeakVirtualMemorySizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PeakWorkingSetSizeInBytes = " & ProcessMemoryUsageReport.PeakWorkingSetSizeInBytes
            Debug.Print "ProcessMemoryUsageReport.PrivatePageCount = " & ProcessMemoryUsageReport.PrivatePageCount
            Debug.Print "ProcessMemoryUsageReport.VirtualMemorySizeInBytes = " & ProcessMemoryUsageReport.VirtualMemorySizeInBytes
            Debug.Print "ProcessMemoryUsageReport.WorkingSetSizeInBytes = " & ProcessMemoryUsageReport.WorkingSetSizeInBytes
        End If
        Dim ProcessDiskUsageReport As ProcessDiskUsageReport
        Set ProcessDiskUsageReport = ProcessDiagnosticInfo.DiskUsage.GetReport
        If IsNotNothing(ProcessDiskUsageReport) Then
            Debug.Print "ProcessDiskUsageReport.BytesReadCount = " & ProcessDiskUsageReport.BytesReadCount
            Debug.Print "ProcessDiskUsageReport.BytesWrittenCount = " & ProcessDiskUsageReport.BytesWrittenCount
            Debug.Print "ProcessDiskUsageReport.OtherBytesCount = " & ProcessDiskUsageReport.OtherBytesCount
            Debug.Print "ProcessDiskUsageReport.OtherOperationCount = " & ProcessDiskUsageReport.OtherOperationCount
            Debug.Print "ProcessDiskUsageReport.ReadOperationCount = " & ProcessDiskUsageReport.ReadOperationCount
            Debug.Print "ProcessDiskUsageReport.WriteOperationCount = " & ProcessDiskUsageReport.WriteOperationCount
        End If
    End If
End Sub

Private Sub Command33_Click()
    Dim AppDiagnosticInfo As AppDiagnosticInfo
    Set AppDiagnosticInfo = Windows.System.AppDiagnosticInfo
    If IsNotNothing(AppDiagnosticInfo) Then
        Debug.Print "AppDiagnosticInfo.RequestAccessAsync = " & AppDiagnosticInfo.RequestAccessAsync
        Dim AppDiagnosticInfoList As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
        Set AppDiagnosticInfoList = AppDiagnosticInfo.RequestInfoAsync
        If IsNotNothing(AppDiagnosticInfoList) Then
            Dim AppDiagnosticInfoCount As Long
            AppDiagnosticInfoCount = AppDiagnosticInfoList.Size
            If AppDiagnosticInfoCount > 0& Then
                Dim AppDiagnosticInfoItem As Long
                Dim AppInfo As AppInfo
                For AppDiagnosticInfoItem = 0 To AppDiagnosticInfoCount - 1
                    Set AppInfo = AppDiagnosticInfoList.GetAt(AppDiagnosticInfoItem).AppInfo
                    If IsNotNothing(AppInfo) Then
                        Debug.Print "--- AppDiagnosticInfo" & CStr(AppDiagnosticInfoItem + 1) & " ---"
                        Debug.Print "AppInfo.PackageFamilyName = " & AppInfo.PackageFamilyName
                        Debug.Print "AppInfo.AppUserModelId = " & AppInfo.AppUserModelId
                        Debug.Print "AppInfo.Id = " & AppInfo.Id
                        Debug.Print "AppInfo.ExecutionContext = " & AppInfo.ExecutionContext
                        Dim AppDisplayInfo As AppDisplayInfo
                        Set AppDisplayInfo = AppInfo.DisplayInfo
                        If IsNotNothing(AppDisplayInfo) Then
                            Debug.Print "AppDisplayInfo.Description = " & AppDisplayInfo.Description
                            Debug.Print "AppDisplayInfo.DisplayName = " & AppDisplayInfo.DisplayName
                            Set AppDisplayInfo = Nothing
                        End If
                        Dim Package As Package
                        Set Package = AppInfo.Package
                        If IsNotNothing(Package) Then
                            Debug.Print "Package.Description = " & Package.Description
                            Debug.Print "Package.DisplayName = " & Package.DisplayName
                            Debug.Print "Package.InstalledDate = " & Package.InstalledDate.VbDate
                            Debug.Print "Package.InstalledPath = " & Package.InstalledPath
                            Dim PackageId As PackageId
                            Set PackageId = Package.Id
                            If IsNotNothing(PackageId) Then
                                Debug.Print "PackageId.Architecture = " & PackageId.Architecture
                                Debug.Print "PackageId.Author = " & PackageId.Author
                                Debug.Print "PackageId.FamilyName = " & PackageId.FamilyName
                                Debug.Print "PackageId.FullName = " & PackageId.FullName
                                Debug.Print "PackageId.Name = " & PackageId.Name
                                Debug.Print "PackageId.ProductId = " & PackageId.ProductId
                                Debug.Print "PackageId.Publisher = " & PackageId.Publisher
                                Debug.Print "PackageId.PublisherId = " & PackageId.PublisherId
                                Debug.Print "PackageId.ResourceId = " & PackageId.ResourceId
                                Debug.Print "PackageId.Version = " & PackageId.version.ToString
                                Set PackageId = Nothing
                            End If
                            Set Package = Nothing
                        End If
                        Debug.Print "------------------------------------"
                        Set AppInfo = Nothing
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub Command34_Click()
    Load frmFace
    frmFace.Show vbModal, Me
    Unload frmFace
End Sub

Private Sub Command35_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = FileOpenPicker.FileTypeFilter
        Call FileTypeFilter.Append(".jpg")
        Call FileTypeFilter.Append(".gif")
        Call FileTypeFilter.Append(".bmp")
        Dim StorageFile As StorageFile
        Set StorageFile = FileOpenPicker.PickSingleFileAsync
        If IsNotNothing(StorageFile) Then
            Dim InputStream As RandomAccessStream
            Set InputStream = StorageFile.OpenAsync(FileAccessMode_Read)
            If IsNotNothing(InputStream) Then
                Dim SoftwareBitmap As SoftwareBitmap
                Set SoftwareBitmap = Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(InputStream).GetSoftwareBitmapAsync
                If IsNotNothing(SoftwareBitmap) Then
                    Dim FolderPicker As FolderPicker
                    Set FolderPicker = Windows.Storage.Pickers.FolderPicker
                    If IsNotNothing(FolderPicker) Then
                        FolderPicker.ParentHwnd = Me.hwnd
                        Dim StorageFolder As StorageFolder
                        Set StorageFolder = FolderPicker.PickSingleFolderAsync
                        If IsNotNothing(StorageFolder) Then
                            Dim OutputStream As RandomAccessStream
                            Set OutputStream = StorageFolder.CreateFileAsync("Converted.bmp", CreationCollisionOption_GenerateUniqueName).OpenAsync(FileAccessMode_ReadWrite)
                            If IsNotNothing(OutputStream) Then
                                Dim BitmapEncoder As BitmapEncoder
                                Set BitmapEncoder = Windows.Graphics.Imaging.BitmapEncoder
                                If IsNotNothing(BitmapEncoder) Then
                                    Dim Encoder As BitmapEncoder
                                    Set Encoder = BitmapEncoder.CreateAsync(BitmapEncoder.BmpEncoderId, OutputStream)
                                    If IsNotNothing(Encoder) Then
                                        If Encoder.SetSoftwareBitmap(SoftwareBitmap) Then
'                                            Encoder.BitmapTransform.ScaledWidth = 300
'                                            Encoder.BitmapTransform.ScaledHeight = 300
'                                            Encoder.BitmapTransform.InterpolationMode = BitmapInterpolationMode_Fant
'                                            Encoder.BitmapTransform.Rotation = BitmapRotation_Clockwise90Degrees
'                                            Encoder.BitmapTransform.Flip = BitmapFlip_Horizontal
'                                            Encoder.BitmapTransform.Bounds = Windows.Graphics.Imaging.BitmapBounds(0, 0, 100, 100)
                                            If Encoder.FlushAsync Then
                                                Debug.Print "Image converted"
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            Debug.Print "FolderPicker = Cancel"
                        End If
                    End If
                End If
            End If
        Else
            Debug.Print "FileOpenPicker = Cancel"
        End If
    End If
End Sub

Private Sub Command36_Click()
    Dim CodecInformation As ReadOnlyList_1 'ReadOnlyList_BitmapCodecInformation
    Set CodecInformation = Windows.Graphics.Imaging.BitmapDecoder.GetDecoderInformationEnumerator
    If IsNotNothing(CodecInformation) Then
        If CodecInformation.Size > 0 Then
            Dim CodecItem As Long
            Dim Information As BitmapCodecInformation
            For CodecItem = 0 To CodecInformation.Size - 1
                Set Information = CodecInformation.GetAt(CodecItem)
                If IsNotNothing(Information) Then
                    Debug.Print "---=== BitmapDecoder " & CStr(CodecItem + 1) & " ===---"
                    Debug.Print "BitmapDecoder.FriendlyName = " & Information.FriendlyName
                    Debug.Print "BitmapDecoder.CodecID = " & Guid2Str(Information.CodecID)
                    Dim MimeTypes As ReadOnlyList_1 'ReadOnlyList_String
                    Dim StrMimeTypes As String
                    Set MimeTypes = Information.MimeTypes
                    If IsNotNothing(MimeTypes) Then
                        If MimeTypes.Size > 0 Then
                            Dim MimeTypesItem As Long
                            For MimeTypesItem = 0 To MimeTypes.Size - 1
                                StrMimeTypes = StrMimeTypes & MimeTypes.GetAt(MimeTypesItem) & ", "
                            Next
                            Debug.Print "BitmapDecoder.MimeTypes = " & StrMimeTypes
                            StrMimeTypes = vbNullString
                        End If
                    End If
                    Dim FileExtensions As ReadOnlyList_1 'ReadOnlyList_String
                    Dim StrFileExtensions As String
                    Set FileExtensions = Information.FileExtensions
                    If IsNotNothing(FileExtensions) Then
                        If FileExtensions.Size > 0 Then
                            Dim FileExtensionsItem As Long
                            For FileExtensionsItem = 0 To FileExtensions.Size - 1
                                StrFileExtensions = StrFileExtensions & FileExtensions.GetAt(FileExtensionsItem) & ", "
                            Next
                            Debug.Print "BitmapDecoder.FileExtensions = " & StrFileExtensions
                            StrFileExtensions = vbNullString
                        End If
                    End If
                    Set Information = Nothing
                End If
            Next
        End If
    End If
End Sub

Private Sub Command37_Click()
    Dim CodecInformation As ReadOnlyList_1 'ReadOnlyList_BitmapCodecInformation
    Set CodecInformation = Windows.Graphics.Imaging.BitmapEncoder.GetEncoderInformationEnumerator
    If IsNotNothing(CodecInformation) Then
        If CodecInformation.Size > 0 Then
            Dim CodecItem As Long
            Dim Information As BitmapCodecInformation
            For CodecItem = 0 To CodecInformation.Size - 1
                Set Information = CodecInformation.GetAt(CodecItem)
                If IsNotNothing(Information) Then
                    Debug.Print "---=== BitmapEncoder " & CStr(CodecItem + 1) & " ===---"
                    Debug.Print "BitmapEncoder.FriendlyName = " & Information.FriendlyName
                    Debug.Print "BitmapEncoder.CodecID = " & Guid2Str(Information.CodecID)
                    Dim MimeTypes As ReadOnlyList_1 'ReadOnlyList_String
                    Dim StrMimeTypes As String
                    Set MimeTypes = Information.MimeTypes
                    If IsNotNothing(MimeTypes) Then
                        If MimeTypes.Size > 0 Then
                            Dim MimeTypesItem As Long
                            For MimeTypesItem = 0 To MimeTypes.Size - 1
                                StrMimeTypes = StrMimeTypes & MimeTypes.GetAt(MimeTypesItem) & ", "
                            Next
                            Debug.Print "BitmapEncoder.MimeTypes = " & StrMimeTypes
                            StrMimeTypes = vbNullString
                        End If
                    End If
                    Dim FileExtensions As ReadOnlyList_1 'ReadOnlyList_String
                    Dim StrFileExtensions As String
                    Set FileExtensions = Information.FileExtensions
                    If IsNotNothing(FileExtensions) Then
                        If FileExtensions.Size > 0 Then
                            Dim FileExtensionsItem As Long
                            For FileExtensionsItem = 0 To FileExtensions.Size - 1
                                StrFileExtensions = StrFileExtensions & FileExtensions.GetAt(FileExtensionsItem) & ", "
                            Next
                            Debug.Print "BitmapEncoder.FileExtensions = " & StrFileExtensions
                            StrFileExtensions = vbNullString
                        End If
                    End If
                    Set Information = Nothing
                End If
            Next
        End If
    End If
End Sub

Private Sub Command38_Click()
    Dim FileSavePicker As FileSavePicker
    Set FileSavePicker = Windows.Storage.Pickers.FileSavePicker
    If IsNotNothing(FileSavePicker) Then
        Dim FileTypeFilter As New List_String
        Call FileTypeFilter.Append(".txt")
        Call FileSavePicker.FileTypeChoices.Insert("Plain Text", FileTypeFilter)
        FileSavePicker.SuggestedFileName = "NewFile"
        FileSavePicker.DefaultFileExtension = ".txt"
        FileSavePicker.ParentHwnd = Me.hwnd
        Dim file As StorageFile
        Set file = FileSavePicker.PickSaveFileAsync
        If IsNotNothing(file) Then
            Debug.Print file.Path
        Else
            Debug.Print "FileSavePicker = Cancel"
        End If
    End If
End Sub

Private Sub Command39_Click()
    ' see https://learn.microsoft.com/en-us/windows/uwp/maps-and-location/get-location
    Dim Geolocator As Geolocator
    Set Geolocator = Windows.Devices.Geolocation.Geolocator
    If IsNotNothing(Geolocator) Then
        Select Case Geolocator.RequestAccessAsync
        Case GeolocationAccessStatus.GeolocationAccessStatus_Allowed
            Debug.Print "Access to location is allowed."
            
            ' untested
            Dim Geoposition As Geoposition
            Set Geoposition = Geolocator.GetGeopositionAsync
            If IsNotNothing(Geoposition) Then
                Select Case Geoposition.Coordinate.PositionSource
                    Case PositionSource.PositionSource_Cellular
                        Debug.Print "PositionSource_Cellular"
                    Case PositionSource.PositionSource_Satellite
                        Debug.Print "PositionSource_Satellite"
                    Case PositionSource.PositionSource_WiFi
                        Debug.Print "PositionSource_WiFi"
                    Case PositionSource.PositionSource_IPAddress
                        Debug.Print "PositionSource_IPAddress"
                    Case PositionSource.PositionSource_Unknown
                        Debug.Print "PositionSource_Unknown"
                    Case PositionSource.PositionSource_Default
                        Debug.Print "PositionSource_Default"
                    Case PositionSource.PositionSource_Obfuscated
                        Debug.Print "PositionSource_Obfuscated"
                End Select
                Debug.Print "Position: " & Geoposition.Coordinate.Point.Position.ToString
            End If
            
        Case GeolocationAccessStatus.GeolocationAccessStatus_Denied
            Debug.Print "Access to location is denied."
        Case GeolocationAccessStatus.GeolocationAccessStatus_Unspecified
            Debug.Print "Unspecified error."
        End Select
    End If
End Sub

Private Sub Command40_Click()
    Dim WordsSegmenter As WordsSegmenter
    Set WordsSegmenter = Windows.Data.Text.WordsSegmenter.CreateWithLanguage("de-DE")
    If IsNotNothing(WordsSegmenter) Then
        Debug.Print "ResolvedLanguage: " & WordsSegmenter.ResolvedLanguage
        Dim WordSegments As ReadOnlyList_1 'ReadOnlyList_WordSegment
        Set WordSegments = WordsSegmenter.Tokens("Dieser Text kostet 10€ und wurde um 06:30 Uhr am 12.08.2023 erstellt. Der Text enthält 18 Wörter.")
        If IsNotNothing(WordSegments) Then
            Dim WordSegmentCount As Long
            WordSegmentCount = WordSegments.Size
            Debug.Print "WordSegments: " & CStr(WordSegmentCount)
            If WordSegmentCount > 0 Then
                Dim WordSegmentItem As Long
                For WordSegmentItem = 0 To WordSegmentCount - 1
                    Dim WordSegment As WordSegment
                    Set WordSegment = WordSegments.GetAt(WordSegmentItem)
                    If IsNotNothing(WordSegment) Then
                        Debug.Print "WordSegment " & CStr(WordSegmentItem) & " Text: '" & WordSegment.Text & "'"
                        Debug.Print vbTab & WordSegment.SourceTextSegment.ToString
                        Dim AlternateWordForms As ReadOnlyList_1 'ReadOnlyList_AlternateWordForm
                        Set AlternateWordForms = WordSegment.AlternateForms
                        If IsNotNothing(AlternateWordForms) Then
                            Dim AlternateWordFormCount As Long
                            AlternateWordFormCount = AlternateWordForms.Size
                            If AlternateWordFormCount > 0 Then
                                Dim AlternateWordFormItem As Long
                                For AlternateWordFormItem = 0 To AlternateWordFormCount - 1
                                    Dim AlternateWordForm As AlternateWordForm
                                    Set AlternateWordForm = AlternateWordForms.GetAt(AlternateWordFormItem)
                                    If IsNotNothing(AlternateWordForm) Then
                                        Debug.Print vbTab & "AlternateWordForm.AlternateText: '" & AlternateWordForm.AlternateText & "'"
                                        Debug.Print vbTab & "AlternateWordForm.SourceTextSegment: " & AlternateWordForm.SourceTextSegment.ToString
                                        Select Case AlternateWordForm.NormalizationFormat
                                            Case AlternateNormalizationFormat.AlternateNormalizationFormat_NotNormalized
                                                Debug.Print vbTab & "AlternateWordForm.NormalizationFormat: NotNormalized"
                                            Case AlternateNormalizationFormat.AlternateNormalizationFormat_Number
                                                Debug.Print vbTab & "AlternateWordForm.NormalizationFormat: Number"
                                            Case AlternateNormalizationFormat.AlternateNormalizationFormat_Currency
                                                Debug.Print vbTab & "AlternateWordForm.NormalizationFormat: Currency"
                                            Case AlternateNormalizationFormat.AlternateNormalizationFormat_Date
                                                Debug.Print vbTab & "AlternateWordForm.NormalizationFormat: Date"
                                            Case AlternateNormalizationFormat.AlternateNormalizationFormat_Time
                                                Debug.Print vbTab & "AlternateWordForm.NormalizationFormat: Time"
                                        End Select
                                        Set AlternateWordForm = Nothing
                                    End If
                                Next
                            End If
                            Set AlternateWordForms = Nothing
                        End If
                        Set WordSegment = Nothing
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub Command41_Click()
    Dim SelectableWordsSegmenter As SelectableWordsSegmenter
    Set SelectableWordsSegmenter = Windows.Data.Text.SelectableWordsSegmenter.CreateWithLanguage("de-DE")
    If IsNotNothing(SelectableWordsSegmenter) Then
        Debug.Print "ResolvedLanguage: " & SelectableWordsSegmenter.ResolvedLanguage
        Dim SelectableWordSegments As ReadOnlyList_1 'ReadOnlyList_SelectableWordSegment
        Set SelectableWordSegments = SelectableWordsSegmenter.Tokens("Das ist ein sinnloser Text.")
        If IsNotNothing(SelectableWordSegments) Then
            Dim SelectableWordSegmentCount As Long
            SelectableWordSegmentCount = SelectableWordSegments.Size
            Debug.Print "SelectableWordSegments: " & CStr(SelectableWordSegmentCount)
            If SelectableWordSegmentCount > 0 Then
                Dim SelectableWordSegmentItem As Long
                For SelectableWordSegmentItem = 0 To SelectableWordSegmentCount - 1
                    Dim SelectableWordSegment As SelectableWordSegment
                    Set SelectableWordSegment = SelectableWordSegments.GetAt(SelectableWordSegmentItem)
                    If IsNotNothing(SelectableWordSegment) Then
                        Debug.Print "SelectableWordSegment " & CStr(SelectableWordSegmentItem) & " Text: '" & SelectableWordSegment.Text & "'"
                        Debug.Print vbTab & SelectableWordSegment.SourceTextSegment.ToString
                        Set SelectableWordSegment = Nothing
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub Command42_Click()
    Dim XmlDocument As XmlDocument
    Set XmlDocument = Windows.Data.Xml.Dom.XmlDocument
    If XmlDocument.LoadXML("<MyXML></MyXML>") Then
        Debug.Print "Befor: " & vbNewLine & GetXMLFormated(XmlDocument.GetXml)
        Dim ToastNode As XmlNode
        Set ToastNode = XmlDocument.SelectSingleNode("/MyXML")
        Call ToastNode.SetAttribute("Att1", "Val1")
        Call ToastNode.SetAttribute("Attr2", "Val2")
        Dim XmlElement1 As XmlElement
        Set XmlElement1 = XmlDocument.CreateElement("Element0")
        Call XmlElement1.SetAttribute("Attr3", "Val3")
        Call XmlElement1.SetAttribute("Attr4", "Val4")
        Call ToastNode.AppendChild(XmlElement1)
        Dim XmlElements As XmlElement
        Set XmlElements = XmlDocument.CreateElement("Elements")
        Call ToastNode.AppendChild(XmlElements)
        Dim XmlElement2 As XmlElement
        Set XmlElement2 = XmlDocument.CreateElement("Element1")
        Call XmlElement2.SetAttribute("Attr5", "Val5")
        Call XmlElement2.SetAttribute("Attr6", "Val6")
        Dim XmlElement3 As XmlElement
        Set XmlElement3 = XmlDocument.CreateElement("Element2")
        Call XmlElement3.SetAttribute("Attr7", "Val7")
        Call XmlElement3.SetAttribute("Attr8", "Val8")
        Dim XmlElement4 As XmlElement
        Set XmlElement4 = XmlDocument.CreateElement("Element3")
        Call XmlElement4.SetAttribute("Attr9", "Val9")
        Call XmlElement4.SetAttribute("Attr10", "Val10")
        Dim XmlNode As XmlNode
        Set XmlNode = ToastNode.SelectSingleNode("/MyXML/Elements")
        Call XmlNode.AppendChild(XmlElement2)
        Call XmlNode.AppendChild(XmlElement3)
        Call XmlNode.AppendChild(XmlElement4)
        Debug.Print "After: " & vbNewLine & GetXMLFormated(XmlDocument.GetXml)
    End If
End Sub

Private Sub Command43_Click()

    ' For a complete ToastNotification example, see:
    ' http://www.activevb.de/cgi-bin/upload/download.pl?id=3918
    
    Dim ToastXML As XmlDocument
    Dim ToastTemplate As ToastTemplateType
    ToastTemplate = ToastImageAndText04
    Set ToastXML = Windows.UI.Notifications.ToastNotificationManager.GetTemplateContent(ToastTemplate)
    Dim ToastNode As XmlNode
    Set ToastNode = ToastXML.SelectSingleNode("/toast")
    Call ToastNode.SetAttribute("launch", "toast://my_arguments")
    Call ToastNode.SetAttribute("duration", "short")
    Call ToastNode.SetAttribute("scenario", "incomingCall")
    Call ToastNode.SetAttribute("useButtonStyle", "true")
    Dim VisualNode As XmlNode
    Set VisualNode = ToastXML.SelectSingleNode("/toast/visual")
    Dim BindingNode As XmlNode
    Set BindingNode = ToastXML.SelectSingleNode("/toast/visual/binding")
    If ToastTemplate <= ToastImageAndText04 Then
        Dim ToastImageElements As XmlNodeList
        Set ToastImageElements = ToastXML.GetElementsByTagName("image")
        Call ToastImageElements.Item(0).SetAttribute("alt", "AVB")
        Call ToastImageElements.Item(0).SetAttribute("src", "file:///" & App.Path & "\avb.jpg")
        Call ToastImageElements.Item(0).SetAttribute("hint-crop", "circle")
    End If
    Dim ToastTextElements As XmlNodeList
    Set ToastTextElements = ToastXML.GetElementsByTagName("text")
    Select Case ToastTemplate
        Case ToastTemplateType.ToastImageAndText01, ToastTemplateType.ToastText01
            Call ToastTextElements.Item(0).AppendChild(ToastXML.CreateTextNode("Text 1"))
        Case ToastTemplateType.ToastImageAndText02, ToastTemplateType.ToastText02
            Call ToastTextElements.Item(0).AppendChild(ToastXML.CreateTextNode("Titel"))
            Call ToastTextElements.Item(1).AppendChild(ToastXML.CreateTextNode("Text 1"))
        Case ToastTemplateType.ToastImageAndText03, ToastTemplateType.ToastText03
            Call ToastTextElements.Item(0).AppendChild(ToastXML.CreateTextNode("The title will wrap onto a second line if it is too long."))
            Call ToastTextElements.Item(1).AppendChild(ToastXML.CreateTextNode("Text 1"))
        Case ToastTemplateType.ToastImageAndText04, ToastTemplateType.ToastText04
            Call ToastTextElements.Item(0).AppendChild(ToastXML.CreateTextNode("Incoming Video Call " & ChrW(&H260E) & ChrW(&HFE0F)))
            Call ToastTextElements.Item(1).AppendChild(ToastXML.CreateTextNode("from 32168 (Rosi) " & ChrW(&HD83D) & ChrW(&HDC69)))
            Call ToastTextElements.Item(2).AppendChild(ToastXML.CreateTextNode("for you. " & ChrW(&HD83E) & ChrW(&HDDD4)))
    End Select
    Dim ToastAudioElement As XmlElement
    Set ToastAudioElement = ToastXML.CreateElement("audio")
    Call ToastAudioElement.SetAttribute("src", "ms-winsoundevent:Notification.Looping.Call10")
    Call ToastAudioElement.SetAttribute("loop", "true")
    Call ToastAudioElement.SetAttribute("silent", "false")
    Call ToastNode.AppendChild(ToastAudioElement)
    Dim ToastCommandsElement As XmlElement
    Set ToastCommandsElement = ToastXML.CreateElement("commands")
    Call ToastCommandsElement.SetAttribute("scenario", "incomingCall")
    Call ToastNode.AppendChild(ToastCommandsElement)
    Dim ToastCommandElement0 As XmlElement
    Set ToastCommandElement0 = ToastXML.CreateElement("command")
    Call ToastCommandElement0.SetAttribute("id", "video")
    Call ToastCommandElement0.SetAttribute("arguments", "start video call")
    Dim ToastCommandElement1 As XmlElement
    Set ToastCommandElement1 = ToastXML.CreateElement("command")
    Call ToastCommandElement1.SetAttribute("id", "voice")
    Call ToastCommandElement1.SetAttribute("arguments", "start voice call")
    Dim ToastCommandElement2 As XmlElement
    Set ToastCommandElement2 = ToastXML.CreateElement("command")
    Call ToastCommandElement2.SetAttribute("id", "decline")
    Call ToastCommandElement2.SetAttribute("arguments", "call declined")
    Dim ToastCommandsNode As XmlNode
    Set ToastCommandsNode = ToastXML.SelectSingleNode("/toast/commands")
    Call ToastCommandsNode.AppendChild(ToastCommandElement0)
    Call ToastCommandsNode.AppendChild(ToastCommandElement1)
    Call ToastCommandsNode.AppendChild(ToastCommandElement2)
    Dim Toast As ToastNotification
    Set Toast = Windows.UI.Notifications.ToastNotification.CreateToastNotification(ToastXML)
    With Toast
        .Group = "MyToastGroup"
        .Tag = "StandardToast"
    End With
    Debug.Print GetXMLFormated(ToastXML.GetXml)
    If Windows.UI.Notifications.ToastNotificationManager.CreateToastNotifierWithId("VBC ToastNotification").Show(Toast) Then
        Toast.AddToastFailedEventHandler = ITEH_ToastFailedEvent.Create(Me)
        Toast.AddToastActivatedEventHandler = ITEH_ToastActivatedEvent.Create(Me)
        Toast.AddToastDismissedEventHandler = ITEH_ToastDismissedEvent.Create(Me)
    End If
End Sub

Private Sub Command44_Click()
    Dim Users As ReadOnlyList_1 'ReadOnlyList_User
    Set Users = Windows.System.User.FindAllAsync
    If IsNotNothing(Users) Then
        Dim UsersCount As Long
        UsersCount = Users.Size
        If UsersCount > 0& Then
            Dim UsersItem As Long
            For UsersItem = 0 To UsersCount - 1
                Dim User As User
                Set User = Users.GetAt(UsersItem)
                Debug.Print "NonRoamableId: " & User.NonRoamableId
                Select Case User.UserType
                    Case UserType.UserType_LocalUser
                        Debug.Print "UserType: LocalUser"
                    Case UserType.UserType_RemoteUser
                        Debug.Print "UserType: RemoteUser"
                    Case UserType.UserType_LocalGuest
                        Debug.Print "UserType: LocalGuest"
                    Case UserType.UserType_RemoteGuest
                        Debug.Print "UserType: RemoteGuest"
                    Case UserType.UserType_SystemManaged
                        Debug.Print "UserType: SystemManaged"
                End Select
                Select Case User.AuthenticationStatus
                    Case UserAuthenticationStatus.UserAuthenticationStatus_Unauthenticated
                        Debug.Print "UserAuthenticationStatus: Unauthenticated"
                    Case UserAuthenticationStatus.UserAuthenticationStatus_LocallyAuthenticated
                        Debug.Print "UserAuthenticationStatus: LocallyAuthenticated"
                    Case UserAuthenticationStatus.UserAuthenticationStatus_RemotelyAuthenticated
                        Debug.Print "UserAuthenticationStatus: RemotelyAuthenticated"
                End Select
                Dim KnownUserProperties As KnownUserProperties
                Set KnownUserProperties = Windows.System.KnownUserProperties
                Debug.Print "DisplayName: " & User.GetPropertyAsync(KnownUserProperties.DisplayName).GetString
                Debug.Print "FirstName: " & User.GetPropertyAsync(KnownUserProperties.FirstName).GetString
                Debug.Print "LastName: " & User.GetPropertyAsync(KnownUserProperties.LastName).GetString
                Debug.Print "ProviderName: " & User.GetPropertyAsync(KnownUserProperties.ProviderName).GetString
                Debug.Print "AccountName: " & User.GetPropertyAsync(KnownUserProperties.AccountName).GetString
                Debug.Print "GuestHost: " & User.GetPropertyAsync(KnownUserProperties.GuestHost).GetString
                Debug.Print "PrincipalName: " & User.GetPropertyAsync(KnownUserProperties.PrincipalName).GetString
                Debug.Print "DomainName: " & User.GetPropertyAsync(KnownUserProperties.DomainName).GetString
                Debug.Print "SessionInitiationProtocolUri: " & User.GetPropertyAsync(KnownUserProperties.SessionInitiationProtocolUri).GetString
                Debug.Print "AgeEnforcementRegion: " & User.GetPropertyAsync(KnownUserProperties.AgeEnforcementRegion).GetString
                Select Case User.CheckUserAgeConsentGroupAsync(UserAgeConsentGroup_Adult)
                    Case UserAgeConsentResult.UserAgeConsentResult_NotEnforced
                        Debug.Print "UserAgeConsentGroup Adult: NotEnforced"
                    Case UserAgeConsentResult.UserAgeConsentResult_Included
                        Debug.Print "UserAgeConsentGroup Adult: Included"
                    Case UserAgeConsentResult.UserAgeConsentResult_NotIncluded
                        Debug.Print "UserAgeConsentGroup Adult: NotIncluded"
                    Case UserAgeConsentResult.UserAgeConsentResult_Unknown
                        Debug.Print "UserAgeConsentGroup Adult: Unknown"
                    Case UserAgeConsentResult.UserAgeConsentResult_Ambiguous
                        Debug.Print "UserAgeConsentGroup Adult: Ambiguous"
                End Select
                Select Case User.CheckUserAgeConsentGroupAsync(UserAgeConsentGroup_Child)
                    Case UserAgeConsentResult.UserAgeConsentResult_NotEnforced
                        Debug.Print "UserAgeConsentGroup Child: NotEnforced"
                    Case UserAgeConsentResult.UserAgeConsentResult_Included
                        Debug.Print "UserAgeConsentGroup Child: Included"
                    Case UserAgeConsentResult.UserAgeConsentResult_NotIncluded
                        Debug.Print "UserAgeConsentGroup Child: NotIncluded"
                    Case UserAgeConsentResult.UserAgeConsentResult_Unknown
                        Debug.Print "UserAgeConsentGroup Child: Unknown"
                    Case UserAgeConsentResult.UserAgeConsentResult_Ambiguous
                        Debug.Print "UserAgeConsentGroup Child: Ambiguous"
                End Select
                Select Case User.CheckUserAgeConsentGroupAsync(UserAgeConsentGroup_Minor)
                    Case UserAgeConsentResult.UserAgeConsentResult_NotEnforced
                        Debug.Print "UserAgeConsentGroup Minor: NotEnforced"
                    Case UserAgeConsentResult.UserAgeConsentResult_Included
                        Debug.Print "UserAgeConsentGroup Minor: Included"
                    Case UserAgeConsentResult.UserAgeConsentResult_NotIncluded
                        Debug.Print "UserAgeConsentGroup Minor: NotIncluded"
                    Case UserAgeConsentResult.UserAgeConsentResult_Unknown
                        Debug.Print "UserAgeConsentGroup Minor: Unknown"
                    Case UserAgeConsentResult.UserAgeConsentResult_Ambiguous
                        Debug.Print "UserAgeConsentGroup Minor: Ambiguous"
                End Select
                Dim RandomAccessStreamWithContentType As RandomAccessStreamWithContentType
                Set RandomAccessStreamWithContentType = User.GetPictureAsync(UserPictureSize_Size424x424).OpenReadAsync
                If IsNotNothing(RandomAccessStreamWithContentType) Then
                    Dim pIStream As Long
                    pIStream = RandomAccessStreamWithContentType.ToIStream
                    Load frmUserPicture
                    frmUserPicture.Picture = GetPictureFromIStream(pIStream)
                    Call ReleaseIfc(pIStream)
                    Set RandomAccessStreamWithContentType = Nothing
                    frmUserPicture.Show vbModal, Me
                    Unload frmUserPicture
                End If
                Set User = Nothing
            Next
        End If
    End If
End Sub

Private Sub Command45_Click()
    Load frmGlobalSystemMediaTransport
    frmGlobalSystemMediaTransport.Show vbModal, Me
    Unload frmGlobalSystemMediaTransport
End Sub

Private Sub Command46_Click()
    Load frmMediaPlayer
    frmMediaPlayer.Show vbModal, Me
    Unload frmMediaPlayer
End Sub

Private Sub Command47_Click()
    Load frmSpeechSynthesizer
    frmSpeechSynthesizer.Show vbModal, Me
    Unload frmSpeechSynthesizer
End Sub

Private Sub Command48_Click()
    Load frmSpeechRecognizer
    frmSpeechRecognizer.Show vbModal, Me
    Unload frmSpeechRecognizer
End Sub

Private Sub Command49_Click()
    Debug.Print "KeyboardIsPresent = " & CStr(CBool(Windows.Devices.InputDevice.KeyboardCapabilities.KeyboardPresent))
End Sub

Private Sub Command50_Click()
    Dim MouseCapabilities As MouseCapabilities
    Set MouseCapabilities = Windows.Devices.InputDevice.MouseCapabilities
    If IsNotNothing(MouseCapabilities) Then
        If MouseCapabilities.MousePresent <> 0& Then
            Debug.Print "MouseIsPresent = True"
            Debug.Print "HorizontalWheelIsPresent = " & CStr(CBool(MouseCapabilities.HorizontalWheelPresent))
            Debug.Print "VerticalWheelIsPresent = " & CStr(CBool(MouseCapabilities.VerticalWheelPresent))
            Debug.Print "SwapButtons = " & MouseCapabilities.SwapButtons
            Debug.Print "NumberOfButtons = " & MouseCapabilities.NumberOfButtons
        Else
            Debug.Print "MouseIsPresent = False"
        End If
    End If
End Sub

Private Sub Command51_Click()
    Dim DateTimeNow As DateTime
    Dim DateTimeFormatter As DateTimeFormatter
    Set DateTimeNow = Windows.Foundation.DateTime.Now
    ' https://learn.microsoft.com/en-us/uwp/api/windows.globalization.datetimeformatting.datetimeformatter?view=winrt-22621
    'Set DateTimeFormatter = Windows.Globalization.DateTimeFormatting.DateTimeFormatter.CreateDateTimeFormatter("shorttime")
    'Set DateTimeFormatter = Windows.Globalization.DateTimeFormatting.DateTimeFormatter.CreateDateTimeFormatter("longtime")
    'Set DateTimeFormatter = Windows.Globalization.DateTimeFormatting.DateTimeFormatter.CreateDateTimeFormatter("shortdate")
    'Set DateTimeFormatter = Windows.Globalization.DateTimeFormatting.DateTimeFormatter.CreateDateTimeFormatter("longdate")
    'Set DateTimeFormatter = Windows.Globalization.DateTimeFormatting.DateTimeFormatter.CreateDateTimeFormatter("month day dayofweek year")
    Set DateTimeFormatter = Windows.Globalization.DateTimeFormatting.DateTimeFormatter.CreateDateTimeFormatter("longdate longtime")
    If IsNotNothing(DateTimeFormatter) Then
        Debug.Print "DateTimeFormatter.Template: " & DateTimeFormatter.Template
        Dim PatternList As ReadOnlyList_1 'ReadOnlyList_String
        Set PatternList = DateTimeFormatter.Patterns
        If IsNotNothing(PatternList) Then
            Dim PatternListCount As Long
            PatternListCount = PatternList.Size
            If PatternListCount > 0& Then
                Dim PatternListItem As Long
                For PatternListItem = 0 To PatternListCount - 1
                    Debug.Print "DateTimeFormatter.Pattern(" & CStr(PatternListItem) & "): " & Replace$(PatternList.GetAt(PatternListItem), ChrW(8206), vbNullString)
                Next
            End If
        End If
        Debug.Print "DateTimeFormatter.Format: " & Replace$(DateTimeFormatter.Format(DateTimeNow), ChrW(8206), vbNullString)
        ' TimeZone -> https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
        Debug.Print "DateTimeFormatter.FormatUsingTimeZone -> Etc/GMT-0: " & Replace$(DateTimeFormatter.FormatUsingTimeZone(DateTimeNow, "Etc/GMT-0"), ChrW(8206), vbNullString)
        Debug.Print "DateTimeFormatter.FormatUsingTimeZone -> Europe/Berlin: " & Replace$(DateTimeFormatter.FormatUsingTimeZone(DateTimeNow, "Europe/Berlin"), ChrW(8206), vbNullString)
        Debug.Print "DateTimeFormatter.FormatUsingTimeZone -> America/Los_Angeles: " & Replace$(DateTimeFormatter.FormatUsingTimeZone(DateTimeNow, "America/Los_Angeles"), ChrW(8206), vbNullString)
        Debug.Print "DateTimeFormatter.FormatUsingTimeZone -> Antarctica/South_Pole: " & Replace$(DateTimeFormatter.FormatUsingTimeZone(DateTimeNow, "Antarctica/South_Pole"), ChrW(8206), vbNullString)
        Debug.Print "DateTimeFormatter.Calendar: " & DateTimeFormatter.Calendar
        Debug.Print "DateTimeFormatter.Clock: " & DateTimeFormatter.Clock
        Debug.Print "DateTimeFormatter.NumeralSystem: " & DateTimeFormatter.NumeralSystem
        Debug.Print "DateTimeFormatter.GeographicRegion: " & DateTimeFormatter.GeographicRegion
        Debug.Print "DateTimeFormatter.ResolvedGeographicRegion: " & DateTimeFormatter.ResolvedGeographicRegion
        Debug.Print "DateTimeFormatter.ResolvedLanguage: " & DateTimeFormatter.ResolvedLanguage
        Dim LanguageList As ReadOnlyList_1 'ReadOnlyList_String
        Set LanguageList = DateTimeFormatter.Languages
        If IsNotNothing(LanguageList) Then
            Dim LanguageListCount As Long
            LanguageListCount = LanguageList.Size
            If LanguageListCount > 0& Then
                Dim LanguageListItem As Long
                For LanguageListItem = 0 To LanguageListCount - 1
                    Debug.Print "DateTimeFormatter.Language(" & CStr(LanguageListItem) & "): " & LanguageList.GetAt(LanguageListItem)
                Next
            End If
        End If
    End If
End Sub

Private Sub Command52_Click()
    Dim DoubleValue As Double
    DoubleValue = 67123.45
    
    Dim CurrencyFormatter As CurrencyFormatter
    'Set CurrencyFormatter = Windows.Globalization.NumberFormatting.CurrencyFormatter.CreateCurrencyFormatterCode("USD")
    Set CurrencyFormatter = Windows.Globalization.NumberFormatting.CurrencyFormatter.CreateCurrencyFormatterCode("EUR")
    If IsNotNothing(CurrencyFormatter) Then
        CurrencyFormatter.Mode = CurrencyFormatterMode_UseSymbol
        'CurrencyFormatter.Mode = CurrencyFormatterMode_UseCurrencyCode
    
        Debug.Print "CurrencyFormatter.Currency: " & CurrencyFormatter.CurrencyFormat
        Debug.Print "CurrencyFormatter.GeographicRegion: " & CurrencyFormatter.GeographicRegion
        Debug.Print "CurrencyFormatter.ResolvedGeographicRegion: " & CurrencyFormatter.ResolvedGeographicRegion
        Debug.Print "CurrencyFormatter.NumeralSystem: " & CurrencyFormatter.NumeralSystem
        
        Dim LanguageList As ReadOnlyList_1 'ReadOnlyList_String
        Set LanguageList = CurrencyFormatter.Languages
        If IsNotNothing(LanguageList) Then
            Dim LanguageListCount As Long
            LanguageListCount = LanguageList.Size
            If LanguageListCount > 0& Then
                Dim LanguageListItem As Long
                For LanguageListItem = 0 To LanguageListCount - 1
                    Debug.Print "CurrencyFormatter.Language(" & CStr(LanguageListItem) & "): " & LanguageList.GetAt(LanguageListItem)
                Next
            End If
        End If
        
        Debug.Print "CurrencyFormatter.FormatDouble: " & CurrencyFormatter.FormatDouble(DoubleValue)
        Call CurrencyFormatter.ApplyRoundingForCurrency(RoundingAlgorithm_RoundDown)
        Debug.Print "CurrencyFormatter.FormatDouble RoundDown: " & CurrencyFormatter.FormatDouble(DoubleValue)
        Call CurrencyFormatter.ApplyRoundingForCurrency(RoundingAlgorithm_RoundUp)
        Debug.Print "CurrencyFormatter.FormatDouble RoundUp: " & CurrencyFormatter.FormatDouble(DoubleValue)
        
    End If
End Sub

Private Sub Command53_Click()
    Dim DoubleValue As Double
    DoubleValue = 67123.45
    Dim DecimalFormatter As DecimalFormatter
    Set DecimalFormatter = Windows.Globalization.NumberFormatting.DecimalFormatter
    If IsNotNothing(DecimalFormatter) Then
    
        Debug.Print "DecimalFormatter.GeographicRegion: " & DecimalFormatter.GeographicRegion
        Debug.Print "DecimalFormatter.ResolvedGeographicRegion: " & DecimalFormatter.ResolvedGeographicRegion
        Debug.Print "DecimalFormatter.NumeralSystem: " & DecimalFormatter.NumeralSystem
        
        Dim LanguageList As ReadOnlyList_1 'ReadOnlyList_String
        Set LanguageList = DecimalFormatter.Languages
        If IsNotNothing(LanguageList) Then
            Dim LanguageListCount As Long
            LanguageListCount = LanguageList.Size
            If LanguageListCount > 0& Then
                Dim LanguageListItem As Long
                For LanguageListItem = 0 To LanguageListCount - 1
                    Debug.Print "DecimalFormatter.Language(" & CStr(LanguageListItem) & "): " & LanguageList.GetAt(LanguageListItem)
                Next
            End If
        End If
    
        Debug.Print "DecimalFormatter.FormatDouble: " & DecimalFormatter.FormatDouble(DoubleValue)
        Debug.Print "DecimalFormatter.ParseDouble: " & DecimalFormatter.ParseDouble("1234,55")
    End If
End Sub

Private Sub Command54_Click()
    Dim DoubleValue As Double
    DoubleValue = 0.5
    Dim PercentFormatter As PercentFormatter
    Set PercentFormatter = Windows.Globalization.NumberFormatting.PercentFormatter
    If IsNotNothing(PercentFormatter) Then
    
        Debug.Print "PercentFormatter.GeographicRegion: " & PercentFormatter.GeographicRegion
        Debug.Print "PercentFormatter.ResolvedGeographicRegion: " & PercentFormatter.ResolvedGeographicRegion
        Debug.Print "PercentFormatter.NumeralSystem: " & PercentFormatter.NumeralSystem
        
        Dim LanguageList As ReadOnlyList_1 'ReadOnlyList_String
        Set LanguageList = PercentFormatter.Languages
        If IsNotNothing(LanguageList) Then
            Dim LanguageListCount As Long
            LanguageListCount = LanguageList.Size
            If LanguageListCount > 0& Then
                Dim LanguageListItem As Long
                For LanguageListItem = 0 To LanguageListCount - 1
                    Debug.Print "PercentFormatter.Language(" & CStr(LanguageListItem) & "): " & LanguageList.GetAt(LanguageListItem)
                Next
            End If
        End If
    
        Debug.Print "PercentFormatter.FormatDouble: " & PercentFormatter.FormatDouble(DoubleValue)
        Debug.Print "PercentFormatter.ParseDouble: " & PercentFormatter.ParseDouble("98 %")
    End If
End Sub

Private Sub Command55_Click()
    Dim DoubleValue As Double
    DoubleValue = 0.1
    Dim PermilleFormatter As PermilleFormatter
    Set PermilleFormatter = Windows.Globalization.NumberFormatting.PermilleFormatter
    If IsNotNothing(PermilleFormatter) Then
    
        Debug.Print "PermilleFormatter.GeographicRegion: " & PermilleFormatter.GeographicRegion
        Debug.Print "PermilleFormatter.ResolvedGeographicRegion: " & PermilleFormatter.ResolvedGeographicRegion
        Debug.Print "PermilleFormatter.NumeralSystem: " & PermilleFormatter.NumeralSystem
        
        Dim LanguageList As ReadOnlyList_1 'ReadOnlyList_String
        Set LanguageList = PermilleFormatter.Languages
        If IsNotNothing(LanguageList) Then
            Dim LanguageListCount As Long
            LanguageListCount = LanguageList.Size
            If LanguageListCount > 0& Then
                Dim LanguageListItem As Long
                For LanguageListItem = 0 To LanguageListCount - 1
                    Debug.Print "PermilleFormatter.Language(" & CStr(LanguageListItem) & "): " & LanguageList.GetAt(LanguageListItem)
                Next
            End If
        End If
    
        Debug.Print "PermilleFormatter.FormatDouble: " & PermilleFormatter.FormatDouble(DoubleValue)
        Debug.Print "PermilleFormatter.ParseDouble: " & PermilleFormatter.ParseDouble("100 " & ChrW(8240))
    End If
End Sub

Private Sub Command56_Click()
    Dim NumeralSystemTranslator As NumeralSystemTranslator
    Set NumeralSystemTranslator = Windows.Globalization.NumberFormatting.NumeralSystemTranslator
    If IsNotNothing(NumeralSystemTranslator) Then
        
        Dim LanguageList As ReadOnlyList_1 'ReadOnlyList_String
        Set LanguageList = NumeralSystemTranslator.Languages
        If IsNotNothing(LanguageList) Then
            Dim LanguageListCount As Long
            LanguageListCount = LanguageList.Size
            If LanguageListCount > 0& Then
                Dim LanguageListItem As Long
                For LanguageListItem = 0 To LanguageListCount - 1
                    Debug.Print "NumeralSystemTranslator.Language(" & CStr(LanguageListItem) & "): " & LanguageList.GetAt(LanguageListItem)
                Next
            End If
        End If
        Debug.Print "NumeralSystemTranslator.ResolvedLanguage: " & NumeralSystemTranslator.ResolvedLanguage
        Debug.Print "NumeralSystemTranslator.NumeralSystem: " & NumeralSystemTranslator.NumeralSystem
        Debug.Print "NumeralSystemTranslator.TranslateNumerals: " & NumeralSystemTranslator.TranslateNumerals("123.45")
    End If
End Sub

Private Sub Command57_Click()
    Dim DoubleValue As Double
    DoubleValue = 123.45
    Dim SignificantDigitsNumberRounder As SignificantDigitsNumberRounder
    Set SignificantDigitsNumberRounder = Windows.Globalization.NumberFormatting.SignificantDigitsNumberRounder
    If IsNotNothing(SignificantDigitsNumberRounder) Then
        Debug.Print "SignificantDigitsNumberRounder.RoundingAlgorithm: " & SignificantDigitsNumberRounder.RoundingAlgorithm
        Debug.Print "SignificantDigitsNumberRounder.SignificantDigits: " & SignificantDigitsNumberRounder.SignificantDigits
        Debug.Print "SignificantDigitsNumberRounder.RoundDouble: " & SignificantDigitsNumberRounder.RoundDouble(DoubleValue)
        SignificantDigitsNumberRounder.RoundingAlgorithm = RoundingAlgorithm_None
        SignificantDigitsNumberRounder.SignificantDigits = 5
        Debug.Print "Set SignificantDigitsNumberRounder.RoundingAlgorithm = RoundingAlgorithm_None"
        Debug.Print "Set SignificantDigitsNumberRounder.SignificantDigits = 5"
        Debug.Print "SignificantDigitsNumberRounder.RoundDouble: " & SignificantDigitsNumberRounder.RoundDouble(DoubleValue)
    End If
End Sub

Private Sub Command58_Click()
    Dim DoubleValue As Double
    DoubleValue = 123.45
    Dim IncrementNumberRounder As IncrementNumberRounder
    Set IncrementNumberRounder = Windows.Globalization.NumberFormatting.IncrementNumberRounder
    If IsNotNothing(IncrementNumberRounder) Then
        Debug.Print "IncrementNumberRounder.RoundingAlgorithm: " & IncrementNumberRounder.RoundingAlgorithm
        Debug.Print "IncrementNumberRounder.Increment: " & IncrementNumberRounder.Increment
        Debug.Print "IncrementNumberRounder.RoundDouble: " & IncrementNumberRounder.RoundDouble(DoubleValue)
        IncrementNumberRounder.RoundingAlgorithm = RoundingAlgorithm_None
        IncrementNumberRounder.Increment = 2
        Debug.Print "Set IncrementNumberRounder.RoundingAlgorithm = RoundingAlgorithm_None"
        Debug.Print "Set IncrementNumberRounder.SIncrement = 2"
        Debug.Print "IncrementNumberRounder.RoundDouble: " & IncrementNumberRounder.RoundDouble(DoubleValue)
    End If
End Sub

Private Sub Command59_Click()
    Dim PhoneNumberInfo As PhoneNumberInfo
    Set PhoneNumberInfo = Windows.Globalization.PhoneNumberFormatting.PhoneNumberInfo.Create("01780000000")
    If IsNotNothing(PhoneNumberInfo) Then
        Debug.Print "PhoneNumberInfo.PhoneNumber: " & PhoneNumberInfo.PhoneNumber
        Debug.Print "PhoneNumberInfo.ToString: " & PhoneNumberInfo.ToString
        Debug.Print "PhoneNumberInfo.CountryCode: " & PhoneNumberInfo.CountryCode
        Debug.Print "PhoneNumberInfo.GetGeographicRegionCode: " & PhoneNumberInfo.GetGeographicRegionCode
        Debug.Print "PhoneNumberInfo.GetLengthOfGeographicalAreaCode: " & PhoneNumberInfo.GetLengthOfGeographicalAreaCode
        Debug.Print "PhoneNumberInfo.GetLengthOfNationalDestinationCode: " & PhoneNumberInfo.GetLengthOfNationalDestinationCode
        Debug.Print "PhoneNumberInfo.GetNationalSignificantNumber: " & PhoneNumberInfo.GetNationalSignificantNumber
        
        Dim strPredictNumberKind As String
        Select Case PhoneNumberInfo.PredictNumberKind
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_FixedLine
                strPredictNumberKind = "FixedLine"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Mobile
                strPredictNumberKind = "Mobile"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_FixedLineOrMobile
                strPredictNumberKind = "FixedLineOrMobile"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_TollFree
                strPredictNumberKind = "TollFree"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_PremiumRate
                strPredictNumberKind = "PremiumRate"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_SharedCost
                strPredictNumberKind = "SharedCost"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Voip
                strPredictNumberKind = "Voip"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_PersonalNumber
                strPredictNumberKind = "PersonalNumber"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Pager
                strPredictNumberKind = "Pager"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_UniversalAccountNumber
                strPredictNumberKind = "UniversalAccountNumber"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Voicemail
                strPredictNumberKind = "Voicemail"
            Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Unknown
                strPredictNumberKind = "Unknown"
        End Select
        Debug.Print "PhoneNumberInfo.PredictNumberKind: " & strPredictNumberKind
        
        Dim strNumberMatch As String
        Select Case PhoneNumberInfo.CheckNumberMatch(Windows.Globalization.PhoneNumberFormatting.PhoneNumberInfo.Create("+491780000000"))
            Case PhoneNumberMatchResult.PhoneNumberMatchResult_NoMatch
                strNumberMatch = "NoMatch"
            Case PhoneNumberMatchResult.PhoneNumberMatchResult_ShortNationalSignificantNumberMatch
                strNumberMatch = "ShortNationalSignificantNumberMatch"
            Case PhoneNumberMatchResult.PhoneNumberMatchResult_NationalSignificantNumberMatch
                strNumberMatch = "NationalSignificantNumberMatch"
            Case PhoneNumberMatchResult.PhoneNumberMatchResult_ExactMatch
                strNumberMatch = "ExactMatch"
        End Select
        Debug.Print "PhoneNumberInfo.CheckNumberMatch: " & strNumberMatch
    Else
        Debug.Print "Invalid phone number."
    End If
End Sub

Private Sub Command60_Click()
    Dim TelNumber As String
    Dim RegionCode As String
    RegionCode = "DE"
    TelNumber = "01780000000"
    
    Dim PhoneNumberFormatter As PhoneNumberFormatter
    Set PhoneNumberFormatter = Windows.Globalization.PhoneNumberFormatting.PhoneNumberFormatter.TryCreate(RegionCode)
    If IsNotNothing(PhoneNumberFormatter) Then
        Debug.Print "PhoneNumberFormatter.GetCountryCodeForRegion: " & PhoneNumberFormatter.GetCountryCodeForRegion(RegionCode)
        Debug.Print "PhoneNumberFormatter.GetNationalDirectDialingPrefixForRegion StripNonDigit = False: " & PhoneNumberFormatter.GetNationalDirectDialingPrefixForRegion(RegionCode, False)
        Debug.Print "PhoneNumberFormatter.GetNationalDirectDialingPrefixForRegion StripNonDigit = True: " & PhoneNumberFormatter.GetNationalDirectDialingPrefixForRegion(RegionCode, True)
        Debug.Print "PhoneNumberFormatter.WrapWithLeftToRightMarkers: " & PhoneNumberFormatter.WrapWithLeftToRightMarkers(TelNumber)
        Dim PhoneNumberInfo As PhoneNumberInfo
        Set PhoneNumberInfo = Windows.Globalization.PhoneNumberFormatting.PhoneNumberInfo.Create(TelNumber)
        If IsNotNothing(PhoneNumberInfo) Then
            Debug.Print "PhoneNumberFormatter.FormatStringWithLeftToRightMarkers: " & PhoneNumberFormatter.GetCountryCodeForRegion(RegionCode)
            Debug.Print "PhoneNumberFormatter.Format: " & PhoneNumberFormatter.Format(PhoneNumberInfo)
            Debug.Print "PhoneNumberFormatter.FormatWithOutputFormat International: " & PhoneNumberFormatter.FormatWithOutputFormat(PhoneNumberInfo, PhoneNumberFormat_International)
            Debug.Print "PhoneNumberFormatter.FormatPartialString: " & PhoneNumberFormatter.FormatPartialString(TelNumber)
            Debug.Print "PhoneNumberFormatter.FormatString: " & PhoneNumberFormatter.FormatString(TelNumber)
            Debug.Print "PhoneNumberFormatter.FormatStringWithLeftToRightMarkers: " & PhoneNumberFormatter.FormatStringWithLeftToRightMarkers(TelNumber)
            
            Debug.Print "PhoneNumberInfo.PhoneNumber: " & PhoneNumberInfo.PhoneNumber
            Debug.Print "PhoneNumberInfo.ToString: " & PhoneNumberInfo.ToString
            Debug.Print "PhoneNumberInfo.CountryCode: " & PhoneNumberInfo.CountryCode
            Debug.Print "PhoneNumberInfo.GetGeographicRegionCode: " & PhoneNumberInfo.GetGeographicRegionCode
            Debug.Print "PhoneNumberInfo.GetLengthOfGeographicalAreaCode: " & PhoneNumberInfo.GetLengthOfGeographicalAreaCode
            Debug.Print "PhoneNumberInfo.GetLengthOfNationalDestinationCode: " & PhoneNumberInfo.GetLengthOfNationalDestinationCode
            Debug.Print "PhoneNumberInfo.GetNationalSignificantNumber: " & PhoneNumberInfo.GetNationalSignificantNumber
            
            Dim strPredictNumberKind As String
            Select Case PhoneNumberInfo.PredictNumberKind
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_FixedLine
                    strPredictNumberKind = "FixedLine"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Mobile
                    strPredictNumberKind = "Mobile"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_FixedLineOrMobile
                    strPredictNumberKind = "FixedLineOrMobile"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_TollFree
                    strPredictNumberKind = "TollFree"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_PremiumRate
                    strPredictNumberKind = "PremiumRate"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_SharedCost
                    strPredictNumberKind = "SharedCost"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Voip
                    strPredictNumberKind = "Voip"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_PersonalNumber
                    strPredictNumberKind = "PersonalNumber"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Pager
                    strPredictNumberKind = "Pager"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_UniversalAccountNumber
                    strPredictNumberKind = "UniversalAccountNumber"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Voicemail
                    strPredictNumberKind = "Voicemail"
                Case PredictedPhoneNumberKind.PredictedPhoneNumberKind_Unknown
                    strPredictNumberKind = "Unknown"
            End Select
            Debug.Print "PhoneNumberInfo.PredictNumberKind: " & strPredictNumberKind
        Else
            Debug.Print "Invalid phone number."
        End If
    Else
        Debug.Print "Invalid RegionCode."
    End If
End Sub

Private Sub Command61_Click()
    Dim GeographicRegion As GeographicRegion
    Set GeographicRegion = Windows.Globalization.GeographicRegion
    If IsNotNothing(GeographicRegion) Then
        Debug.Print "GeographicRegion.Code: " & GeographicRegion.Code
        Debug.Print "GeographicRegion.CodeTwoLetter: " & GeographicRegion.CodeTwoLetter
        Debug.Print "GeographicRegion.CodeThreeLetter: " & GeographicRegion.CodeThreeLetter
        Debug.Print "GeographicRegion.CodeThreeDigit: " & GeographicRegion.CodeThreeDigit
        Debug.Print "GeographicRegion.DisplayName: " & GeographicRegion.DisplayName
        Debug.Print "GeographicRegion.NativeName: " & GeographicRegion.NativeName
        Dim CurrenciesInUseList As ReadOnlyList_1 'ReadOnlyList_String
        Set CurrenciesInUseList = GeographicRegion.CurrenciesInUse
        If IsNotNothing(CurrenciesInUseList) Then
            Dim CurrenciesInUseListCount As Long
            CurrenciesInUseListCount = CurrenciesInUseList.Size
            If CurrenciesInUseListCount > 0& Then
                Dim CurrenciesInUseListItem As Long
                For CurrenciesInUseListItem = 0 To CurrenciesInUseListCount - 1
                    Debug.Print "CurrenciesInUse(" & CStr(CurrenciesInUseListItem) & "): " & CurrenciesInUseList.GetAt(CurrenciesInUseListItem)
                Next
            End If
        End If
        Debug.Print String$(50, "-")
        Dim IsSupported As Boolean
        IsSupported = GeographicRegion.IsSupported("US")
        Debug.Print "GeographicRegion.IsSupported('US'): " & IsSupported
        If IsSupported Then
            Dim GeographicRegionUS As GeographicRegion
            Set GeographicRegionUS = Windows.Globalization.GeographicRegion.CreateGeographicRegion("US")
            Debug.Print "GeographicRegionUS.Code: " & GeographicRegionUS.Code
            Debug.Print "GeographicRegionUS.CodeTwoLetter: " & GeographicRegionUS.CodeTwoLetter
            Debug.Print "GeographicRegionUS.CodeThreeLetter: " & GeographicRegionUS.CodeThreeLetter
            Debug.Print "GeographicRegionUS.CodeThreeDigit: " & GeographicRegionUS.CodeThreeDigit
            Debug.Print "GeographicRegionUS.DisplayName: " & GeographicRegionUS.DisplayName
            Debug.Print "GeographicRegionUS.NativeName: " & GeographicRegionUS.NativeName
            Dim CurrenciesInUseListUS As ReadOnlyList_1 'ReadOnlyList_String
            Set CurrenciesInUseListUS = GeographicRegionUS.CurrenciesInUse
            If IsNotNothing(CurrenciesInUseListUS) Then
                Dim CurrenciesInUseListUSCount As Long
                CurrenciesInUseListUSCount = CurrenciesInUseListUS.Size
                If CurrenciesInUseListUSCount > 0& Then
                    Dim CurrenciesInUseListUSItem As Long
                    For CurrenciesInUseListUSItem = 0 To CurrenciesInUseListUSCount - 1
                        Debug.Print "CurrenciesInUse(" & CStr(CurrenciesInUseListUSItem) & "): " & CurrenciesInUseListUS.GetAt(CurrenciesInUseListUSItem)
                    Next
                End If
            End If
        End If
    End If
End Sub

Private Sub Command62_Click()
    Dim LanguageFontGroup As LanguageFontGroup
    Set LanguageFontGroup = Windows.Globalization.Fonts.LanguageFontGroup.CreateLanguageFontGroup("en")
    If IsNotNothing(LanguageFontGroup) Then
        Dim LanguageFont As LanguageFont
        Set LanguageFont = LanguageFontGroup.UITextFont
        If IsNotNothing(LanguageFont) Then
            Debug.Print "LanguageFontGroup.UITextFont.FontFamily: " & LanguageFont.FontFamily
            Debug.Print "LanguageFontGroup.UITextFont.FontWeight: " & LanguageFont.FontWeight.Weight
            Debug.Print "LanguageFontGroup.UITextFont.ScaleFactor: " & LanguageFont.ScaleFactor
            Dim strFontStretch As String
            Select Case LanguageFont.FontStretch
                Case FontStretch.FontStretch_Undefined
                    strFontStretch = "Undefined"
                Case FontStretch.FontStretch_UltraCondensed
                    strFontStretch = "UltraCondensed"
                Case FontStretch.FontStretch_ExtraCondensed
                    strFontStretch = "ExtraCondensed"
                Case FontStretch.FontStretch_Condensed
                    strFontStretch = "Condensed"
                Case FontStretch.FontStretch_SemiCondensed
                    strFontStretch = "SemiCondensed"
                Case FontStretch.FontStretch_Normal
                    strFontStretch = "Normal"
                Case FontStretch.FontStretch_SemiExpanded
                    strFontStretch = "SemiExpanded"
                Case FontStretch.FontStretch_Expanded
                    strFontStretch = "Expanded"
                Case FontStretch.FontStretch_ExtraExpanded
                    strFontStretch = "ExtraExpanded"
                Case FontStretch.FontStretch_UltraExpanded
                    strFontStretch = "UltraExpanded"
            End Select
            Debug.Print "LanguageFontGroup.UITextFont.FontStretch: " & strFontStretch
            Dim strFontStyle As String
            Select Case LanguageFont.FontStyle
                Case FontStyle.FontStyle_Normal
                    strFontStyle = "Normal"
                Case FontStyle.FontStyle_Oblique
                    strFontStyle = "Oblique"
                Case FontStyle.FontStyle_Italic
                    strFontStyle = "Italic"
            End Select
            Debug.Print "LanguageFontGroup.UITextFont.FontStyle: " & strFontStyle
        End If
    End If
End Sub

Private Sub Command63_Click()
    Dim KnownFolders As KnownFolders
    Set KnownFolders = Windows.Storage.KnownFolders
    If IsNotNothing(KnownFolders) Then
        Dim StorageFolder As StorageFolder
        Set StorageFolder = KnownFolders.DocumentsLibrary
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.DocumentsLibrary.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.HomeGroup
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.HomeGroup.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.MediaServerDevices
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.MediaServerDevices.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.MusicLibrary
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.MusicLibrary.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.PicturesLibrary
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.PicturesLibrary.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.RemovableDevices
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.RemovableDevices.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.VideosLibrary
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.VideosLibrary.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.Objects3D
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.Objects3D.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.AppCaptures
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.AppCaptures.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
        Set StorageFolder = KnownFolders.RecordedCalls
        If IsNotNothing(StorageFolder) Then
            Debug.Print "KnownFolders.RecordedCalls.Name: " & StorageFolder.Name
            Set StorageFolder = Nothing
        End If
    End If
End Sub

Private Sub Command64_Click()
    Dim UserDataPaths As UserDataPaths
    Set UserDataPaths = Windows.Storage.UserDataPaths.GetDefault
    If IsNotNothing(UserDataPaths) Then
        Debug.Print "UserDataPaths.Startup: " & UserDataPaths.Startup
        Debug.Print "UserDataPaths.StartMenu: " & UserDataPaths.StartMenu
        Debug.Print "UserDataPaths.Desktop: " & UserDataPaths.Desktop
        ' ...
    End If
End Sub

Private Sub Command65_Click()
    If Windows.Graphics.Capture.GraphicsCaptureSession.IsSupported Then
        Debug.Print "GraphicsCaptureSession.IsSupported: True"
        Dim GraphicsCapturePicker As GraphicsCapturePicker
        Set GraphicsCapturePicker = Windows.Graphics.Capture.GraphicsCapturePicker
        If IsNotNothing(GraphicsCapturePicker) Then
            GraphicsCapturePicker.ParentHwnd = Me.hwnd
            Dim GraphicsCaptureItem As GraphicsCaptureItem
            Set GraphicsCaptureItem = GraphicsCapturePicker.PickSingleItemAsync
            If IsNotNothing(GraphicsCaptureItem) Then
                Debug.Print "GraphicsCaptureItem.DisplayName: " & GraphicsCaptureItem.DisplayName
                Debug.Print "GraphicsCaptureItem.Size: " & GraphicsCaptureItem.Size.ToString
                Debug.Print "GraphicsCaptureItem.IsMinimized: " & GraphicsCaptureItem.IsMinimized
            End If
        End If
    Else
        Debug.Print "GraphicsCaptureSession.IsSupported: False"
    End If
End Sub

Private Sub Command66_Click()
'
End Sub

' ----==== Events ====----
Public Sub ToastActivatedEvent(ByVal sender As ToastNotification, _
                               ByVal args As ToastActivatedEventArgs)
    Debug.Print "ToastActivatedEvent:" & vbNewLine & _
                vbTab & "ToastNotification.Tag = " & sender.Tag & vbNewLine & _
                vbTab & "ToastNotification.ExpirationTime = " & sender.ExpirationTime.VbDate & vbNewLine & _
                vbTab & "ToastActivatedEventArgs.Arguments = " & args.Arguments & vbNewLine

    If args.UserInput.Size > 0 Then

        Debug.Print vbTab & "ToastActivatedEventArgs.UserInput.Size = " & args.UserInput.Size & vbNewLine & _
                    vbTab & "ToastActivatedEventArgs.UserInput.Test = " & args.UserInput.Test & vbNewLine

    End If
    
End Sub

Public Sub ToastDismissedEvent(ByVal sender As ToastNotification, _
                               ByVal args As ToastDismissedEventArgs)
    Dim ReasonStr As String
    Select Case args.Reason
        Case ToastDismissalReason.ToastDismissalReason_UserCanceled
            ReasonStr = "The user dismissed the toast."
        Case ToastDismissalReason.ToastDismissalReason_ApplicationHidden
            ReasonStr = "The app hide the toast using ToastNotifier.Hide."
        Case ToastDismissalReason.ToastDismissalReason_TimedOut
            ReasonStr = "The toast has timed out."
        Case Else
            ReasonStr = "Unknown reason."
    End Select
    Debug.Print "ToastDismissedEvent:" & vbNewLine & _
                vbTab & "ToastNotification.Tag = " & sender.Tag & vbNewLine & _
                vbTab & "ToastNotification.ExpirationTime = " & sender.ExpirationTime.VbDate & vbNewLine & _
                vbTab & "ToastDismissedEventArgs.Reason = " & ReasonStr & vbNewLine & vbNewLine
                         
End Sub

Public Sub ToastFailedEvent(ByVal sender As ToastNotification, _
                            ByVal args As ToastFailedEventArgs)
    Debug.Print "ToastFailedEvent:" & vbNewLine & _
                vbTab & "ToastNotification.Tag = " & sender.Tag & vbNewLine & _
                vbTab & "ToastNotification.ExpirationTime = " & sender.ExpirationTime.VbDate & vbNewLine & _
                vbTab & "ToastFailedEventArgs.ErrorCode = 0x" & Hex$(args.ErrorCode) & vbNewLine & vbNewLine
End Sub

Private Sub AsyncActionWithProgress_Double_Completed(ByVal asyncInfo As Long, ByVal asyncStat As Long)
    If IsNotNothing(PrepareTranscodeResult) Then
        Debug.Print "Transcoding Completed, PrepareTranscodeResult.FailureReason: " & CStr(PrepareTranscodeResult.FailureReason)
    End If
    Set PrepareTranscodeResult = Nothing
    Set AsyncActionWithProgress_Double = Nothing
    Command6.Enabled = True
    Command7.Enabled = True
End Sub

Private Sub AsyncActionWithProgress_Double_Progress(ByVal asyncInfo As Long, ByVal progressInfo As Double)
    Debug.Print "Transcoding Progress: " & CStr(CLng(progressInfo)) & "%"
End Sub

Private Sub ThreadPoolTimer_TimerDestroyed(ByVal pTimer As Long)
    If IsNotNothing(ThreadPoolTimer) Then
        If pTimer = ThreadPoolTimer.Ifc Then
            Debug.Print "ThreadPoolTimer TimerDestroyed"
        End If
    End If
End Sub

Private Sub ThreadPoolTimer_TimerElapsed(ByVal pTimer As Long)
    If IsNotNothing(ThreadPoolTimer) Then
        If pTimer = ThreadPoolTimer.Ifc Then
            Debug.Print "ThreadPoolTimer TimerElapsed"
        End If
    End If
End Sub

Private Sub UICommandInvokedHandler_CommandInvoked(ByVal command As UICommand)
    If IsNotNothing(command) Then
        Debug.Print "Return from UICommandInvokedHandler_CommandInvoked"
        Debug.Print "UICommand.Label = " & command.Label
        If VarType(command.Id) <> vbEmpty Then Debug.Print "UICommand.Id = " & command.Id
        Debug.Print
    End If
End Sub

Public Sub CoreWindowPopupShowingEvent(ByVal sender As Long, ByVal args As CoreWindowPopupShowingEventArgs)
    If IsNotNothing(args) Then
        Dim Size As New Size
        Size.Width = 200
        Size.Height = 100
        Call args.SetDesiredSize(Size)
    End If
End Sub

Public Sub DeviceSelectedEvent(ByVal sender As Long, ByVal args As DeviceSelectedEventArgs)
    If IsNotNothing(args) Then
        Debug.Print "DeviceSelectedEvent: " & args.SelectedDevice.Name
    End If
End Sub

Public Sub DeviceDisconnectButtonClickedEvent(ByVal sender As Long, ByVal args As DeviceDisconnectButtonClickedEventArgs)
    If IsNotNothing(args) Then
        Debug.Print "DeviceDisconnectButtonClickedEvent: " & args.Device.Name
    End If
End Sub

Public Sub DevicePickerDismissedEvent(ByVal sender As Long, ByVal args As Inspectable)
    If IsNotNothing(args) Then
        Debug.Print "DevicePickerDismissedEvent:"
    End If
End Sub

Public Sub CastingDeviceSelectedEvent(ByVal sender As Long, ByVal args As CastingDeviceSelectedEventArgs)
    If IsNotNothing(args) Then
        Debug.Print "CastingDeviceSelectedEvent: " & args.SelectedCastingDevice.FriendlyName
    End If
End Sub

Public Sub CastingDevicePickerDismissedEvent(ByVal sender As Long, ByVal args As Inspectable)
    If IsNotNothing(args) Then
        Debug.Print "CastingDevicePickerDismissedEvent:"
    End If
End Sub
