Attribute VB_Name = "Win32"
' Autor: F. Schüler (frank@activevb.de)
' Datum: 04/2023

Option Explicit

' ----==== Const ====----
Public Const S_OK As Long = &H0&
Private Const API_TRUE As Long = &H1&
Private Const CC_STDCALL As Long = &H4&
Private Const GdiPlusVersion As Long = &H1&
Private Const IUnknown_Release As Long = &H8&
Public Const E_NOINTERFACE As Long = &H80004002
Private Const CLSCTX_INPROC_SERVER As Long = &H1&

Public Const IID_IUnknown As String = "{00000000-0000-0000-c000-000000000046}"
Private Const IID_IPicture As String = "{7bf80981-bf32-101a-8bbb-00aa00300cab}"
Private Const IID_IClosable As String = "{30d5a829-7fa4-4026-83bb-d75bae4ea99e}"
Private Const IID_IAsyncInfo As String = "{00000036-0000-0000-c000-000000000046}"
Private Const IID_IInitializeWithWindow As String = "{3e68d4bd-7135-4d10-8018-9fb6d9f33fa1}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IUnknown
    IUnknown_QueryInterface = 0

    ' IInspectable
    IInspectable_GetIids = 3
    IInspectable_GetRuntimeClassName = 4
    IInspectable_GetTrustLevel = 5
    
    ' IAsyncInfo
    IAsyncInfo_GetStatus = 7
    IAsyncInfo_Cancel = 9
    IAsyncInfo_Close = 10

    ' IAsyncOperation
'    IAsyncOperation_PutCompleted = 6
'    IAsyncOperation_GetCompleted = 7
    IAsyncOperation_GetResults = 8
    
    ' IClosable
    IClosable_Close = 6

    ' IInitializeWithWindow
    IInitializeWithWindow_Initialize = 3

End Enum

Private Enum AsyncStatus
    Started = 0
    Completed = 1
    Canceled = 2
    Error = 3
End Enum

Public Enum TrustLevel
    BaseTrust = 0
    PartialTrust = 1
    FullTrust = 2
End Enum

Private Enum GdipStatus
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

' ----==== Enums ShCore ====----
Public Enum BSOS_OPTIONS
    BSOS_DEFAULT = 0
    BSOS_PREFERDESTINATIONSTREAM = 1
End Enum

' ----==== Types ====----
Public Type GUID
    data1 As Long
    data2 As Long
    data3 As Long
    data4 As Long
End Type
                
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

' ----==== Ole32.dll Declarations ====----
Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
                         ByVal pstring As Long, _
                         ByRef pCLSID As GUID) As Long
                         
Private Declare Function StringFromGUID Lib "ole32.dll" _
                         Alias "StringFromGUID2" ( _
                         ByRef rguid As GUID, _
                         ByVal lpsz As Long, _
                         ByVal cchMax As Long) As Long
                         
Private Declare Function CoCreateInstance Lib "ole32.dll" ( _
                         ByVal rclsid As Long, _
                         ByVal pUnkOuter As Long, _
                         ByVal dwClsContext As Long, _
                         ByVal riid As Long, _
                         ByRef ppv As Long) As Long
                         
Public Declare Sub CoTaskMemFree Lib "ole32.dll" ( _
                   ByVal hMem As Long)
                         
' ----==== Oleaut32.dll Declarations ====----
Private Declare Function DispCallFunc Lib "Oleaut32.dll" ( _
                         ByVal pvInstance As Long, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long

Private Declare Function OleCreatePictureIndirect Lib "Oleaut32.dll" ( _
                         ByRef lpPictDesc As PictDesc, _
                         ByRef riid As GUID, _
                         ByVal fOwn As Long, _
                         ByRef lplpvObj As Object) As Long

' ----==== Kernel32.dll Declarations ====----
Public Declare Sub CopyMemory Lib "kernel32.dll" _
                   Alias "RtlMoveMemory" ( _
                   ByRef pTo As Any, _
                   ByRef uFrom As Any, _
                   ByVal lSize As Long)

Private Declare Function lstrlenW Lib "kernel32.dll" ( _
                         ByVal lpString As Long) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" ( _
                         ByRef lpFileTime As FILETIME, _
                         ByRef lpSystemTime As SYSTEMTIME) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" ( _
                         ByRef lpSystemTime As SYSTEMTIME, _
                         ByRef lpFileTime As FILETIME) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" ( _
                         ByRef lpFileTime As FILETIME, _
                         ByRef lpLocalFileTime As FILETIME) As Long

Private Declare Function LocalFileTimeToFileTime Lib "kernel32.dll" ( _
                         ByRef lpLocalFileTime As FILETIME, _
                         ByRef lpFileTime As FILETIME) As Long

' ----==== Combase.dll Declarations ====----
Private Declare Function WindowsCreateString Lib "Combase.dll" ( _
                         ByVal sourceString As Long, _
                         ByVal lenght As Long, _
                         ByRef hString As Long) As Long

Private Declare Function WindowsDeleteString Lib "Combase.dll" ( _
                         ByVal hString As Long) As Long

Private Declare Function WindowsGetStringRawBuffer Lib "Combase.dll" ( _
                         ByVal hString As Long, _
                         ByRef Length As Long) As Long

Private Declare Function RoGetActivationFactory Lib "Combase.dll" ( _
                         ByVal activatableClassId As Long, _
                         ByRef riid As GUID, _
                         ByRef factory As Long) As Long

Private Declare Function RoActivateInstance Lib "Combase.dll" ( _
                         ByVal activatableClassId As Long, _
                         ByRef instance As Long) As Long

' ----==== Shlwapi.dll Declarations ====----
Public Declare Function SHCreateMemStream Lib "Shlwapi.dll" ( _
                        ByVal pInit As Long, _
                        ByVal cbInit As Long) As Long

' ----==== Gdiplus.dll Declarations ====----
Private Declare Function GdiplusShutdown Lib "GdiPlus.dll" ( _
                         ByVal token As Long) As GdipStatus
                         
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" ( _
                         ByRef token As Long, _
                         ByRef lpInput As GDIPlusStartupInput, _
                         ByRef lpOutput As GdiplusStartupOutput) As GdipStatus

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus.dll" ( _
                         ByVal bitmap As Long, _
                         ByRef hbmReturn As Long, _
                         ByVal background As Long) As GdipStatus

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" ( _
                         ByVal hbm As Long, _
                         ByVal hpal As Long, _
                         ByRef bitmap As Long) As GdipStatus

Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" ( _
                         ByVal image As Long) As GdipStatus

Private Declare Function GdipCreateBitmapFromStream Lib "GdiPlus.dll" ( _
                         ByVal mStream As Long, _
                         ByRef mBitmap As Long) As GdipStatus
                         
Private Declare Function GdipSaveImageToFile Lib "GdiPlus.dll" ( _
                         ByVal image As Long, _
                         ByVal fileName As Long, _
                         ByRef clsidEncoder As GUID, _
                         ByRef encoderParams As Any) As GdipStatus

' ----==== ShCore.dll Declarations ====----
Public Declare Function CreateRandomAccessStreamOverStream Lib "ShCore.dll" ( _
                        ByVal pStream As Long, _
                        ByVal Options As BSOS_OPTIONS, _
                        ByRef riid As GUID, _
                        ByRef ppv As Long) As Long

Public Declare Function CreateStreamOverRandomAccessStream Lib "ShCore.dll" ( _
                        ByVal pRandomAccessStream As Long, _
                        ByRef riid As GUID, _
                        ByRef ppv As Long) As Long
                         
Public Declare Function CreateRandomAccessStreamOnFile Lib "ShCore.dll" ( _
                        ByVal filePath As Long, _
                        ByVal accessMode As FileAccessMode, _
                        ByRef riid As GUID, _
                        ByRef ppv As Long) As Long

' ----==== WinRT Functions ====----
Public Function GetActivationFactory(ByVal className As String, _
                                     ByVal iid As String, _
                                     ByRef pFactory As Long) As Boolean
    Dim hString As Long
    hString = CreateWindowsString(className)
    If RoGetActivationFactory(hString, Str2Guid(iid), pFactory) = S_OK Then
        GetActivationFactory = True
    End If
    Call DeleteWindowsString(hString)
End Function

Public Function GetActivateInstance(ByVal className As String, _
                                    ByVal iid As String, _
                                    ByRef pInstance As Long) As Boolean
    Dim hString As Long
    hString = CreateWindowsString(className)
    Dim pIInspectable As Long
    If RoActivateInstance(hString, pIInspectable) = S_OK Then
        If pIInspectable <> 0& Then
            If QueryIfc(pIInspectable, iid, pInstance) Then
                If pInstance <> 0& Then GetActivateInstance = True
            End If
            Call ReleaseIfc(pIInspectable)
        End If
    End If
    Call DeleteWindowsString(hString)
End Function

Public Function CreateWindowsString(ByVal value As String) As Long
    Call WindowsCreateString(StrPtr(value), Len(value), CreateWindowsString)
End Function

Public Function GetWindowsString(ByVal hString As Long) As String
    GetWindowsString = Ptr2Str(WindowsGetStringRawBuffer(hString, 0&))
    'Call DeleteWindowsString(hString)
End Function

Public Function DeleteWindowsString(ByRef hString As Long) As Boolean
    If WindowsDeleteString(hString) = S_OK Then
        hString = 0&
        DeleteWindowsString = True
    End If
End Function

Public Function InitWithWindow(ByVal pInterface As Long, _
                               ByVal ownerHwnd As Long) As Boolean
    If pInterface <> 0& And ownerHwnd <> 0& Then
        Dim pIInitializeWithWindow As Long
        If QueryIfc(pInterface, _
                    IID_IInitializeWithWindow, _
                    pIInitializeWithWindow) Then
            If InvokeIfc(pIInitializeWithWindow, _
                         IInitializeWithWindow_Initialize, _
                         ownerHwnd) = S_OK Then
                InitWithWindow = True
            End If
            Call ReleaseIfc(pIInitializeWithWindow)
        End If
    End If
End Function

Public Function Await(ByRef pInterface As Long, _
                      Optional ByVal WithResult As Boolean = True) As Boolean
    If pInterface <> 0& Then
        Dim pIAsyncInfo As Long
        If QueryIfc(pInterface, IID_IAsyncInfo, pIAsyncInfo) Then
            Dim eStatus As AsyncStatus
            Do
                If InvokeIfc(pIAsyncInfo, IAsyncInfo_GetStatus, VarPtr(eStatus)) = S_OK Then
                    DoEvents
                End If
            Loop While eStatus = Started
            If eStatus = Completed Then
                If WithResult Then
                    Dim pRet As Long
                    If InvokeIfc(pInterface, IAsyncOperation_GetResults, VarPtr(pRet)) = S_OK Then
                        If pRet <> 0& Then
                            Call ReleaseIfc(pInterface)
                            pInterface = pRet
                            Await = True
                        Else
                            Call ReleaseIfc(pInterface)
                        End If
                    Else
                        Call ReleaseIfc(pInterface)
                    End If
                Else
                    If InvokeIfc(pInterface, IAsyncOperation_GetResults) = S_OK Then
                        Await = True
                        Call ReleaseIfc(pInterface)
                    Else
                        Call ReleaseIfc(pInterface)
                    End If
                End If
            Else
                Call ReleaseIfc(pInterface)
            End If
            Call InvokeIfc(pIAsyncInfo, IAsyncInfo_Close)
            Call ReleaseIfc(pIAsyncInfo)
        End If
    End If
End Function

Public Sub GetInspectableInfo(ByVal pInspectableInterface As Long, _
                              Optional ByVal interfaceName As String = vbNullString)
    If pInspectableInterface <> 0& Then
        If interfaceName <> vbNullString Then
            Debug.Print "Inspectable Info for: " & interfaceName
        End If
        Debug.Print "RuntimeClassName = " & GetRuntimeClassName(pInspectableInterface)
        Dim eTrustLevel As TrustLevel
        eTrustLevel = GetTrustLevel(pInspectableInterface)
        Select Case eTrustLevel
            Case TrustLevel.BaseTrust
                Debug.Print "TrustLevel = BaseTrust"
            Case TrustLevel.FullTrust
                Debug.Print "TrustLevel = FullTrust"
            Case TrustLevel.PartialTrust
                Debug.Print "TrustLevel = PartialTrust"
            Case Else
                Debug.Print "TrustLevel = UnknownTrust"
        End Select
        Dim tIIDs() As GUID
        Dim lngIIDsCount As Long
        lngIIDsCount = GetIids(pInspectableInterface, tIIDs)
        Debug.Print "AviableIIDsCount = " & CStr(lngIIDsCount)
        If lngIIDsCount > 0 Then
            Dim lngIIDsItem As Long
            For lngIIDsItem = 0 To lngIIDsCount - 1
                Debug.Print "    AviableIID " & CStr(lngIIDsItem + 1) & _
                            " = " & LCase$(Guid2Str(tIIDs(lngIIDsItem)))
            Next
        End If
        Debug.Print String(60, "-")
    End If
End Sub

Private Function GetIids(ByVal pInterface As Long, _
                         ByRef iids() As GUID) As Long
    If pInterface <> 0& Then
        Dim count As Long
        Dim pIIDs As Long
        If InvokeIfc(pInterface, IInspectable_GetIids, VarPtr(count), VarPtr(pIIDs)) = S_OK Then
            If count > 0& Then
                Dim bytes As Long
                Dim Item As Long
                ReDim iids(count - 1)
                For Item = 0 To count - 1
                    bytes = Len(iids(Item))
                    Call CopyMemory(iids(Item), ByVal pIIDs + (bytes * Item), bytes)
                Next
                GetIids = count
            End If
        End If
    End If
End Function

Private Function GetRuntimeClassName(ByVal pInterface As Long) As String
    If pInterface <> 0& Then
        Dim hString As Long
        If InvokeIfc(pInterface, IInspectable_GetRuntimeClassName, VarPtr(hString)) = S_OK Then
            If hString <> 0& Then
                GetRuntimeClassName = GetWindowsString(hString)
            End If
        End If
    End If
End Function

Private Function GetTrustLevel(ByVal pInterface As Long) As TrustLevel
    If pInterface <> 0& Then
        Dim value As Long
        If InvokeIfc(pInterface, IInspectable_GetTrustLevel, VarPtr(value)) = S_OK Then
            GetTrustLevel = value
        End If
    End If
End Function

Public Sub DisposeIfc(ByRef pInterface As Long)
    If pInterface <> 0& Then
        Dim pIClosable As Long
        If InvokeIfc(pInterface, IUnknown_QueryInterface, _
                     VarPtr(Str2Guid(IID_IClosable)), _
                     VarPtr(pIClosable)) = S_OK Then
            Call InvokeIfc(pIClosable, IClosable_Close)
            Call ReleaseIfc(pIClosable)
            Call ReleaseIfc(pInterface)
            pInterface = 0&
        End If
    End If
End Sub

' ----==== Interface Functions ====----
Public Function CreateIfc(ByVal clsid As String, _
                          ByVal iid As String, _
                          ByRef pInterface As Long) As Boolean
    Dim Ret As Boolean
    If CoCreateInstance(VarPtr(Str2Guid(clsid)), _
                        0&, _
                        CLSCTX_INPROC_SERVER, _
                        VarPtr(Str2Guid(iid)), _
                        pInterface) = S_OK Then
        Ret = True
    End If
    CreateIfc = Ret
End Function

Public Function QueryIfc(ByVal pInterface As Long, _
                         ByVal iid As String, _
                         ByRef ppInterface As Long) As Boolean
    If pInterface <> 0& Then
        If InvokeIfc(pInterface, IUnknown_QueryInterface, _
                     VarPtr(Str2Guid(iid)), VarPtr(ppInterface)) = S_OK Then
            If ppInterface <> 0& Then QueryIfc = True
        End If
    End If
End Function

Public Sub ReleaseIfc(ByRef pInterface As Long)
    If pInterface <> 0& Then
        Dim Ret As Long
        If DispCallFunc(pInterface, IUnknown_Release, _
                        CC_STDCALL, vbLong, 0&, 0&, _
                        0&, Ret) = S_OK Then
            pInterface = 0&
        End If
    End If
End Sub

Private Function InvokeIfc(ByVal pInterface As Long, _
                           ByVal vtb As vtb_Interfaces, _
                           ParamArray var()) As Variant
    If pInterface <> 0& Then
        InvokeIfc = OleInvoke(pInterface, vtb, var)
    End If
End Function

Public Function OleInvoke(ByVal pInterface As Long, _
                          ByVal lngCmd As Long, _
                          ParamArray aParam()) As Variant
    Dim Ret As Variant
    If pInterface <> 0& Then
        Dim varParam As Variant
        Dim lngCount As Long
        Dim lngItem As Long
        Dim olePtr(10) As Long
        Dim oleTyp(10) As Integer
        If UBound(aParam) >= 0& Then
            varParam = aParam
            If IsArray(varParam) Then varParam = varParam(0)
            lngCount = UBound(varParam)
            For lngItem = 0& To lngCount
                oleTyp(lngItem) = VarType(varParam(lngItem))
                olePtr(lngItem) = VarPtr(varParam(lngItem))
            Next
        End If
        If DispCallFunc(pInterface, lngCmd * 4, CC_STDCALL, _
                        vbLong, lngItem, oleTyp(0), _
                        olePtr(0), Ret) <> S_OK Then
            Debug.Print "Fehler beim Aufrufen der Interface-Funktion!"
        End If
'        If Ret <> S_OK Then
'            Debug.Print "0x" & Hex$(Ret)
'        End If
    End If
    OleInvoke = Ret
End Function

' ----==== Helper Functions ====----
Public Function ProcPtr(ByVal ptr As Long) As Long
    ProcPtr = ptr
End Function

Public Function Str2Guid(ByVal strGUID As String) As GUID
    If Len(strGUID) = 38 Then
        Call CLSIDFromString(StrPtr(strGUID), Str2Guid)
    End If
End Function

Public Function Guid2Str(ByRef tGUID As GUID) As String
    Dim lngRet As Long
    Dim strGUID As String
    strGUID = String(40, 0)
    lngRet = StringFromGUID(tGUID, StrPtr(strGUID), Len(strGUID))
    If lngRet > 0& Then
        Guid2Str = Left$(strGUID, lngRet - 1)
    Else
        Guid2Str = vbNullString
    End If
End Function

Public Function Ptr2Str(ByVal lpStr As Long) As String
    If lpStr <> 0& Then
        Dim lngLen As Long
        Dim bytBuffer() As Byte
        lngLen = lstrlenW(lpStr) * 2
        If lngLen > 0 Then
            ReDim bytBuffer(lngLen - 1)
            Call CopyMemory(bytBuffer(0), ByVal lpStr, lngLen)
            Call CoTaskMemFree(lpStr)
            Ptr2Str = bytBuffer
        End If
    End If
End Function

Public Function IsNothing(ByVal obj As Object) As Boolean
    If obj Is Nothing Then IsNothing = True
End Function

Public Function IsNotNothing(ByVal obj As Object) As Boolean
    If Not obj Is Nothing Then IsNotNothing = True
End Function

Public Function DateTime2VBDate(ByVal value As Currency, _
                                Optional ByVal ToLocalFileTime As Boolean = True) As Date
    If value <> CCur(0) Then
        Dim tFILETIME As FILETIME
        Dim tSYSTEMTIME As SYSTEMTIME
        Call CopyMemory(tFILETIME, value, 8)
        If ToLocalFileTime Then
            Call FileTimeToLocalFileTime(tFILETIME, tFILETIME)
        End If
        If CBool(FileTimeToSystemTime(tFILETIME, tSYSTEMTIME)) Then
            With tSYSTEMTIME
                DateTime2VBDate = DateSerial(.wYear, .wMonth, .wDay) + _
                                  TimeSerial(.wHour, .wMinute, .wSecond)
            End With
        End If
    End If
End Function

Public Function VBDate2DateTime(ByVal value As Date) As Currency
    Dim tFILETIME As FILETIME
    Dim tSYSTEMTIME As SYSTEMTIME
    With tSYSTEMTIME
        .wYear = Year(value)
        .wMonth = Month(value)
        .wDay = Day(value)
        .wHour = Hour(value)
        .wMinute = Minute(value)
        .wSecond = Second(value)
    End With
    If CBool(SystemTimeToFileTime(tSYSTEMTIME, tFILETIME)) Then
        If CBool(LocalFileTimeToFileTime(tFILETIME, tFILETIME)) Then
            Call CopyMemory(VBDate2DateTime, tFILETIME, 8)
        End If
    End If
End Function

Public Function TimeSpan2VBDate(ByVal value As Currency) As Date
    If value <> CCur(0) Then
        Dim tFILETIME As FILETIME
        Dim tSYSTEMTIME As SYSTEMTIME
        Call CopyMemory(tFILETIME, value, 8)
        If CBool(FileTimeToSystemTime(tFILETIME, tSYSTEMTIME)) Then
            With tSYSTEMTIME
                TimeSpan2VBDate = TimeSerial(.wHour, .wMinute, .wSecond)
            End With
        End If
    End If
End Function

Public Function GetPictureFromRandomAccessStreamReference(ByVal RandomAccessStreamRef As RandomAccessStreamReference) As StdPicture
    Dim Ret As StdPicture
    If IsNotNothing(RandomAccessStreamRef) Then
        Dim RandomAccessStreamWithContentType As RandomAccessStreamWithContentType
        Set RandomAccessStreamWithContentType = RandomAccessStreamRef.OpenReadAsync
        If IsNotNothing(RandomAccessStreamWithContentType) Then
            Dim pIStream As Long
            pIStream = RandomAccessStreamWithContentType.ToIStream
            If pIStream <> 0& Then
                Set Ret = GetPictureFromIStream(pIStream)
                Call ReleaseIfc(pIStream)
            End If
            Set RandomAccessStreamWithContentType = Nothing
        End If
    End If
    Set GetPictureFromRandomAccessStreamReference = Ret
End Function

Public Function GetPictureFromIStream(ByVal pIStream As Long) As StdPicture
    Dim Ret As StdPicture
    If pIStream <> 0& Then
        Dim GdipToken As Long
        Dim tGdipStartupInput As GDIPlusStartupInput
        Dim tGdipStartupOutput As GdiplusStartupOutput
        tGdipStartupInput.GdiPlusVersion = GdiPlusVersion
        If GdiplusStartup(GdipToken, _
                          tGdipStartupInput, _
                          tGdipStartupOutput) = OK Then
            Dim pBitmap As Long
            If GdipCreateBitmapFromStream(pIStream, _
                                          pBitmap) = OK Then
                Dim hBitmap As Long
                If GdipCreateHBITMAPFromBitmap(pBitmap, _
                                               hBitmap, 0&) = OK Then
                    Dim tPictDesc As PictDesc
                    With tPictDesc
                        .cbSizeofStruct = Len(tPictDesc)
                        .picType = vbPicTypeBitmap
                        .hImage = hBitmap
                    End With
                    Dim oIPicture As IPicture
                    If OleCreatePictureIndirect(tPictDesc, _
                                                Str2Guid(IID_IPicture), _
                                                API_TRUE, _
                                                oIPicture) = S_OK Then
                        If IsNotNothing(oIPicture) Then
                            Set Ret = oIPicture
                            Set oIPicture = Nothing
                        End If
                    End If
                End If
                If GdipDisposeImage(pBitmap) = OK Then pBitmap = 0&
            End If
            If GdiplusShutdown(GdipToken) = OK Then GdipToken = 0&
        End If
    End If
    Set GetPictureFromIStream = Ret
End Function

Public Function GetXMLFormated(ByVal Xml As String) As String
    Dim strXML As String
    Dim arrXML() As String
    Dim lngItem As Long
    Dim lngItems As Long
    Dim lngTab As Long
    arrXML = Split(Replace$(Xml, "><", ">" & vbNewLine & "<"), vbNewLine)
    lngItems = UBound(arrXML)
    For lngItem = 0 To lngItems
        If Len(arrXML(lngItem)) >= 2 Then
            If lngTab > 0 And Left$(arrXML(lngItem), 2) = "</" Then lngTab = lngTab - 1
            strXML = strXML & String$(lngTab, vbTab) & arrXML(lngItem) & vbNewLine
            If Right$(arrXML(lngItem), 2) <> "/>" And _
               InStr(1, arrXML(lngItem), "</") = 0 Then lngTab = lngTab + 1
        End If
    Next
    GetXMLFormated = strXML
End Function

Public Function New_ReadOnlyList_1(ByVal Of As OfType_xxx, _
                                   ByVal pIVectorView As Long) As ReadOnlyList_1
    Set New_ReadOnlyList_1 = New ReadOnlyList_1
    New_ReadOnlyList_1.Ifc = pIVectorView
    New_ReadOnlyList_1.Of = Of
End Function

Public Function New_ReadOnlyList_2(ByVal Of As OfType_xxx_yyy, _
                                   ByVal pIMapView As Long) As ReadOnlyList_2
    Set New_ReadOnlyList_2 = New ReadOnlyList_2
    New_ReadOnlyList_2.Ifc = pIMapView
    New_ReadOnlyList_2.Of = Of
End Function
