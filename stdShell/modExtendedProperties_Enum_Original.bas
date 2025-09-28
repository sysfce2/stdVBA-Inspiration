Attribute VB_Name = "modExtendedProperties_Enum"
Private Enum ExtProps_PropGroup
    Advanced
    Audio
    Calendar
    Camera
    Contact
    Content
    Description
    FileSystem
    General
    GPS
    Image
    Media
    MediaAdvanced
    message
    Music
    Origin
    PhotoAdvanced
    RecordedTV
    Video
End Enum

Private Enum ExtProps_Video
    Compression
    Director
    EncodingBitrate
    FourCC
    FrameHeight
    FrameRate
    FrameWidth
    HorizontalAspectRatio
    IsSpherical
    IsStereo
    Orientation
    SampleSize
    StreamName
    StreamNumber
    TotalBitrate
    TranscodedForSync
    VerticalAspectRatio
End Enum

Private Enum ExtProps_Photo
    Aperture
    ApertureDenominator
    ApertureNumerator
    Brightness
    BrightnessDenominator
    BrightnessNumerator
    CameraManufacturer
    CameraModel
    CameraSerialNumber
    Contrast
    ContrastText
    DateTaken
    DigitalZoom
    DigitalZoomDenominator
    DigitalZoomNumerator
    [_Event]
    EXIFVersion
    ExposureBias
    ExposureBiasDenominator
    ExposureBiasNumerator
    ExposureIndex
    ExposureIndexDenominator
    ExposureIndexNumerator
    ExposureProgram
    ExposureProgramText
    ExposureTime
    ExposureTimeDenominator
    ExposureTimeNumerator
    Flash
    FlashEnergy
    FlashEnergyDenominator
    FlashEnergyNumerator
    FlashManufacturer
    FlashModel
    FlashText
    FNumber
    FNumberDenominator
    FNumberNumerator
    FocalLength
    FocalLengthDenominator
    FocalLengthInFilm
    FocalLengthNumerator
    FocalPlaneXResolution
    FocalPlaneXResolutionDenominator
    FocalPlaneXResolutionNumerator
    FocalPlaneYResolution
    FocalPlaneYResolutionDenominator
    FocalPlaneYResolutionNumerator
    GainControl
    GainControlDenominator
    GainControlNumerator
    GainControlText
    ISOSpeed
    LensManufacturer
    LensModel
    LightSource
    MakerNote
    MakerNoteOffset
    MaxAperture
    MaxApertureDenominator
    MaxApertureNumerator
    MeteringMode
    MeteringModeText
    Orientation
    OrientationText
    PeopleNames
    PhotometricInterpretation
    PhotometricInterpretationText
    ProgramMode
    ProgramModeText
    RelatedSoundFile
    Saturation
    SaturationText
    Sharpness
    SharpnessText
    ShutterSpeed
    ShutterSpeedDenominator
    ShutterSpeedNumerator
    SubjectDistance
    SubjectDistanceDenominator
    SubjectDistanceNumerator
    TagViewAggregate
    TranscodedForSync
    WhiteBalance
    WhiteBalanceText
End Enum
Private Enum ExtProps_Music
    AlbumArtist
    AlbumArtistSortOverride
    AlbumID
    AlbumTitle
    AlbumTitleSortOverride
    Artist
    ArtistSortOverride
    BeatsPerMinute
    Composer
    ComposerSortOverride
    Conductor
    ContentGroupDescription
    DiscNumber
    DisplayArtist
    Genre
    InitialKey
    IsCompilation
    Lyrics
    Mood
    PartOfSet
    period
    SynchronizedLyrics
    TrackNumber
End Enum
Private Enum ExtProps_Message
    AttachmentContents
    AttachmentNames
    BccAddress
    BccName
    CcAddress
    CcName
    ConversationID
    ConversationIndex
    DateReceived
    DateSent
    Flags
    FromAddress
    FromName
    HasAttachments
    IsFwdOrReply
    MessageClass
    Participants
    ProofInProgress
    SenderAddress
    SenderName
    Store
    ToAddress
    ToDoFlags
    ToDoTitle
    ToName
End Enum
Private Enum ExtProps_Media
    AuthorUrl
    AverageLevel
    ClassPrimaryID
    ClassSecondaryID
    CollectionGroupID
    CollectionID
    ContentDistributor
    ContentID
    CreatorApplication
    CreatorApplicationVersion
    DateEncoded
    DateReleased
    DlnaProfileID
    duration
    DVDID
    EncodedBy
    EncodingSettings
    EpisodeNumber
    FrameCount
    MCDI
    MetadataContentProvider
    Producer
    PromotionUrl
    ProtectionType
    ProviderRating
    ProviderStyle
    Publisher
    SeasonNumber
    SeriesName
    SubscriptionContentId
    SubTitle
    ThumbnailLargePath
    ThumbnailLargeUri
    ThumbnailSmallPath
    ThumbnailSmallUri
    UniqueFileIdentifier
    UserNoAutoInfo
    UserWebUrl
    Writer
    Year
End Enum
Private Enum ExtProps_Image
    BitDepth
    ColorSpace
    CompressedBitsPerPixel
    CompressedBitsPerPixelDenominator
    CompressedBitsPerPixelNumerator
    Compression
    CompressionText
    Dimensions
    HorizontalResolution
    HorizontalSize
    ImageID
    ResolutionUnit
    VerticalResolution
    VerticalSize
End Enum
Private Enum ExtProps_Document
    ByteCount
    CharacterCount
    ClientID
    Contributor
    DateCreated
    DatePrinted
    DateSaved
    Division
    DocumentID
    HiddenSlideCount
    LastAuthor
    LineCount
    Manager
    MultimediaClipCount
    NoteCount
    PageCount
    ParagraphCount
    PresentationFormat
    RevisionNumber
    Security
    SlideCount
    TEMPLATE
    TotalEditingTime
    Version
    WordCount
End Enum
Private Enum ExtProps_Core
    AcquisitionID
    ApplicationDefinedProperties
    ApplicationName
    AppZoneIdentifier
    Author
    CachedFileUpdaterContentIdForConflictResolution
    CachedFileUpdaterContentIdForStream
    Capacity
    Category
    Comment
    Company
    ComputerName
    ContainedItems
    ContentStatus
    ContentType
    Copyright
    CreatorAppId
    CreatorOpenWithUIOptions
    DataObjectFormat
    DateAccessed
    DateAcquired
    DateArchived
    DateCompleted
    DateCreated
    DateImported
    DateModified
    DefaultSaveLocationDisplay
    DueDate
    EndDate
    ExpandoProperties
    FileAllocationSize
    FileAttributes
    FileCount
    FileDescription
    FileExtension
    FileFRN
    Filename
    FileOfflineAvailabilityStatus
    FileOwner
    FilePlaceholderStatus
    FileVersion
    FindData
    FlagColor
    FlagColorText
    FlagStatus
    FlagStatusText
    FolderKind
    FolderNameDisplay
    FreeSpace
    FullText
    HighKeywords
    ImageParsingName
    Importance
    ImportanceText
    IsAttachment
    IsDefaultNonOwnerSaveLocation
    IsDefaultSaveLocation
    IsDeleted
    IsEncrypted
    IsFlagged
    IsFlaggedComplete
    IsIncomplete
    IsLocationSupported
    IsPinnedToNameSpaceTree
    IsRead
    IsSearchOnlyItem
    IsSendToTarget
    IsShared
    ItemAuthors
    ItemClassType
    ItemDate
    ItemFolderNameDisplay
    ItemFolderPathDisplay
    ItemFolderPathDisplayNarrow
    ItemName
    ItemNameDisplay
    ItemNameDisplayWithoutExtension
    ItemNamePrefix
    ItemNameSortOverride
    ItemParticipants
    ItemPathDisplay
    ItemPathDisplayNarrow
    ItemSubType
    ItemType
    ItemTypeText
    ItemUrl
    KEYWORDS
    Kind
    KindText
    language
    LastSyncError
    LastWriterPackageFamilyName
    LowKeywords
    MediumKeywords
    MileageInformation
    MIMEType
    [_Null]
    OfflineAvailability
    OfflineStatus
    OriginalFileName
    OwnerSID
    ParentalRating
    ParentalRatingReason
    ParentalRatingsOrganization
    ParsingBindContext
    ParsingName
    ParsingPath
    PerceivedType
    PercentFull
    Priority
    PriorityText
    Project
    ProviderItemID
    Rating
    RatingText
    RemoteConflictingFile
    Sensitivity
    SensitivityText
    SFGAOFlags
    SharedWith
    ShareUserRating
    SharingStatus
    shell
    SimpleRating
    Size
    SoftwareUsed
    SourceItem
    SourcePackageFamilyName
    StartDate
    Status
    StorageProviderCallerVersionInformation
    StorageProviderError
    StorageProviderFileChecksum
    StorageProviderFileIdentifier
    StorageProviderFileRemoteUri
    StorageProviderFileVersion
    StorageProviderFileVersionWaterline
    StorageProviderId
    StorageProviderShareStatuses
    StorageProviderSharingStatus
    StorageProviderStatus
    Subject
    SyncTransferStatus
    Thumbnail
    ThumbnailCacheId
    ThumbnailStream
    Title
    TitleSortOverride
    TotalFileSize
    Trademarks
    TransferOrder
    TransferPosition
    TransferSize
    VolumeId
    ZoneIdentifier
End Enum
Const MakeEnum  As Boolean = False
Sub BuildEnum()
    Dim List As Variant
    Dim ListType As String
    Dim IndexNo As Long
    Dim Counter As Long
    Dim Item As String
    Dim CodeStrong As String
    List = Application.Transpose(Application.Transpose(Application.Transpose(Selection.Value)))
    If UBound(Split(List(1), ".")) = 2 Then
    
        ListType = Split(List(1), ".")(1)
        IndexNo = 2
    Else
        ListType = "Core"
        IndexNo = 1
    End If
    'Stop
    If MakeEnum Then
        codestring = "Private Enum ExtProps_" & ListType & vbNewLine
        For Counter = LBound(List) To UBound(List)
            Item = Split(List(Counter), ".")(IndexNo)
            codestring = codestring & vbTab & Item & vbNewLine
        Next
        Debug.Print codestring & "End Enum"
    Else
        codestring = "PropertiesArray_" & ListType & " = Array("
        For Counter = LBound(List) To UBound(List)
            Item = Split(List(Counter), ".")(IndexNo)
            codestring = codestring & Chr(34) & Item & Chr(34) & ", "
        Next
        Debug.Print Left(codestring, Len(codestring) - 2) & ")"
    End If
End Sub


Function GetProps() As Variant

PropertiesArray_Core = Array("AcquisitionID", "ApplicationDefinedProperties", "ApplicationName", "AppZoneIdentifier", "Author", "CachedFileUpdaterContentIdForConflictResolution", "CachedFileUpdaterContentIdForStream", "Capacity", "Category", "Comment", "Company", "ComputerName", "ContainedItems", "ContentStatus", "ContentType", "Copyright", "CreatorAppId", "CreatorOpenWithUIOptions", "DataObjectFormat", "DateAccessed", "DateAcquired", "DateArchived", "DateCompleted", "DateCreated", "DateImported", "DateModified", "DefaultSaveLocationDisplay", "DueDate", "EndDate", "ExpandoProperties", "FileAllocationSize", "FileAttributes", "FileCount", "FileDescription", "FileExtension", "FileFRN", "FileName", "FileOfflineAvailabilityStatus", "FileOwner", "FilePlaceholderStatus", "FileVersion", "FindData", "FlagColor", "FlagColorText", "FlagStatus", "FlagStatusText", "FolderKind", "FolderNameDisplay", "FreeSpace", "FullText", "HighKeywords", "ImageParsingName", "Importance", "ImportanceText", _
    "IsAttachment", "IsDefaultNonOwnerSaveLocation", "IsDefaultSaveLocation", "IsDeleted", "IsEncrypted", "IsFlagged", "IsFlaggedComplete", "IsIncomplete", "IsLocationSupported", "IsPinnedToNameSpaceTree", "IsRead", "IsSearchOnlyItem", "IsSendToTarget", "IsShared", "ItemAuthors", "ItemClassType", "ItemDate", "ItemFolderNameDisplay", "ItemFolderPathDisplay", "ItemFolderPathDisplayNarrow", "ItemName", "ItemNameDisplay", "ItemNameDisplayWithoutExtension", "ItemNamePrefix", "ItemNameSortOverride", "ItemParticipants", "ItemPathDisplay", "ItemPathDisplayNarrow", "ItemSubType", "ItemType", "ItemTypeText", "ItemUrl", "Keywords", "Kind", "KindText", "Language", "LastSyncError", "LastWriterPackageFamilyName", "LowKeywords", "MediumKeywords", "MileageInformation", "MIMEType", "Null", "OfflineAvailability", "OfflineStatus", "OriginalFileName", "OwnerSID", "ParentalRating", "ParentalRatingReason", "ParentalRatingsOrganization", "ParsingBindContext", "ParsingName", _
    "ParsingPath", "PerceivedType", "PercentFull", "Priority", "PriorityText", "Project", "ProviderItemID", "Rating", "RatingText", "RemoteConflictingFile", "Sensitivity", "SensitivityText", "SFGAOFlags", "SharedWith", "ShareUserRating", "SharingStatus", "Shell", "SimpleRating", "Size", "SoftwareUsed", "SourceItem", "SourcePackageFamilyName", "StartDate", "Status", "StorageProviderCallerVersionInformation", "StorageProviderError", "StorageProviderFileChecksum", "StorageProviderFileIdentifier", "StorageProviderFileRemoteUri", "StorageProviderFileVersion", "StorageProviderFileVersionWaterline", "StorageProviderId", "StorageProviderShareStatuses", "StorageProviderSharingStatus", "StorageProviderStatus", "Subject", "SyncTransferStatus", "Thumbnail", "ThumbnailCacheId", "ThumbnailStream", "Title", "TitleSortOverride", "TotalFileSize", "Trademarks", "TransferOrder", "TransferPosition", "TransferSize", "VolumeId", "ZoneIdentifier")

PropertiesArray_Document = Array("ByteCount", "CharacterCount", "ClientID", "Contributor", "DateCreated", "DatePrinted", "DateSaved", "Division", "DocumentID", "HiddenSlideCount", "LastAuthor", "LineCount", "Manager", "MultimediaClipCount", "NoteCount", "PageCount", "ParagraphCount", "PresentationFormat", "RevisionNumber", "Security", "SlideCount", "Template", "TotalEditingTime", "Version", "WordCount")
PropertiesArray_Image = Array("BitDepth", "ColorSpace", "CompressedBitsPerPixel", "CompressedBitsPerPixelDenominator", "CompressedBitsPerPixelNumerator", "Compression", "CompressionText", "Dimensions", "HorizontalResolution", "HorizontalSize", "ImageID", "ResolutionUnit", "VerticalResolution", "VerticalSize")
PropertiesArray_Media = Array("AuthorUrl", "AverageLevel", "ClassPrimaryID", "ClassSecondaryID", "CollectionGroupID", "CollectionID", "ContentDistributor", "ContentID", "CreatorApplication", "CreatorApplicationVersion", "DateEncoded", "DateReleased", "DlnaProfileID", "Duration", "DVDID", "EncodedBy", "EncodingSettings", "EpisodeNumber", "FrameCount", "MCDI", "MetadataContentProvider", "Producer", "PromotionUrl", "ProtectionType", "ProviderRating", "ProviderStyle", "Publisher", "SeasonNumber", "SeriesName", "SubscriptionContentId", "SubTitle", "ThumbnailLargePath", "ThumbnailLargeUri", "ThumbnailSmallPath", "ThumbnailSmallUri", "UniqueFileIdentifier", "UserNoAutoInfo", "UserWebUrl", "Writer", "Year")
PropertiesArray_Message = Array("AttachmentContents", "AttachmentNames", "BccAddress", "BccName", "CcAddress", "CcName", "ConversationID", "ConversationIndex", "DateReceived", "DateSent", "Flags", "FromAddress", "FromName", "HasAttachments", "IsFwdOrReply", "MessageClass", "Participants", "ProofInProgress", "SenderAddress", "SenderName", "Store", "ToAddress", "ToDoFlags", "ToDoTitle", "ToName")
PropertiesArray_Music = Array("AlbumArtist", "AlbumArtistSortOverride", "AlbumID", "AlbumTitle", "AlbumTitleSortOverride", "Artist", "ArtistSortOverride", "BeatsPerMinute", "Composer", "ComposerSortOverride", "Conductor", "ContentGroupDescription", "DiscNumber", "DisplayArtist", "Genre", "InitialKey", "IsCompilation", "Lyrics", "Mood", "PartOfSet", "Period", "SynchronizedLyrics", "TrackNumber")
PropertiesArray_Photo = Array("Aperture", "ApertureDenominator", "ApertureNumerator", "Brightness", "BrightnessDenominator", "BrightnessNumerator", "CameraManufacturer", "CameraModel", "CameraSerialNumber", "Contrast", "ContrastText", "DateTaken", "DigitalZoom", "DigitalZoomDenominator", "DigitalZoomNumerator", "Event", "EXIFVersion", "ExposureBias", "ExposureBiasDenominator", "ExposureBiasNumerator", "ExposureIndex", "ExposureIndexDenominator", "ExposureIndexNumerator", "ExposureProgram", "ExposureProgramText", "ExposureTime", "ExposureTimeDenominator", "ExposureTimeNumerator", "Flash", "FlashEnergy", "FlashEnergyDenominator", "FlashEnergyNumerator", "FlashManufacturer", "FlashModel", "FlashText", "FNumber", "FNumberDenominator", "FNumberNumerator", "FocalLength", "FocalLengthDenominator", "FocalLengthInFilm", "FocalLengthNumerator", "FocalPlaneXResolution", "FocalPlaneXResolutionDenominator", "FocalPlaneXResolutionNumerator", "FocalPlaneYResolution", "FocalPlaneYResolutionDenominator", _
    "FocalPlaneYResolutionNumerator", "GainControl", "GainControlDenominator", "GainControlNumerator", "GainControlText", "ISOSpeed", "LensManufacturer", "LensModel", "LightSource", "MakerNote", "MakerNoteOffset", "MaxAperture", "MaxApertureDenominator", "MaxApertureNumerator", "MeteringMode", "MeteringModeText", "Orientation", "OrientationText", "PeopleNames", "PhotometricInterpretation", "PhotometricInterpretationText", "ProgramMode", "ProgramModeText", "RelatedSoundFile", "Saturation", "SaturationText", "Sharpness", "SharpnessText", "ShutterSpeed", "ShutterSpeedDenominator", "ShutterSpeedNumerator", "SubjectDistance", "SubjectDistanceDenominator", "SubjectDistanceNumerator", "TagViewAggregate", "TranscodedForSync", "WhiteBalance", "WhiteBalanceText")
PropertiesArray_PropGroup = Array("Advanced", "Audio", "Calendar", "Camera", "Contact", "Content", "Description", "FileSystem", "General", "GPS", "Image", "Media", "MediaAdvanced", "Message", "Music", "Origin", "PhotoAdvanced", "RecordedTV", "Video")
PropertiesArray_Video = Array("Compression", "Director", "EncodingBitrate", "FourCC", "FrameHeight", "FrameRate", "FrameWidth", "HorizontalAspectRatio", "IsSpherical", "IsStereo", "Orientation", "SampleSize", "StreamName", "StreamNumber", "TotalBitrate", "TranscodedForSync", "VerticalAspectRatio")
End Function
