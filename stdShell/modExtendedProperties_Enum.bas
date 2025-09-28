Attribute VB_Name = "modExtendedProperties_Enum"


',Advanced,Audio,Calendar,Camera,Contact,Content,Description,FileSystem,General,GPS,Image,Media,MediaAdvanced,message,Music,Origin,PhotoAdvanced,RecordedTV,Video
Private Enum stdShellExtendedProps_PropGroup
    KindAdvanced
    KindAudio
    KindCalendar
    KindCamera
    KindContact
    KindContent
    KindDescription
    KindFileSystem
    KindGeneral
    KindGPS
    KindImage
    KindMedia
    KindMediaAdvanced
    Kindmessage
    KindMusic
    KindOrigin
    KindPhotoAdvanced
    KindRecordedTV
    KindVideo
End Enum

Private Enum stdShellExtendedProps
    [_stdShellExtendedPropsStart] = 0
    GeneralAcquisitionID = [_stdShellExtendedPropsStart]
    GeneralApplicationDefinedProperties
    GeneralApplicationName
    GeneralAppZoneIdentifier
    GeneralAuthor
    GeneralCachedFileUpdaterContentIdForConflictResolution
    GeneralCachedFileUpdaterContentIdForStream
    GeneralCapacity
    GeneralCategory
    GeneralComment
    GeneralCompany
    GeneralComputerName
    GeneralContainedItems
    GeneralContentStatus
    GeneralContentType
    GeneralCopyright
    GeneralCreatorAppId
    GeneralCreatorOpenWithUIOptions
    GeneralDataObjectFormat
    GeneralDateAccessed
    GeneralDateAcquired
    GeneralDateArchived
    GeneralDateCompleted
    GeneralDateCreated
    GeneralDateImported
    GeneralDateModified
    GeneralDefaultSaveLocationDisplay
    GeneralDueDate
    GeneralEndDate
    GeneralExpandoProperties
    GeneralFileAllocationSize
    GeneralFileAttributes
    GeneralFileCount
    GeneralFileDescription
    GeneralFileExtension
    GeneralFileFRN
    GeneralFilename
    GeneralFileOfflineAvailabilityStatus
    GeneralFileOwner
    GeneralFilePlaceholderStatus
    GeneralFileVersion
    GeneralFindData
    GeneralFlagColor
    GeneralFlagColorText
    GeneralFlagStatus
    GeneralFlagStatusText
    GeneralFolderKind
    GeneralFolderNameDisplay
    GeneralFreeSpace
    GeneralFullText
    GeneralHighKeywords
    GeneralImageParsingName
    GeneralImportance
    GeneralImportanceText
    GeneralIsAttachment
    GeneralIsDefaultNonOwnerSaveLocation
    GeneralIsDefaultSaveLocation
    GeneralIsDeleted
    GeneralIsEncrypted
    GeneralIsFlagged
    GeneralIsFlaggedComplete
    GeneralIsIncomplete
    GeneralIsLocationSupported
    GeneralIsPinnedToNameSpaceTree
    GeneralIsRead
    GeneralIsSearchOnlyItem
    GeneralIsSendToTarget
    GeneralIsShared
    GeneralItemAuthors
    GeneralItemClassType
    GeneralItemDate
    GeneralItemFolderNameDisplay
    GeneralItemFolderPathDisplay
    GeneralItemFolderPathDisplayNarrow
    GeneralItemName
    GeneralItemNameDisplay
    GeneralItemNameDisplayWithoutExtension
    GeneralItemNamePrefix
    GeneralItemNameSortOverride
    GeneralItemParticipants
    GeneralItemPathDisplay
    GeneralItemPathDisplayNarrow
    GeneralItemSubType
    GeneralItemType
    GeneralItemTypeText
    GeneralItemUrl
    GeneralKEYWORDS
    GeneralKind
    GeneralKindText
    Generallanguage
    GeneralLastSyncError
    GeneralLastWriterPackageFamilyName
    GeneralLowKeywords
    GeneralMediumKeywords
    GeneralMileageInformation
    GeneralMIMEType
    [_Null]
    GeneralOfflineAvailability
    GeneralOfflineStatus
    GeneralOriginalFileName
    GeneralOwnerSID
    GeneralParentalRating
    GeneralParentalRatingReason
    GeneralParentalRatingsOrganization
    GeneralParsingBindContext
    GeneralParsingName
    GeneralParsingPath
    GeneralPerceivedType
    GeneralPercentFull
    GeneralPriority
    GeneralPriorityText
    GeneralProject
    GeneralProviderItemID
    GeneralRating
    GeneralRatingText
    GeneralRemoteConflictingFile
    GeneralSensitivity
    GeneralSensitivityText
    GeneralSFGAOFlags
    GeneralSharedWith
    GeneralShareUserRating
    GeneralSharingStatus
    Generalshell
    GeneralSimpleRating
    GeneralSize
    GeneralSoftwareUsed
    GeneralSourceItem
    GeneralSourcePackageFamilyName
    GeneralStartDate
    GeneralStatus
    GeneralStorageProviderCallerVersionInformation
    GeneralStorageProviderError
    GeneralStorageProviderFileChecksum
    GeneralStorageProviderFileIdentifier
    GeneralStorageProviderFileRemoteUri
    GeneralStorageProviderFileVersion
    GeneralStorageProviderFileVersionWaterline
    GeneralStorageProviderId
    GeneralStorageProviderShareStatuses
    GeneralStorageProviderSharingStatus
    GeneralStorageProviderStatus
    GeneralSubject
    GeneralSyncTransferStatus
    GeneralThumbnail
    GeneralThumbnailCacheId
    GeneralThumbnailStream
    GeneralTitle
    GeneralTitleSortOverride
    GeneralTotalFileSize
    GeneralTrademarks
    GeneralTransferOrder
    GeneralTransferPosition
    GeneralTransferSize
    GeneralVolumeId
    GeneralZoneIdentifier
    DocumentByteCount
    DocumentCharacterCount
    DocumentClientID
    DocumentContributor
    DocumentDateCreated
    DocumentDatePrinted
    DocumentDateSaved
    DocumentDivision
    DocumentDocumentID
    DocumentHiddenSlideCount
    DocumentLastAuthor
    DocumentLineCount
    DocumentManager
    DocumentMultimediaClipCount
    DocumentNoteCount
    DocumentPageCount
    DocumentParagraphCount
    DocumentPresentationFormat
    DocumentRevisionNumber
    DocumentSecurity
    DocumentSlideCount
    DocumentTEMPLATE
    DocumentTotalEditingTime
    DocumentVersion
    DocumentWordCount
    VideoCompression
    VideoDirector
    VideoEncodingBitrate
    VideoFourCC
    VideoFrameHeight
    VideoFrameRate
    VideoFrameWidth
    VideoHorizontalAspectRatio
    VideoIsSpherical
    VideoIsStereo
    VideoOrientation
    VideoSampleSize
    VideoStreamName
    VideoStreamNumber
    VideoTotalBitrate
    VideoTranscodedForSync
    VideoVerticalAspectRatio
    PhotoAperture
    PhotoApertureDenominator
    PhotoApertureNumerator
    PhotoBrightness
    PhotoBrightnessDenominator
    PhotoBrightnessNumerator
    PhotoCameraManufacturer
    PhotoCameraModel
    PhotoCameraSerialNumber
    PhotoContrast
    PhotoContrastText
    PhotoDateTaken
    PhotoDigitalZoom
    PhotoDigitalZoomDenominator
    PhotoDigitalZoomNumerator
    [_Event]
    PhotoEXIFVersion
    PhotoExposureBias
    PhotoExposureBiasDenominator
    PhotoExposureBiasNumerator
    PhotoExposureIndex
    PhotoExposureIndexDenominator
    PhotoExposureIndexNumerator
    PhotoExposureProgram
    PhotoExposureProgramText
    PhotoExposureTime
    PhotoExposureTimeDenominator
    PhotoExposureTimeNumerator
    PhotoFlash
    PhotoFlashEnergy
    PhotoFlashEnergyDenominator
    PhotoFlashEnergyNumerator
    PhotoFlashManufacturer
    PhotoFlashModel
    PhotoFlashText
    PhotoFNumber
    PhotoFNumberDenominator
    PhotoFNumberNumerator
    PhotoFocalLength
    PhotoFocalLengthDenominator
    PhotoFocalLengthInFilm
    PhotoFocalLengthNumerator
    PhotoFocalPlaneXResolution
    PhotoFocalPlaneXResolutionDenominator
    PhotoFocalPlaneXResolutionNumerator
    PhotoFocalPlaneYResolution
    PhotoFocalPlaneYResolutionDenominator
    PhotoFocalPlaneYResolutionNumerator
    PhotoGainControl
    PhotoGainControlDenominator
    PhotoGainControlNumerator
    PhotoGainControlText
    PhotoISOSpeed
    PhotoLensManufacturer
    PhotoLensModel
    PhotoLightSource
    PhotoMakerNote
    PhotoMakerNoteOffset
    PhotoMaxAperture
    PhotoMaxApertureDenominator
    PhotoMaxApertureNumerator
    PhotoMeteringMode
    PhotoMeteringModeText
    PhotoOrientation
    PhotoOrientationText
    PhotoPeopleNames
    PhotoPhotometricInterpretation
    PhotoPhotometricInterpretationText
    PhotoProgramMode
    PhotoProgramModeText
    PhotoRelatedSoundFile
    PhotoSaturation
    PhotoSaturationText
    PhotoSharpness
    PhotoSharpnessText
    PhotoShutterSpeed
    PhotoShutterSpeedDenominator
    PhotoShutterSpeedNumerator
    PhotoSubjectDistance
    PhotoSubjectDistanceDenominator
    PhotoSubjectDistanceNumerator
    PhotoTagViewAggregate
    PhotoTranscodedForSync
    PhotoWhiteBalance
    PhotoWhiteBalanceText
    MusicAlbumArtist
    MusicAlbumArtistSortOverride
    MusicAlbumID
    MusicAlbumTitle
    MusicAlbumTitleSortOverride
    MusicArtist
    MusicArtistSortOverride
    MusicBeatsPerMinute
    MusicComposer
    MusicComposerSortOverride
    MusicConductor
    MusicContentGroupDescription
    MusicDiscNumber
    MusicDisplayArtist
    MusicGenre
    MusicInitialKey
    MusicIsCompilation
    MusicLyrics
    MusicMood
    MusicPartOfSet
    Musicperiod
    MusicSynchronizedLyrics
    MusicTrackNumber
    MailAttachmentContents
    MailAttachmentNames
    MailBccAddress
    MailBccName
    MailCcAddress
    MailCcName
    MailConversationID
    MailConversationIndex
    MailDateReceived
    MailDateSent
    MailFlags
    MailFromAddress
    MailFromName
    MailHasAttachments
    MailIsFwdOrReply
    MailMessageClass
    MailParticipants
    MailProofInProgress
    MailSenderAddress
    MailSenderName
    MailStore
    MailToAddress
    MailToDoFlags
    MailToDoTitle
    MailToName
    MediaAuthorUrl
    MediaAverageLevel
    MediaClassPrimaryID
    MediaClassSecondaryID
    MediaCollectionGroupID
    MediaCollectionID
    MediaContentDistributor
    MediaContentID
    MediaCreatorApplication
    MediaCreatorApplicationVersion
    MediaDateEncoded
    MediaDateReleased
    MediaDlnaProfileID
    Mediaduration
    MediaDVDID
    MediaEncodedBy
    MediaEncodingSettings
    MediaEpisodeNumber
    MediaFrameCount
    MediaMCDI
    MediaMetadataContentProvider
    MediaProducer
    MediaPromotionUrl
    MediaProtectionType
    MediaProviderRating
    MediaProviderStyle
    MediaPublisher
    MediaSeasonNumber
    MediaSeriesName
    MediaSubscriptionContentId
    MediaSubTitle
    MediaThumbnailLargePath
    MediaThumbnailLargeUri
    MediaThumbnailSmallPath
    MediaThumbnailSmallUri
    MediaUniqueFileIdentifier
    MediaUserNoAutoInfo
    MediaUserWebUrl
    MediaWriter
    MediaYear
    ImageBitDepth
    ImageColorSpace
    ImageCompressedBitsPerPixel
    ImageCompressedBitsPerPixelDenominator
    ImageCompressedBitsPerPixelNumerator
    ImageCompression
    ImageCompressionText
    ImageDimensions
    ImageHorizontalResolution
    ImageHorizontalSize
    ImageImageID
    ImageResolutionUnit
    ImageVerticalResolution
    ImageVerticalSize

    [_stdShellExtendedPropsEnd]
End Enum

'Obtain the value of a property from the extended properties
'@param str - The property name to obtain
'@return Variant<Variant|Array<Variant>> - The value of the property
Public Property Get PropertyFromString(ByVal str as string) as Variant
  If Exists then
    PropertyFromString = CreateObject("Shell.Application").Namespace(0).ParseName(This.path).ExtendedProperty(str)
  End If
End Property

'Obtain the value of a property from the extended properties
'@param iProp - The property to obtain
'@return Variant<Variant|Array<Variant>> - The value of the property
Public Property Get PropertyFromEnum(ByVal iProp as EShellProperties) as Variant
  PropertyFromEnum = PropertyFromString(stdShell.protEnumToPropName(iProp))
End Property

'Obtain the name of a property from the extended properties
'@protected
'@param prop - The property to obtain
'@return String - The name of the property as utilised in Shell.Application
Public Function protEnumToPropName(ByVal prop as stdShellExtendedProps) As String
    static map as variant
    if isEmpty(map) Then
        Dim s as string
        'General
        s = s & "AcquisitionID,ApplicationDefinedProperties,ApplicationName,AppZoneIdentifier,Author,CachedFileUpdaterContentIdForConflictResolution"
        s = s & ",CachedFileUpdaterContentIdForStream,Capacity,Category,Comment,Company,ComputerName,ContainedItems,ContentStatus,ContentType,Copyright"
        s = s & ",CreatorAppId,CreatorOpenWithUIOptions,DataObjectFormat,DateAccessed,DateAcquired,DateArchived,DateCompleted,DateCreated,DateImported"
        s = s & ",DateModified,DefaultSaveLocationDisplay,DueDate,EndDate,ExpandoProperties,FileAllocationSize,FileAttributes,FileCount,FileDescription"
        s = s & ",FileExtension,FileFRN,FileName,FileOfflineAvailabilityStatus,FileOwner,FilePlaceholderStatus,FileVersion,FindData,FlagColor,FlagColorText"
        s = s & ",FlagStatus,FlagStatusText,FolderKind,FolderNameDisplay,FreeSpace,FullText,HighKeywords,ImageParsingName,Importance,ImportanceText,IsAttachment"
        s = s & ",IsDefaultNonOwnerSaveLocation,IsDefaultSaveLocation,IsDeleted,IsEncrypted,IsFlagged,IsFlaggedComplete,IsIncomplete,IsLocationSupported"
        s = s & ",IsPinnedToNameSpaceTree,IsRead,IsSearchOnlyItem,IsSendToTarget,IsShared,ItemAuthors,ItemClassType,ItemDate,ItemFolderNameDisplay"
        s = s & ",ItemFolderPathDisplay,ItemFolderPathDisplayNarrow,ItemName,ItemNameDisplay,ItemNameDisplayWithoutExtension,ItemNamePrefix"
        s = s & ",ItemNameSortOverride,ItemParticipants,ItemPathDisplay,ItemPathDisplayNarrow,ItemSubType,ItemType,ItemTypeText,ItemUrl,Keywords,Kind,KindText"
        s = s & ",Language,LastSyncError,LastWriterPackageFamilyName,LowKeywords,MediumKeywords,MileageInformation,MIMEType,Null,OfflineAvailability,OfflineStatus"
        s = s & ",OriginalFileName,OwnerSID,ParentalRating,ParentalRatingReason,ParentalRatingsOrganization,ParsingBindContext,ParsingName,ParsingPath,PerceivedType"
        s = s & ",PercentFull,Priority,PriorityText,Project,ProviderItemID,Rating,RatingText,RemoteConflictingFile,Sensitivity,SensitivityText,SFGAOFlags,SharedWith"
        s = s & ",ShareUserRating,SharingStatus,Shell,SimpleRating,Size,SoftwareUsed,SourceItem,SourcePackageFamilyName,StartDate,Status"
        s = s & ",StorageProviderCallerVersionInformation,StorageProviderError,StorageProviderFileChecksum,StorageProviderFileIdentifier,StorageProviderFileRemoteUri"
        s = s & ",StorageProviderFileVersion,StorageProviderFileVersionWaterline,StorageProviderId,StorageProviderShareStatuses,StorageProviderSharingStatus"
        s = s & ",StorageProviderStatus,Subject,SyncTransferStatus,Thumbnail,ThumbnailCacheId,ThumbnailStream,Title,TitleSortOverride,TotalFileSize,Trademarks"
        s = s & ",TransferOrder,TransferPosition,TransferSize,VolumeId,ZoneIdentifier"
        
        'Document
        s = s & ",ByteCount,CharacterCount,ClientID,Contributor,DateCreated,DatePrinted,DateSaved,Division,DocumentID,HiddenSlideCount,LastAuthor,LineCount,Manager"
        s = s & ",MultimediaClipCount,NoteCount,PageCount,ParagraphCount,PresentationFormat,RevisionNumber,Security,SlideCount,Template,TotalEditingTime,Version,WordCount"
        
        'Video
        s = s & ",Compression,Director,EncodingBitrate,FourCC,FrameHeight,FrameRate,FrameWidth,HorizontalAspectRatio,IsSpherical,IsStereo,Orientation,SampleSize"
        s = s & ",StreamName,StreamNumber,TotalBitrate,TranscodedForSync,VerticalAspectRatio"

        'Photo
        s = s & ",Aperture,ApertureDenominator,ApertureNumerator,Brightness,BrightnessDenominator,BrightnessNumerator,CameraManufacturer,CameraModel"
        s = s & ",CameraSerialNumber,Contrast,ContrastText,DateTaken,DigitalZoom,DigitalZoomDenominator,DigitalZoomNumerator,Event,EXIFVersion,ExposureBias"
        s = s & ",ExposureBiasDenominator,ExposureBiasNumerator,ExposureIndex,ExposureIndexDenominator,ExposureIndexNumerator,ExposureProgram,ExposureProgramText"
        s = s & ",ExposureTime,ExposureTimeDenominator,ExposureTimeNumerator,Flash,FlashEnergy,FlashEnergyDenominator,FlashEnergyNumerator,FlashManufacturer"
        s = s & ",FlashModel,FlashText,FNumber,FNumberDenominator,FNumberNumerator,FocalLength,FocalLengthDenominator,FocalLengthInFilm,FocalLengthNumerator"
        s = s & ",FocalPlaneXResolution,FocalPlaneXResolutionDenominator,FocalPlaneXResolutionNumerator,FocalPlaneYResolution,FocalPlaneYResolutionDenominator"
        s = s & ",FocalPlaneYResolutionNumerator,GainControl,GainControlDenominator,GainControlNumerator,GainControlText,ISOSpeed,LensManufacturer,LensModel"
        s = s & ",LightSource,MakerNote,MakerNoteOffset,MaxAperture,MaxApertureDenominator,MaxApertureNumerator,MeteringMode,MeteringModeText,Orientation"
        s = s & ",OrientationText,PeopleNames,PhotometricInterpretation,PhotometricInterpretationText,ProgramMode,ProgramModeText,RelatedSoundFile,Saturation"
        s = s & ",SaturationText,Sharpness,SharpnessText,ShutterSpeed,ShutterSpeedDenominator,ShutterSpeedNumerator,SubjectDistance,SubjectDistanceDenominator"
        s = s & ",SubjectDistanceNumerator,TagViewAggregate,TranscodedForSync,WhiteBalance,WhiteBalanceText"

        'Music
        s = s & ",AlbumArtist,AlbumArtistSortOverride,AlbumID,AlbumTitle,AlbumTitleSortOverride,Artist,ArtistSortOverride,BeatsPerMinute,Composer,ComposerSortOverride"
        s = s & ",Conductor,ContentGroupDescription,DiscNumber,DisplayArtist,Genre,InitialKey,IsCompilation,Lyrics,Mood,PartOfSet,Period,SynchronizedLyrics,TrackNumber"

        'EMail/Message
        s = s & ",AttachmentContents,AttachmentNames,BccAddress,BccName,CcAddress,CcName,ConversationID,ConversationIndex,DateReceived,DateSent,Flags,FromAddress"
        s = s & ",FromName,HasAttachments,IsFwdOrReply,MessageClass,Participants,ProofInProgress,SenderAddress,SenderName,Store,ToAddress,ToDoFlags,ToDoTitle,ToName"

        'Media
        s = s & ",AuthorUrl,AverageLevel,ClassPrimaryID,ClassSecondaryID,CollectionGroupID,CollectionID,ContentDistributor,ContentID,CreatorApplication"
        s = s & ",CreatorApplicationVersion,DateEncoded,DateReleased,DlnaProfileID,Duration,DVDID,EncodedBy,EncodingSettings,EpisodeNumber,FrameCount,MCDI"
        s = s & ",MetadataContentProvider,Producer,PromotionUrl,ProtectionType,ProviderRating,ProviderStyle,Publisher,SeasonNumber,SeriesName,SubscriptionContentId"
        s = s & ",SubTitle,ThumbnailLargePath,ThumbnailLargeUri,ThumbnailSmallPath,ThumbnailSmallUri,UniqueFileIdentifier,UserNoAutoInfo,UserWebUrl,Writer,Year"

        'Image
        s = s & ",BitDepth,ColorSpace,CompressedBitsPerPixel,CompressedBitsPerPixelDenominator,CompressedBitsPerPixelNumerator,Compression,CompressionText,Dimensions"
        s = s & ",HorizontalResolution,HorizontalSize,ImageID,ResolutionUnit,VerticalResolution,VerticalSize"
        
        map = split(s, ",")
    end if
    protEnumToPropName = map(prop)
Exit Function













