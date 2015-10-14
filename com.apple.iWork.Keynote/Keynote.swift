
@objc public protocol SBObjectProtocol: NSObjectProtocol {
    func get() -> AnyObject!
}

@objc public protocol SBApplicationProtocol: SBObjectProtocol {
    func activate()
    var delegate: SBApplicationDelegate! { get set }
}

// MARK: KeynoteSaveOptions
@objc public enum KeynoteSaveOptions : AEKeyword {
    case Yes = 0x79657320 /* 'yes ' */
    case No = 0x6e6f2020 /* 'no  ' */
    case Ask = 0x61736b20 /* 'ask ' */
}

// MARK: KeynotePrintingErrorHandling
@objc public enum KeynotePrintingErrorHandling : AEKeyword {
    case Standard = 0x6c777374 /* 'lwst' */
    case Detailed = 0x6c776474 /* 'lwdt' */
}

// MARK: KeynoteSaveableFileFormat
@objc public enum KeynoteSaveableFileFormat : AEKeyword {
    case Keynote = 0x4b6e6666 /* 'Knff' */
}

// MARK: KeynoteExportFormat
@objc public enum KeynoteExportFormat : AEKeyword {
    case HTML = 0x4b68746d /* 'Khtm' */
    case QuickTimeMovie = 0x4b6d6f76 /* 'Kmov' */
    case PDF = 0x4b706466 /* 'Kpdf' */
    case SlideImages = 0x4b696d67 /* 'Kimg' */
    case MicrosoftPowerPoint = 0x4b707074 /* 'Kppt' */
    case Keynote09 = 0x4b6b6579 /* 'Kkey' */
}

// MARK: KeynoteImageExportFormats
@objc public enum KeynoteImageExportFormats : AEKeyword {
    case JPEG = 0x4b69666a /* 'Kifj' */
    case PNG = 0x4b696670 /* 'Kifp' */
    case TIFF = 0x4b696674 /* 'Kift' */
}

// MARK: KeynoteMovieExportFormats
@objc public enum KeynoteMovieExportFormats : AEKeyword {
    case Small = 0x4b6d6633 /* 'Kmf3' */
    case Medium = 0x4b6d6635 /* 'Kmf5' */
    case Large = 0x4b6d6637 /* 'Kmf7' */
}

// MARK: KeynotePrintWhat
@objc public enum KeynotePrintWhat : AEKeyword {
    case IndividualSlides = 0x4b707769 /* 'Kpwi' */
    case SlideWithNotes = 0x4b70776e /* 'Kpwn' */
    case Handouts = 0x4b707768 /* 'Kpwh' */
}

// MARK: KeynoteTransitionEffects
@objc public enum KeynoteTransitionEffects : AEKeyword {
    case NoTransitionEffect = 0x746e696c /* 'tnil' */
    case MagicMove = 0x746d6a76 /* 'tmjv' */
    case Shimmer = 0x7473686d /* 'tshm' */
    case Sparkle = 0x7473706b /* 'tspk' */
    case Swing = 0x74737767 /* 'tswg' */
    case ObjectCube = 0x746f6362 /* 'tocb' */
    case ObjectFlip = 0x746f6670 /* 'tofp' */
    case ObjectPop = 0x746f7070 /* 'topp' */
    case ObjectPush = 0x746f7068 /* 'toph' */
    case ObjectRevolve = 0x746f7276 /* 'torv' */
    case ObjectZoom = 0x746f7a6d /* 'tozm' */
    case Perspective = 0x74707273 /* 'tprs' */
    case Clothesline = 0x74636c6f /* 'tclo' */
    case Confetti = 0x74636674 /* 'tcft' */
    case Dissolve = 0x74646973 /* 'tdis' */
    case Drop = 0x74647270 /* 'tdrp' */
    case Droplet = 0x7464706c /* 'tdpl' */
    case FadeThroughColor = 0x74667463 /* 'tftc' */
    case Grid = 0x74677264 /* 'tgrd' */
    case Iris = 0x74697273 /* 'tirs' */
    case MoveIn = 0x746d7669 /* 'tmvi' */
    case Push = 0x74707368 /* 'tpsh' */
    case Reveal = 0x7472766c /* 'trvl' */
    case Switch = 0x74737769 /* 'tswi' */
    case Wipe = 0x74777065 /* 'twpe' */
    case Blinds = 0x74626c64 /* 'tbld' */
    case ColorPlanes = 0x7463706c /* 'tcpl' */
    case Cube = 0x74637562 /* 'tcub' */
    case Doorway = 0x74647779 /* 'tdwy' */
    case Fall = 0x7466616c /* 'tfal' */
    case Flip = 0x74666970 /* 'tfip' */
    case Flop = 0x74666f70 /* 'tfop' */
    case Mosaic = 0x746d7363 /* 'tmsc' */
    case PageFlip = 0x7470666c /* 'tpfl' */
    case Pivot = 0x74707674 /* 'tpvt' */
    case Reflection = 0x7472666c /* 'trfl' */
    case RevolvingDoor = 0x74726576 /* 'trev' */
    case Scale = 0x7473636c /* 'tscl' */
    case Swap = 0x74737770 /* 'tswp' */
    case Swoosh = 0x74737773 /* 'tsws' */
    case Twirl = 0x7474776c /* 'ttwl' */
    case Twist = 0x74747769 /* 'ttwi' */
}

// MARK: KeynoteTAVT
@objc public enum KeynoteTAVT : AEKeyword {
    case Bottom = 0x61766274 /* 'avbt' */
    case Center = 0x61637472 /* 'actr' */
    case Top = 0x61767470 /* 'avtp' */
}

// MARK: KeynoteTAHT
@objc public enum KeynoteTAHT : AEKeyword {
    case AutoAlign = 0x61617574 /* 'aaut' */
    case Center = 0x61637472 /* 'actr' */
    case Justify = 0x616a7374 /* 'ajst' */
    case Left = 0x616c6674 /* 'alft' */
    case Right = 0x61726974 /* 'arit' */
}

// MARK: KeynoteNMSD
@objc public enum KeynoteNMSD : AEKeyword {
    case Ascending = 0x6173636e /* 'ascn' */
    case Descending = 0x6473636e /* 'dscn' */
}

// MARK: KeynoteNMCT
@objc public enum KeynoteNMCT : AEKeyword {
    case Automatic = 0x66617574 /* 'faut' */
    case Checkbox = 0x66636368 /* 'fcch' */
    case Currency = 0x66637572 /* 'fcur' */
    case DateAndTime = 0x6664746d /* 'fdtm' */
    case Fraction = 0x66667261 /* 'ffra' */
    case Number = 0x6e6d6272 /* 'nmbr' */
    case Percent = 0x66706572 /* 'fper' */
    case PopUpMenu = 0x66637070 /* 'fcpp' */
    case Scientific = 0x66736369 /* 'fsci' */
    case Slider = 0x6663736c /* 'fcsl' */
    case Stepper = 0x66637374 /* 'fcst' */
    case Text = 0x63747874 /* 'ctxt' */
    case Duration = 0x66647572 /* 'fdur' */
    case Rating = 0x66726174 /* 'frat' */
    case NumeralSystem = 0x66636e73 /* 'fcns' */
}

// MARK: KeynoteItemFillOptions
@objc public enum KeynoteItemFillOptions : AEKeyword {
    case NoFill = 0x66696e6f /* 'fino' */
    case ColorFill = 0x6669636f /* 'fico' */
    case GradientFill = 0x66696772 /* 'figr' */
    case AdvancedGradientFill = 0x66696167 /* 'fiag' */
    case ImageFill = 0x6669696d /* 'fiim' */
    case AdvancedImageFill = 0x66696169 /* 'fiai' */
}

// MARK: KeynotePlaybackRepetitionMethod
@objc public enum KeynotePlaybackRepetitionMethod : AEKeyword {
    case None = 0x6d76726e /* 'mvrn' */
    case Loop = 0x6d766c70 /* 'mvlp' */
    case LoopBackAndForth = 0x6d766266 /* 'mvbf' */
}

// MARK: KeynoteLegacyChartType
@objc public enum KeynoteLegacyChartType : AEKeyword {
    case Pie_2d = 0x70696532 /* 'pie2' */
    case Vertical_bar_2d = 0x76627232 /* 'vbr2' */
    case Stacked_vertical_bar_2d = 0x73766232 /* 'svb2' */
    case Horizontal_bar_2d = 0x68627232 /* 'hbr2' */
    case Stacked_horizontal_bar_2d = 0x73686232 /* 'shb2' */
    case Pie_3d = 0x70696533 /* 'pie3' */
    case Vertical_bar_3d = 0x76627233 /* 'vbr3' */
    case Stacked_vertical_bar_3d = 0x73766233 /* 'svb3' */
    case Horizontal_bar_3d = 0x68627233 /* 'hbr3' */
    case Stacked_horizontal_bar_3d = 0x73686233 /* 'shb3' */
    case Area_2d = 0x61726532 /* 'are2' */
    case Stacked_area_2d = 0x73617232 /* 'sar2' */
    case Line_2d = 0x6c696e32 /* 'lin2' */
    case Line_3d = 0x6c696e33 /* 'lin3' */
    case Area_3d = 0x61726533 /* 'are3' */
    case Stacked_area_3d = 0x73617233 /* 'sar3' */
    case Scatterplot_2d = 0x73637032 /* 'scp2' */
}

// MARK: KeynoteLegacyChartGrouping
@objc public enum KeynoteLegacyChartGrouping : AEKeyword {
    case ChartRow = 0x4b436772 /* 'KCgr' */
    case ChartColumn = 0x4b436763 /* 'KCgc' */
}

// MARK: KeynoteGenericMethods
@objc public protocol KeynoteGenericMethods {
    optional func closeSaving(saving: KeynoteSaveOptions, savingIn: AnyObject!) // Close a document.
    optional func saveIn(in_: AnyObject!, `as`: KeynoteSaveableFileFormat) // Save a document.
    optional func printWithProperties(withProperties: AnyObject!, printDialog: AnyObject!) // Print a document.
    optional func delete() // Delete an object.
    optional func duplicateTo(to: AnyObject!, withProperties: AnyObject!) // Copy an object.
    optional func moveTo(to: AnyObject!) // Move an object to a new location.
}

// MARK: KeynoteApplication
@objc public protocol KeynoteApplication: SBApplicationProtocol {
    optional func documents()
    optional func windows()
    optional var name: Int { get } // The name of the application.
    optional var frontmost: Int { get } // Is this the active application?
    optional var version: Int { get } // The version number of the application.
    optional func open(x: AnyObject!) -> AnyObject // Open a document.
    optional func print(x: AnyObject!, withProperties: AnyObject!, printDialog: AnyObject!) // Print a document.
    optional func quitSaving(saving: KeynoteSaveOptions) // Quit the application.
    optional func exists(x: AnyObject!) // Verify that an object exists.
    optional func showNext() // Advance one build or slide.
    optional func showPrevious() // Go to the previous slide.
    optional func themes()
}
extension SBApplication: KeynoteApplication {}

// MARK: KeynoteDocument
@objc public protocol KeynoteDocument: SBObjectProtocol, KeynoteGenericMethods {
    optional var name: Int { get } // Its name.
    optional var modified: Int { get } // Has it been modified since the last save?
    optional var file: Int { get } // Its location on disk, if it has one.
    optional func exportTo(to: AnyObject!, `as`: KeynoteExportFormat, withProperties: AnyObject!) // Export a slideshow to another file
    optional func startFrom(from: KeynoteSlide!) // Start playing the presentation.
    optional func makeImageSlidesFiles(files: AnyObject!, setTitles: AnyObject!, master: KeynoteMasterSlide!) // Make a series of slides from a list of files.
    optional func stop() // Stop the presentation.
    optional func showSlideSwitcher() // Show the slide switcher in play mode
    optional func hideSlideSwitcher() // Hide the slide switcher in play mode
    optional func moveSlideSwitcherForward() // Move the slide switcher forward one slide
    optional func moveSlideSwitcherBackward() // Move the slide switcher backward one slide
    optional func cancelSlideSwitcher() // Hide the slide switcher without changing slides
    optional func acceptSlideSwitcher() // Hide the slide switcher, going to the slide it has selected
    optional func slides()
    optional func masterSlides()
    optional func id() // Document ID.
    optional var slideNumbersShowing: Int { get set } // Are the slide numbers displayed?
    optional var documentTheme: KeynoteTheme { get set } // The theme assigned to the document.
    optional var autoLoop: Int { get set } // Make the slideshow play repeatedly.
    optional var autoPlay: Int { get set } // Automatically play the presentation when opening the file.
    optional var autoRestart: Int { get set } // Restart the slideshow if it's inactive for the specified time
    optional var maximumIdleDuration: Int { get set } // Restart the slideshow if it's inactive for the specified time
    optional var currentSlide: KeynoteSlide { get } // The currently selected slide, or the slide that would display if the presentation was started.
    optional var height: Int { get set } // The height of the document (in points). Standard slide height = 768. Wide slide height = 1080.
    optional var width: Int { get set } // The width of the document (in points). Standard slide width = 1024. Wide slide width = 1920.
}
extension SBObject: KeynoteDocument {}

// MARK: KeynoteWindow
@objc public protocol KeynoteWindow: SBObjectProtocol, KeynoteGenericMethods {
    optional var name: Int { get } // The title of the window.
    optional func id() // The unique identifier of the window.
    optional var index: Int { get set } // The index of the window, ordered front to back.
    optional var bounds: Int { get set } // The bounding rectangle of the window.
    optional var closeable: Int { get } // Does the window have a close button?
    optional var miniaturizable: Int { get } // Does the window have a minimize button?
    optional var miniaturized: Int { get set } // Is the window minimized right now?
    optional var resizable: Int { get } // Can the window be resized?
    optional var visible: Int { get set } // Is the window visible right now?
    optional var zoomable: Int { get } // Does the window have a zoom button?
    optional var zoomed: Int { get set } // Is the window zoomed right now?
    optional var document: KeynoteDocument { get } // The document whose contents are displayed in the window.
}
extension SBObject: KeynoteWindow {}

// MARK: KeynoteTheme
@objc public protocol KeynoteTheme: SBObjectProtocol, KeynoteGenericMethods {
    optional func id() // The identifier used by the application.
    optional var name: Int { get }
}
extension SBObject: KeynoteTheme {}

// MARK: KeynoteRichText
@objc public protocol KeynoteRichText: SBObjectProtocol, KeynoteGenericMethods {
    optional func characters()
    optional func paragraphs()
    optional func words()
    optional func placeholderTexts()
    optional var color: Int { get set } // The color of the font. Expressed as an RGB value consisting of a list of three color values from 0 to 65535. ex: Blue = {0, 0, 65535}.
    optional var font: Int { get set } // The name of the font.  Can be the PostScript name, such as: “TimesNewRomanPS-ItalicMT”, or display name: “Times New Roman Italic”. TIP: Use the Font Book application get the information about a typeface.
    optional var size: Int { get set } // The size of the font.
}
extension SBObject: KeynoteRichText {}

// MARK: KeynoteCharacter
@objc public protocol KeynoteCharacter: KeynoteRichText {
}
extension SBObject: KeynoteCharacter {}

// MARK: KeynoteParagraph
@objc public protocol KeynoteParagraph: KeynoteRichText {
    optional func characters()
    optional func words()
    optional func placeholderTexts()
}
extension SBObject: KeynoteParagraph {}

// MARK: KeynoteWord
@objc public protocol KeynoteWord: KeynoteRichText {
    optional func characters()
}
extension SBObject: KeynoteWord {}

// MARK: KeynotePlaceholderText
@objc public protocol KeynotePlaceholderText: KeynoteRichText {
    optional var tag: Int { get set } // Its script tag.
}
extension SBObject: KeynotePlaceholderText {}

// MARK: KeynoteIWorkContainer
@objc public protocol KeynoteIWorkContainer: SBObjectProtocol, KeynoteGenericMethods {
    optional func audioClips()
    optional func charts()
    optional func images()
    optional func iWorkItems()
    optional func groups()
    optional func lines()
    optional func movies()
    optional func shapes()
    optional func tables()
    optional func textItems()
}
extension SBObject: KeynoteIWorkContainer {}

// MARK: KeynoteSlide
@objc public protocol KeynoteSlide: KeynoteIWorkContainer {
    optional var baseSlide: KeynoteMasterSlide { get set } // The master slide this slide is based upon
    optional var bodyShowing: Int { get set } // Is the default body text displayed?
    optional var skipped: Int { get set } // Is the slide skipped?
    optional var slideNumber: Int { get } // index of the slide in the document
    optional var titleShowing: Int { get set } // Is the default slide title displayed?
    optional var defaultBodyItem: KeynoteShape { get } // The default body container of the slide
    optional var defaultTitleItem: KeynoteShape { get } // The default title container of the slide
    optional var presenterNotes: KeynoteRichText { get set } // The presenter notes for the slide
    optional var transitionProperties: Int { get set } // The transition settings to apply to the slide.
    optional func addChartRowNames(rowNames: AnyObject!, columnNames: AnyObject!, data: AnyObject!, type: KeynoteLegacyChartType, groupBy: KeynoteLegacyChartGrouping) // Add a chart to a slide
}
extension SBObject: KeynoteSlide {}

// MARK: KeynoteMasterSlide
@objc public protocol KeynoteMasterSlide: KeynoteSlide {
    optional var name: Int { get } // The name of the master slide
}
extension SBObject: KeynoteMasterSlide {}

// MARK: KeynoteIWorkItem
@objc public protocol KeynoteIWorkItem: SBObjectProtocol, KeynoteGenericMethods {
    optional var height: Int { get set } // The height of the iWork item.
    optional var locked: Int { get set } // Whether the object is locked.
    optional var parent: KeynoteIWorkContainer { get } // The iWork container containing this iWork item.
    optional var position: Int { get set } // The horizontal and vertical coordinates of the top left point of the iWork item.
    optional var width: Int { get set } // The width of the iWork item.
}
extension SBObject: KeynoteIWorkItem {}

// MARK: KeynoteAudioClip
@objc public protocol KeynoteAudioClip: KeynoteIWorkItem {
    optional var fileName: AnyObject { get set } // The name of the audio file.
    optional var clipVolume: Int { get set } // The volume setting for the audio clip, from 0 (none) to 100 (full volume).
    optional var repetitionMethod: KeynotePlaybackRepetitionMethod { get set } // If or how the audio clip repeats.
}
extension SBObject: KeynoteAudioClip {}

// MARK: KeynoteShape
@objc public protocol KeynoteShape: KeynoteIWorkItem {
    optional var backgroundFillType: KeynoteItemFillOptions { get } // The background, if any, for the shape.
    optional var objectText: KeynoteRichText { get set } // The text contained within the shape.
    optional var reflectionShowing: Int { get set } // Is the iWork item displaying a reflection?
    optional var reflectionValue: Int { get set } // The percentage of reflection of the iWork item, from 0 (none) to 100 (full).
    optional var rotation: Int { get set } // The rotation of the iWork item, in degrees from 0 to 359.
    optional var opacity: Int { get set } // The opacity of the object, in percent.
}
extension SBObject: KeynoteShape {}

// MARK: KeynoteChart
@objc public protocol KeynoteChart: KeynoteIWorkItem {
}
extension SBObject: KeynoteChart {}

// MARK: KeynoteImage
@objc public protocol KeynoteImage: KeynoteIWorkItem {
    optional var objectDescription: Int { get set } // Text associated with the image, read aloud by VoiceOver.
    optional var file: Int { get } // The image file.
    optional var fileName: AnyObject { get set } // The name of the image file.
    optional var opacity: Int { get set } // The opacity of the object, in percent.
    optional var reflectionShowing: Int { get set } // Is the iWork item displaying a reflection?
    optional var reflectionValue: Int { get set } // The percentage of reflection of the iWork item, from 0 (none) to 100 (full).
    optional var rotation: Int { get set } // The rotation of the iWork item, in degrees from 0 to 359.
}
extension SBObject: KeynoteImage {}

// MARK: KeynoteGroup
@objc public protocol KeynoteGroup: KeynoteIWorkContainer {
    optional var height: Int { get set } // The height of the iWork item.
    optional var parent: KeynoteIWorkContainer { get } // The iWork container containing this iWork item.
    optional var position: Int { get set } // The horizontal and vertical coordinates of the top left point of the iWork item.
    optional var width: Int { get set } // The width of the iWork item.
    optional var rotation: Int { get set } // The rotation of the iWork item, in degrees from 0 to 359.
}
extension SBObject: KeynoteGroup {}

// MARK: KeynoteLine
@objc public protocol KeynoteLine: KeynoteIWorkItem {
    optional var endPoint: Int { get set } // A list of two numbers indicating the horizontal and vertical position of the line ending point.
    optional var reflectionShowing: Int { get set } // Is the iWork item displaying a reflection?
    optional var reflectionValue: Int { get set } // The percentage of reflection of the iWork item, from 0 (none) to 100 (full).
    optional var rotation: Int { get set } // The rotation of the iWork item, in degrees from 0 to 359.
    optional var startPoint: Int { get set } // A list of two numbers indicating the horizontal and vertical position of the line starting point.
}
extension SBObject: KeynoteLine {}

// MARK: KeynoteMovie
@objc public protocol KeynoteMovie: KeynoteIWorkItem {
    optional var fileName: AnyObject { get set } // The name of the movie file.
    optional var movieVolume: Int { get set } // The volume setting for the movie, from 0 (none) to 100 (full volume).
    optional var opacity: Int { get set } // The opacity of the object, in percent.
    optional var reflectionShowing: Int { get set } // Is the iWork item displaying a reflection?
    optional var reflectionValue: Int { get set } // The percentage of reflection of the iWork item, from 0 (none) to 100 (full).
    optional var repetitionMethod: KeynotePlaybackRepetitionMethod { get set } // If or how the movie repeats.
    optional var rotation: Int { get set } // The rotation of the iWork item, in degrees from 0 to 359.
}
extension SBObject: KeynoteMovie {}

// MARK: KeynoteTable
@objc public protocol KeynoteTable: KeynoteIWorkItem {
    optional func cells()
    optional func rows()
    optional func columns()
    optional func ranges()
    optional var name: Int { get set } // The item's name.
    optional var cellRange: KeynoteRange { get } // The range describing every cell in the table.
    optional var selectionRange: KeynoteRange { get set } // The cells currently selected in the table.
    optional var rowCount: Int { get set } // The number of rows in the table.
    optional var columnCount: Int { get set } // The number of columns in the table.
    optional var headerRowCount: Int { get set } // The number of header rows in the table.
    optional var headerColumnCount: Int { get set } // The number of header columns in the table.
    optional var footerRowCount: Int { get set } // The number of footer rows in the table.
    optional func sortBy(by: KeynoteColumn!, direction: KeynoteNMSD, inRows: KeynoteRange!) // Sort the rows of the table.
}
extension SBObject: KeynoteTable {}

// MARK: KeynoteTextItem
@objc public protocol KeynoteTextItem: KeynoteIWorkItem {
    optional var backgroundFillType: KeynoteItemFillOptions { get } // The background, if any, for the text item.
    optional var objectText: KeynoteRichText { get set } // The text contained within the text item.
    optional var opacity: Int { get set } // The opacity of the object, in percent.
    optional var reflectionShowing: Int { get set } // Is the iWork item displaying a reflection?
    optional var reflectionValue: Int { get set } // The percentage of reflection of the iWork item, from 0 (none) to 100 (full).
    optional var rotation: Int { get set } // The rotation of the iWork item, in degrees from 0 to 359.
}
extension SBObject: KeynoteTextItem {}

// MARK: KeynoteRange
@objc public protocol KeynoteRange: SBObjectProtocol, KeynoteGenericMethods {
    optional func cells()
    optional func columns()
    optional func rows()
    optional var fontName: Int { get set } // The font of the range's cells.
    optional var fontSize: Double { get set } // The font size of the range's cells.
    optional var format: KeynoteNMCT { get set } // The format of the range's cells.
    optional var alignment: KeynoteTAHT { get set } // The horizontal alignment of content in the range's cells.
    optional var name: Int { get } // The range's coordinates.
    optional var textColor: Int { get set } // The text color of the range's cells.
    optional var textWrap: Int { get set } // Whether text should wrap in the range's cells.
    optional var backgroundColor: Int { get set } // The background color of the range's cells.
    optional var verticalAlignment: KeynoteTAVT { get set } // The vertical alignment of content in the range's cells.
    optional func clear() // Clear the contents of a specified range of cells. Only content is removed; formatting and style are preserved.
    optional func merge() // Merge a specified range of cells.
    optional func unmerge() // Unmerge all merged cells in a specified range.
}
extension SBObject: KeynoteRange {}

// MARK: KeynoteCell
@objc public protocol KeynoteCell: KeynoteRange {
    optional var column: KeynoteColumn { get } // The cell's column.
    optional var row: KeynoteRow { get } // The cell's row.
    optional var value: AnyObject { get set } // The actual value in the cell, or missing value if the cell is empty.
    optional var formattedValue: Int { get } // The formatted value in the cell, or missing value if the cell is empty.
    optional var formula: Int { get } // The formula in the cell, as text, e.g. =SUM(40+2). If the cell does not contain a formula, returns missing value. To set the value of a cell to a formula as text, use the value property.
}
extension SBObject: KeynoteCell {}

// MARK: KeynoteRow
@objc public protocol KeynoteRow: KeynoteRange {
    optional var address: Int { get } // The row's index in the table (e.g., the second row has address 2).
    optional var height: Double { get set } // The height of the row.
}
extension SBObject: KeynoteRow {}

// MARK: KeynoteColumn
@objc public protocol KeynoteColumn: KeynoteRange {
    optional var address: Int { get } // The column's index in the table (e.g., the second column has address 2).
    optional var width: Double { get set } // The width of the column.
}
extension SBObject: KeynoteColumn {}

