/*
 * PowerPoint.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class PowerPointApplication, PowerPointDocument, PowerPointWindow, PowerPointCommandBarControl, PowerPointCommandBarButton, PowerPointCommandBarCombobox, PowerPointCommandBarPopup, PowerPointCommandBar, PowerPointDocumentProperty, PowerPointCustomDocumentProperty, PowerPointWebPageFont, PowerPointActionSetting, PowerPointAddIn, PowerPointAnimationBehavior, PowerPointAnimationPoint, PowerPointAnimationSettings, PowerPointApplication, PowerPointAutocorrectEntry, PowerPointAutocorrect, PowerPointBroadcast, PowerPointBulletFormat, PowerPointColorScheme, PowerPointColorsEffect, PowerPointCommandEffect, PowerPointCustomLayout, PowerPointDefaultWebOptions, PowerPointDesign, PowerPointDocumentWindow, PowerPointEffectInformation, PowerPointEffectParameters, PowerPointEffect, PowerPointFilterEffect, PowerPointFirstLetterException, PowerPointFont, PowerPointHeaderOrFooter, PowerPointHeadersAndFooters, PowerPointHyperlink, PowerPointMaster, PowerPointMotionEffect, PowerPointNamedSlideShow, PowerPointPageSetup, PowerPointPane, PowerPointParagraphFormat, PowerPointPlaySettings, PowerPointPlayer, PowerPointPreferences, PowerPointPresentation, PowerPointPresenterTool, PowerPointPresenterViewWindow, PowerPointPrintOptions, PowerPointPrintRange, PowerPointPropertyEffect, PowerPointRotatingEffect, PowerPointRulerLevel, PowerPointRuler, PowerPointSaveAsMovieSettings, PowerPointScaleEffect, PowerPointSectionProperties, PowerPointSelection, PowerPointSequence, PowerPointSetEffect, PowerPointSlideRange, PowerPointSlideShowSettings, PowerPointSlideShowTransition, PowerPointSlideShowView, PowerPointSlideShowWindow, PowerPointSlide, PowerPointSoundEffect, PowerPointTabStop, PowerPointTextStyleLevel, PowerPointTextStyle, PowerPointTimeline, PowerPointTiming, PowerPointTwoInitialCapsException, PowerPointView, PowerPointWebOptions, PowerPointAdjustment, PowerPointCalloutFormat, PowerPointConnectorFormat, PowerPointFillFormat, PowerPointGlowFormat, PowerPointGradientStop, PowerPointLineFormat, PowerPointLinkFormat, PowerPointOfficeTheme, PowerPointPictureFormat, PowerPointPlaceholderFormat, PowerPointReflectionFormat, PowerPointShadowFormat, PowerPointShapeRange, PowerPointShape, PowerPointCallout, PowerPointComment, PowerPointConnector, PowerPointLineShape, PowerPointMediaObject, PowerPointMedia2Object, PowerPointPicture, PowerPointPlaceHolder, PowerPointShapeTable, PowerPointSoftEdgeFormat, PowerPointTextBox, PowerPointTextColumn, PowerPointTextFrame, PowerPointThemeColorScheme, PowerPointThemeColor, PowerPointThemeEffectScheme, PowerPointThemeFontScheme, PowerPointThemeFont, PowerPointMajorThemeFont, PowerPointMinorThemeFont, PowerPointThreeDFormat, PowerPointWordArtFormat, PowerPointTextRange, PowerPointCharacter, PowerPointLine, PowerPointParagraph, PowerPointSentence, PowerPointTextFlow, PowerPointWord, PowerPointCell, PowerPointColumn, PowerPointRow, PowerPointTable;

enum PowerPointSaveOptions {
	PowerPointSaveOptionsYes = 'yes ' /* Save the file. */,
	PowerPointSaveOptionsNo = 'no  ' /* Do not save the file. */,
	PowerPointSaveOptionsAsk = 'ask ' /* Ask the user whether or not to save the file. */
};
typedef enum PowerPointSaveOptions PowerPointSaveOptions;

enum PowerPointPrintingErrorHandling {
	PowerPointPrintingErrorHandlingStandard = 'lwst' /* Standard PostScript error handling */,
	PowerPointPrintingErrorHandlingDetailed = 'lwdt' /* print a detailed report of PostScript errors */
};
typedef enum PowerPointPrintingErrorHandling PowerPointPrintingErrorHandling;

enum PowerPointMsoLineDashStyle {
	PowerPointMsoLineDashStyleLineDashStyleUnset = '\000\222\377\376',
	PowerPointMsoLineDashStyleLineDashStyleSolid = '\000\223\000\001',
	PowerPointMsoLineDashStyleLineDashStyleSquareDot = '\000\223\000\002',
	PowerPointMsoLineDashStyleLineDashStyleRoundDot = '\000\223\000\003',
	PowerPointMsoLineDashStyleLineDashStyleDash = '\000\223\000\004',
	PowerPointMsoLineDashStyleLineDashStyleDashDot = '\000\223\000\005',
	PowerPointMsoLineDashStyleLineDashStyleDashDotDot = '\000\223\000\006',
	PowerPointMsoLineDashStyleLineDashStyleLongDash = '\000\223\000\007',
	PowerPointMsoLineDashStyleLineDashStyleLongDashDot = '\000\223\000\010',
	PowerPointMsoLineDashStyleLineDashStyleLongDashDotDot = '\000\223\000\011',
	PowerPointMsoLineDashStyleLineDashStyleSystemDash = '\000\223\000\012',
	PowerPointMsoLineDashStyleLineDashStyleSystemDot = '\000\223\000\013',
	PowerPointMsoLineDashStyleLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum PowerPointMsoLineDashStyle PowerPointMsoLineDashStyle;

enum PowerPointMsoLineStyle {
	PowerPointMsoLineStyleLineStyleUnset = '\000\224\377\376',
	PowerPointMsoLineStyleSingleLine = '\000\225\000\001',
	PowerPointMsoLineStyleThinThinLine = '\000\225\000\002',
	PowerPointMsoLineStyleThinThickLine = '\000\225\000\003',
	PowerPointMsoLineStyleThickThinLine = '\000\225\000\004',
	PowerPointMsoLineStyleThickBetweenThinLine = '\000\225\000\005'
};
typedef enum PowerPointMsoLineStyle PowerPointMsoLineStyle;

enum PowerPointMsoArrowheadStyle {
	PowerPointMsoArrowheadStyleArrowheadStyleUnset = '\000\221\377\376',
	PowerPointMsoArrowheadStyleNoArrowhead = '\000\222\000\001',
	PowerPointMsoArrowheadStyleTriangleArrowhead = '\000\222\000\002',
	PowerPointMsoArrowheadStyleOpen_arrowhead = '\000\222\000\003',
	PowerPointMsoArrowheadStyleStealthArrowhead = '\000\222\000\004',
	PowerPointMsoArrowheadStyleDiamondArrowhead = '\000\222\000\005',
	PowerPointMsoArrowheadStyleOvalArrowhead = '\000\222\000\006'
};
typedef enum PowerPointMsoArrowheadStyle PowerPointMsoArrowheadStyle;

enum PowerPointMsoArrowheadWidth {
	PowerPointMsoArrowheadWidthArrowheadWidthUnset = '\000\220\377\376',
	PowerPointMsoArrowheadWidthNarrowWidthArrowhead = '\000\221\000\001',
	PowerPointMsoArrowheadWidthMediumWidthArrowhead = '\000\221\000\002',
	PowerPointMsoArrowheadWidthWideArrowhead = '\000\221\000\003'
};
typedef enum PowerPointMsoArrowheadWidth PowerPointMsoArrowheadWidth;

enum PowerPointMsoArrowheadLength {
	PowerPointMsoArrowheadLengthArrowheadLengthUnset = '\000\223\377\376',
	PowerPointMsoArrowheadLengthShortArrowhead = '\000\224\000\001',
	PowerPointMsoArrowheadLengthMediumArrowhead = '\000\224\000\002',
	PowerPointMsoArrowheadLengthLongArrowhead = '\000\224\000\003'
};
typedef enum PowerPointMsoArrowheadLength PowerPointMsoArrowheadLength;

enum PowerPointMsoFillType {
	PowerPointMsoFillTypeFillUnset = '\000c\377\376',
	PowerPointMsoFillTypeFillSolid = '\000d\000\001',
	PowerPointMsoFillTypeFillPatterned = '\000d\000\002',
	PowerPointMsoFillTypeFillGradient = '\000d\000\003',
	PowerPointMsoFillTypeFillTextured = '\000d\000\004',
	PowerPointMsoFillTypeFillBackground = '\000d\000\005',
	PowerPointMsoFillTypeFillPicture = '\000d\000\006'
};
typedef enum PowerPointMsoFillType PowerPointMsoFillType;

enum PowerPointMsoGradientStyle {
	PowerPointMsoGradientStyleGradientUnset = '\000d\377\376',
	PowerPointMsoGradientStyleHorizontalGradient = '\000e\000\001',
	PowerPointMsoGradientStyleVerticalGradient = '\000e\000\002',
	PowerPointMsoGradientStyleDiagonalUpGradient = '\000e\000\003',
	PowerPointMsoGradientStyleDiagonalDownGradient = '\000e\000\004',
	PowerPointMsoGradientStyleFromCornerGradient = '\000e\000\005',
	PowerPointMsoGradientStyleFromTitleGradient = '\000e\000\006',
	PowerPointMsoGradientStyleFromCenterGradient = '\000e\000\007'
};
typedef enum PowerPointMsoGradientStyle PowerPointMsoGradientStyle;

enum PowerPointMsoGradientColorType {
	PowerPointMsoGradientColorTypeGradientTypeUnset = '\003\357\377\376',
	PowerPointMsoGradientColorTypeSingleShadeGradientType = '\003\360\000\001',
	PowerPointMsoGradientColorTypeTwoColorsGradientType = '\003\360\000\002',
	PowerPointMsoGradientColorTypePresetColorsGradientType = '\003\360\000\003',
	PowerPointMsoGradientColorTypeMultiColorsGradientType = '\003\360\000\004'
};
typedef enum PowerPointMsoGradientColorType PowerPointMsoGradientColorType;

enum PowerPointMsoTextureType {
	PowerPointMsoTextureTypeTextureTypeTextureTypeUnset = '\003\360\377\376',
	PowerPointMsoTextureTypeTextureTypePresetTexture = '\003\361\000\001',
	PowerPointMsoTextureTypeTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum PowerPointMsoTextureType PowerPointMsoTextureType;

enum PowerPointMsoPresetTexture {
	PowerPointMsoPresetTexturePresetTextureUnset = '\000e\377\376',
	PowerPointMsoPresetTextureTexturePapyrus = '\000f\000\001',
	PowerPointMsoPresetTextureTextureCanvas = '\000f\000\002',
	PowerPointMsoPresetTextureTextureDenim = '\000f\000\003',
	PowerPointMsoPresetTextureTextureWovenMat = '\000f\000\004',
	PowerPointMsoPresetTextureTextureWaterDroplets = '\000f\000\005',
	PowerPointMsoPresetTextureTexturePaperBag = '\000f\000\006',
	PowerPointMsoPresetTextureTextureFishFossil = '\000f\000\007',
	PowerPointMsoPresetTextureTextureSand = '\000f\000\010',
	PowerPointMsoPresetTextureTextureGreenMarble = '\000f\000\011',
	PowerPointMsoPresetTextureTextureWhiteMarble = '\000f\000\012',
	PowerPointMsoPresetTextureTextureBrownMarble = '\000f\000\013',
	PowerPointMsoPresetTextureTextureGranite = '\000f\000\014',
	PowerPointMsoPresetTextureTextureNewsprint = '\000f\000\015',
	PowerPointMsoPresetTextureTextureRecycledPaper = '\000f\000\016',
	PowerPointMsoPresetTextureTextureParchment = '\000f\000\017',
	PowerPointMsoPresetTextureTextureStationery = '\000f\000\020',
	PowerPointMsoPresetTextureTextureBlueTissuePaper = '\000f\000\021',
	PowerPointMsoPresetTextureTexturePinkTissuePaper = '\000f\000\022',
	PowerPointMsoPresetTextureTexturePurpleMesh = '\000f\000\023',
	PowerPointMsoPresetTextureTextureBouquet = '\000f\000\024',
	PowerPointMsoPresetTextureTextureCork = '\000f\000\025',
	PowerPointMsoPresetTextureTextureWalnut = '\000f\000\026',
	PowerPointMsoPresetTextureTextureOak = '\000f\000\027',
	PowerPointMsoPresetTextureTextureMediumWood = '\000f\000\030'
};
typedef enum PowerPointMsoPresetTexture PowerPointMsoPresetTexture;

enum PowerPointMsoPatternType {
	PowerPointMsoPatternTypePatternUnset = '\000f\377\376',
	PowerPointMsoPatternTypeFivePercentPattern = '\000g\000\001',
	PowerPointMsoPatternTypeTenPercentPattern = '\000g\000\002',
	PowerPointMsoPatternTypeTwentyPercentPattern = '\000g\000\003',
	PowerPointMsoPatternTypeTwentyFivePercentPattern = '\000g\000\004',
	PowerPointMsoPatternTypeThirtyPercentPattern = '\000g\000\005',
	PowerPointMsoPatternTypeFortyPercentPattern = '\000g\000\006',
	PowerPointMsoPatternTypeFiftyPercentPattern = '\000g\000\007',
	PowerPointMsoPatternTypeSixtyPercentPattern = '\000g\000\010',
	PowerPointMsoPatternTypeSeventyPercentPattern = '\000g\000\011',
	PowerPointMsoPatternTypeSeventyFivePercentPattern = '\000g\000\012',
	PowerPointMsoPatternTypeEightyPercentPattern = '\000g\000\013',
	PowerPointMsoPatternTypeNinetyPercentPattern = '\000g\000\014',
	PowerPointMsoPatternTypeDarkHorizontalPattern = '\000g\000\015',
	PowerPointMsoPatternTypeDarkVerticalPattern = '\000g\000\016',
	PowerPointMsoPatternTypeDarkDownwardDiagonalPattern = '\000g\000\017',
	PowerPointMsoPatternTypeDarkUpwardDiagonalPattern = '\000g\000\020',
	PowerPointMsoPatternTypeSmallCheckerBoardPattern = '\000g\000\021',
	PowerPointMsoPatternTypeTrellisPattern = '\000g\000\022',
	PowerPointMsoPatternTypeLightHorizontalPattern = '\000g\000\023',
	PowerPointMsoPatternTypeLightVerticalPattern = '\000g\000\024',
	PowerPointMsoPatternTypeLightDownwardDiagonalPattern = '\000g\000\025',
	PowerPointMsoPatternTypeLightUpwardDiagonalPattern = '\000g\000\026',
	PowerPointMsoPatternTypeSmallGridPattern = '\000g\000\027',
	PowerPointMsoPatternTypeDottedDiamondPattern = '\000g\000\030',
	PowerPointMsoPatternTypeWideDownwardDiagonal = '\000g\000\031',
	PowerPointMsoPatternTypeWideUpwardDiagonalPattern = '\000g\000\032',
	PowerPointMsoPatternTypeDashedUpwardDiagonalPattern = '\000g\000\033',
	PowerPointMsoPatternTypeDashedDownwardDiagonalPattern = '\000g\000\034',
	PowerPointMsoPatternTypeNarrowVerticalPattern = '\000g\000\035',
	PowerPointMsoPatternTypeNarrowHorizontalPattern = '\000g\000\036',
	PowerPointMsoPatternTypeDashedVerticalPattern = '\000g\000\037',
	PowerPointMsoPatternTypeDashedHorizontalPattern = '\000g\000 ',
	PowerPointMsoPatternTypeLargeConfettiPattern = '\000g\000!',
	PowerPointMsoPatternTypeLargeGridPattern = '\000g\000\"',
	PowerPointMsoPatternTypeHorizontalBrickPattern = '\000g\000#',
	PowerPointMsoPatternTypeLargeCheckerBoardPattern = '\000g\000$',
	PowerPointMsoPatternTypeSmallConfettiPattern = '\000g\000%',
	PowerPointMsoPatternTypeZigZagPattern = '\000g\000&',
	PowerPointMsoPatternTypeSolidDiamondPattern = '\000g\000\'',
	PowerPointMsoPatternTypeDiagonalBrickPattern = '\000g\000(',
	PowerPointMsoPatternTypeOutlinedDiamondPattern = '\000g\000)',
	PowerPointMsoPatternTypePlaidPattern = '\000g\000*',
	PowerPointMsoPatternTypeSpherePattern = '\000g\000+',
	PowerPointMsoPatternTypeWeavePattern = '\000g\000,',
	PowerPointMsoPatternTypeDottedGridPattern = '\000g\000-',
	PowerPointMsoPatternTypeDivotPattern = '\000g\000.',
	PowerPointMsoPatternTypeShinglePattern = '\000g\000/',
	PowerPointMsoPatternTypeWavePattern = '\000g\0000',
	PowerPointMsoPatternTypeHorizontalPattern = '\000g\0001',
	PowerPointMsoPatternTypeVerticalPattern = '\000g\0002',
	PowerPointMsoPatternTypeCrossPattern = '\000g\0003',
	PowerPointMsoPatternTypeDownwardDiagonalPattern = '\000g\0004',
	PowerPointMsoPatternTypeUpwardDiagonalPattern = '\000g\0005',
	PowerPointMsoPatternTypeDiagonalCrossPattern = '\000g\0005'
};
typedef enum PowerPointMsoPatternType PowerPointMsoPatternType;

enum PowerPointMsoPresetGradientType {
	PowerPointMsoPresetGradientTypePresetGradientUnset = '\000g\377\376',
	PowerPointMsoPresetGradientTypeGradientEarlySunset = '\000h\000\001',
	PowerPointMsoPresetGradientTypeGradientLateSunset = '\000h\000\002',
	PowerPointMsoPresetGradientTypeGradientNightfall = '\000h\000\003',
	PowerPointMsoPresetGradientTypeGradientDaybreak = '\000h\000\004',
	PowerPointMsoPresetGradientTypeGradientHorizon = '\000h\000\005',
	PowerPointMsoPresetGradientTypeGradientDesert = '\000h\000\006',
	PowerPointMsoPresetGradientTypeGradientOcean = '\000h\000\007',
	PowerPointMsoPresetGradientTypeGradientCalmWater = '\000h\000\010',
	PowerPointMsoPresetGradientTypeGradientFire = '\000h\000\011',
	PowerPointMsoPresetGradientTypeGradientFog = '\000h\000\012',
	PowerPointMsoPresetGradientTypeGradientMoss = '\000h\000\013',
	PowerPointMsoPresetGradientTypeGradientPeacock = '\000h\000\014',
	PowerPointMsoPresetGradientTypeGradientWheat = '\000h\000\015',
	PowerPointMsoPresetGradientTypeGradientParchment = '\000h\000\016',
	PowerPointMsoPresetGradientTypeGradientMahogany = '\000h\000\017',
	PowerPointMsoPresetGradientTypeGradientRainbow = '\000h\000\020',
	PowerPointMsoPresetGradientTypeGradientRainbow2 = '\000h\000\021',
	PowerPointMsoPresetGradientTypeGradientGold = '\000h\000\022',
	PowerPointMsoPresetGradientTypeGradientGold2 = '\000h\000\023',
	PowerPointMsoPresetGradientTypeGradientBrass = '\000h\000\024',
	PowerPointMsoPresetGradientTypeGradientChrome = '\000h\000\025',
	PowerPointMsoPresetGradientTypeGradientChrome2 = '\000h\000\026',
	PowerPointMsoPresetGradientTypeGradientSilver = '\000h\000\027',
	PowerPointMsoPresetGradientTypeGradientSapphire = '\000h\000\030'
};
typedef enum PowerPointMsoPresetGradientType PowerPointMsoPresetGradientType;

enum PowerPointMsoShadowType {
	PowerPointMsoShadowTypeShadowUnset = '\003_\377\376',
	PowerPointMsoShadowTypeShadow1 = '\003`\000\001',
	PowerPointMsoShadowTypeShadow2 = '\003`\000\002',
	PowerPointMsoShadowTypeShadow3 = '\003`\000\003',
	PowerPointMsoShadowTypeShadow4 = '\003`\000\004',
	PowerPointMsoShadowTypeShadow5 = '\003`\000\005',
	PowerPointMsoShadowTypeShadow6 = '\003`\000\006',
	PowerPointMsoShadowTypeShadow7 = '\003`\000\007',
	PowerPointMsoShadowTypeShadow8 = '\003`\000\010',
	PowerPointMsoShadowTypeShadow9 = '\003`\000\011',
	PowerPointMsoShadowTypeShadow10 = '\003`\000\012',
	PowerPointMsoShadowTypeShadow11 = '\003`\000\013',
	PowerPointMsoShadowTypeShadow12 = '\003`\000\014',
	PowerPointMsoShadowTypeShadow13 = '\003`\000\015',
	PowerPointMsoShadowTypeShadow14 = '\003`\000\016',
	PowerPointMsoShadowTypeShadow15 = '\003`\000\017',
	PowerPointMsoShadowTypeShadow16 = '\003`\000\020',
	PowerPointMsoShadowTypeShadow17 = '\003`\000\021',
	PowerPointMsoShadowTypeShadow18 = '\003`\000\022',
	PowerPointMsoShadowTypeShadow19 = '\003`\000\023',
	PowerPointMsoShadowTypeShadow20 = '\003`\000\024',
	PowerPointMsoShadowTypeShadow21 = '\003`\000\025',
	PowerPointMsoShadowTypeShadow22 = '\003`\000\026',
	PowerPointMsoShadowTypeShadow23 = '\003`\000\027',
	PowerPointMsoShadowTypeShadow24 = '\003`\000\030',
	PowerPointMsoShadowTypeShadow25 = '\003`\000\031',
	PowerPointMsoShadowTypeShadow26 = '\003`\000\032',
	PowerPointMsoShadowTypeShadow27 = '\003`\000\033',
	PowerPointMsoShadowTypeShadow28 = '\003`\000\034',
	PowerPointMsoShadowTypeShadow29 = '\003`\000\035',
	PowerPointMsoShadowTypeShadow30 = '\003`\000\036',
	PowerPointMsoShadowTypeShadow31 = '\003`\000\037',
	PowerPointMsoShadowTypeShadow32 = '\003`\000 ',
	PowerPointMsoShadowTypeShadow33 = '\003`\000!',
	PowerPointMsoShadowTypeShadow34 = '\003`\000\"',
	PowerPointMsoShadowTypeShadow35 = '\003`\000#',
	PowerPointMsoShadowTypeShadow36 = '\003`\000$',
	PowerPointMsoShadowTypeShadow37 = '\003`\000%',
	PowerPointMsoShadowTypeShadow38 = '\003`\000&',
	PowerPointMsoShadowTypeShadow39 = '\003`\000\'',
	PowerPointMsoShadowTypeShadow40 = '\003`\000(',
	PowerPointMsoShadowTypeShadow41 = '\003`\000)',
	PowerPointMsoShadowTypeShadow42 = '\003`\000*',
	PowerPointMsoShadowTypeShadow43 = '\003`\000+'
};
typedef enum PowerPointMsoShadowType PowerPointMsoShadowType;

enum PowerPointMsoPresetTextEffect {
	PowerPointMsoPresetTextEffectWordartFormatUnset = '\003\361\377\376',
	PowerPointMsoPresetTextEffectWordartFormat1 = '\003\362\000\000',
	PowerPointMsoPresetTextEffectWordartFormat2 = '\003\362\000\001',
	PowerPointMsoPresetTextEffectWordartFormat3 = '\003\362\000\002',
	PowerPointMsoPresetTextEffectWordartFormat4 = '\003\362\000\003',
	PowerPointMsoPresetTextEffectWordartFormat5 = '\003\362\000\004',
	PowerPointMsoPresetTextEffectWordartFormat6 = '\003\362\000\005',
	PowerPointMsoPresetTextEffectWordartFormat7 = '\003\362\000\006',
	PowerPointMsoPresetTextEffectWordartFormat8 = '\003\362\000\007',
	PowerPointMsoPresetTextEffectWordartFormat9 = '\003\362\000\010',
	PowerPointMsoPresetTextEffectWordartFormat10 = '\003\362\000\011',
	PowerPointMsoPresetTextEffectWordartFormat11 = '\003\362\000\012',
	PowerPointMsoPresetTextEffectWordartFormat12 = '\003\362\000\013',
	PowerPointMsoPresetTextEffectWordartFormat13 = '\003\362\000\014',
	PowerPointMsoPresetTextEffectWordartFormat14 = '\003\362\000\015',
	PowerPointMsoPresetTextEffectWordartFormat15 = '\003\362\000\016',
	PowerPointMsoPresetTextEffectWordartFormat16 = '\003\362\000\017',
	PowerPointMsoPresetTextEffectWordartFormat17 = '\003\362\000\020',
	PowerPointMsoPresetTextEffectWordartFormat18 = '\003\362\000\021',
	PowerPointMsoPresetTextEffectWordartFormat19 = '\003\362\000\022',
	PowerPointMsoPresetTextEffectWordartFormat20 = '\003\362\000\023',
	PowerPointMsoPresetTextEffectWordartFormat21 = '\003\362\000\024',
	PowerPointMsoPresetTextEffectWordartFormat22 = '\003\362\000\025',
	PowerPointMsoPresetTextEffectWordartFormat23 = '\003\362\000\026',
	PowerPointMsoPresetTextEffectWordartFormat24 = '\003\362\000\027',
	PowerPointMsoPresetTextEffectWordartFormat25 = '\003\362\000\030',
	PowerPointMsoPresetTextEffectWordartFormat26 = '\003\362\000\031',
	PowerPointMsoPresetTextEffectWordartFormat27 = '\003\362\000\032',
	PowerPointMsoPresetTextEffectWordartFormat28 = '\003\362\000\033',
	PowerPointMsoPresetTextEffectWordartFormat29 = '\003\362\000\034',
	PowerPointMsoPresetTextEffectWordartFormat30 = '\003\362\000\035',
	PowerPointMsoPresetTextEffectWordartFormat31 = '\003\362\000\036',
	PowerPointMsoPresetTextEffectWordartFormat32 = '\003\362\000\037',
	PowerPointMsoPresetTextEffectWordartFormat33 = '\003\362\000 ',
	PowerPointMsoPresetTextEffectWordartFormat34 = '\003\362\000!',
	PowerPointMsoPresetTextEffectWordartFormat35 = '\003\362\000\"',
	PowerPointMsoPresetTextEffectWordartFormat36 = '\003\362\000#',
	PowerPointMsoPresetTextEffectWordartFormat37 = '\003\362\000$',
	PowerPointMsoPresetTextEffectWordartFormat38 = '\003\362\000%',
	PowerPointMsoPresetTextEffectWordartFormat39 = '\003\362\000&',
	PowerPointMsoPresetTextEffectWordartFormat40 = '\003\362\000\'',
	PowerPointMsoPresetTextEffectWordartFormat41 = '\003\362\000(',
	PowerPointMsoPresetTextEffectWordartFormat42 = '\003\362\000)',
	PowerPointMsoPresetTextEffectWordartFormat43 = '\003\362\000*',
	PowerPointMsoPresetTextEffectWordartFormat44 = '\003\362\000+',
	PowerPointMsoPresetTextEffectWordartFormat45 = '\003\362\000,',
	PowerPointMsoPresetTextEffectWordartFormat46 = '\003\362\000-',
	PowerPointMsoPresetTextEffectWordartFormat47 = '\003\362\000.',
	PowerPointMsoPresetTextEffectWordartFormat48 = '\003\362\000/',
	PowerPointMsoPresetTextEffectWordartFormat49 = '\003\362\0000',
	PowerPointMsoPresetTextEffectWordartFormat50 = '\003\362\0001'
};
typedef enum PowerPointMsoPresetTextEffect PowerPointMsoPresetTextEffect;

enum PowerPointMsoPresetTextEffectShape {
	PowerPointMsoPresetTextEffectShapeTextEffectShapeUnset = '\000\227\377\376',
	PowerPointMsoPresetTextEffectShapePlainText = '\000\230\000\001',
	PowerPointMsoPresetTextEffectShapeStop = '\000\230\000\002',
	PowerPointMsoPresetTextEffectShapeTriangleUp = '\000\230\000\003',
	PowerPointMsoPresetTextEffectShapeTriangleDown = '\000\230\000\004',
	PowerPointMsoPresetTextEffectShapeChevronUp = '\000\230\000\005',
	PowerPointMsoPresetTextEffectShapeChevronDown = '\000\230\000\006',
	PowerPointMsoPresetTextEffectShapeRingInside = '\000\230\000\007',
	PowerPointMsoPresetTextEffectShapeRingOutside = '\000\230\000\010',
	PowerPointMsoPresetTextEffectShapeArchUpCurve = '\000\230\000\011',
	PowerPointMsoPresetTextEffectShapeArchDownCurve = '\000\230\000\012',
	PowerPointMsoPresetTextEffectShapeCircleCurve = '\000\230\000\013',
	PowerPointMsoPresetTextEffectShapeButtonCurve = '\000\230\000\014',
	PowerPointMsoPresetTextEffectShapeArchUpPour = '\000\230\000\015',
	PowerPointMsoPresetTextEffectShapeArchDownPour = '\000\230\000\016',
	PowerPointMsoPresetTextEffectShapeCirclePour = '\000\230\000\017',
	PowerPointMsoPresetTextEffectShapeButtonPour = '\000\230\000\020',
	PowerPointMsoPresetTextEffectShapeCurveUp = '\000\230\000\021',
	PowerPointMsoPresetTextEffectShapeCurveDown = '\000\230\000\022',
	PowerPointMsoPresetTextEffectShapeCanUp = '\000\230\000\023',
	PowerPointMsoPresetTextEffectShapeCanDown = '\000\230\000\024',
	PowerPointMsoPresetTextEffectShapeWave1 = '\000\230\000\025',
	PowerPointMsoPresetTextEffectShapeWave2 = '\000\230\000\026',
	PowerPointMsoPresetTextEffectShapeDoubleWave1 = '\000\230\000\027',
	PowerPointMsoPresetTextEffectShapeDoubleWave2 = '\000\230\000\030',
	PowerPointMsoPresetTextEffectShapeInflate = '\000\230\000\031',
	PowerPointMsoPresetTextEffectShapeDeflate = '\000\230\000\032',
	PowerPointMsoPresetTextEffectShapeInflateBottom = '\000\230\000\033',
	PowerPointMsoPresetTextEffectShapeDeflateBottom = '\000\230\000\034',
	PowerPointMsoPresetTextEffectShapeInflateTop = '\000\230\000\035',
	PowerPointMsoPresetTextEffectShapeDeflateTop = '\000\230\000\036',
	PowerPointMsoPresetTextEffectShapeDeflateInflate = '\000\230\000\037',
	PowerPointMsoPresetTextEffectShapeDeflateInflateDeflate = '\000\230\000 ',
	PowerPointMsoPresetTextEffectShapeFadeRight = '\000\230\000!',
	PowerPointMsoPresetTextEffectShapeFadeLeft = '\000\230\000\"',
	PowerPointMsoPresetTextEffectShapeFadeUp = '\000\230\000#',
	PowerPointMsoPresetTextEffectShapeFadeDown = '\000\230\000$',
	PowerPointMsoPresetTextEffectShapeSlantUp = '\000\230\000%',
	PowerPointMsoPresetTextEffectShapeSlantDown = '\000\230\000&',
	PowerPointMsoPresetTextEffectShapeCascadeUp = '\000\230\000\'',
	PowerPointMsoPresetTextEffectShapeCascadeDown = '\000\230\000('
};
typedef enum PowerPointMsoPresetTextEffectShape PowerPointMsoPresetTextEffectShape;

enum PowerPointMsoTextEffectAlignment {
	PowerPointMsoTextEffectAlignmentTextEffectAlignmentUnset = '\000\226\377\376',
	PowerPointMsoTextEffectAlignmentLeftTextEffectAlignment = '\000\227\000\001',
	PowerPointMsoTextEffectAlignmentCenteredTextEffectAlignment = '\000\227\000\002',
	PowerPointMsoTextEffectAlignmentRightTextEffectAlignment = '\000\227\000\003',
	PowerPointMsoTextEffectAlignmentJustifyTextEffectAlignment = '\000\227\000\004',
	PowerPointMsoTextEffectAlignmentWordJustifyTextEffectAlignment = '\000\227\000\005',
	PowerPointMsoTextEffectAlignmentStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum PowerPointMsoTextEffectAlignment PowerPointMsoTextEffectAlignment;

enum PowerPointMsoPresetLightingDirection {
	PowerPointMsoPresetLightingDirectionPresetLightingDirectionUnset = '\000\233\377\376',
	PowerPointMsoPresetLightingDirectionLightFromTopLeft = '\000\234\000\001',
	PowerPointMsoPresetLightingDirectionLightFromTop = '\000\234\000\002',
	PowerPointMsoPresetLightingDirectionLightFromTopRight = '\000\234\000\003',
	PowerPointMsoPresetLightingDirectionLightFromLeft = '\000\234\000\004',
	PowerPointMsoPresetLightingDirectionLightFromNone = '\000\234\000\005',
	PowerPointMsoPresetLightingDirectionLightFromRight = '\000\234\000\006',
	PowerPointMsoPresetLightingDirectionLightFromBottomLeft = '\000\234\000\007',
	PowerPointMsoPresetLightingDirectionLightFromBottom = '\000\234\000\010',
	PowerPointMsoPresetLightingDirectionLightFromBottomRight = '\000\234\000\011'
};
typedef enum PowerPointMsoPresetLightingDirection PowerPointMsoPresetLightingDirection;

enum PowerPointMsoPresetLightingSoftness {
	PowerPointMsoPresetLightingSoftnessLightingSoftnessUnset = '\000\234\377\376',
	PowerPointMsoPresetLightingSoftnessLightingDim = '\000\235\000\001',
	PowerPointMsoPresetLightingSoftnessLightingNormal = '\000\235\000\002',
	PowerPointMsoPresetLightingSoftnessLightingBright = '\000\235\000\003'
};
typedef enum PowerPointMsoPresetLightingSoftness PowerPointMsoPresetLightingSoftness;

enum PowerPointMsoPresetMaterial {
	PowerPointMsoPresetMaterialPresetMaterialUnset = '\000\235\377\376',
	PowerPointMsoPresetMaterialMatte = '\000\236\000\001',
	PowerPointMsoPresetMaterialPlastic = '\000\236\000\002',
	PowerPointMsoPresetMaterialMetal = '\000\236\000\003',
	PowerPointMsoPresetMaterialWireframe = '\000\236\000\004',
	PowerPointMsoPresetMaterialMatte2 = '\000\236\000\005',
	PowerPointMsoPresetMaterialPlastic2 = '\000\236\000\006',
	PowerPointMsoPresetMaterialMetal2 = '\000\236\000\007',
	PowerPointMsoPresetMaterialWarmMatte = '\000\236\000\010',
	PowerPointMsoPresetMaterialTranslucentPowder = '\000\236\000\011',
	PowerPointMsoPresetMaterialPowder = '\000\236\000\012',
	PowerPointMsoPresetMaterialDarkEdge = '\000\236\000\013',
	PowerPointMsoPresetMaterialSoftEdge = '\000\236\000\014',
	PowerPointMsoPresetMaterialMaterialClear = '\000\236\000\015',
	PowerPointMsoPresetMaterialFlat = '\000\236\000\016',
	PowerPointMsoPresetMaterialSoftMetal = '\000\236\000\017'
};
typedef enum PowerPointMsoPresetMaterial PowerPointMsoPresetMaterial;

enum PowerPointMsoPresetExtrusionDirection {
	PowerPointMsoPresetExtrusionDirectionPresetExtrusionDirectionUnset = '\000\231\377\376',
	PowerPointMsoPresetExtrusionDirectionExtrudeBottomRight = '\000\232\000\001',
	PowerPointMsoPresetExtrusionDirectionExtrudeBottom = '\000\232\000\002',
	PowerPointMsoPresetExtrusionDirectionExtrudeBottomLeft = '\000\232\000\003',
	PowerPointMsoPresetExtrusionDirectionExtrudeRight = '\000\232\000\004',
	PowerPointMsoPresetExtrusionDirectionExtrudeNone = '\000\232\000\005',
	PowerPointMsoPresetExtrusionDirectionExtrudeLeft = '\000\232\000\006',
	PowerPointMsoPresetExtrusionDirectionExtrudeTopRight = '\000\232\000\007',
	PowerPointMsoPresetExtrusionDirectionExtrudeTop = '\000\232\000\010',
	PowerPointMsoPresetExtrusionDirectionExtrudeTopLeft = '\000\232\000\011'
};
typedef enum PowerPointMsoPresetExtrusionDirection PowerPointMsoPresetExtrusionDirection;

enum PowerPointMsoPresetThreeDFormat {
	PowerPointMsoPresetThreeDFormatPresetThreeDFormatUnset = '\000\230\377\376',
	PowerPointMsoPresetThreeDFormatFormat1 = '\000\231\000\001',
	PowerPointMsoPresetThreeDFormatFormat2 = '\000\231\000\002',
	PowerPointMsoPresetThreeDFormatFormat3 = '\000\231\000\003',
	PowerPointMsoPresetThreeDFormatFormat4 = '\000\231\000\004',
	PowerPointMsoPresetThreeDFormatFormat5 = '\000\231\000\005',
	PowerPointMsoPresetThreeDFormatFormat6 = '\000\231\000\006',
	PowerPointMsoPresetThreeDFormatFormat7 = '\000\231\000\007',
	PowerPointMsoPresetThreeDFormatFormat8 = '\000\231\000\010',
	PowerPointMsoPresetThreeDFormatFormat9 = '\000\231\000\011',
	PowerPointMsoPresetThreeDFormatFormat10 = '\000\231\000\012',
	PowerPointMsoPresetThreeDFormatFormat11 = '\000\231\000\013',
	PowerPointMsoPresetThreeDFormatFormat12 = '\000\231\000\014',
	PowerPointMsoPresetThreeDFormatFormat13 = '\000\231\000\015',
	PowerPointMsoPresetThreeDFormatFormat14 = '\000\231\000\016',
	PowerPointMsoPresetThreeDFormatFormat15 = '\000\231\000\017',
	PowerPointMsoPresetThreeDFormatFormat16 = '\000\231\000\020',
	PowerPointMsoPresetThreeDFormatFormat17 = '\000\231\000\021',
	PowerPointMsoPresetThreeDFormatFormat18 = '\000\231\000\022',
	PowerPointMsoPresetThreeDFormatFormat19 = '\000\231\000\023',
	PowerPointMsoPresetThreeDFormatFormat20 = '\000\231\000\024'
};
typedef enum PowerPointMsoPresetThreeDFormat PowerPointMsoPresetThreeDFormat;

enum PowerPointMsoExtrusionColorType {
	PowerPointMsoExtrusionColorTypeExtrusionColorUnset = '\000\232\377\376',
	PowerPointMsoExtrusionColorTypeExtrusionColorAutomatic = '\000\233\000\001',
	PowerPointMsoExtrusionColorTypeExtrusionColorCustom = '\000\233\000\002'
};
typedef enum PowerPointMsoExtrusionColorType PowerPointMsoExtrusionColorType;

enum PowerPointMsoConnectorType {
	PowerPointMsoConnectorTypeConnectorTypeUnset = '\000h\377\376',
	PowerPointMsoConnectorTypeStraight = '\000i\000\001',
	PowerPointMsoConnectorTypeElbow = '\000i\000\002',
	PowerPointMsoConnectorTypeCurve = '\000i\000\003'
};
typedef enum PowerPointMsoConnectorType PowerPointMsoConnectorType;

enum PowerPointMsoHorizontalAnchor {
	PowerPointMsoHorizontalAnchorHorizontalAnchorUnset = '\000\236\377\376',
	PowerPointMsoHorizontalAnchorHorizontalAnchorNone = '\000\237\000\001',
	PowerPointMsoHorizontalAnchorHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum PowerPointMsoHorizontalAnchor PowerPointMsoHorizontalAnchor;

enum PowerPointMsoVerticalAnchor {
	PowerPointMsoVerticalAnchorVerticalAnchorUnset = '\000\237\377\376',
	PowerPointMsoVerticalAnchorAnchorTop = '\000\240\000\001',
	PowerPointMsoVerticalAnchorAnchorTopBaseline = '\000\240\000\002',
	PowerPointMsoVerticalAnchorAnchorMiddle = '\000\240\000\003',
	PowerPointMsoVerticalAnchorAnchorBottom = '\000\240\000\004',
	PowerPointMsoVerticalAnchorAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum PowerPointMsoVerticalAnchor PowerPointMsoVerticalAnchor;

enum PowerPointMsoAutoShapeType {
	PowerPointMsoAutoShapeTypeAutoshapeShapeTypeUnset = '\000i\377\376',
	PowerPointMsoAutoShapeTypeAutoshapeRectangle = '\000j\000\001',
	PowerPointMsoAutoShapeTypeAutoshapeParallelogram = '\000j\000\002',
	PowerPointMsoAutoShapeTypeAutoshapeTrapezoid = '\000j\000\003',
	PowerPointMsoAutoShapeTypeAutoshapeDiamond = '\000j\000\004',
	PowerPointMsoAutoShapeTypeAutoshapeRoundedRectangle = '\000j\000\005',
	PowerPointMsoAutoShapeTypeAutoshapeOctagon = '\000j\000\006',
	PowerPointMsoAutoShapeTypeAutoshapeIsoscelesTriangle = '\000j\000\007',
	PowerPointMsoAutoShapeTypeAutoshapeRightTriangle = '\000j\000\010',
	PowerPointMsoAutoShapeTypeAutoshapeOval = '\000j\000\011',
	PowerPointMsoAutoShapeTypeAutoshapeHexagon = '\000j\000\012',
	PowerPointMsoAutoShapeTypeAutoshapeCross = '\000j\000\013',
	PowerPointMsoAutoShapeTypeAutoshapeRegularPentagon = '\000j\000\014',
	PowerPointMsoAutoShapeTypeAutoshapeCan = '\000j\000\015',
	PowerPointMsoAutoShapeTypeAutoshapeCube = '\000j\000\016',
	PowerPointMsoAutoShapeTypeAutoshapeBevel = '\000j\000\017',
	PowerPointMsoAutoShapeTypeAutoshapeFoldedCorner = '\000j\000\020',
	PowerPointMsoAutoShapeTypeAutoshapeSmileyFace = '\000j\000\021',
	PowerPointMsoAutoShapeTypeAutoshapeDonut = '\000j\000\022',
	PowerPointMsoAutoShapeTypeAutoshapeNoSymbol = '\000j\000\023',
	PowerPointMsoAutoShapeTypeAutoshapeBlockArc = '\000j\000\024',
	PowerPointMsoAutoShapeTypeAutoshapeHeart = '\000j\000\025',
	PowerPointMsoAutoShapeTypeAutoshapeLightningBolt = '\000j\000\026',
	PowerPointMsoAutoShapeTypeAutoshapeSun = '\000j\000\027',
	PowerPointMsoAutoShapeTypeAutoshapeMoon = '\000j\000\030',
	PowerPointMsoAutoShapeTypeAutoshapeArc = '\000j\000\031',
	PowerPointMsoAutoShapeTypeAutoshapeDoubleBracket = '\000j\000\032',
	PowerPointMsoAutoShapeTypeAutoshapeDoubleBrace = '\000j\000\033',
	PowerPointMsoAutoShapeTypeAutoshapePlaque = '\000j\000\034',
	PowerPointMsoAutoShapeTypeAutoshapeLeftBracket = '\000j\000\035',
	PowerPointMsoAutoShapeTypeAutoshapeRightBracket = '\000j\000\036',
	PowerPointMsoAutoShapeTypeAutoshapeLeftBrace = '\000j\000\037',
	PowerPointMsoAutoShapeTypeAutoshapeRightBrace = '\000j\000 ',
	PowerPointMsoAutoShapeTypeAutoshapeRightArrow = '\000j\000!',
	PowerPointMsoAutoShapeTypeAutoshapeLeftArrow = '\000j\000\"',
	PowerPointMsoAutoShapeTypeAutoshapeUpArrow = '\000j\000#',
	PowerPointMsoAutoShapeTypeAutoshapeDownArrow = '\000j\000$',
	PowerPointMsoAutoShapeTypeAutoshapeLeftRightArrow = '\000j\000%',
	PowerPointMsoAutoShapeTypeAutoshapeUpDownArrow = '\000j\000&',
	PowerPointMsoAutoShapeTypeAutoshapeQuadArrow = '\000j\000\'',
	PowerPointMsoAutoShapeTypeAutoshapeLeftRightUpArrow = '\000j\000(',
	PowerPointMsoAutoShapeTypeAutoshapeBentArrow = '\000j\000)',
	PowerPointMsoAutoShapeTypeAutoshapeUTurnArrow = '\000j\000*',
	PowerPointMsoAutoShapeTypeAutoshapeLeftUpArrow = '\000j\000+',
	PowerPointMsoAutoShapeTypeAutoshapeBentUpArrow = '\000j\000,',
	PowerPointMsoAutoShapeTypeAutoshapeCurvedRightArrow = '\000j\000-',
	PowerPointMsoAutoShapeTypeAutoshapeCurvedLeftArrow = '\000j\000.',
	PowerPointMsoAutoShapeTypeAutoshapeCurvedUpArrow = '\000j\000/',
	PowerPointMsoAutoShapeTypeAutoshapeCurvedDownArrow = '\000j\0000',
	PowerPointMsoAutoShapeTypeAutoshapeStripedRightArrow = '\000j\0001',
	PowerPointMsoAutoShapeTypeAutoshapeNotchedRightArrow = '\000j\0002',
	PowerPointMsoAutoShapeTypeAutoshapePentagon = '\000j\0003',
	PowerPointMsoAutoShapeTypeAutoshapeChevron = '\000j\0004',
	PowerPointMsoAutoShapeTypeAutoshapeRightArrowCallout = '\000j\0005',
	PowerPointMsoAutoShapeTypeAutoshapeLeftArrowCallout = '\000j\0006',
	PowerPointMsoAutoShapeTypeAutoshapeUpArrowCallout = '\000j\0007',
	PowerPointMsoAutoShapeTypeAutoshapeDownArrowCallout = '\000j\0008',
	PowerPointMsoAutoShapeTypeAutoshapeLeftRightArrowCallout = '\000j\0009',
	PowerPointMsoAutoShapeTypeAutoshapeUpDownArrowCallout = '\000j\000:',
	PowerPointMsoAutoShapeTypeAutoshapeQuadArrowCallout = '\000j\000;',
	PowerPointMsoAutoShapeTypeAutoshapeCircularArrow = '\000j\000<',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartProcess = '\000j\000=',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartAlternateProcess = '\000j\000>',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartDecision = '\000j\000\?',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartData = '\000j\000@',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartInternalStorage = '\000j\000B',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartDocument = '\000j\000C',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartMultiDocument = '\000j\000D',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartTerminator = '\000j\000E',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartPreparation = '\000j\000F',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartManualInput = '\000j\000G',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartManualOperation = '\000j\000H',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartConnector = '\000j\000I',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartOffpageConnector = '\000j\000J',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartCard = '\000j\000K',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartPunchedTape = '\000j\000L',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartSummingJunction = '\000j\000M',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartOr = '\000j\000N',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartCollate = '\000j\000O',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartSort = '\000j\000P',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartExtract = '\000j\000Q',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartMerge = '\000j\000R',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartStoredData = '\000j\000S',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartDelay = '\000j\000T',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartMagneticDisk = '\000j\000V',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartDisplay = '\000j\000X',
	PowerPointMsoAutoShapeTypeAutoshapeExplosionOne = '\000j\000Y',
	PowerPointMsoAutoShapeTypeAutoshapeExplosionTwo = '\000j\000Z',
	PowerPointMsoAutoShapeTypeAutoshapeFourPointStar = '\000j\000[',
	PowerPointMsoAutoShapeTypeAutoshapeFivePointStar = '\000j\000\\',
	PowerPointMsoAutoShapeTypeAutoshapeEightPointStar = '\000j\000]',
	PowerPointMsoAutoShapeTypeAutoshapeSixteenPointStar = '\000j\000^',
	PowerPointMsoAutoShapeTypeAutoshapeTwentyFourPointStar = '\000j\000_',
	PowerPointMsoAutoShapeTypeAutoshapeThirtyTwoPointStar = '\000j\000`',
	PowerPointMsoAutoShapeTypeAutoshapeUpRibbon = '\000j\000a',
	PowerPointMsoAutoShapeTypeAutoshapeDownRibbon = '\000j\000b',
	PowerPointMsoAutoShapeTypeAutoshapeCurvedUpRibbon = '\000j\000c',
	PowerPointMsoAutoShapeTypeAutoshapeCurvedDownRibbon = '\000j\000d',
	PowerPointMsoAutoShapeTypeAutoshapeVerticalScroll = '\000j\000e',
	PowerPointMsoAutoShapeTypeAutoshapeHorizontalScroll = '\000j\000f',
	PowerPointMsoAutoShapeTypeAutoshapeWave = '\000j\000g',
	PowerPointMsoAutoShapeTypeAutoshapeDoubleWave = '\000j\000h',
	PowerPointMsoAutoShapeTypeAutoshapeRectangularCallout = '\000j\000i',
	PowerPointMsoAutoShapeTypeAutoshapeRoundedRectangularCallout = '\000j\000j',
	PowerPointMsoAutoShapeTypeAutoshapeOvalCallout = '\000j\000k',
	PowerPointMsoAutoShapeTypeAutoshapeCloudCallout = '\000j\000l',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutOne = '\000j\000m',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutTwo = '\000j\000n',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutThree = '\000j\000o',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutFour = '\000j\000p',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	PowerPointMsoAutoShapeTypeAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	PowerPointMsoAutoShapeTypeAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	PowerPointMsoAutoShapeTypeAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	PowerPointMsoAutoShapeTypeAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	PowerPointMsoAutoShapeTypeAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonCustom = '\000j\000}',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonHome = '\000j\000~',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonHelp = '\000j\000\177',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonInformation = '\000j\000\200',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonBeginning = '\000j\000\203',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonEnd = '\000j\000\204',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonReturn = '\000j\000\205',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonDocument = '\000j\000\206',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonSound = '\000j\000\207',
	PowerPointMsoAutoShapeTypeAutoshapeActionButtonMovie = '\000j\000\210',
	PowerPointMsoAutoShapeTypeAutoshapeBalloon = '\000j\000\211',
	PowerPointMsoAutoShapeTypeAutoshapeNotPrimitive = '\000j\000\212',
	PowerPointMsoAutoShapeTypeAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	PowerPointMsoAutoShapeTypeAutoshapeLeftRightRibbon = '\000j\000\214',
	PowerPointMsoAutoShapeTypeAutoshapeDiagonalStripe = '\000j\000\215',
	PowerPointMsoAutoShapeTypeAutoshapePie = '\000j\000\216',
	PowerPointMsoAutoShapeTypeAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	PowerPointMsoAutoShapeTypeAutoshapeDecagon = '\000j\000\220',
	PowerPointMsoAutoShapeTypeAutoshapeHeptagon = '\000j\000\221',
	PowerPointMsoAutoShapeTypeAutoshapeDodecagon = '\000j\000\222',
	PowerPointMsoAutoShapeTypeAutoshapeSixPointsStar = '\000j\000\223',
	PowerPointMsoAutoShapeTypeAutoshapeSevenPointsStar = '\000j\000\224',
	PowerPointMsoAutoShapeTypeAutoshapeTenPointsStar = '\000j\000\225',
	PowerPointMsoAutoShapeTypeAutoshapeTwelvePointsStar = '\000j\000\226',
	PowerPointMsoAutoShapeTypeAutoshapeRoundOneRectangle = '\000j\000\227',
	PowerPointMsoAutoShapeTypeAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	PowerPointMsoAutoShapeTypeAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	PowerPointMsoAutoShapeTypeAutoshapeSnipRoundRectangle = '\000j\000\232',
	PowerPointMsoAutoShapeTypeAutoshapeSnipOneRectangle = '\000j\000\233',
	PowerPointMsoAutoShapeTypeAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	PowerPointMsoAutoShapeTypeAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	PowerPointMsoAutoShapeTypeAutoshapeFrame = '\000j\000\236',
	PowerPointMsoAutoShapeTypeAutoshapeHalfFrame = '\000j\000\237',
	PowerPointMsoAutoShapeTypeAutoshapeTear = '\000j\000\240',
	PowerPointMsoAutoShapeTypeAutoshapeChord = '\000j\000\241',
	PowerPointMsoAutoShapeTypeAutoshapeCorner = '\000j\000\242',
	PowerPointMsoAutoShapeTypeAutoshapeMathPlus = '\000j\000\243',
	PowerPointMsoAutoShapeTypeAutoshapeMathMinus = '\000j\000\244',
	PowerPointMsoAutoShapeTypeAutoshapeMathMultiply = '\000j\000\245',
	PowerPointMsoAutoShapeTypeAutoshapeMathDivide = '\000j\000\246',
	PowerPointMsoAutoShapeTypeAutoshapeMathEqual = '\000j\000\247',
	PowerPointMsoAutoShapeTypeAutoshapeMathNotEqual = '\000j\000\250',
	PowerPointMsoAutoShapeTypeAutoshapeCornerTabs = '\000j\000\251',
	PowerPointMsoAutoShapeTypeAutoshapeSquareTabs = '\000j\000\252',
	PowerPointMsoAutoShapeTypeAutoshapePlaqueTabs = '\000j\000\253',
	PowerPointMsoAutoShapeTypeAutoshapeGearSix = '\000j\000\254',
	PowerPointMsoAutoShapeTypeAutoshapeGearNine = '\000j\000\255',
	PowerPointMsoAutoShapeTypeAutoshapeFunnel = '\000j\000\256',
	PowerPointMsoAutoShapeTypeAutoshapePieWedge = '\000j\000\257',
	PowerPointMsoAutoShapeTypeAutoshapeLeftCircularArrow = '\000j\000\260',
	PowerPointMsoAutoShapeTypeAutoshapeLeftRightCircularArrow = '\000j\000\261',
	PowerPointMsoAutoShapeTypeAutoshapeSwooshArrow = '\000j\000\262',
	PowerPointMsoAutoShapeTypeAutoshapeCloud = '\000j\000\263',
	PowerPointMsoAutoShapeTypeAutoshapeChartX = '\000j\000\264',
	PowerPointMsoAutoShapeTypeAutoshapeChartStar = '\000j\000\265',
	PowerPointMsoAutoShapeTypeAutoshapeChartPlus = '\000j\000\266',
	PowerPointMsoAutoShapeTypeAutoshapeLineInverse = '\000j\000\267'
};
typedef enum PowerPointMsoAutoShapeType PowerPointMsoAutoShapeType;

enum PowerPointMsoShapeType {
	PowerPointMsoShapeTypeShapeTypeUnset = '\000\213\377\376',
	PowerPointMsoShapeTypeShapeTypeAuto = '\000\214\000\001',
	PowerPointMsoShapeTypeShapeTypeCallout = '\000\214\000\002',
	PowerPointMsoShapeTypeShapeTypeChart = '\000\214\000\003',
	PowerPointMsoShapeTypeShapeTypeComment = '\000\214\000\004',
	PowerPointMsoShapeTypeShapeTypeFreeForm = '\000\214\000\005',
	PowerPointMsoShapeTypeShapeTypeGroup = '\000\214\000\006',
	PowerPointMsoShapeTypeShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	PowerPointMsoShapeTypeShapeTypeFormControl = '\000\214\000\010',
	PowerPointMsoShapeTypeShapeTypeLine = '\000\214\000\011',
	PowerPointMsoShapeTypeShapeTypeLinkedOLEObject = '\000\214\000\012',
	PowerPointMsoShapeTypeShapeTypeLinkedPicture = '\000\214\000\013',
	PowerPointMsoShapeTypeShapeTypeOLEControl = '\000\214\000\014',
	PowerPointMsoShapeTypeShapeTypePicture = '\000\214\000\015',
	PowerPointMsoShapeTypeShapeTypePlaceHolder = '\000\214\000\016',
	PowerPointMsoShapeTypeShapeTypeWordArt = '\000\214\000\017',
	PowerPointMsoShapeTypeShapeTypeMedia = '\000\214\000\020',
	PowerPointMsoShapeTypeShapeTypeTextBox = '\000\214\000\021',
	PowerPointMsoShapeTypeShapeTypeTable = '\000\214\000\022',
	PowerPointMsoShapeTypeShapeTypeCanvas = '\000\214\000\023',
	PowerPointMsoShapeTypeShapeTypeDiagram = '\000\214\000\024',
	PowerPointMsoShapeTypeShapeTypeInk = '\000\214\000\025',
	PowerPointMsoShapeTypeShapeTypeInkComment = '\000\214\000\026',
	PowerPointMsoShapeTypeShapeTypeSmartartGraphic = '\000\214\000\027',
	PowerPointMsoShapeTypeShapeTypeSlicer = '\000\214\000\030',
	PowerPointMsoShapeTypeShapeTypeWebVideo = '\000\214\000\031',
	PowerPointMsoShapeTypeShapeTypeContentApplication = '\000\214\000\032'
};
typedef enum PowerPointMsoShapeType PowerPointMsoShapeType;

enum PowerPointMsoColorType {
	PowerPointMsoColorTypeColorTypeUnset = '\000j\377\376',
	PowerPointMsoColorTypeRGB = '\000k\000\001',
	PowerPointMsoColorTypeScheme = '\000k\000\002'
};
typedef enum PowerPointMsoColorType PowerPointMsoColorType;

enum PowerPointMsoPictureColorType {
	PowerPointMsoPictureColorTypePictureColorTypeUnset = '\000\265\377\376',
	PowerPointMsoPictureColorTypePictureColorAutomatic = '\000\266\000\001',
	PowerPointMsoPictureColorTypePictureColorGrayScale = '\000\266\000\002',
	PowerPointMsoPictureColorTypePictureColorBlackAndWhite = '\000\266\000\003',
	PowerPointMsoPictureColorTypePictureColorWatermark = '\000\266\000\004'
};
typedef enum PowerPointMsoPictureColorType PowerPointMsoPictureColorType;

enum PowerPointMsoCalloutAngleType {
	PowerPointMsoCalloutAngleTypeAngleUnset = '\000k\377\376',
	PowerPointMsoCalloutAngleTypeAngleAutomatic = '\000l\000\001',
	PowerPointMsoCalloutAngleTypeAngle30 = '\000l\000\002',
	PowerPointMsoCalloutAngleTypeAngle45 = '\000l\000\003',
	PowerPointMsoCalloutAngleTypeAngle60 = '\000l\000\004',
	PowerPointMsoCalloutAngleTypeAngle90 = '\000l\000\005'
};
typedef enum PowerPointMsoCalloutAngleType PowerPointMsoCalloutAngleType;

enum PowerPointMsoCalloutDropType {
	PowerPointMsoCalloutDropTypeDropUnset = '\000l\377\376',
	PowerPointMsoCalloutDropTypeDropCustom = '\000m\000\001',
	PowerPointMsoCalloutDropTypeDropTop = '\000m\000\002',
	PowerPointMsoCalloutDropTypeDropCenter = '\000m\000\003',
	PowerPointMsoCalloutDropTypeDropBottom = '\000m\000\004'
};
typedef enum PowerPointMsoCalloutDropType PowerPointMsoCalloutDropType;

enum PowerPointMsoCalloutType {
	PowerPointMsoCalloutTypeCalloutUnset = '\000m\377\376',
	PowerPointMsoCalloutTypeCalloutOne = '\000n\000\001',
	PowerPointMsoCalloutTypeCalloutTwo = '\000n\000\002',
	PowerPointMsoCalloutTypeCalloutThree = '\000n\000\003',
	PowerPointMsoCalloutTypeCalloutFour = '\000n\000\004'
};
typedef enum PowerPointMsoCalloutType PowerPointMsoCalloutType;

enum PowerPointMsoTextOrientation {
	PowerPointMsoTextOrientationTextOrientationUnset = '\000\215\377\376',
	PowerPointMsoTextOrientationHorizontal = '\000\216\000\001',
	PowerPointMsoTextOrientationUpward = '\000\216\000\002',
	PowerPointMsoTextOrientationDownward = '\000\216\000\003',
	PowerPointMsoTextOrientationVerticalEastAsian = '\000\216\000\004',
	PowerPointMsoTextOrientationVertical = '\000\216\000\005',
	PowerPointMsoTextOrientationHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum PowerPointMsoTextOrientation PowerPointMsoTextOrientation;

enum PowerPointMsoScaleFrom {
	PowerPointMsoScaleFromScaleFromTopLeft = '\000o\000\000',
	PowerPointMsoScaleFromScaleFromMiddle = '\000o\000\001',
	PowerPointMsoScaleFromScaleFromBottomRight = '\000o\000\002'
};
typedef enum PowerPointMsoScaleFrom PowerPointMsoScaleFrom;

enum PowerPointMsoPresetCamera {
	PowerPointMsoPresetCameraPresetCameraUnset = '\000\256\377\376',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromTop = '\000\257\000\002',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromTopright = '\000\257\000\003',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromLeft = '\000\257\000\004',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromFront = '\000\257\000\005',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromRight = '\000\257\000\006',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromBottom = '\000\257\000\010',
	PowerPointMsoPresetCameraCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	PowerPointMsoPresetCameraCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	PowerPointMsoPresetCameraCameraOrthographic = '\000\257\000\023',
	PowerPointMsoPresetCameraCameraIsometricFromTopUp = '\000\257\000\024',
	PowerPointMsoPresetCameraCameraIsometricFromTopDown = '\000\257\000\025',
	PowerPointMsoPresetCameraCameraIsometricFromBottomUp = '\000\257\000\026',
	PowerPointMsoPresetCameraCameraIsometricFromBottomDown = '\000\257\000\027',
	PowerPointMsoPresetCameraCameraIsometricFromLeftUp = '\000\257\000\030',
	PowerPointMsoPresetCameraCameraIsometricFromLeftDown = '\000\257\000\031',
	PowerPointMsoPresetCameraCameraIsometricFromRightUp = '\000\257\000\032',
	PowerPointMsoPresetCameraCameraIsometricFromRightDown = '\000\257\000\033',
	PowerPointMsoPresetCameraCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	PowerPointMsoPresetCameraCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	PowerPointMsoPresetCameraCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	PowerPointMsoPresetCameraCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	PowerPointMsoPresetCameraCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	PowerPointMsoPresetCameraCameraIsometricOffAxis2FromTop = '\000\257\000!',
	PowerPointMsoPresetCameraCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	PowerPointMsoPresetCameraCameraIsometricOffAxis3FromRight = '\000\257\000#',
	PowerPointMsoPresetCameraCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	PowerPointMsoPresetCameraCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	PowerPointMsoPresetCameraCameraIsometricOffAxis4FromRight = '\000\257\000&',
	PowerPointMsoPresetCameraCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	PowerPointMsoPresetCameraCameraObliqueFromTopLeft = '\000\257\000(',
	PowerPointMsoPresetCameraCameraObliqueFromTop = '\000\257\000)',
	PowerPointMsoPresetCameraCameraObliqueFromTopRight = '\000\257\000*',
	PowerPointMsoPresetCameraCameraObliqueFromLeft = '\000\257\000+',
	PowerPointMsoPresetCameraCameraObliqueFromRight = '\000\257\000,',
	PowerPointMsoPresetCameraCameraObliqueFromBottomLeft = '\000\257\000-',
	PowerPointMsoPresetCameraCameraObliqueFromBottom = '\000\257\000.',
	PowerPointMsoPresetCameraCameraObliqueFromBottomRight = '\000\257\000/',
	PowerPointMsoPresetCameraCameraPerspectiveFromFront = '\000\257\0000',
	PowerPointMsoPresetCameraCameraPerspectiveFromLeft = '\000\257\0001',
	PowerPointMsoPresetCameraCameraPerspectiveFromRight = '\000\257\0002',
	PowerPointMsoPresetCameraCameraPerspectiveFromAbove = '\000\257\0003',
	PowerPointMsoPresetCameraCameraPerspectiveFromBelow = '\000\257\0004',
	PowerPointMsoPresetCameraCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	PowerPointMsoPresetCameraCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	PowerPointMsoPresetCameraCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	PowerPointMsoPresetCameraCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	PowerPointMsoPresetCameraCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	PowerPointMsoPresetCameraCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	PowerPointMsoPresetCameraCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	PowerPointMsoPresetCameraCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	PowerPointMsoPresetCameraCameraPerspectiveRelaxed = '\000\257\000=',
	PowerPointMsoPresetCameraCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum PowerPointMsoPresetCamera PowerPointMsoPresetCamera;

enum PowerPointMsoLightRigType {
	PowerPointMsoLightRigTypeLightRigUnset = '\000\257\377\376',
	PowerPointMsoLightRigTypeLightRigFlat1 = '\000\260\000\001',
	PowerPointMsoLightRigTypeLightRigFlat2 = '\000\260\000\002',
	PowerPointMsoLightRigTypeLightRigFlat3 = '\000\260\000\003',
	PowerPointMsoLightRigTypeLightRigFlat4 = '\000\260\000\004',
	PowerPointMsoLightRigTypeLightRigNormal1 = '\000\260\000\005',
	PowerPointMsoLightRigTypeLightRigNormal2 = '\000\260\000\006',
	PowerPointMsoLightRigTypeLightRigNormal3 = '\000\260\000\007',
	PowerPointMsoLightRigTypeLightRigNormal4 = '\000\260\000\010',
	PowerPointMsoLightRigTypeLightRigHarsh1 = '\000\260\000\011',
	PowerPointMsoLightRigTypeLightRigHarsh2 = '\000\260\000\012',
	PowerPointMsoLightRigTypeLightRigHarsh3 = '\000\260\000\013',
	PowerPointMsoLightRigTypeLightRigHarsh4 = '\000\260\000\014',
	PowerPointMsoLightRigTypeLightRigThreePoint = '\000\260\000\015',
	PowerPointMsoLightRigTypeLightRigBalanced = '\000\260\000\016',
	PowerPointMsoLightRigTypeLightRigSoft = '\000\260\000\017',
	PowerPointMsoLightRigTypeLightRigHarsh = '\000\260\000\020',
	PowerPointMsoLightRigTypeLightRigFlood = '\000\260\000\021',
	PowerPointMsoLightRigTypeLightRigContrasting = '\000\260\000\022',
	PowerPointMsoLightRigTypeLightRigMorning = '\000\260\000\023',
	PowerPointMsoLightRigTypeLightRigSunrise = '\000\260\000\024',
	PowerPointMsoLightRigTypeLightRigSunset = '\000\260\000\025',
	PowerPointMsoLightRigTypeLightRigChilly = '\000\260\000\026',
	PowerPointMsoLightRigTypeLightRigFreezing = '\000\260\000\027',
	PowerPointMsoLightRigTypeLightRigFlat = '\000\260\000\030',
	PowerPointMsoLightRigTypeLightRigTwoPoint = '\000\260\000\031',
	PowerPointMsoLightRigTypeLightRigGlow = '\000\260\000\032',
	PowerPointMsoLightRigTypeLightRigBrightRoom = '\000\260\000\033'
};
typedef enum PowerPointMsoLightRigType PowerPointMsoLightRigType;

enum PowerPointMsoBevelType {
	PowerPointMsoBevelTypeBevelTypeUnset = '\000\260\377\376',
	PowerPointMsoBevelTypeBevelNone = '\000\261\000\001',
	PowerPointMsoBevelTypeBevelRelaxedInset = '\000\261\000\002',
	PowerPointMsoBevelTypeBevelCircle = '\000\261\000\003',
	PowerPointMsoBevelTypeBevelSlope = '\000\261\000\004',
	PowerPointMsoBevelTypeBevelCross = '\000\261\000\005',
	PowerPointMsoBevelTypeBevelAngle = '\000\261\000\006',
	PowerPointMsoBevelTypeBevelSoftRound = '\000\261\000\007',
	PowerPointMsoBevelTypeBevelConvex = '\000\261\000\010',
	PowerPointMsoBevelTypeBevelCoolSlant = '\000\261\000\011',
	PowerPointMsoBevelTypeBevelDivot = '\000\261\000\012',
	PowerPointMsoBevelTypeBevelRiblet = '\000\261\000\013',
	PowerPointMsoBevelTypeBevelHardEdge = '\000\261\000\014',
	PowerPointMsoBevelTypeBevelArtDeco = '\000\261\000\015'
};
typedef enum PowerPointMsoBevelType PowerPointMsoBevelType;

enum PowerPointMsoShadowStyle {
	PowerPointMsoShadowStyleShadowStyleUnset = '\000\261\377\376',
	PowerPointMsoShadowStyleShadowStyleInner = '\000\262\000\001',
	PowerPointMsoShadowStyleShadowStyleOuter = '\000\262\000\002'
};
typedef enum PowerPointMsoShadowStyle PowerPointMsoShadowStyle;

enum PowerPointMsoParagraphAlignment {
	PowerPointMsoParagraphAlignmentParagraphAlignmentUnset = '\000\346\377\376',
	PowerPointMsoParagraphAlignmentParagraphAlignLeft = '\000\347\000\000',
	PowerPointMsoParagraphAlignmentParagraphAlignCenter = '\000\347\000\001',
	PowerPointMsoParagraphAlignmentParagraphAlignRight = '\000\347\000\002',
	PowerPointMsoParagraphAlignmentParagraphAlignJustify = '\000\347\000\003',
	PowerPointMsoParagraphAlignmentParagraphAlignDistribute = '\000\347\000\004',
	PowerPointMsoParagraphAlignmentParagraphAlignThai = '\000\347\000\005',
	PowerPointMsoParagraphAlignmentParagraphAlignJustifyLow = '\000\347\000\006'
};
typedef enum PowerPointMsoParagraphAlignment PowerPointMsoParagraphAlignment;

enum PowerPointMsoTextStrike {
	PowerPointMsoTextStrikeStrikeUnset = '\000\263\377\376',
	PowerPointMsoTextStrikeNoStrike = '\000\264\000\000',
	PowerPointMsoTextStrikeSingleStrike = '\000\264\000\001',
	PowerPointMsoTextStrikeDoubleStrike = '\000\264\000\002'
};
typedef enum PowerPointMsoTextStrike PowerPointMsoTextStrike;

enum PowerPointMsoTextCaps {
	PowerPointMsoTextCapsCapsUnset = '\000\264\377\376',
	PowerPointMsoTextCapsNoCaps = '\000\265\000\000',
	PowerPointMsoTextCapsSmallCaps = '\000\265\000\001',
	PowerPointMsoTextCapsAllCaps = '\000\265\000\002'
};
typedef enum PowerPointMsoTextCaps PowerPointMsoTextCaps;

enum PowerPointMsoTextUnderlineType {
	PowerPointMsoTextUnderlineTypeUnderlineUnset = '\003\356\377\376',
	PowerPointMsoTextUnderlineTypeNoUnderline = '\003\357\000\000',
	PowerPointMsoTextUnderlineTypeUnderlineWordsOnly = '\003\357\000\001',
	PowerPointMsoTextUnderlineTypeUnderlineSingleLine = '\003\357\000\002',
	PowerPointMsoTextUnderlineTypeUnderlineDoubleLine = '\003\357\000\003',
	PowerPointMsoTextUnderlineTypeUnderlineHeavyLine = '\003\357\000\004',
	PowerPointMsoTextUnderlineTypeUnderlineDottedLine = '\003\357\000\005',
	PowerPointMsoTextUnderlineTypeUnderlineHeavyDottedLine = '\003\357\000\006',
	PowerPointMsoTextUnderlineTypeUnderlineDashLine = '\003\357\000\007',
	PowerPointMsoTextUnderlineTypeUnderlineHeavyDashLine = '\003\357\000\010',
	PowerPointMsoTextUnderlineTypeUnderlineLongDashLine = '\003\357\000\011',
	PowerPointMsoTextUnderlineTypeUnderlineHeavyLongDashLine = '\003\357\000\012',
	PowerPointMsoTextUnderlineTypeUnderlineDotDashLine = '\003\357\000\013',
	PowerPointMsoTextUnderlineTypeUnderlineHeavyDotDashLine = '\003\357\000\014',
	PowerPointMsoTextUnderlineTypeUnderlineDotDotDashLine = '\003\357\000\015',
	PowerPointMsoTextUnderlineTypeUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	PowerPointMsoTextUnderlineTypeUnderlineWavyLine = '\003\357\000\017',
	PowerPointMsoTextUnderlineTypeUnderlineHeavyWavyLine = '\003\357\000\020',
	PowerPointMsoTextUnderlineTypeUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum PowerPointMsoTextUnderlineType PowerPointMsoTextUnderlineType;

enum PowerPointMsoTextTabAlign {
	PowerPointMsoTextTabAlignTabUnset = '\000\266\377\376',
	PowerPointMsoTextTabAlignLeftTab = '\000\267\000\000',
	PowerPointMsoTextTabAlignCenterTab = '\000\267\000\001',
	PowerPointMsoTextTabAlignRightTab = '\000\267\000\002',
	PowerPointMsoTextTabAlignDecimalTab = '\000\267\000\003'
};
typedef enum PowerPointMsoTextTabAlign PowerPointMsoTextTabAlign;

enum PowerPointMsoTextCharWrap {
	PowerPointMsoTextCharWrapCharacterWrapUnset = '\000\267\377\376',
	PowerPointMsoTextCharWrapNoCharacterWrap = '\000\270\000\000',
	PowerPointMsoTextCharWrapStandardCharacterWrap = '\000\270\000\001',
	PowerPointMsoTextCharWrapStrictCharacterWrap = '\000\270\000\002',
	PowerPointMsoTextCharWrapCustomCharacterWrap = '\000\270\000\003'
};
typedef enum PowerPointMsoTextCharWrap PowerPointMsoTextCharWrap;

enum PowerPointMsoTextFontAlign {
	PowerPointMsoTextFontAlignFontAlignmentUnset = '\000\270\377\376',
	PowerPointMsoTextFontAlignAutomaticAlignment = '\000\271\000\000',
	PowerPointMsoTextFontAlignTopAlignment = '\000\271\000\001',
	PowerPointMsoTextFontAlignCenterAlignment = '\000\271\000\002',
	PowerPointMsoTextFontAlignBaselineAlignment = '\000\271\000\003',
	PowerPointMsoTextFontAlignBottomAlignment = '\000\271\000\004'
};
typedef enum PowerPointMsoTextFontAlign PowerPointMsoTextFontAlign;

enum PowerPointMsoAutoSize {
	PowerPointMsoAutoSizeAutoSizeUnset = '\000\344\377\376',
	PowerPointMsoAutoSizeAutoSizeNone = '\000\345\000\000',
	PowerPointMsoAutoSizeShapeToFitText = '\000\345\000\001',
	PowerPointMsoAutoSizeTextToFitShape = '\000\345\000\002'
};
typedef enum PowerPointMsoAutoSize PowerPointMsoAutoSize;

enum PowerPointMsoPathFormat {
	PowerPointMsoPathFormatPathTypeUnset = '\000\272\377\376',
	PowerPointMsoPathFormatNoPathType = '\000\273\000\000',
	PowerPointMsoPathFormatPathType1 = '\000\273\000\001',
	PowerPointMsoPathFormatPathType2 = '\000\273\000\002',
	PowerPointMsoPathFormatPathType3 = '\000\273\000\003',
	PowerPointMsoPathFormatPathType4 = '\000\273\000\004'
};
typedef enum PowerPointMsoPathFormat PowerPointMsoPathFormat;

enum PowerPointMsoWarpFormat {
	PowerPointMsoWarpFormatWarpFormatUnset = '\000\273\377\376',
	PowerPointMsoWarpFormatWarpFormat1 = '\000\274\000\000',
	PowerPointMsoWarpFormatWarpFormat2 = '\000\274\000\001',
	PowerPointMsoWarpFormatWarpFormat3 = '\000\274\000\002',
	PowerPointMsoWarpFormatWarpFormat4 = '\000\274\000\003',
	PowerPointMsoWarpFormatWarpFormat5 = '\000\274\000\004',
	PowerPointMsoWarpFormatWarpFormat6 = '\000\274\000\005',
	PowerPointMsoWarpFormatWarpFormat7 = '\000\274\000\006',
	PowerPointMsoWarpFormatWarpFormat8 = '\000\274\000\007',
	PowerPointMsoWarpFormatWarpFormat9 = '\000\274\000\010',
	PowerPointMsoWarpFormatWarpFormat10 = '\000\274\000\011',
	PowerPointMsoWarpFormatWarpFormat11 = '\000\274\000\012',
	PowerPointMsoWarpFormatWarpFormat12 = '\000\274\000\013',
	PowerPointMsoWarpFormatWarpFormat13 = '\000\274\000\014',
	PowerPointMsoWarpFormatWarpFormat14 = '\000\274\000\015',
	PowerPointMsoWarpFormatWarpFormat15 = '\000\274\000\016',
	PowerPointMsoWarpFormatWarpFormat16 = '\000\274\000\017',
	PowerPointMsoWarpFormatWarpFormat17 = '\000\274\000\020',
	PowerPointMsoWarpFormatWarpFormat18 = '\000\274\000\021',
	PowerPointMsoWarpFormatWarpFormat19 = '\000\274\000\022',
	PowerPointMsoWarpFormatWarpFormat20 = '\000\274\000\023',
	PowerPointMsoWarpFormatWarpFormat21 = '\000\274\000\024',
	PowerPointMsoWarpFormatWarpFormat22 = '\000\274\000\025',
	PowerPointMsoWarpFormatWarpFormat23 = '\000\274\000\026',
	PowerPointMsoWarpFormatWarpFormat24 = '\000\274\000\027',
	PowerPointMsoWarpFormatWarpFormat25 = '\000\274\000\030',
	PowerPointMsoWarpFormatWarpFormat26 = '\000\274\000\031',
	PowerPointMsoWarpFormatWarpFormat27 = '\000\274\000\032',
	PowerPointMsoWarpFormatWarpFormat28 = '\000\274\000\033',
	PowerPointMsoWarpFormatWarpFormat29 = '\000\274\000\034',
	PowerPointMsoWarpFormatWarpFormat30 = '\000\274\000\035',
	PowerPointMsoWarpFormatWarpFormat31 = '\000\274\000\036',
	PowerPointMsoWarpFormatWarpFormat32 = '\000\274\000\037',
	PowerPointMsoWarpFormatWarpFormat33 = '\000\274\000 ',
	PowerPointMsoWarpFormatWarpFormat34 = '\000\274\000!',
	PowerPointMsoWarpFormatWarpFormat35 = '\000\274\000\"',
	PowerPointMsoWarpFormatWarpFormat36 = '\000\274\000#',
	PowerPointMsoWarpFormatWarpFormat37 = '\000\274\000$'
};
typedef enum PowerPointMsoWarpFormat PowerPointMsoWarpFormat;

enum PowerPointMsoTextChangeCase {
	PowerPointMsoTextChangeCaseCaseSentence = '\000\344\000\001',
	PowerPointMsoTextChangeCaseCaseLower = '\000\344\000\002',
	PowerPointMsoTextChangeCaseCaseUpper = '\000\344\000\003',
	PowerPointMsoTextChangeCaseCaseTitle = '\000\344\000\004',
	PowerPointMsoTextChangeCaseCaseToggle = '\000\344\000\005'
};
typedef enum PowerPointMsoTextChangeCase PowerPointMsoTextChangeCase;

enum PowerPointMsoDateTimeFormat {
	PowerPointMsoDateTimeFormatDateTimeFormatUnset = '\000\275\377\376',
	PowerPointMsoDateTimeFormatDateTimeFormatMdyy = '\000\276\000\001',
	PowerPointMsoDateTimeFormatDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	PowerPointMsoDateTimeFormatDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	PowerPointMsoDateTimeFormatDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	PowerPointMsoDateTimeFormatDateTimeFormatDMMMyy = '\000\276\000\005',
	PowerPointMsoDateTimeFormatDateTimeFormatMMMMyy = '\000\276\000\006',
	PowerPointMsoDateTimeFormatDateTimeFormatMMyy = '\000\276\000\007',
	PowerPointMsoDateTimeFormatDateTimeFormatMMddyyHmm = '\000\276\000\010',
	PowerPointMsoDateTimeFormatDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	PowerPointMsoDateTimeFormatDateTimeFormatHmm = '\000\276\000\012',
	PowerPointMsoDateTimeFormatDateTimeFormatHmmss = '\000\276\000\013',
	PowerPointMsoDateTimeFormatDateTimeFormatHmmAMPM = '\000\276\000\014',
	PowerPointMsoDateTimeFormatDateTimeFormatHmmssAMPM = '\000\276\000\015',
	PowerPointMsoDateTimeFormatDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum PowerPointMsoDateTimeFormat PowerPointMsoDateTimeFormat;

enum PowerPointMsoSoftEdgeType {
	PowerPointMsoSoftEdgeTypeSoftEdgeUnset = '\000\277\377\376',
	PowerPointMsoSoftEdgeTypeNoSoftEdge = '\000\300\000\000',
	PowerPointMsoSoftEdgeTypeSoftEdgeType1 = '\000\300\000\001',
	PowerPointMsoSoftEdgeTypeSoftEdgeType2 = '\000\300\000\002',
	PowerPointMsoSoftEdgeTypeSoftEdgeType3 = '\000\300\000\003',
	PowerPointMsoSoftEdgeTypeSoftEdgeType4 = '\000\300\000\004',
	PowerPointMsoSoftEdgeTypeSoftEdgeType5 = '\000\300\000\005',
	PowerPointMsoSoftEdgeTypeSoftEdgeType6 = '\000\300\000\006'
};
typedef enum PowerPointMsoSoftEdgeType PowerPointMsoSoftEdgeType;

enum PowerPointMsoThemeColorSchemeIndex {
	PowerPointMsoThemeColorSchemeIndexFirstDarkSchemeColor = '\000\301\000\001',
	PowerPointMsoThemeColorSchemeIndexFirstLightSchemeColor = '\000\301\000\002',
	PowerPointMsoThemeColorSchemeIndexSecondDarkSchemeColor = '\000\301\000\003',
	PowerPointMsoThemeColorSchemeIndexSecondLightSchemeColor = '\000\301\000\004',
	PowerPointMsoThemeColorSchemeIndexFirstAccentSchemeColor = '\000\301\000\005',
	PowerPointMsoThemeColorSchemeIndexSecondAccentSchemeColor = '\000\301\000\006',
	PowerPointMsoThemeColorSchemeIndexThirdAccentSchemeColor = '\000\301\000\007',
	PowerPointMsoThemeColorSchemeIndexFourthAccentSchemeColor = '\000\301\000\010',
	PowerPointMsoThemeColorSchemeIndexFifthAccentSchemeColor = '\000\301\000\011',
	PowerPointMsoThemeColorSchemeIndexSixthAccentSchemeColor = '\000\301\000\012',
	PowerPointMsoThemeColorSchemeIndexHyperlinkSchemeColor = '\000\301\000\013',
	PowerPointMsoThemeColorSchemeIndexFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum PowerPointMsoThemeColorSchemeIndex PowerPointMsoThemeColorSchemeIndex;

enum PowerPointMsoThemeColorIndex {
	PowerPointMsoThemeColorIndexThemeColorUnset = '\000\301\377\376',
	PowerPointMsoThemeColorIndexNoThemeColor = '\000\302\000\000',
	PowerPointMsoThemeColorIndexFirstDarkThemeColor = '\000\302\000\001',
	PowerPointMsoThemeColorIndexFirstLightThemeColor = '\000\302\000\002',
	PowerPointMsoThemeColorIndexSecondDarkThemeColor = '\000\302\000\003',
	PowerPointMsoThemeColorIndexSecondLightThemeColor = '\000\302\000\004',
	PowerPointMsoThemeColorIndexFirstAccentThemeColor = '\000\302\000\005',
	PowerPointMsoThemeColorIndexSecondAccentThemeColor = '\000\302\000\006',
	PowerPointMsoThemeColorIndexThirdAccentThemeColor = '\000\302\000\007',
	PowerPointMsoThemeColorIndexFourthAccentThemeColor = '\000\302\000\010',
	PowerPointMsoThemeColorIndexFifthAccentThemeColor = '\000\302\000\011',
	PowerPointMsoThemeColorIndexSixthAccentThemeColor = '\000\302\000\012',
	PowerPointMsoThemeColorIndexHyperlinkThemeColor = '\000\302\000\013',
	PowerPointMsoThemeColorIndexFollowedHyperlinkThemeColor = '\000\302\000\014',
	PowerPointMsoThemeColorIndexFirstTextThemeColor = '\000\302\000\015',
	PowerPointMsoThemeColorIndexFirstBackgroundThemeColor = '\000\302\000\016',
	PowerPointMsoThemeColorIndexSecondTextThemeColor = '\000\302\000\017',
	PowerPointMsoThemeColorIndexSecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum PowerPointMsoThemeColorIndex PowerPointMsoThemeColorIndex;

enum PowerPointMsoFontLanguageIndex {
	PowerPointMsoFontLanguageIndexThemeFontLatin = '\000\303\000\001',
	PowerPointMsoFontLanguageIndexThemeFontComplexScript = '\000\303\000\002',
	PowerPointMsoFontLanguageIndexThemeFontHighAnsi = '\000\303\000\003',
	PowerPointMsoFontLanguageIndexThemeFontEastAsian = '\000\303\000\004'
};
typedef enum PowerPointMsoFontLanguageIndex PowerPointMsoFontLanguageIndex;

enum PowerPointMsoShapeStyleIndex {
	PowerPointMsoShapeStyleIndexShapeStyleUnset = '\000\303\377\376',
	PowerPointMsoShapeStyleIndexShapeNotAPreset = '\000\304\000\000',
	PowerPointMsoShapeStyleIndexShapePreset1 = '\000\304\000\001',
	PowerPointMsoShapeStyleIndexShapePreset2 = '\000\304\000\002',
	PowerPointMsoShapeStyleIndexShapePreset3 = '\000\304\000\003',
	PowerPointMsoShapeStyleIndexShapePreset4 = '\000\304\000\004',
	PowerPointMsoShapeStyleIndexShapePreset5 = '\000\304\000\005',
	PowerPointMsoShapeStyleIndexShapePreset6 = '\000\304\000\006',
	PowerPointMsoShapeStyleIndexShapePreset7 = '\000\304\000\007',
	PowerPointMsoShapeStyleIndexShapePreset8 = '\000\304\000\010',
	PowerPointMsoShapeStyleIndexShapePreset9 = '\000\304\000\011',
	PowerPointMsoShapeStyleIndexShapePreset10 = '\000\304\000\012',
	PowerPointMsoShapeStyleIndexShapePreset11 = '\000\304\000\013',
	PowerPointMsoShapeStyleIndexShapePreset12 = '\000\304\000\014',
	PowerPointMsoShapeStyleIndexShapePreset13 = '\000\304\000\015',
	PowerPointMsoShapeStyleIndexShapePreset14 = '\000\304\000\016',
	PowerPointMsoShapeStyleIndexShapePreset15 = '\000\304\000\017',
	PowerPointMsoShapeStyleIndexShapePreset16 = '\000\304\000\020',
	PowerPointMsoShapeStyleIndexShapePreset17 = '\000\304\000\021',
	PowerPointMsoShapeStyleIndexShapePreset18 = '\000\304\000\022',
	PowerPointMsoShapeStyleIndexShapePreset19 = '\000\304\000\023',
	PowerPointMsoShapeStyleIndexShapePreset20 = '\000\304\000\024',
	PowerPointMsoShapeStyleIndexShapePreset21 = '\000\304\000\025',
	PowerPointMsoShapeStyleIndexShapePreset22 = '\000\304\000\026',
	PowerPointMsoShapeStyleIndexShapePreset23 = '\000\304\000\027',
	PowerPointMsoShapeStyleIndexShapePreset24 = '\000\304\000\030',
	PowerPointMsoShapeStyleIndexShapePreset25 = '\000\304\000\031',
	PowerPointMsoShapeStyleIndexShapePreset26 = '\000\304\000\032',
	PowerPointMsoShapeStyleIndexShapePreset27 = '\000\304\000\033',
	PowerPointMsoShapeStyleIndexShapePreset28 = '\000\304\000\034',
	PowerPointMsoShapeStyleIndexShapePreset29 = '\000\304\000\035',
	PowerPointMsoShapeStyleIndexShapePreset30 = '\000\304\000\036',
	PowerPointMsoShapeStyleIndexShapePreset31 = '\000\304\000\037',
	PowerPointMsoShapeStyleIndexShapePreset32 = '\000\304\000 ',
	PowerPointMsoShapeStyleIndexShapePreset33 = '\000\304\000!',
	PowerPointMsoShapeStyleIndexShapePreset34 = '\000\304\000\"',
	PowerPointMsoShapeStyleIndexShapePreset35 = '\000\304\000#',
	PowerPointMsoShapeStyleIndexShapePreset36 = '\000\304\000$',
	PowerPointMsoShapeStyleIndexShapePreset37 = '\000\304\000%',
	PowerPointMsoShapeStyleIndexShapePreset38 = '\000\304\000&',
	PowerPointMsoShapeStyleIndexShapePreset39 = '\000\304\000\'',
	PowerPointMsoShapeStyleIndexShapePreset40 = '\000\304\000(',
	PowerPointMsoShapeStyleIndexShapePreset41 = '\000\304\000)',
	PowerPointMsoShapeStyleIndexShapePreset42 = '\000\304\000*',
	PowerPointMsoShapeStyleIndexLinePreset1 = '\000\304\'\021',
	PowerPointMsoShapeStyleIndexLinePreset2 = '\000\304\'\022',
	PowerPointMsoShapeStyleIndexLinePreset3 = '\000\304\'\023',
	PowerPointMsoShapeStyleIndexLinePreset4 = '\000\304\'\024',
	PowerPointMsoShapeStyleIndexLinePreset5 = '\000\304\'\025',
	PowerPointMsoShapeStyleIndexLinePreset6 = '\000\304\'\026',
	PowerPointMsoShapeStyleIndexLinePreset7 = '\000\304\'\027',
	PowerPointMsoShapeStyleIndexLinePreset8 = '\000\304\'\030',
	PowerPointMsoShapeStyleIndexLinePreset9 = '\000\304\'\031',
	PowerPointMsoShapeStyleIndexLinePreset10 = '\000\304\'\032',
	PowerPointMsoShapeStyleIndexLinePreset11 = '\000\304\'\033',
	PowerPointMsoShapeStyleIndexLinePreset12 = '\000\304\'\034',
	PowerPointMsoShapeStyleIndexLinePreset13 = '\000\304\'\035',
	PowerPointMsoShapeStyleIndexLinePreset14 = '\000\304\'\036',
	PowerPointMsoShapeStyleIndexLinePreset15 = '\000\304\'\037',
	PowerPointMsoShapeStyleIndexLinePreset16 = '\000\304\' ',
	PowerPointMsoShapeStyleIndexLinePreset17 = '\000\304\'!',
	PowerPointMsoShapeStyleIndexLinePreset18 = '\000\304\'\"',
	PowerPointMsoShapeStyleIndexLinePreset19 = '\000\304\'#',
	PowerPointMsoShapeStyleIndexLinePreset20 = '\000\304\'$',
	PowerPointMsoShapeStyleIndexLinePreset21 = '\000\304\'%'
};
typedef enum PowerPointMsoShapeStyleIndex PowerPointMsoShapeStyleIndex;

enum PowerPointMsoBackgroundStyleIndex {
	PowerPointMsoBackgroundStyleIndexBackgroundUnset = '\000\304\377\376',
	PowerPointMsoBackgroundStyleIndexBackgroundNotAPreset = '\000\305\000\000',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset1 = '\000\305\000\001',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset2 = '\000\305\000\002',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset3 = '\000\305\000\003',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset4 = '\000\305\000\004',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset5 = '\000\305\000\005',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset6 = '\000\305\000\006',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset7 = '\000\305\000\007',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset8 = '\000\305\000\010',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset9 = '\000\305\000\011',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset10 = '\000\305\000\012',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset11 = '\000\305\000\013',
	PowerPointMsoBackgroundStyleIndexBackgroundPreset12 = '\000\305\000\014'
};
typedef enum PowerPointMsoBackgroundStyleIndex PowerPointMsoBackgroundStyleIndex;

enum PowerPointMsoTextDirection {
	PowerPointMsoTextDirectionTextDirectionUnset = '\000\352\377\376',
	PowerPointMsoTextDirectionLeftToRight = '\000\353\000\001',
	PowerPointMsoTextDirectionRightToLeft = '\000\353\000\002'
};
typedef enum PowerPointMsoTextDirection PowerPointMsoTextDirection;

enum PowerPointMsoBulletType {
	PowerPointMsoBulletTypeBulletTypeUnset = '\000\347\377\376',
	PowerPointMsoBulletTypeBulletTypeNone = '\000\350\000\000',
	PowerPointMsoBulletTypeBulletTypeUnnumbered = '\000\350\000\001',
	PowerPointMsoBulletTypeBulletTypeNumbered = '\000\350\000\002',
	PowerPointMsoBulletTypePictureBulletType = '\000\350\000\003'
};
typedef enum PowerPointMsoBulletType PowerPointMsoBulletType;

enum PowerPointMsoNumberedBulletStyle {
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleUnset = '\000\350\377\376',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicPlain = '\000\351\000\015',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	PowerPointMsoNumberedBulletStyleNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum PowerPointMsoNumberedBulletStyle PowerPointMsoNumberedBulletStyle;

enum PowerPointMsoTabStopType {
	PowerPointMsoTabStopTypeTabstopUnset = '\000\364\377\376',
	PowerPointMsoTabStopTypeTabstopLeft = '\000\365\000\001',
	PowerPointMsoTabStopTypeTabstopCenter = '\000\365\000\002',
	PowerPointMsoTabStopTypeTabstopRight = '\000\365\000\003',
	PowerPointMsoTabStopTypeTabstopDecimal = '\000\365\000\004'
};
typedef enum PowerPointMsoTabStopType PowerPointMsoTabStopType;

enum PowerPointMsoReflectionType {
	PowerPointMsoReflectionTypeReflectionUnset = '\003\351\377\376',
	PowerPointMsoReflectionTypeReflectionTypeNone = '\003\352\000\000',
	PowerPointMsoReflectionTypeReflectionType1 = '\003\352\000\001',
	PowerPointMsoReflectionTypeReflectionType2 = '\003\352\000\002',
	PowerPointMsoReflectionTypeReflectionType3 = '\003\352\000\003',
	PowerPointMsoReflectionTypeReflectionType4 = '\003\352\000\004',
	PowerPointMsoReflectionTypeReflectionType5 = '\003\352\000\005',
	PowerPointMsoReflectionTypeReflectionType6 = '\003\352\000\006',
	PowerPointMsoReflectionTypeReflectionType7 = '\003\352\000\007',
	PowerPointMsoReflectionTypeReflectionType8 = '\003\352\000\010',
	PowerPointMsoReflectionTypeReflectionType9 = '\003\352\000\011'
};
typedef enum PowerPointMsoReflectionType PowerPointMsoReflectionType;

enum PowerPointMsoTextureAlignment {
	PowerPointMsoTextureAlignmentTextureUnset = '\003\352\377\376',
	PowerPointMsoTextureAlignmentTextureTopLeft = '\003\353\000\000',
	PowerPointMsoTextureAlignmentTextureTop = '\003\353\000\001',
	PowerPointMsoTextureAlignmentTextureTopRight = '\003\353\000\002',
	PowerPointMsoTextureAlignmentTextureLeft = '\003\353\000\003',
	PowerPointMsoTextureAlignmentTextureCenter = '\003\353\000\004',
	PowerPointMsoTextureAlignmentTextureRight = '\003\353\000\005',
	PowerPointMsoTextureAlignmentTextureBottomLeft = '\003\353\000\006',
	PowerPointMsoTextureAlignmentTextureBotton = '\003\353\000\007',
	PowerPointMsoTextureAlignmentTextureBottomRight = '\003\353\000\010'
};
typedef enum PowerPointMsoTextureAlignment PowerPointMsoTextureAlignment;

enum PowerPointMsoBaselineAlignment {
	PowerPointMsoBaselineAlignmentTextBaselineAlignmentUnset = '\003\353\377\376',
	PowerPointMsoBaselineAlignmentTextBaselineAlignBaseline = '\003\354\000\001',
	PowerPointMsoBaselineAlignmentTextBaselineAlignTop = '\003\354\000\002',
	PowerPointMsoBaselineAlignmentTextBaselineAlignCenter = '\003\354\000\003',
	PowerPointMsoBaselineAlignmentTextBaselineAlignEastAsian50 = '\003\354\000\004',
	PowerPointMsoBaselineAlignmentTextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum PowerPointMsoBaselineAlignment PowerPointMsoBaselineAlignment;

enum PowerPointMsoClipboardFormat {
	PowerPointMsoClipboardFormatClipboardFormatUnset = '\003\354\377\376',
	PowerPointMsoClipboardFormatNativeClipboardFormat = '\003\355\000\001',
	PowerPointMsoClipboardFormatHTMlClipboardFormat = '\003\355\000\002',
	PowerPointMsoClipboardFormatRTFClipboardFormat = '\003\355\000\003',
	PowerPointMsoClipboardFormatPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum PowerPointMsoClipboardFormat PowerPointMsoClipboardFormat;

enum PowerPointMsoTextRangeInsertPosition {
	PowerPointMsoTextRangeInsertPositionInsertBefore = '\003\356\000\000',
	PowerPointMsoTextRangeInsertPositionInsertAfter = '\003\356\000\001'
};
typedef enum PowerPointMsoTextRangeInsertPosition PowerPointMsoTextRangeInsertPosition;

enum PowerPointMsoPictureType {
	PowerPointMsoPictureTypeSaveAsDefault = '\003\362\377\376',
	PowerPointMsoPictureTypeSaveAsPNGFile = '\003\363\000\000',
	PowerPointMsoPictureTypeSaveAsBMPFile = '\003\363\000\001',
	PowerPointMsoPictureTypeSaveAsGIFFile = '\003\363\000\002',
	PowerPointMsoPictureTypeSaveAsJPGFile = '\003\363\000\003',
	PowerPointMsoPictureTypeSaveAsPDFFile = '\003\363\000\004'
};
typedef enum PowerPointMsoPictureType PowerPointMsoPictureType;

enum PowerPointMsoPictureEffectType {
	PowerPointMsoPictureEffectTypeNoEffect = '\003\364\000\000',
	PowerPointMsoPictureEffectTypeEffectBackgroundRemoval = '\003\364\000\001',
	PowerPointMsoPictureEffectTypeEffectBlur = '\003\364\000\002',
	PowerPointMsoPictureEffectTypeEffectBrightnessContrast = '\003\364\000\003',
	PowerPointMsoPictureEffectTypeEffectCement = '\003\364\000\004',
	PowerPointMsoPictureEffectTypeEffectCrisscrossEtching = '\003\364\000\005',
	PowerPointMsoPictureEffectTypeEffectChalkSketch = '\003\364\000\006',
	PowerPointMsoPictureEffectTypeEffectColorTemperature = '\003\364\000\007',
	PowerPointMsoPictureEffectTypeEffectCutout = '\003\364\000\010',
	PowerPointMsoPictureEffectTypeEffectFilmGrain = '\003\364\000\011',
	PowerPointMsoPictureEffectTypeEffectGlass = '\003\364\000\012',
	PowerPointMsoPictureEffectTypeEffectGlowDiffused = '\003\364\000\013',
	PowerPointMsoPictureEffectTypeEffectGlowEdges = '\003\364\000\014',
	PowerPointMsoPictureEffectTypeEffectLightScreen = '\003\364\000\015',
	PowerPointMsoPictureEffectTypeEffectLineDrawing = '\003\364\000\016',
	PowerPointMsoPictureEffectTypeEffectMarker = '\003\364\000\017',
	PowerPointMsoPictureEffectTypeEffectMosiaicBubbles = '\003\364\000\020',
	PowerPointMsoPictureEffectTypeEffectPaintBrush = '\003\364\000\021',
	PowerPointMsoPictureEffectTypeEffectPaintStrokes = '\003\364\000\022',
	PowerPointMsoPictureEffectTypeEffectPastelsSmooth = '\003\364\000\023',
	PowerPointMsoPictureEffectTypeEffectPencilGrayscale = '\003\364\000\024',
	PowerPointMsoPictureEffectTypeEffectPencilSketch = '\003\364\000\025',
	PowerPointMsoPictureEffectTypeEffectPhotocopy = '\003\364\000\026',
	PowerPointMsoPictureEffectTypeEffectPlasticWrap = '\003\364\000\027',
	PowerPointMsoPictureEffectTypeEffectSaturation = '\003\364\000\030',
	PowerPointMsoPictureEffectTypeEffectSharpenSoften = '\003\364\000\031',
	PowerPointMsoPictureEffectTypeEffectTexturizer = '\003\364\000\032',
	PowerPointMsoPictureEffectTypeEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum PowerPointMsoPictureEffectType PowerPointMsoPictureEffectType;

enum PowerPointMsoSegmentType {
	PowerPointMsoSegmentTypeLine = '\000\217\000\000',
	PowerPointMsoSegmentTypeCurve = '\000\217\000\001'
};
typedef enum PowerPointMsoSegmentType PowerPointMsoSegmentType;

enum PowerPointMsoEditingType {
	PowerPointMsoEditingTypeAuto = '\000\220\000\000',
	PowerPointMsoEditingTypeCorner = '\000\220\000\001',
	PowerPointMsoEditingTypeSmooth = '\000\220\000\002',
	PowerPointMsoEditingTypeSymmetric = '\000\220\000\003'
};
typedef enum PowerPointMsoEditingType PowerPointMsoEditingType;

enum PowerPointMsoSmartArtNodePosition {
	PowerPointMsoSmartArtNodePositionDefaultNodePosition = '\003\365\000\001',
	PowerPointMsoSmartArtNodePositionAfterNode = '\003\365\000\002',
	PowerPointMsoSmartArtNodePositionBeforeNode = '\003\365\000\003',
	PowerPointMsoSmartArtNodePositionAboveNode = '\003\365\000\004',
	PowerPointMsoSmartArtNodePositionBelowNode = '\003\365\000\005'
};
typedef enum PowerPointMsoSmartArtNodePosition PowerPointMsoSmartArtNodePosition;

enum PowerPointMsoSmartArtNodeType {
	PowerPointMsoSmartArtNodeTypeDefaultNode = '\003\366\000\001',
	PowerPointMsoSmartArtNodeTypeAssistantNode = '\003\366\000\002'
};
typedef enum PowerPointMsoSmartArtNodeType PowerPointMsoSmartArtNodeType;

enum PowerPointMsoOrgChartLayoutType {
	PowerPointMsoOrgChartLayoutTypeOrgChartLayoutUnset = '\003\366\377\376',
	PowerPointMsoOrgChartLayoutTypeOrgChartLayoutStandard = '\003\367\000\001',
	PowerPointMsoOrgChartLayoutTypeOrgChartLayoutBothHanging = '\003\367\000\002',
	PowerPointMsoOrgChartLayoutTypeOrgChartLayoutLeftHanging = '\003\367\000\003',
	PowerPointMsoOrgChartLayoutTypeOrgChartLayoutRightHanging = '\003\367\000\004',
	PowerPointMsoOrgChartLayoutTypeOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum PowerPointMsoOrgChartLayoutType PowerPointMsoOrgChartLayoutType;

enum PowerPointMsoAlignCmd {
	PowerPointMsoAlignCmdAlignLefts = '\000\000\000\000',
	PowerPointMsoAlignCmdAlignCenters = '\000\000\000\001',
	PowerPointMsoAlignCmdAlignRights = '\000\000\000\002',
	PowerPointMsoAlignCmdAlignTops = '\000\000\000\003',
	PowerPointMsoAlignCmdAlignMiddles = '\000\000\000\004',
	PowerPointMsoAlignCmdAlignBottoms = '\000\000\000\005'
};
typedef enum PowerPointMsoAlignCmd PowerPointMsoAlignCmd;

enum PowerPointMsoDistributeCmd {
	PowerPointMsoDistributeCmdDistributeHorizontally = '\000\000\000\000',
	PowerPointMsoDistributeCmdDistributeVertically = '\000\000\000\001'
};
typedef enum PowerPointMsoDistributeCmd PowerPointMsoDistributeCmd;

enum PowerPointMsoOrientation {
	PowerPointMsoOrientationOrientationUnset = '\000\214\377\376',
	PowerPointMsoOrientationHorizontalOrientation = '\000\215\000\001',
	PowerPointMsoOrientationVerticalOrientation = '\000\215\000\002'
};
typedef enum PowerPointMsoOrientation PowerPointMsoOrientation;

enum PowerPointMsoZOrderCmd {
	PowerPointMsoZOrderCmdBringShapeToFront = '\000p\000\000',
	PowerPointMsoZOrderCmdSendShapeToBack = '\000p\000\001',
	PowerPointMsoZOrderCmdBringShapeForward = '\000p\000\002',
	PowerPointMsoZOrderCmdSendShapeBackward = '\000p\000\003',
	PowerPointMsoZOrderCmdBringShapeInFrontOfText = '\000p\000\004',
	PowerPointMsoZOrderCmdSendShapeBehindText = '\000p\000\005'
};
typedef enum PowerPointMsoZOrderCmd PowerPointMsoZOrderCmd;

enum PowerPointMsoScriptLanguage {
	PowerPointMsoScriptLanguageBringShapeToFront = '\000o\000\001',
	PowerPointMsoScriptLanguageSendShapeToBack = '\000o\000\002',
	PowerPointMsoScriptLanguageBringShapeForward = '\000o\000\003',
	PowerPointMsoScriptLanguageSendShapeBackward = '\000o\000\004'
};
typedef enum PowerPointMsoScriptLanguage PowerPointMsoScriptLanguage;

enum PowerPointMsoFlipCmd {
	PowerPointMsoFlipCmdFlipHorizontal = '\000q\000\000',
	PowerPointMsoFlipCmdFlipVertical = '\000q\000\001'
};
typedef enum PowerPointMsoFlipCmd PowerPointMsoFlipCmd;

enum PowerPointMsoTriState {
	PowerPointMsoTriStateTrue = '\000\240\377\377',
	PowerPointMsoTriStateFalse = '\000\241\000\000',
	PowerPointMsoTriStateCTrue = '\000\241\000\001',
	PowerPointMsoTriStateToggle = '\000\240\377\375',
	PowerPointMsoTriStateTriStateUnset = '\000\240\377\376'
};
typedef enum PowerPointMsoTriState PowerPointMsoTriState;

enum PowerPointMsoBlackWhiteMode {
	PowerPointMsoBlackWhiteModeBlackAndWhiteUnset = '\000\254\377\376',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeAutomatic = '\000\255\000\001',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeGrayScale = '\000\255\000\002',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeHighContrast = '\000\255\000\007',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeBlack = '\000\255\000\010',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeWhite = '\000\255\000\011',
	PowerPointMsoBlackWhiteModeBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum PowerPointMsoBlackWhiteMode PowerPointMsoBlackWhiteMode;

enum PowerPointMsoBarPosition {
	PowerPointMsoBarPositionBarLeft = '\000r\000\000',
	PowerPointMsoBarPositionBarTop = '\000r\000\001',
	PowerPointMsoBarPositionBarRight = '\000r\000\002',
	PowerPointMsoBarPositionBarBottom = '\000r\000\003',
	PowerPointMsoBarPositionBarFloating = '\000r\000\004',
	PowerPointMsoBarPositionBarPopUp = '\000r\000\005',
	PowerPointMsoBarPositionBarMenu = '\000r\000\006'
};
typedef enum PowerPointMsoBarPosition PowerPointMsoBarPosition;

enum PowerPointMsoBarProtection {
	PowerPointMsoBarProtectionNoProtection = '\000s\000\000',
	PowerPointMsoBarProtectionNoCustomize = '\000s\000\001',
	PowerPointMsoBarProtectionNoResize = '\000s\000\002',
	PowerPointMsoBarProtectionNoMove = '\000s\000\004',
	PowerPointMsoBarProtectionNoChangeVisible = '\000s\000\010',
	PowerPointMsoBarProtectionNoChangeDock = '\000s\000\020',
	PowerPointMsoBarProtectionNoVerticalDock = '\000s\000 ',
	PowerPointMsoBarProtectionNoHorizontalDock = '\000s\000@'
};
typedef enum PowerPointMsoBarProtection PowerPointMsoBarProtection;

enum PowerPointMsoBarType {
	PowerPointMsoBarTypeNormalCommandBar = '\000t\000\000',
	PowerPointMsoBarTypeMenubarCommandBar = '\000t\000\001',
	PowerPointMsoBarTypePopupCommandBar = '\000t\000\002'
};
typedef enum PowerPointMsoBarType PowerPointMsoBarType;

enum PowerPointMsoControlType {
	PowerPointMsoControlTypeControlCustom = '\000u\000\000',
	PowerPointMsoControlTypeControlButton = '\000u\000\001',
	PowerPointMsoControlTypeControlEdit = '\000u\000\002',
	PowerPointMsoControlTypeControlDropDown = '\000u\000\003',
	PowerPointMsoControlTypeControlCombobox = '\000u\000\004',
	PowerPointMsoControlTypeButtonDropDown = '\000u\000\005',
	PowerPointMsoControlTypeSplitDropDown = '\000u\000\006',
	PowerPointMsoControlTypeOCXDropDown = '\000u\000\007',
	PowerPointMsoControlTypeGenericDropDown = '\000u\000\010',
	PowerPointMsoControlTypeGraphicDropDown = '\000u\000\011',
	PowerPointMsoControlTypeControlPopup = '\000u\000\012',
	PowerPointMsoControlTypeGraphicPopup = '\000u\000\013',
	PowerPointMsoControlTypeButtonPopup = '\000u\000\014',
	PowerPointMsoControlTypeSplitButtonPopup = '\000u\000\015',
	PowerPointMsoControlTypeSplitButtonMRUPopup = '\000u\000\016',
	PowerPointMsoControlTypeControlLabel = '\000u\000\017',
	PowerPointMsoControlTypeExpandingGrid = '\000u\000\020',
	PowerPointMsoControlTypeSplitExpandingGrid = '\000u\000\021',
	PowerPointMsoControlTypeControlGrid = '\000u\000\022',
	PowerPointMsoControlTypeControlGauge = '\000u\000\023',
	PowerPointMsoControlTypeGraphicCombobox = '\000u\000\024',
	PowerPointMsoControlTypeControlPane = '\000u\000\025',
	PowerPointMsoControlTypeActiveX = '\000u\000\026',
	PowerPointMsoControlTypeControlGroup = '\000u\000\027',
	PowerPointMsoControlTypeControlTab = '\000u\000\030',
	PowerPointMsoControlTypeControlSpinner = '\000u\000\031'
};
typedef enum PowerPointMsoControlType PowerPointMsoControlType;

enum PowerPointMsoButtonState {
	PowerPointMsoButtonStateButtonStateUp = '\000v\000\000',
	PowerPointMsoButtonStateButtonStateDown = '\000u\377\377',
	PowerPointMsoButtonStateButtonStateUnset = '\000v\000\002'
};
typedef enum PowerPointMsoButtonState PowerPointMsoButtonState;

enum PowerPointMsoControlOLEUsage {
	PowerPointMsoControlOLEUsageNeither = '\000w\000\000',
	PowerPointMsoControlOLEUsageServer = '\000w\000\001',
	PowerPointMsoControlOLEUsageClient = '\000w\000\002',
	PowerPointMsoControlOLEUsageBoth = '\000w\000\003'
};
typedef enum PowerPointMsoControlOLEUsage PowerPointMsoControlOLEUsage;

enum PowerPointMsoButtonStyle {
	PowerPointMsoButtonStyleButtonAutomatic = '\000x\000\000',
	PowerPointMsoButtonStyleButtonIcon = '\000x\000\001',
	PowerPointMsoButtonStyleButtonCaption = '\000x\000\002',
	PowerPointMsoButtonStyleButtonIconAndCaption = '\000x\000\003'
};
typedef enum PowerPointMsoButtonStyle PowerPointMsoButtonStyle;

enum PowerPointMsoComboStyle {
	PowerPointMsoComboStyleComboboxStyleNormal = '\000y\000\000',
	PowerPointMsoComboStyleComboboxStyleLabel = '\000y\000\001'
};
typedef enum PowerPointMsoComboStyle PowerPointMsoComboStyle;

enum PowerPointMsoMenuAnimation {
	PowerPointMsoMenuAnimationNone = '\000{\000\000',
	PowerPointMsoMenuAnimationRandom = '\000{\000\001',
	PowerPointMsoMenuAnimationUnfold = '\000{\000\002',
	PowerPointMsoMenuAnimationSlide = '\000{\000\003'
};
typedef enum PowerPointMsoMenuAnimation PowerPointMsoMenuAnimation;

enum PowerPointMsoHyperlinkType {
	PowerPointMsoHyperlinkTypeHyperlinkTypeTextRange = '\000\226\000\000',
	PowerPointMsoHyperlinkTypeHyperlinkTypeShape = '\000\226\000\001',
	PowerPointMsoHyperlinkTypeHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum PowerPointMsoHyperlinkType PowerPointMsoHyperlinkType;

enum PowerPointMsoExtraInfoMethod {
	PowerPointMsoExtraInfoMethodAppendString = '\000\256\000\000',
	PowerPointMsoExtraInfoMethodPostString = '\000\256\000\001'
};
typedef enum PowerPointMsoExtraInfoMethod PowerPointMsoExtraInfoMethod;

enum PowerPointMsoAnimationType {
	PowerPointMsoAnimationTypeIdle = '\000|\000\001',
	PowerPointMsoAnimationTypeGreeting = '\000|\000\002',
	PowerPointMsoAnimationTypeGoodbye = '\000|\000\003',
	PowerPointMsoAnimationTypeBeginSpeaking = '\000|\000\004',
	PowerPointMsoAnimationTypeCharacterSuccessMajor = '\000|\000\006',
	PowerPointMsoAnimationTypeGetAttentionMajor = '\000|\000\013',
	PowerPointMsoAnimationTypeGetAttentionMinor = '\000|\000\014',
	PowerPointMsoAnimationTypeSearching = '\000|\000\015',
	PowerPointMsoAnimationTypePrinting = '\000|\000\022',
	PowerPointMsoAnimationTypeGestureRight = '\000|\000\023',
	PowerPointMsoAnimationTypeWritingNotingSomething = '\000|\000\026',
	PowerPointMsoAnimationTypeWorkingAtSomething = '\000|\000\027',
	PowerPointMsoAnimationTypeThinking = '\000|\000\030',
	PowerPointMsoAnimationTypeSendingMail = '\000|\000\031',
	PowerPointMsoAnimationTypeListensToComputer = '\000|\000\032',
	PowerPointMsoAnimationTypeDisappear = '\000|\000\037',
	PowerPointMsoAnimationTypeAppear = '\000|\000 ',
	PowerPointMsoAnimationTypeGetArtsy = '\000|\000d',
	PowerPointMsoAnimationTypeGetTechy = '\000|\000e',
	PowerPointMsoAnimationTypeGetWizardy = '\000|\000f',
	PowerPointMsoAnimationTypeCheckingSomething = '\000|\000g',
	PowerPointMsoAnimationTypeLookDown = '\000|\000h',
	PowerPointMsoAnimationTypeLookDownLeft = '\000|\000i',
	PowerPointMsoAnimationTypeLookDownRight = '\000|\000j',
	PowerPointMsoAnimationTypeLookLeft = '\000|\000k',
	PowerPointMsoAnimationTypeLookRight = '\000|\000l',
	PowerPointMsoAnimationTypeLookUp = '\000|\000m',
	PowerPointMsoAnimationTypeLookUpLeft = '\000|\000n',
	PowerPointMsoAnimationTypeLookUpRight = '\000|\000o',
	PowerPointMsoAnimationTypeSaving = '\000|\000p',
	PowerPointMsoAnimationTypeGestureDown = '\000|\000q',
	PowerPointMsoAnimationTypeGestureLeft = '\000|\000r',
	PowerPointMsoAnimationTypeGestureUp = '\000|\000s',
	PowerPointMsoAnimationTypeEmptyTrash = '\000|\000t'
};
typedef enum PowerPointMsoAnimationType PowerPointMsoAnimationType;

enum PowerPointMsoButtonSetType {
	PowerPointMsoButtonSetTypeButtonNone = '\000}\000\000',
	PowerPointMsoButtonSetTypeButtonOk = '\000}\000\001',
	PowerPointMsoButtonSetTypeButtonCancel = '\000}\000\002',
	PowerPointMsoButtonSetTypeButtonsOkCancel = '\000}\000\003',
	PowerPointMsoButtonSetTypeButtonsYesNo = '\000}\000\004',
	PowerPointMsoButtonSetTypeButtonsYesNoCancel = '\000}\000\005',
	PowerPointMsoButtonSetTypeButtonsBackClose = '\000}\000\006',
	PowerPointMsoButtonSetTypeButtonsNextClose = '\000}\000\007',
	PowerPointMsoButtonSetTypeButtonsBackNextClose = '\000}\000\010',
	PowerPointMsoButtonSetTypeButtonsRetryCancel = '\000}\000\011',
	PowerPointMsoButtonSetTypeButtonsAbortRetryIgnore = '\000}\000\012',
	PowerPointMsoButtonSetTypeButtonsSearchClose = '\000}\000\013',
	PowerPointMsoButtonSetTypeButtonsBackNextSnooze = '\000}\000\014',
	PowerPointMsoButtonSetTypeButtonsTipsOptionsClose = '\000}\000\015',
	PowerPointMsoButtonSetTypeButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum PowerPointMsoButtonSetType PowerPointMsoButtonSetType;

enum PowerPointMsoIconType {
	PowerPointMsoIconTypeIconNone = '\000~\000\000',
	PowerPointMsoIconTypeIconApplication = '\000~\000\001',
	PowerPointMsoIconTypeIconAlert = '\000~\000\002',
	PowerPointMsoIconTypeIconTip = '\000~\000\003',
	PowerPointMsoIconTypeIconAlertCritical = '\000~\000e',
	PowerPointMsoIconTypeIconAlertWarning = '\000~\000g',
	PowerPointMsoIconTypeIconAlertInfo = '\000~\000h'
};
typedef enum PowerPointMsoIconType PowerPointMsoIconType;

enum PowerPointMsoWizardActType {
	PowerPointMsoWizardActTypeInactive = '\000\202\000\000',
	PowerPointMsoWizardActTypeActive = '\000\202\000\001',
	PowerPointMsoWizardActTypeSuspend = '\000\202\000\002',
	PowerPointMsoWizardActTypeResume = '\000\202\000\003'
};
typedef enum PowerPointMsoWizardActType PowerPointMsoWizardActType;

enum PowerPointMsoDocProperties {
	PowerPointMsoDocPropertiesPropertyTypeNumber = '\000\242\000\001',
	PowerPointMsoDocPropertiesPropertyTypeBoolean = '\000\242\000\002',
	PowerPointMsoDocPropertiesPropertyTypeDate = '\000\242\000\003',
	PowerPointMsoDocPropertiesPropertyTypeString = '\000\242\000\004',
	PowerPointMsoDocPropertiesPropertyTypeFloat = '\000\242\000\005'
};
typedef enum PowerPointMsoDocProperties PowerPointMsoDocProperties;

enum PowerPointMsoAutomationSecurity {
	PowerPointMsoAutomationSecurityMsoAutomationSecurityLow = '\000\243\000\001',
	PowerPointMsoAutomationSecurityMsoAutomationSecurityByUI = '\000\243\000\002',
	PowerPointMsoAutomationSecurityMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum PowerPointMsoAutomationSecurity PowerPointMsoAutomationSecurity;

enum PowerPointMsoScreenSize {
	PowerPointMsoScreenSizeResolution544x376 = '\000\204\000\000',
	PowerPointMsoScreenSizeResolution640x480 = '\000\204\000\001',
	PowerPointMsoScreenSizeResolution720x512 = '\000\204\000\002',
	PowerPointMsoScreenSizeResolution800x600 = '\000\204\000\003',
	PowerPointMsoScreenSizeResolution1024x768 = '\000\204\000\004',
	PowerPointMsoScreenSizeResolution1152x882 = '\000\204\000\005',
	PowerPointMsoScreenSizeResolution1152x900 = '\000\204\000\006',
	PowerPointMsoScreenSizeResolution1280x1024 = '\000\204\000\007',
	PowerPointMsoScreenSizeResolution1600x1200 = '\000\204\000\010',
	PowerPointMsoScreenSizeResolution1800x1440 = '\000\204\000\011',
	PowerPointMsoScreenSizeResolution1920x1200 = '\000\204\000\012'
};
typedef enum PowerPointMsoScreenSize PowerPointMsoScreenSize;

enum PowerPointMsoCharacterSet {
	PowerPointMsoCharacterSetArabicCharacterSet = '\000\205\000\001',
	PowerPointMsoCharacterSetCyrillicCharacterSet = '\000\205\000\002',
	PowerPointMsoCharacterSetEnglishCharacterSet = '\000\205\000\003',
	PowerPointMsoCharacterSetGreekCharacterSet = '\000\205\000\004',
	PowerPointMsoCharacterSetHebrewCharacterSet = '\000\205\000\005',
	PowerPointMsoCharacterSetJapaneseCharacterSet = '\000\205\000\006',
	PowerPointMsoCharacterSetKoreanCharacterSet = '\000\205\000\007',
	PowerPointMsoCharacterSetMultilingualUnicodeCharacterSet = '\000\205\000\010',
	PowerPointMsoCharacterSetSimplifiedChineseCharacterSet = '\000\205\000\011',
	PowerPointMsoCharacterSetThaiCharacterSet = '\000\205\000\012',
	PowerPointMsoCharacterSetTraditionalChineseCharacterSet = '\000\205\000\013',
	PowerPointMsoCharacterSetVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum PowerPointMsoCharacterSet PowerPointMsoCharacterSet;

enum PowerPointMsoEncoding {
	PowerPointMsoEncodingEncodingThai = '\000\213\003j',
	PowerPointMsoEncodingEncodingJapaneseShiftJIS = '\000\213\003\244',
	PowerPointMsoEncodingEncodingSimplifiedChinese = '\000\213\003\250',
	PowerPointMsoEncodingEncodingKorean = '\000\213\003\265',
	PowerPointMsoEncodingEncodingBig5TraditionalChinese = '\000\213\003\266',
	PowerPointMsoEncodingEncodingLittleEndian = '\000\213\004\260',
	PowerPointMsoEncodingEncodingBigEndian = '\000\213\004\261',
	PowerPointMsoEncodingEncodingCentralEuropean = '\000\213\004\342',
	PowerPointMsoEncodingEncodingCyrillic = '\000\213\004\343',
	PowerPointMsoEncodingEncodingWestern = '\000\213\004\344',
	PowerPointMsoEncodingEncodingGreek = '\000\213\004\345',
	PowerPointMsoEncodingEncodingTurkish = '\000\213\004\346',
	PowerPointMsoEncodingEncodingHebrew = '\000\213\004\347',
	PowerPointMsoEncodingEncodingArabic = '\000\213\004\350',
	PowerPointMsoEncodingEncodingBaltic = '\000\213\004\351',
	PowerPointMsoEncodingEncodingVietnamese = '\000\213\004\352',
	PowerPointMsoEncodingEncodingISO88591Latin1 = '\000\213o\257',
	PowerPointMsoEncodingEncodingISO88592CentralEurope = '\000\213o\260',
	PowerPointMsoEncodingEncodingISO88593Latin3 = '\000\213o\261',
	PowerPointMsoEncodingEncodingISO88594Baltic = '\000\213o\262',
	PowerPointMsoEncodingEncodingISO88595Cyrillic = '\000\213o\263',
	PowerPointMsoEncodingEncodingISO88596Arabic = '\000\213o\264',
	PowerPointMsoEncodingEncodingISO88597Greek = '\000\213o\265',
	PowerPointMsoEncodingEncodingISO88598Hebrew = '\000\213o\266',
	PowerPointMsoEncodingEncodingISO88599Turkish = '\000\213o\267',
	PowerPointMsoEncodingEncodingISO885915Latin9 = '\000\213o\275',
	PowerPointMsoEncodingEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	PowerPointMsoEncodingEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	PowerPointMsoEncodingEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	PowerPointMsoEncodingEncodingISO2022KR = '\000\213\3041',
	PowerPointMsoEncodingEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	PowerPointMsoEncodingEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	PowerPointMsoEncodingEncodingMacRoman = '\000\213\'\020',
	PowerPointMsoEncodingEncodingMacJapanese = '\000\213\'\021',
	PowerPointMsoEncodingEncodingMacTraditionalChinese = '\000\213\'\022',
	PowerPointMsoEncodingEncodingMacKorean = '\000\213\'\023',
	PowerPointMsoEncodingEncodingMacArabic = '\000\213\'\024',
	PowerPointMsoEncodingEncodingMacHebrew = '\000\213\'\025',
	PowerPointMsoEncodingEncodingMacGreek1 = '\000\213\'\026',
	PowerPointMsoEncodingEncodingMacCyrillic = '\000\213\'\027',
	PowerPointMsoEncodingEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	PowerPointMsoEncodingEncodingMacRomania = '\000\213\'\032',
	PowerPointMsoEncodingEncodingMacUkraine = '\000\213\'!',
	PowerPointMsoEncodingEncodingMacLatin2 = '\000\213\'-',
	PowerPointMsoEncodingEncodingMacIcelandic = '\000\213\'_',
	PowerPointMsoEncodingEncodingMacTurkish = '\000\213\'a',
	PowerPointMsoEncodingEncodingMacCroatia = '\000\213\'b',
	PowerPointMsoEncodingEncodingEBCDICUSCanada = '\000\213\000%',
	PowerPointMsoEncodingEncodingEBCDICInternational = '\000\213\001\364',
	PowerPointMsoEncodingEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	PowerPointMsoEncodingEncodingEBCDICGreekModern = '\000\213\003k',
	PowerPointMsoEncodingEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	PowerPointMsoEncodingEncodingEBCDICGermany = '\000\213O1',
	PowerPointMsoEncodingEncodingEBCDICDenmarkNorway = '\000\213O5',
	PowerPointMsoEncodingEncodingEBCDICFinlandSweden = '\000\213O6',
	PowerPointMsoEncodingEncodingEBCDICItaly = '\000\213O8',
	PowerPointMsoEncodingEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	PowerPointMsoEncodingEncodingEBCDICUnitedKingdom = '\000\213O=',
	PowerPointMsoEncodingEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	PowerPointMsoEncodingEncodingEBCDICFrance = '\000\213OI',
	PowerPointMsoEncodingEncodingEBCDICArabic = '\000\213O\304',
	PowerPointMsoEncodingEncodingEBCDICGreek = '\000\213O\307',
	PowerPointMsoEncodingEncodingEBCDICHebrew = '\000\213O\310',
	PowerPointMsoEncodingEncodingEBCDICKoreanExtended = '\000\213Qa',
	PowerPointMsoEncodingEncodingEBCDICThai = '\000\213Qf',
	PowerPointMsoEncodingEncodingEBCDICIcelandic = '\000\213Q\207',
	PowerPointMsoEncodingEncodingEBCDICTurkish = '\000\213Q\251',
	PowerPointMsoEncodingEncodingEBCDICRussian = '\000\213Q\220',
	PowerPointMsoEncodingEncodingEBCDICSerbianBulgarian = '\000\213R!',
	PowerPointMsoEncodingEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	PowerPointMsoEncodingEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	PowerPointMsoEncodingEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	PowerPointMsoEncodingEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	PowerPointMsoEncodingEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	PowerPointMsoEncodingEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	PowerPointMsoEncodingEncodingOEMUnitedStates = '\000\213\001\265',
	PowerPointMsoEncodingEncodingOEMGreek = '\000\213\002\341',
	PowerPointMsoEncodingEncodingOEMBaltic = '\000\213\003\007',
	PowerPointMsoEncodingEncodingOEMMultilingualLatinI = '\000\213\003R',
	PowerPointMsoEncodingEncodingOEMMultilingualLatinII = '\000\213\003T',
	PowerPointMsoEncodingEncodingOEMCyrillic = '\000\213\003W',
	PowerPointMsoEncodingEncodingOEMTurkish = '\000\213\003Y',
	PowerPointMsoEncodingEncodingOEMPortuguese = '\000\213\003\\',
	PowerPointMsoEncodingEncodingOEMIcelandic = '\000\213\003]',
	PowerPointMsoEncodingEncodingOEMHebrew = '\000\213\003^',
	PowerPointMsoEncodingEncodingOEMCanadianFrench = '\000\213\003_',
	PowerPointMsoEncodingEncodingOEMArabic = '\000\213\003`',
	PowerPointMsoEncodingEncodingOEMNordic = '\000\213\003a',
	PowerPointMsoEncodingEncodingOEMCyrillicII = '\000\213\003b',
	PowerPointMsoEncodingEncodingOEMModernGreek = '\000\213\003e',
	PowerPointMsoEncodingEncodingEUCJapanese = '\000\213\312\334',
	PowerPointMsoEncodingEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	PowerPointMsoEncodingEncodingEUCKorean = '\000\213\312\355',
	PowerPointMsoEncodingEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	PowerPointMsoEncodingEncodingDevanagari = '\000\213\336\252',
	PowerPointMsoEncodingEncodingBengali = '\000\213\336\253',
	PowerPointMsoEncodingEncodingTamil = '\000\213\336\254',
	PowerPointMsoEncodingEncodingTelugu = '\000\213\336\255',
	PowerPointMsoEncodingEncodingAssamese = '\000\213\336\256',
	PowerPointMsoEncodingEncodingOriya = '\000\213\336\257',
	PowerPointMsoEncodingEncodingKannada = '\000\213\336\260',
	PowerPointMsoEncodingEncodingMalayalam = '\000\213\336\261',
	PowerPointMsoEncodingEncodingGujarati = '\000\213\336\262',
	PowerPointMsoEncodingEncodingPunjabi = '\000\213\336\263',
	PowerPointMsoEncodingEncodingArabicASMO = '\000\213\002\304',
	PowerPointMsoEncodingEncodingArabicTransparentASMO = '\000\213\002\320',
	PowerPointMsoEncodingEncodingKoreanJohab = '\000\213\005Q',
	PowerPointMsoEncodingEncodingTaiwanCNS = '\000\213N ',
	PowerPointMsoEncodingEncodingTaiwanTCA = '\000\213N!',
	PowerPointMsoEncodingEncodingTaiwanEten = '\000\213N\"',
	PowerPointMsoEncodingEncodingTaiwanIBM5550 = '\000\213N#',
	PowerPointMsoEncodingEncodingTaiwanTeletext = '\000\213N$',
	PowerPointMsoEncodingEncodingTaiwanWang = '\000\213N%',
	PowerPointMsoEncodingEncodingIA5IRV = '\000\213N\211',
	PowerPointMsoEncodingEncodingIA5German = '\000\213N\212',
	PowerPointMsoEncodingEncodingIA5Swedish = '\000\213N\213',
	PowerPointMsoEncodingEncodingIA5Norwegian = '\000\213N\214',
	PowerPointMsoEncodingEncodingUSASCII = '\000\213N\237',
	PowerPointMsoEncodingEncodingT61 = '\000\213O%',
	PowerPointMsoEncodingEncodingISO6937NonspacingAccent = '\000\213O-',
	PowerPointMsoEncodingEncodingKOI8R = '\000\213Q\202',
	PowerPointMsoEncodingEncodingExtAlphaLowercase = '\000\213R#',
	PowerPointMsoEncodingEncodingKOI8U = '\000\213Uj',
	PowerPointMsoEncodingEncodingEuropa3 = '\000\213qI',
	PowerPointMsoEncodingEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	PowerPointMsoEncodingEncodingUTF7 = '\000\213\375\350',
	PowerPointMsoEncodingEncodingUTF8 = '\000\213\375\351'
};
typedef enum PowerPointMsoEncoding PowerPointMsoEncoding;

enum PowerPointReset {
	PowerPointResetCommandBar = 'msCB',
	PowerPointResetCommandBarControl = 'mCBC'
};
typedef enum PowerPointReset PowerPointReset;

enum PowerPointEPPWindowState {
	PowerPointEPPWindowStateWindowNormal = '\000\311\000\001',
	PowerPointEPPWindowStateWindowMinimized = '\000\311\000\002'
};
typedef enum PowerPointEPPWindowState PowerPointEPPWindowState;

enum PowerPointEPPArrangeStyle {
	PowerPointEPPArrangeStyleArrangeTiled = '\000\321\000\001',
	PowerPointEPPArrangeStyleArrangeCascade = '\000\321\000\002'
};
typedef enum PowerPointEPPArrangeStyle PowerPointEPPArrangeStyle;

enum PowerPointEPPViewType {
	PowerPointEPPViewTypeSlideView = '\000\312\000\001',
	PowerPointEPPViewTypeMasterView = '\000\312\000\002',
	PowerPointEPPViewTypePageView = '\000\312\000\003',
	PowerPointEPPViewTypeHandoutMasterView = '\000\312\000\004',
	PowerPointEPPViewTypeNotesMasterView = '\000\312\000\005',
	PowerPointEPPViewTypeOutlineView = '\000\312\000\006',
	PowerPointEPPViewTypeSlideSorterView = '\000\312\000\007',
	PowerPointEPPViewTypeTitleMasterView = '\000\312\000\010',
	PowerPointEPPViewTypeNormalView = '\000\312\000\011',
	PowerPointEPPViewTypePrintPreview = '\000\312\000\012',
	PowerPointEPPViewTypeThumbnailView = '\000\312\000\013',
	PowerPointEPPViewTypeThumbnailMasterView = '\000\312\000\014'
};
typedef enum PowerPointEPPViewType PowerPointEPPViewType;

enum PowerPointEPPColorSchemeIndex {
	PowerPointEPPColorSchemeIndexSchemeColorUnset = '\000\362\377\376',
	PowerPointEPPColorSchemeIndexNotASchemeColor = '\000\363\000\000',
	PowerPointEPPColorSchemeIndexBackgroundScheme = '\000\363\000\001',
	PowerPointEPPColorSchemeIndexForegroundScheme = '\000\363\000\002',
	PowerPointEPPColorSchemeIndexShadowScheme = '\000\363\000\003',
	PowerPointEPPColorSchemeIndexTitleScheme = '\000\363\000\004',
	PowerPointEPPColorSchemeIndexFillScheme = '\000\363\000\005',
	PowerPointEPPColorSchemeIndexAccent1Scheme = '\000\363\000\006',
	PowerPointEPPColorSchemeIndexAccent2Scheme = '\000\363\000\007',
	PowerPointEPPColorSchemeIndexAccent3Scheme = '\000\363\000\010'
};
typedef enum PowerPointEPPColorSchemeIndex PowerPointEPPColorSchemeIndex;

enum PowerPointEPPSlideSizeType {
	PowerPointEPPSlideSizeTypeSlideSizeOnScreen = '\000\313\000\001',
	PowerPointEPPSlideSizeTypeSlideSizeLetterPaper = '\000\313\000\002',
	PowerPointEPPSlideSizeTypeSlideSizeA4Paper = '\000\313\000\003',
	PowerPointEPPSlideSizeTypeSlideSize35MM = '\000\313\000\004',
	PowerPointEPPSlideSizeTypeSlideSizeOverhead = '\000\313\000\005',
	PowerPointEPPSlideSizeTypeSlideSizeBanner = '\000\313\000\006',
	PowerPointEPPSlideSizeTypeSlideSizeCustom = '\000\313\000\007',
	PowerPointEPPSlideSizeTypeSlideSizeLedgerPaper = '\000\313\000\010',
	PowerPointEPPSlideSizeTypeSlideSizeA3Paper = '\000\313\000\011',
	PowerPointEPPSlideSizeTypeSlideSizeB4ISOPaper = '\000\313\000\012',
	PowerPointEPPSlideSizeTypeSlideSizeB5ISOPaper = '\000\313\000\013',
	PowerPointEPPSlideSizeTypeSlideSizeB4JISPaper = '\000\313\000\014',
	PowerPointEPPSlideSizeTypeSlideSizeB5JISPaper = '\000\313\000\015',
	PowerPointEPPSlideSizeTypeSlideSizeHagakiCard = '\000\313\000\016',
	PowerPointEPPSlideSizeTypeSlideSizeOnScreen16x9 = '\000\313\000\017',
	PowerPointEPPSlideSizeTypeSlideSizeOnScreen16x10 = '\000\313\000\020'
};
typedef enum PowerPointEPPSlideSizeType PowerPointEPPSlideSizeType;

enum PowerPointEPPSaveAsFileType {
	PowerPointEPPSaveAsFileTypeSaveAsPresentation = '\000\314\000\001',
	PowerPointEPPSaveAsFileTypeSaveAsTemplate = '\000\314\000\005',
	PowerPointEPPSaveAsFileTypeSaveAsRTF = '\000\314\000\006',
	PowerPointEPPSaveAsFileTypeSaveAsShow = '\000\314\000\007',
	PowerPointEPPSaveAsFileTypeSaveAsAddIn = '\000\314\000\010',
	PowerPointEPPSaveAsFileTypeSaveAsDefault = '\000\314\000\012',
	PowerPointEPPSaveAsFileTypeSaveAsHTML = '\000\314\000\013',
	PowerPointEPPSaveAsFileTypeSaveAsMovie = '\000\314\000\014',
	PowerPointEPPSaveAsFileTypeSaveAsPackage = '\000\314\000\015',
	PowerPointEPPSaveAsFileTypeSaveAsPDF = '\000\314\000\016',
	PowerPointEPPSaveAsFileTypeSaveAsOpenXMLPresentation = '\000\314\000\017',
	PowerPointEPPSaveAsFileTypeSaveAsOpenXMLPresentationMacroEnabled = '\000\314\000\020',
	PowerPointEPPSaveAsFileTypeSaveAsOpenXMLShow = '\000\314\000\021',
	PowerPointEPPSaveAsFileTypeSaveAsOpenXMLShowMacroEnabled = '\000\314\000\022',
	PowerPointEPPSaveAsFileTypeSaveAsOpenXMLTemplate = '\000\314\000\023',
	PowerPointEPPSaveAsFileTypeSaveAsOpenXMLTemplateMacroEnabled = '\000\314\000\024',
	PowerPointEPPSaveAsFileTypeSaveAsOpenXMLTheme = '\000\314\000\025',
	PowerPointEPPSaveAsFileTypeSaveAsGIF = '\000\314\000\026',
	PowerPointEPPSaveAsFileTypeSaveAsJPG = '\000\314\000\027',
	PowerPointEPPSaveAsFileTypeSaveAsPNG = '\000\314\000\030',
	PowerPointEPPSaveAsFileTypeSaveAsBMP = '\000\314\000\031',
	PowerPointEPPSaveAsFileTypeSaveAsTIF = '\000\314\000\032'
};
typedef enum PowerPointEPPSaveAsFileType PowerPointEPPSaveAsFileType;

enum PowerPointPpTextStyleType {
	PowerPointPpTextStyleTypeTextStyleDefault = '\001*\000\001',
	PowerPointPpTextStyleTypeTextStyleTitle = '\001*\000\002',
	PowerPointPpTextStyleTypeTextStyleBody = '\001*\000\003'
};
typedef enum PowerPointPpTextStyleType PowerPointPpTextStyleType;

enum PowerPointEPPSlideLayout {
	PowerPointEPPSlideLayoutSlideLayoutUnset = '\000\314\377\376',
	PowerPointEPPSlideLayoutSlideLayoutTitleSlide = '\000\315\000\001',
	PowerPointEPPSlideLayoutSlideLayoutTextSlide = '\000\315\000\002',
	PowerPointEPPSlideLayoutSlideLayoutTwoColumnText = '\000\315\000\003',
	PowerPointEPPSlideLayoutSlideLayoutTable = '\000\315\000\004',
	PowerPointEPPSlideLayoutSlideLayoutTextAndChart = '\000\315\000\005',
	PowerPointEPPSlideLayoutSlideLayoutChartAndText = '\000\315\000\006',
	PowerPointEPPSlideLayoutSlideLayoutOrgchart = '\000\315\000\007',
	PowerPointEPPSlideLayoutSlideLayoutChart = '\000\315\000\010',
	PowerPointEPPSlideLayoutSlideLayoutTextAndClipart = '\000\315\000\011',
	PowerPointEPPSlideLayoutSlideLayoutClipartAndText = '\000\315\000\012',
	PowerPointEPPSlideLayoutSlideLayoutTitleOnly = '\000\315\000\013',
	PowerPointEPPSlideLayoutSlideLayoutBlank = '\000\315\000\014',
	PowerPointEPPSlideLayoutSlideLayoutTextAndObject = '\000\315\000\015',
	PowerPointEPPSlideLayoutSlideLayoutObjectAndText = '\000\315\000\016',
	PowerPointEPPSlideLayoutSlideLayoutLargeObject = '\000\315\000\017',
	PowerPointEPPSlideLayoutSlideLayoutObject = '\000\315\000\020',
	PowerPointEPPSlideLayoutSlideLayoutMediaClip = '\000\315\000\021',
	PowerPointEPPSlideLayoutSlideLayoutMediaClipAndText = '\000\315\000\022',
	PowerPointEPPSlideLayoutSlideLayoutObjectOverText = '\000\315\000\023',
	PowerPointEPPSlideLayoutSlideLayoutTextOverObject = '\000\315\000\024',
	PowerPointEPPSlideLayoutSlideLayoutTextAndTwoObjects = '\000\315\000\025',
	PowerPointEPPSlideLayoutSlideLayoutTwoObjectsAndText = '\000\315\000\026',
	PowerPointEPPSlideLayoutSlideLayoutTwoObjectsOverText = '\000\315\000\027',
	PowerPointEPPSlideLayoutSlideLayoutFourObjects = '\000\315\000\030',
	PowerPointEPPSlideLayoutSlideLayoutVerticalText = '\000\315\000\031',
	PowerPointEPPSlideLayoutSlideLayoutClipartAndVerticalText = '\000\315\000\032',
	PowerPointEPPSlideLayoutSlideLayoutVerticalTitleAndText = '\000\315\000\033',
	PowerPointEPPSlideLayoutSlideLayoutVerticalTitleAndTextOverChart = '\000\315\000\034',
	PowerPointEPPSlideLayoutSlideLayoutTwoObjects = '\000\315\000\035',
	PowerPointEPPSlideLayoutSlideLayoutObjectAndTwoObjects = '\000\315\000\036',
	PowerPointEPPSlideLayoutSlideLayoutTwoObjectsAndObject = '\000\315\000\037',
	PowerPointEPPSlideLayoutSlideLayoutCustom = '\000\315\000 ',
	PowerPointEPPSlideLayoutSlideLayoutSectionHeader = '\000\315\000!',
	PowerPointEPPSlideLayoutSlideLayoutComparison = '\000\315\000\"',
	PowerPointEPPSlideLayoutSlideLayoutContentWithCaption = '\000\315\000#',
	PowerPointEPPSlideLayoutSlideLayoutPictureWithCaption = '\000\315\000$'
};
typedef enum PowerPointEPPSlideLayout PowerPointEPPSlideLayout;

enum PowerPointEPPEntryEffect {
	PowerPointEPPEntryEffectEntryEffectAppear = '\000\366\017\004',
	PowerPointEPPEntryEffectEntryEffectBlindsHorizontal = '\000\366\003\001',
	PowerPointEPPEntryEffectEntryEffectBlindsVertical = '\000\366\003\002',
	PowerPointEPPEntryEffectEntryEffectBoxDown = '\000\366\017U',
	PowerPointEPPEntryEffectEntryEffectBoxIn = '\000\366\014\002',
	PowerPointEPPEntryEffectEntryEffectBoxLeft = '\000\366\017R',
	PowerPointEPPEntryEffectEntryEffectBoxOut = '\000\366\014\001',
	PowerPointEPPEntryEffectEntryEffectBoxRight = '\000\366\017T',
	PowerPointEPPEntryEffectEntryEffectBoxUp = '\000\366\017S',
	PowerPointEPPEntryEffectEntryEffectCheckerboardAcross = '\000\366\004\001',
	PowerPointEPPEntryEffectEntryEffectCheckerboardDown = '\000\366\004\002',
	PowerPointEPPEntryEffectEntryEffectCircle = '\000\366\017\005',
	PowerPointEPPEntryEffectEntryEffectCollapseAcross = '\000\366\015\027',
	PowerPointEPPEntryEffectEntryEffectCollapseBottom = '\000\366\015\033',
	PowerPointEPPEntryEffectEntryEffectCollapseLeft = '\000\366\015\030',
	PowerPointEPPEntryEffectEntryEffectCollapseRight = '\000\366\015\032',
	PowerPointEPPEntryEffectEntryEffectCollapseUp = '\000\366\015\031',
	PowerPointEPPEntryEffectEntryEffectCombHorizontal = '\000\366\017\007',
	PowerPointEPPEntryEffectEntryEffectCombVertical = '\000\366\017\010',
	PowerPointEPPEntryEffectEntryEffectConveyorLeft = '\000\366\017*',
	PowerPointEPPEntryEffectEntryEffectConveyorRight = '\000\366\017+',
	PowerPointEPPEntryEffectEntryEffectCoverDown = '\000\366\005\004',
	PowerPointEPPEntryEffectEntryEffectCoverLeftDown = '\000\366\005\007',
	PowerPointEPPEntryEffectEntryEffectCoverLeftUp = '\000\366\005\005',
	PowerPointEPPEntryEffectEntryEffectCoverLeft = '\000\366\005\001',
	PowerPointEPPEntryEffectEntryEffectCoverRightDown = '\000\366\005\010',
	PowerPointEPPEntryEffectEntryEffectCoverRightUp = '\000\366\005\006',
	PowerPointEPPEntryEffectEntryEffectCoverRight = '\000\366\005\003',
	PowerPointEPPEntryEffectEntryEffectCoverUp = '\000\366\005\002',
	PowerPointEPPEntryEffectEntryEffectCrawlFromDown = '\000\366\015\020',
	PowerPointEPPEntryEffectEntryEffectCrawlFromLeft = '\000\366\015\015',
	PowerPointEPPEntryEffectEntryEffectCrawlFromRight = '\000\366\015\017',
	PowerPointEPPEntryEffectEntryEffectCrawlFromUp = '\000\366\015\016',
	PowerPointEPPEntryEffectEntryEffectCubeDown = '\000\366\017M',
	PowerPointEPPEntryEffectEntryEffectCubeLeft = '\000\366\017J',
	PowerPointEPPEntryEffectEntryEffectCubeRight = '\000\366\017K',
	PowerPointEPPEntryEffectEntryEffectCubeUp = '\000\366\017L',
	PowerPointEPPEntryEffectEntryEffectCutBlack = '\000\366\001\002',
	PowerPointEPPEntryEffectEntryEffectCut = '\000\366\001\001',
	PowerPointEPPEntryEffectEntryEffectDiamond = '\000\366\017\006',
	PowerPointEPPEntryEffectEntryEffectDissolve = '\000\366\006\001',
	PowerPointEPPEntryEffectEntryEffectDoorsHorizontal = '\000\366\017-',
	PowerPointEPPEntryEffectEntryEffectDoorsVertical = '\000\366\017,',
	PowerPointEPPEntryEffectEntryEffectFadeBlack = '\000\366\007\001',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromBottomLeft = '\000\366\015$',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromBottomRight = '\000\366\015%',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromBottom = '\000\366\015!',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromLeft = '\000\366\015\036',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromRight = '\000\366\015 ',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromTopLeft = '\000\366\015\"',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromTopRight = '\000\366\015#',
	PowerPointEPPEntryEffectEntryEffectFadeFlyFromTop = '\000\366\015\037',
	PowerPointEPPEntryEffectEntryEffectFadeSmoothly = '\000\366\017\011',
	PowerPointEPPEntryEffectEntryEffectFade = '\000\366\017\011',
	PowerPointEPPEntryEffectEntryEffectFerrisWheelLeft = '\000\366\017;',
	PowerPointEPPEntryEffectEntryEffectFerrisWheelRight = '\000\366\017<',
	PowerPointEPPEntryEffectEntryEffectFlashOnceFast = '\000\366\017\001',
	PowerPointEPPEntryEffectEntryEffectFlashOnceMedium = '\000\366\017\002',
	PowerPointEPPEntryEffectEntryEffectFlashOnceSlow = '\000\366\017\003',
	PowerPointEPPEntryEffectEntryEffectFlashbulb = '\000\366\017E',
	PowerPointEPPEntryEffectEntryEffectFlipDown = '\000\366\017D',
	PowerPointEPPEntryEffectEntryEffectFlipLeft = '\000\366\017A',
	PowerPointEPPEntryEffectEntryEffectFlipRight = '\000\366\017B',
	PowerPointEPPEntryEffectEntryEffectFlipUp = '\000\366\017C',
	PowerPointEPPEntryEffectEntryEffectFlyFromBottomLeft = '\000\366\015\007',
	PowerPointEPPEntryEffectEntryEffectFlyFromBottomRight = '\000\366\015\010',
	PowerPointEPPEntryEffectEntryEffectFlyFromBottom = '\000\366\015\004',
	PowerPointEPPEntryEffectEntryEffectFlyFromLeft = '\000\366\015\001',
	PowerPointEPPEntryEffectEntryEffectFlyFromRight = '\000\366\015\003',
	PowerPointEPPEntryEffectEntryEffectFlyFromTopLeft = '\000\366\015\005',
	PowerPointEPPEntryEffectEntryEffectFlyFromTopRight = '\000\366\015\006',
	PowerPointEPPEntryEffectEntryEffectFlyFromTop = '\000\366\015\002',
	PowerPointEPPEntryEffectEntryEffectFlyThroughInBounce = '\000\366\0174',
	PowerPointEPPEntryEffectEntryEffectFlyThroughIn = '\000\366\0172',
	PowerPointEPPEntryEffectEntryEffectFlyThroughOutBounce = '\000\366\0175',
	PowerPointEPPEntryEffectEntryEffectFlyThroughOut = '\000\366\0173',
	PowerPointEPPEntryEffectEntryEffectGalleryLeft = '\000\366\017(',
	PowerPointEPPEntryEffectEntryEffectGalleryRight = '\000\366\017)',
	PowerPointEPPEntryEffectEntryEffectGlitterDiamondDown = '\000\366\017#',
	PowerPointEPPEntryEffectEntryEffectGlitterDiamondLeft = '\000\366\017 ',
	PowerPointEPPEntryEffectEntryEffectGlitterDiamondRight = '\000\366\017\"',
	PowerPointEPPEntryEffectEntryEffectGlitterDiamondUp = '\000\366\017!',
	PowerPointEPPEntryEffectEntryEffectGlitterHexagonDown = '\000\366\017\'',
	PowerPointEPPEntryEffectEntryEffectGlitterHexagonLeft = '\000\366\017$',
	PowerPointEPPEntryEffectEntryEffectGlitterHexagonRight = '\000\366\017&',
	PowerPointEPPEntryEffectEntryEffectGlitterHexagonUp = '\000\366\017%',
	PowerPointEPPEntryEffectEntryEffectHoneycomb = '\000\366\017:',
	PowerPointEPPEntryEffectEntryEffectNewsFlash = '\000\366\017\012',
	PowerPointEPPEntryEffectEntryEffectNone = '\000\366\000\000',
	PowerPointEPPEntryEffectEntryEffectOrbitDown = '\000\366\017Y',
	PowerPointEPPEntryEffectEntryEffectOrbitLeft = '\000\366\017V',
	PowerPointEPPEntryEffectEntryEffectOrbitRight = '\000\366\017X',
	PowerPointEPPEntryEffectEntryEffectOrbitUp = '\000\366\017W',
	PowerPointEPPEntryEffectEntryEffectPanDown = '\000\366\017]',
	PowerPointEPPEntryEffectEntryEffectPanLeft = '\000\366\017Z',
	PowerPointEPPEntryEffectEntryEffectPanRight = '\000\366\017\\',
	PowerPointEPPEntryEffectEntryEffectPanUp = '\000\366\017[',
	PowerPointEPPEntryEffectEntryEffectPeekFromDown = '\000\366\015\012',
	PowerPointEPPEntryEffectEntryEffectPeekFromLeft = '\000\366\015\011',
	PowerPointEPPEntryEffectEntryEffectPeekFromRight = '\000\366\015\013',
	PowerPointEPPEntryEffectEntryEffectPeekFromUp = '\000\366\015\014',
	PowerPointEPPEntryEffectEntryEffectPlus = '\000\366\017\013',
	PowerPointEPPEntryEffectEntryEffectPushDown = '\000\366\017\014',
	PowerPointEPPEntryEffectEntryEffectPushLeft = '\000\366\017\015',
	PowerPointEPPEntryEffectEntryEffectPushRight = '\000\366\017\016',
	PowerPointEPPEntryEffectEntryEffectPushUp = '\000\366\017\017',
	PowerPointEPPEntryEffectEntryEffectRandomBarsHorizontal = '\000\366\011\001',
	PowerPointEPPEntryEffectEntryEffectRandomBarsVertical = '\000\366\011\002',
	PowerPointEPPEntryEffectEntryEffectRandom = '\000\366\002\001',
	PowerPointEPPEntryEffectEntryEffectRevealBlackLeft = '\000\366\0178',
	PowerPointEPPEntryEffectEntryEffectRevealBlackRight = '\000\366\0179',
	PowerPointEPPEntryEffectEntryEffectRevealSmoothLeft = '\000\366\0176',
	PowerPointEPPEntryEffectEntryEffectRevealSmoothRight = '\000\366\0177',
	PowerPointEPPEntryEffectEntryEffectRippleCenter = '\000\366\017\033',
	PowerPointEPPEntryEffectEntryEffectRippleLeftDown = '\000\366\017\036',
	PowerPointEPPEntryEffectEntryEffectRippleLeftUp = '\000\366\017\035',
	PowerPointEPPEntryEffectEntryEffectRippleRightDown = '\000\366\017\037',
	PowerPointEPPEntryEffectEntryEffectRippleRightUp = '\000\366\017\034',
	PowerPointEPPEntryEffectEntryEffectRotateDown = '\000\366\017Q',
	PowerPointEPPEntryEffectEntryEffectRotateLeft = '\000\366\017N',
	PowerPointEPPEntryEffectEntryEffectRotateRight = '\000\366\017P',
	PowerPointEPPEntryEffectEntryEffectRotateUp = '\000\366\017O',
	PowerPointEPPEntryEffectEntryEffectShredRectangleIn = '\000\366\017H',
	PowerPointEPPEntryEffectEntryEffectShredRectangleOut = '\000\366\017I',
	PowerPointEPPEntryEffectEntryEffectShredStripsIn = '\000\366\017F',
	PowerPointEPPEntryEffectEntryEffectShredStripsOut = '\000\366\017G',
	PowerPointEPPEntryEffectEntryEffectSpinner = '\000\366\017\012',
	PowerPointEPPEntryEffectEntryEffectSpiral = '\000\366\015\035',
	PowerPointEPPEntryEffectEntryEffectSplitHorizontalIn = '\000\366\016\002',
	PowerPointEPPEntryEffectEntryEffectSplitHorizontalOut = '\000\366\016\001',
	PowerPointEPPEntryEffectEntryEffectSplitVerticalIn = '\000\366\016\004',
	PowerPointEPPEntryEffectEntryEffectSplitVerticalOut = '\000\366\016\003',
	PowerPointEPPEntryEffectEntryEffectStretchAcross = '\000\366\015\027',
	PowerPointEPPEntryEffectEntryEffectStretchDown = '\000\366\015\033',
	PowerPointEPPEntryEffectEntryEffectStretchLeft = '\000\366\015\030',
	PowerPointEPPEntryEffectEntryEffectStretchRight = '\000\366\015\032',
	PowerPointEPPEntryEffectEntryEffectStretchUp = '\000\366\015\031',
	PowerPointEPPEntryEffectEntryEffectStripsDownLeft = '\000\366\012\003',
	PowerPointEPPEntryEffectEntryEffectStripsDownRight = '\000\366\012\004',
	PowerPointEPPEntryEffectEntryEffectStripsLeftDown = '\000\366\012\007',
	PowerPointEPPEntryEffectEntryEffectStripsLeftUp = '\000\366\012\005',
	PowerPointEPPEntryEffectEntryEffectStripsRightDown = '\000\366\012\010',
	PowerPointEPPEntryEffectEntryEffectStripsRightUp = '\000\366\012\006',
	PowerPointEPPEntryEffectEntryEffectStripsUpLeft = '\000\366\012\001',
	PowerPointEPPEntryEffectEntryEffectStripsUpRight = '\000\366\012\002',
	PowerPointEPPEntryEffectEntryEffectSwitchDown = '\000\366\017@',
	PowerPointEPPEntryEffectEntryEffectSwitchLeft = '\000\366\017=',
	PowerPointEPPEntryEffectEntryEffectSwitchRight = '\000\366\017\?',
	PowerPointEPPEntryEffectEntryEffectSwitchUp = '\000\366\017>',
	PowerPointEPPEntryEffectEntryEffectSwivel = '\000\366\015\034',
	PowerPointEPPEntryEffectEntryEffectUncoverDown = '\000\366\010\004',
	PowerPointEPPEntryEffectEntryEffectUncoverLeftDown = '\000\366\010\007',
	PowerPointEPPEntryEffectEntryEffectUncoverLeftUp = '\000\366\010\005',
	PowerPointEPPEntryEffectEntryEffectUncoverLeft = '\000\366\010\001',
	PowerPointEPPEntryEffectEntryEffectUncoverRightDown = '\000\366\010\010',
	PowerPointEPPEntryEffectEntryEffectUncoverRightUp = '\000\366\010\006',
	PowerPointEPPEntryEffectEntryEffectUncoverRight = '\000\366\010\003',
	PowerPointEPPEntryEffectEntryEffectUncoverUp = '\000\366\010\002',
	PowerPointEPPEntryEffectEntryEffectUnset = '\000\365\377\376',
	PowerPointEPPEntryEffectEntryEffectVortexDown = '\000\366\017\032',
	PowerPointEPPEntryEffectEntryEffectVortexLeft = '\000\366\017\027',
	PowerPointEPPEntryEffectEntryEffectVortexRight = '\000\366\017\031',
	PowerPointEPPEntryEffectEntryEffectVortexUp = '\000\366\017\030',
	PowerPointEPPEntryEffectEntryEffectWarpIn = '\000\366\0170',
	PowerPointEPPEntryEffectEntryEffectWarpOut = '\000\366\0171',
	PowerPointEPPEntryEffectEntryEffectWedge = '\000\366\017\020',
	PowerPointEPPEntryEffectEntryEffectWheel1Spoke = '\000\366\017\021',
	PowerPointEPPEntryEffectEntryEffectWheel2Spokes = '\000\366\017\022',
	PowerPointEPPEntryEffectEntryEffectWheel3Spokes = '\000\366\017\023',
	PowerPointEPPEntryEffectEntryEffectWheel4Spokes = '\000\366\017\024',
	PowerPointEPPEntryEffectEntryEffectWheel8Spokes = '\000\366\017\025',
	PowerPointEPPEntryEffectEntryEffectWheelReverse1Spoke = '\000\366\017\026',
	PowerPointEPPEntryEffectEntryEffectWindowHorizontal = '\000\366\017/',
	PowerPointEPPEntryEffectEntryEffectWindowVertical = '\000\366\017.',
	PowerPointEPPEntryEffectEntryEffectWipeDown = '\000\366\013\004',
	PowerPointEPPEntryEffectEntryEffectWipeLeft = '\000\366\013\001',
	PowerPointEPPEntryEffectEntryEffectWipeRight = '\000\366\013\003',
	PowerPointEPPEntryEffectEntryEffectWipeUp = '\000\366\013\002',
	PowerPointEPPEntryEffectEntryEffectZoomBottom = '\000\366\015\026',
	PowerPointEPPEntryEffectEntryEffectZoomCenter = '\000\366\015\025',
	PowerPointEPPEntryEffectEntryEffectZoomInSlightly = '\000\366\015\022',
	PowerPointEPPEntryEffectEntryEffectZoomIn = '\000\366\015\021',
	PowerPointEPPEntryEffectEntryEffectZoomOutSlightly = '\000\366\015\024',
	PowerPointEPPEntryEffectEntryEffectZoomOut = '\000\366\015\023'
};
typedef enum PowerPointEPPEntryEffect PowerPointEPPEntryEffect;

enum PowerPointEPPTextLevelEffect {
	PowerPointEPPTextLevelEffectAnimationLevelUnset = '\000\337\377\376',
	PowerPointEPPTextLevelEffectAnimateLevelNone = '\000\340\000\000',
	PowerPointEPPTextLevelEffectAnimateLevelFirstLevel = '\000\340\000\001',
	PowerPointEPPTextLevelEffectAnimateLevelSecondLevel = '\000\340\000\002',
	PowerPointEPPTextLevelEffectAnimateLevelThirdLevel = '\000\340\000\003',
	PowerPointEPPTextLevelEffectAnimateLevelFourthLevel = '\000\340\000\004',
	PowerPointEPPTextLevelEffectAnimateLevelFifthLevel = '\000\340\000\005',
	PowerPointEPPTextLevelEffectAnimateLevelAllLevels = '\000\340\000\020'
};
typedef enum PowerPointEPPTextLevelEffect PowerPointEPPTextLevelEffect;

enum PowerPointEPPTextUnitEffect {
	PowerPointEPPTextUnitEffectAnimationUnitUnset = '\000\340\377\376',
	PowerPointEPPTextUnitEffectTextUnitEffectByParagraph = '\000\341\000\000',
	PowerPointEPPTextUnitEffectTextUnitEffectByWord = '\000\341\000\001',
	PowerPointEPPTextUnitEffectTextUnitEffectByCharacter = '\000\341\000\002'
};
typedef enum PowerPointEPPTextUnitEffect PowerPointEPPTextUnitEffect;

enum PowerPointEPPChartUnitEffect {
	PowerPointEPPChartUnitEffectAnimationChartUnset = '\000\341\377\376',
	PowerPointEPPChartUnitEffectChartUnitEffectBySeries = '\000\342\000\001',
	PowerPointEPPChartUnitEffectChartUnitEffectByCategory = '\000\342\000\002',
	PowerPointEPPChartUnitEffectChartUnitEffectBySeriesElement = '\000\342\000\003'
};
typedef enum PowerPointEPPChartUnitEffect PowerPointEPPChartUnitEffect;

enum PowerPointEPPAfterEffect {
	PowerPointEPPAfterEffectAfterEffectUnset = '\000\363\377\376',
	PowerPointEPPAfterEffectAfterEffectNone = '\000\364\000\000',
	PowerPointEPPAfterEffectAfterEffectHide = '\000\364\000\001',
	PowerPointEPPAfterEffectAfterEffectDim = '\000\364\000\002'
};
typedef enum PowerPointEPPAfterEffect PowerPointEPPAfterEffect;

enum PowerPointEPPAdvanceMode {
	PowerPointEPPAdvanceModeAdvanceModeUnset = '\000\361\377\376',
	PowerPointEPPAdvanceModeAdvanceModeOnClick = '\000\362\000\001'
};
typedef enum PowerPointEPPAdvanceMode PowerPointEPPAdvanceMode;

enum PowerPointEPPSoundEffectType {
	PowerPointEPPSoundEffectTypeSoundEffectUnset = '\000\331\377\376',
	PowerPointEPPSoundEffectTypeSoundEffectNone = '\000\332\000\000',
	PowerPointEPPSoundEffectTypeSoundEffectStopPrevious = '\000\332\000\001',
	PowerPointEPPSoundEffectTypeSoundEffectFile = '\000\332\000\002'
};
typedef enum PowerPointEPPSoundEffectType PowerPointEPPSoundEffectType;

enum PowerPointEPPUpdateOption {
	PowerPointEPPUpdateOptionUpdateOptionUnset = '\000\336\377\376',
	PowerPointEPPUpdateOptionUpdateOptionManual = '\000\337\000\001'
};
typedef enum PowerPointEPPUpdateOption PowerPointEPPUpdateOption;

enum PowerPointEPPChangeCase {
	PowerPointEPPChangeCasePpCaseSentence = '\000\344\000\001',
	PowerPointEPPChangeCasePpCaseLower = '\000\344\000\002',
	PowerPointEPPChangeCasePpCaseUpper = '\000\344\000\003',
	PowerPointEPPChangeCasePpCaseTitle = '\000\344\000\004'
};
typedef enum PowerPointEPPChangeCase PowerPointEPPChangeCase;

enum PowerPointEPPDialogMode {
	PowerPointEPPDialogModeDialogModeUnset = '\000\357\377\376',
	PowerPointEPPDialogModeDialogModeModless = '\000\360\000\000',
	PowerPointEPPDialogModeDialogModeModal = '\000\360\000\001'
};
typedef enum PowerPointEPPDialogMode PowerPointEPPDialogMode;

enum PowerPointEPPDialogStyle {
	PowerPointEPPDialogStyleDialogStyleUnset = '\000\360\377\376',
	PowerPointEPPDialogStyleDialogStyleStandard = '\000\361\000\001'
};
typedef enum PowerPointEPPDialogStyle PowerPointEPPDialogStyle;

enum PowerPointEPPSlideShowPointerType {
	PowerPointEPPSlideShowPointerTypeSlideShowPointerNone = '\000\322\000\000',
	PowerPointEPPSlideShowPointerTypeSlideShowPointerArrow = '\000\322\000\001',
	PowerPointEPPSlideShowPointerTypeSlideShowPointerPen = '\000\322\000\002',
	PowerPointEPPSlideShowPointerTypeSlideShowPointerAlwaysHidden = '\000\322\000\003'
};
typedef enum PowerPointEPPSlideShowPointerType PowerPointEPPSlideShowPointerType;

enum PowerPointEPPSlideShowState {
	PowerPointEPPSlideShowStateSlideShowStateRunning = '\000\323\000\001',
	PowerPointEPPSlideShowStateSlideShowStatePaused = '\000\323\000\002',
	PowerPointEPPSlideShowStateSlideShowStateBlackScreen = '\000\323\000\003',
	PowerPointEPPSlideShowStateSlideShowStateWhiteScreen = '\000\323\000\004',
	PowerPointEPPSlideShowStateSlideShowStateDone = '\000\323\000\005'
};
typedef enum PowerPointEPPSlideShowState PowerPointEPPSlideShowState;

enum PowerPointEPPPlayerState {
	PowerPointEPPPlayerStatePlayerStatePlaying = '\000\365\000\000',
	PowerPointEPPPlayerStatePlayerStatePaused = '\000\365\000\001',
	PowerPointEPPPlayerStatePlayerStateStopped = '\000\365\000\002',
	PowerPointEPPPlayerStatePlayerStateNotReady = '\000\365\000\003'
};
typedef enum PowerPointEPPPlayerState PowerPointEPPPlayerState;

enum PowerPointEPPSlideShowAdvanceMode {
	PowerPointEPPSlideShowAdvanceModeSlideShowAdvanceManualAdvance = '\000\324\000\001',
	PowerPointEPPSlideShowAdvanceModeSlideShowAdvanceUseSlideTimings = '\000\324\000\002'
};
typedef enum PowerPointEPPSlideShowAdvanceMode PowerPointEPPSlideShowAdvanceMode;

enum PowerPointEPPPrintOutputType {
	PowerPointEPPPrintOutputTypePrintSlides = '\000\330\000\001',
	PowerPointEPPPrintOutputTypePrintTwoSlideHandouts = '\000\330\000\002',
	PowerPointEPPPrintOutputTypePrintThreeSlideHandouts = '\000\330\000\003',
	PowerPointEPPPrintOutputTypePrintSixSlideHandouts = '\000\330\000\004',
	PowerPointEPPPrintOutputTypePrintNotesPages = '\000\330\000\005',
	PowerPointEPPPrintOutputTypePrintOutline = '\000\330\000\006',
	PowerPointEPPPrintOutputTypePrintFourSlideHandouts = '\000\330\000\007',
	PowerPointEPPPrintOutputTypePrintNineSlideHandouts = '\000\330\000\010'
};
typedef enum PowerPointEPPPrintOutputType PowerPointEPPPrintOutputType;

enum PowerPointEPPPrintColorType {
	PowerPointEPPPrintColorTypePrintColor = '\000\327\000\001',
	PowerPointEPPPrintColorTypePrintBlackAndWhite = '\000\327\000\002'
};
typedef enum PowerPointEPPPrintColorType PowerPointEPPPrintColorType;

enum PowerPointEPPSelectionType {
	PowerPointEPPSelectionTypeSelectionTypeNone = '\000\316\000\000',
	PowerPointEPPSelectionTypeSelectionTypeSlides = '\000\316\000\001',
	PowerPointEPPSelectionTypeSelectionTypeShapes = '\000\316\000\002',
	PowerPointEPPSelectionTypeSelectionTypeText = '\000\316\000\003'
};
typedef enum PowerPointEPPSelectionType PowerPointEPPSelectionType;

enum PowerPointEPPDirection {
	PowerPointEPPDirectionDirectionUnset = '\000\352\377\376',
	PowerPointEPPDirectionLeftToRight = '\000\353\000\001'
};
typedef enum PowerPointEPPDirection PowerPointEPPDirection;

enum PowerPointEPPDateTimeFormat {
	PowerPointEPPDateTimeFormatUnset = '\000\342\377\376',
	PowerPointEPPDateTimeFormatMdyy = '\000\343\000\001',
	PowerPointEPPDateTimeFormatDdddMMMMddyyyy = '\000\343\000\002',
	PowerPointEPPDateTimeFormatMMMMyyyy = '\000\343\000\003',
	PowerPointEPPDateTimeFormatMMMMdyyyy = '\000\343\000\004',
	PowerPointEPPDateTimeFormatMMMyy = '\000\343\000\005',
	PowerPointEPPDateTimeFormatMMMMyy = '\000\343\000\006',
	PowerPointEPPDateTimeFormatMMyy = '\000\343\000\007',
	PowerPointEPPDateTimeFormatMMddyyHmm = '\000\343\000\010',
	PowerPointEPPDateTimeFormatMddyyhmmAMPM = '\000\343\000\011',
	PowerPointEPPDateTimeFormatHmm = '\000\343\000\012',
	PowerPointEPPDateTimeFormatHmmss = '\000\343\000\013',
	PowerPointEPPDateTimeFormatHmmAMPM = '\000\343\000\014',
	PowerPointEPPDateTimeFormatHmmssAMPM = '\000\343\000\015'
};
typedef enum PowerPointEPPDateTimeFormat PowerPointEPPDateTimeFormat;

enum PowerPointEPPTransitionSpeed {
	PowerPointEPPTransitionSpeedTransitionSpeedUnset = '\000\330\377\376',
	PowerPointEPPTransitionSpeedTransistionSpeedSlow = '\000\331\000\001',
	PowerPointEPPTransitionSpeedTransistionSpeedMedium = '\000\331\000\002'
};
typedef enum PowerPointEPPTransitionSpeed PowerPointEPPTransitionSpeed;

enum PowerPointEPPMouseActivation {
	PowerPointEPPMouseActivationMouseActivationMouseClick = '\000\372\000\001',
	PowerPointEPPMouseActivationMouseActivationMouseOver = '\000\372\000\002'
};
typedef enum PowerPointEPPMouseActivation PowerPointEPPMouseActivation;

enum PowerPointEPPActionType {
	PowerPointEPPActionTypeActionTypeUnset = '\000\345\377\376',
	PowerPointEPPActionTypeActionTypeNone = '\000\346\000\000',
	PowerPointEPPActionTypeActionTypeNextSlide = '\000\346\000\001',
	PowerPointEPPActionTypeActionTypePreviousSlide = '\000\346\000\002',
	PowerPointEPPActionTypeActionTypeFirstSlide = '\000\346\000\003',
	PowerPointEPPActionTypeActionTypeLastSlide = '\000\346\000\004',
	PowerPointEPPActionTypeActionTypeLastSlideViewed = '\000\346\000\005',
	PowerPointEPPActionTypeActionTypeEndShow = '\000\346\000\006',
	PowerPointEPPActionTypeActionTypeHyperlinkAction = '\000\346\000\007',
	PowerPointEPPActionTypeActionTypeRunMacro = '\000\346\000\010',
	PowerPointEPPActionTypeActionTypeRunProgram = '\000\346\000\011',
	PowerPointEPPActionTypeActionTypeNamedSlideshowAction = '\000\346\000\012',
	PowerPointEPPActionTypeActionTypeOLEVerb = '\000\346\000\013'
};
typedef enum PowerPointEPPActionType PowerPointEPPActionType;

enum PowerPointEPPPlaceholderType {
	PowerPointEPPPlaceholderTypePlaceholderTypeUnset = '\000\332\377\376',
	PowerPointEPPPlaceholderTypePlaceholderTypeTitlePlaceholder = '\000\333\000\001',
	PowerPointEPPPlaceholderTypePlaceholderTypeBodyPlaceholder = '\000\333\000\002',
	PowerPointEPPPlaceholderTypePlaceholderTypeCenterTitlePlaceholder = '\000\333\000\003',
	PowerPointEPPPlaceholderTypePlaceholderTypeSubtitlePlaceholder = '\000\333\000\004',
	PowerPointEPPPlaceholderTypePlaceholderTypeVerticalTitlePlaceholder = '\000\333\000\005',
	PowerPointEPPPlaceholderTypePlaceholderTypeVerticalBodyPlaceholder = '\000\333\000\006',
	PowerPointEPPPlaceholderTypePlaceholderTypeObjectPlaceholder = '\000\333\000\007',
	PowerPointEPPPlaceholderTypePlaceholderTypeChartPlaceholder = '\000\333\000\010',
	PowerPointEPPPlaceholderTypePlaceholderTypeBitmapPlaceholder = '\000\333\000\011',
	PowerPointEPPPlaceholderTypePlaceholderTypeMediaClipPlaceholder = '\000\333\000\012',
	PowerPointEPPPlaceholderTypePlaceholderTypeOrgChartPlaceholder = '\000\333\000\013',
	PowerPointEPPPlaceholderTypePlaceholderTypeTablePlaceholder = '\000\333\000\014',
	PowerPointEPPPlaceholderTypePlaceholderTypeSlideNumberPlaceholder = '\000\333\000\015',
	PowerPointEPPPlaceholderTypePlaceholderTypeHeaderPlaceholder = '\000\333\000\016',
	PowerPointEPPPlaceholderTypePlaceholderTypeFooterPlaceholder = '\000\333\000\017',
	PowerPointEPPPlaceholderTypePlaceholderTypeDatePlaceholder = '\000\333\000\020',
	PowerPointEPPPlaceholderTypePlaceholderTypeVerticalObjectPlaceholder = '\000\333\000\021',
	PowerPointEPPPlaceholderTypePlaceholderTypePicturePlaceholder = '\000\333\000\022'
};
typedef enum PowerPointEPPPlaceholderType PowerPointEPPPlaceholderType;

enum PowerPointEPPSlideShowType {
	PowerPointEPPSlideShowTypeSlideShowTypeSpeaker = '\000\325\000\001',
	PowerPointEPPSlideShowTypeSlideShowTypeWindow = '\000\325\000\002',
	PowerPointEPPSlideShowTypeSlideShowTypeKiosk = '\000\325\000\003',
	PowerPointEPPSlideShowTypeSlideShowTypePresenter = '\000\325\000\005'
};
typedef enum PowerPointEPPSlideShowType PowerPointEPPSlideShowType;

enum PowerPointEPPPrintRangeType {
	PowerPointEPPPrintRangeTypePrintRangeAll = '\000\367\000\001',
	PowerPointEPPPrintRangeTypePrintRangeSelection = '\000\367\000\002',
	PowerPointEPPPrintRangeTypePrintRangeCurrent = '\000\367\000\003',
	PowerPointEPPPrintRangeTypePrintRangeSlideRange = '\000\367\000\004',
	PowerPointEPPPrintRangeTypePrintSection = '\000\367\000\005'
};
typedef enum PowerPointEPPPrintRangeType PowerPointEPPPrintRangeType;

enum PowerPointEPPAutoSize {
	PowerPointEPPAutoSizePpAutoSizemixed = '\000\344\377\376',
	PowerPointEPPAutoSizePpAutoSizeNone = '\000\345\000\000',
	PowerPointEPPAutoSizePpAutoSizeShapeToFitText = '\000\345\000\001',
	PowerPointEPPAutoSizePpAutoSizeTextToFitShape = '\000\345\000\002'
};
typedef enum PowerPointEPPAutoSize PowerPointEPPAutoSize;

enum PowerPointEPPMediaType {
	PowerPointEPPMediaTypeMediaTypeUnset = '\000\333\377\376',
	PowerPointEPPMediaTypeMediaTypeOther = '\000\334\000\001',
	PowerPointEPPMediaTypeMediaTypeSound = '\000\334\000\002',
	PowerPointEPPMediaTypeMediaTypeMovie = '\000\334\000\003'
};
typedef enum PowerPointEPPMediaType PowerPointEPPMediaType;

enum PowerPointEPPSoundFormatType {
	PowerPointEPPSoundFormatTypeSoundFormatUnset = '\000\367\377\376',
	PowerPointEPPSoundFormatTypeSoundFormatNone = '\000\370\000\000',
	PowerPointEPPSoundFormatTypeSoundFormatWAV = '\000\370\000\001',
	PowerPointEPPSoundFormatTypeSoundFormatMIDI = '\000\370\000\002'
};
typedef enum PowerPointEPPSoundFormatType PowerPointEPPSoundFormatType;

enum PowerPointEPPFarEastLineBreakLevel {
	PowerPointEPPFarEastLineBreakLevelEastAsianLineBreakNormal = '\000\354\000\001',
	PowerPointEPPFarEastLineBreakLevelEastAsianLineBreakStrict = '\000\354\000\002',
	PowerPointEPPFarEastLineBreakLevelEastAsianLineBreakCustom = '\000\354\000\003'
};
typedef enum PowerPointEPPFarEastLineBreakLevel PowerPointEPPFarEastLineBreakLevel;

enum PowerPointEPPSlideShowRangeType {
	PowerPointEPPSlideShowRangeTypeSlideShowRangeShowAll = '\000\326\000\001',
	PowerPointEPPSlideShowRangeTypeSlideShowRange = '\000\326\000\002',
	PowerPointEPPSlideShowRangeTypeSlideShowRangeNamedSlideshow = '\000\326\000\003'
};
typedef enum PowerPointEPPSlideShowRangeType PowerPointEPPSlideShowRangeType;

enum PowerPointEPPFrameColors {
	PowerPointEPPFrameColorsFrameColorsBrowserColors = '\000\317\000\001',
	PowerPointEPPFrameColorsFrameColorsPresentationSchemeTextColor = '\000\317\000\002',
	PowerPointEPPFrameColorsFrameColorsPresentationSchemeAccentColor = '\000\317\000\003',
	PowerPointEPPFrameColorsFrameColorsWhiteTextOnBlack = '\000\317\000\004',
	PowerPointEPPFrameColorsFrameColorsBlackTextOnWhite = '\000\317\000\005'
};
typedef enum PowerPointEPPFrameColors PowerPointEPPFrameColors;

enum PowerPointEPPMovieOptimization {
	PowerPointEPPMovieOptimizationMovieOptimizationNormal = '\000\317\377\376',
	PowerPointEPPMovieOptimizationMovieOptimizationSize = '\000\320\000\001',
	PowerPointEPPMovieOptimizationMovieOptimizationSpeed = '\000\320\000\002',
	PowerPointEPPMovieOptimizationMovieOptimizationQuality = '\000\320\000\003'
};
typedef enum PowerPointEPPMovieOptimization PowerPointEPPMovieOptimization;

enum PowerPointEPPBulletType {
	PowerPointEPPBulletTypePpBulletmixed = '\000\347\377\376',
	PowerPointEPPBulletTypePpBulletNone = '\000\350\000\000',
	PowerPointEPPBulletTypePpBulletUnnumbered = '\000\350\000\001',
	PowerPointEPPBulletTypePpBulletNumbered = '\000\350\000\002',
	PowerPointEPPBulletTypePpBulletPicture = '\000\350\000\003'
};
typedef enum PowerPointEPPBulletType PowerPointEPPBulletType;

enum PowerPointEPPNumberedBulletStyle {
	PowerPointEPPNumberedBulletStylePpBulletStylemixed = '\000\350\377\376',
	PowerPointEPPNumberedBulletStylePpBulletAlphaLCPeriod = '\000\351\000\000',
	PowerPointEPPNumberedBulletStylePpBulletAlphaUCPeriod = '\000\351\000\001',
	PowerPointEPPNumberedBulletStylePpBulletArabicParenRight = '\000\351\000\002',
	PowerPointEPPNumberedBulletStylePpBulletArabicPeriod = '\000\351\000\003',
	PowerPointEPPNumberedBulletStylePpBulletRomanLCParenBoth = '\000\351\000\004',
	PowerPointEPPNumberedBulletStylePpBulletRomanLCParenRight = '\000\351\000\005',
	PowerPointEPPNumberedBulletStylePpBulletRomanLCPeriod = '\000\351\000\006',
	PowerPointEPPNumberedBulletStylePpBulletRomanUCPeriod = '\000\351\000\007',
	PowerPointEPPNumberedBulletStylePpBulletAlphaLCParenBoth = '\000\351\000\010',
	PowerPointEPPNumberedBulletStylePpBulletAlphaLCParenRight = '\000\351\000\011',
	PowerPointEPPNumberedBulletStylePpBulletAlphaUCParenBoth = '\000\351\000\012',
	PowerPointEPPNumberedBulletStylePpBulletAlphaUCParenRight = '\000\351\000\013',
	PowerPointEPPNumberedBulletStylePpBulletArabicParenBoth = '\000\351\000\014',
	PowerPointEPPNumberedBulletStylePpBulletArabicPlain = '\000\351\000\015',
	PowerPointEPPNumberedBulletStylePpBulletRomanUCParenBoth = '\000\351\000\016',
	PowerPointEPPNumberedBulletStylePpBulletRomanUCParenRight = '\000\351\000\017',
	PowerPointEPPNumberedBulletStylePpBulletSimpChinPlain = '\000\351\000\020',
	PowerPointEPPNumberedBulletStylePpBulletSimpChinPeriod = '\000\351\000\021',
	PowerPointEPPNumberedBulletStylePpBulletCircleNumDBPlain = '\000\351\000\022',
	PowerPointEPPNumberedBulletStylePpBulletCircleNumWDWhitePlain = '\000\351\000\023',
	PowerPointEPPNumberedBulletStylePpBulletCircleNumWDBlackPlain = '\000\351\000\024',
	PowerPointEPPNumberedBulletStylePpBulletTradChinPlain = '\000\351\000\025',
	PowerPointEPPNumberedBulletStylePpBulletTradChinPeriod = '\000\351\000\026',
	PowerPointEPPNumberedBulletStylePpBulletArabicAlphaDash = '\000\351\000\027',
	PowerPointEPPNumberedBulletStylePpBulletArabicAbjadDash = '\000\351\000\030',
	PowerPointEPPNumberedBulletStylePpBulletHebrewAlphaDash = '\000\351\000\031',
	PowerPointEPPNumberedBulletStylePpBulletKanjiKoreanPlain = '\000\351\000\032',
	PowerPointEPPNumberedBulletStylePpBulletKanjiKoreanPeriod = '\000\351\000\033',
	PowerPointEPPNumberedBulletStylePpBulletArabicDBPlain = '\000\351\000\034'
};
typedef enum PowerPointEPPNumberedBulletStyle PowerPointEPPNumberedBulletStyle;

enum PowerPointEPPShapeFormat {
	PowerPointEPPShapeFormatShapeFormatGIF = '\000\335\000\000',
	PowerPointEPPShapeFormatShapeFormatJPEG = '\000\335\000\001',
	PowerPointEPPShapeFormatShapeFormatPNG = '\000\335\000\002',
	PowerPointEPPShapeFormatShapeFormatPICT = '\000\335\000\003'
};
typedef enum PowerPointEPPShapeFormat PowerPointEPPShapeFormat;

enum PowerPointEPPExportMode {
	PowerPointEPPExportModeExportModeRelativeToSlide = '\000\336\000\001',
	PowerPointEPPExportModeExportModeClipRelativeToSlide = '\000\336\000\001',
	PowerPointEPPExportModeExportModeScaleToFit = '\000\336\000\002',
	PowerPointEPPExportModeExportModeScaleXY = '\000\336\000\003'
};
typedef enum PowerPointEPPExportMode PowerPointEPPExportMode;

enum PowerPointPpBorderType {
	PowerPointPpBorderTypeTopBorder = '\001\032\000\001',
	PowerPointPpBorderTypeLeftBorder = '\001\032\000\002',
	PowerPointPpBorderTypeBottomBorder = '\001\032\000\003',
	PowerPointPpBorderTypeRightBorder = '\001\032\000\004',
	PowerPointPpBorderTypeDiagonalDownBorder = '\001\032\000\005',
	PowerPointPpBorderTypeDiagonalUpBorder = '\001\032\000\006'
};
typedef enum PowerPointPpBorderType PowerPointPpBorderType;

enum PowerPointEPPCheckInVersionType {
	PowerPointEPPCheckInVersionTypeMinorVersion = '\002\304\000\000',
	PowerPointEPPCheckInVersionTypeMajorVersion = '\002\304\000\001',
	PowerPointEPPCheckInVersionTypeOverwriteCurrentVersion = '\002\304\000\002'
};
typedef enum PowerPointEPPCheckInVersionType PowerPointEPPCheckInVersionType;

enum PowerPointIppPageLayout {
	PowerPointIppPageLayoutPageLayoutNormal = '\000\355\000\000',
	PowerPointIppPageLayoutPageLayoutFullScreen = '\000\355\000\001'
};
typedef enum PowerPointIppPageLayout PowerPointIppPageLayout;

enum PowerPointIppButtonsType {
	PowerPointIppButtonsTypeRegular = '\000\356\000\001',
	PowerPointIppButtonsTypeFancy = '\000\356\000\002',
	PowerPointIppButtonsTypeTextOnly = '\000\356\000\003'
};
typedef enum PowerPointIppButtonsType PowerPointIppButtonsType;

enum PowerPointIppNavBarPlacement {
	PowerPointIppNavBarPlacementBarPlacementTop = '\000\357\000\001',
	PowerPointIppNavBarPlacementBarPlacementBottom = '\000\357\000\002'
};
typedef enum PowerPointIppNavBarPlacement PowerPointIppNavBarPlacement;

enum PowerPointMsoAnimEffect {
	PowerPointMsoAnimEffectAnimationTypeCustom = '\001\002\000\000',
	PowerPointMsoAnimEffectAnimationTypeAppear = '\001\002\000\001',
	PowerPointMsoAnimEffectAnimationTypeFly = '\001\002\000\002',
	PowerPointMsoAnimEffectAnimationTypeBlinds = '\001\002\000\003',
	PowerPointMsoAnimEffectAnimationTypeBox = '\001\002\000\004',
	PowerPointMsoAnimEffectAnimationTypeCheckerboard = '\001\002\000\005',
	PowerPointMsoAnimEffectAnimationTypeCircle = '\001\002\000\006',
	PowerPointMsoAnimEffectAnimationTypeCrawl = '\001\002\000\007',
	PowerPointMsoAnimEffectAnimationTypeDiamond = '\001\002\000\010',
	PowerPointMsoAnimEffectAnimationTypeDissolve = '\001\002\000\011',
	PowerPointMsoAnimEffectAnimationTypeFade = '\001\002\000\012',
	PowerPointMsoAnimEffectAnimationTypeFlashOnce = '\001\002\000\013',
	PowerPointMsoAnimEffectAnimationTypePeek = '\001\002\000\014',
	PowerPointMsoAnimEffectAnimationTypePlus = '\001\002\000\015',
	PowerPointMsoAnimEffectAnimationTypeRandomBars = '\001\002\000\016',
	PowerPointMsoAnimEffectAnimationTypeSpiral = '\001\002\000\017',
	PowerPointMsoAnimEffectAnimationTypeSplit = '\001\002\000\020',
	PowerPointMsoAnimEffectAnimationTypeStretch = '\001\002\000\021',
	PowerPointMsoAnimEffectAnimationTypeStrips = '\001\002\000\022',
	PowerPointMsoAnimEffectAnimationTypeSwivel = '\001\002\000\023',
	PowerPointMsoAnimEffectAnimationTypeWedge = '\001\002\000\024',
	PowerPointMsoAnimEffectAnimationTypeWheel = '\001\002\000\025',
	PowerPointMsoAnimEffectAnimationTypeWipe = '\001\002\000\026',
	PowerPointMsoAnimEffectAnimationTypeZoom = '\001\002\000\027',
	PowerPointMsoAnimEffectAnimationTypeRandomEffect = '\001\002\000\030',
	PowerPointMsoAnimEffectAnimationTypeBoomerang = '\001\002\000\031',
	PowerPointMsoAnimEffectAnimationTypeBounce = '\001\002\000\032',
	PowerPointMsoAnimEffectAnimationTypeColorReveal = '\001\002\000\033',
	PowerPointMsoAnimEffectAnimationTypeCredits = '\001\002\000\034',
	PowerPointMsoAnimEffectAnimationTypeEaseIn = '\001\002\000\035',
	PowerPointMsoAnimEffectAnimationTypeFloat = '\001\002\000\036',
	PowerPointMsoAnimEffectAnimationTypeGrowAndTurn = '\001\002\000\037',
	PowerPointMsoAnimEffectAnimationTypeLightSpeed = '\001\002\000 ',
	PowerPointMsoAnimEffectAnimationTypePinwheel = '\001\002\000!',
	PowerPointMsoAnimEffectAnimationTypeRiseUp = '\001\002\000\"',
	PowerPointMsoAnimEffectAnimationTypeSwish = '\001\002\000#',
	PowerPointMsoAnimEffectAnimationTypeThinLine = '\001\002\000$',
	PowerPointMsoAnimEffectAnimationTypeUnfold = '\001\002\000%',
	PowerPointMsoAnimEffectAnimationTypeWhip = '\001\002\000&',
	PowerPointMsoAnimEffectAnimationTypeAscend = '\001\002\000\'',
	PowerPointMsoAnimEffectAnimationTypeCenterRevolve = '\001\002\000(',
	PowerPointMsoAnimEffectAnimationTypeFadedSwivel = '\001\002\000)',
	PowerPointMsoAnimEffectAnimationTypeDescend = '\001\002\000*',
	PowerPointMsoAnimEffectAnimationTypeSling = '\001\002\000+',
	PowerPointMsoAnimEffectAnimationTypeSpinner = '\001\002\000,',
	PowerPointMsoAnimEffectAnimationTypeStretchy = '\001\002\000-',
	PowerPointMsoAnimEffectAnimationTypeZip = '\001\002\000.',
	PowerPointMsoAnimEffectAnimationTypeArcUp = '\001\002\000/',
	PowerPointMsoAnimEffectAnimationTypeFadeZoom = '\001\002\0000',
	PowerPointMsoAnimEffectAnimationTypeGlide = '\001\002\0001',
	PowerPointMsoAnimEffectAnimationTypeExpand = '\001\002\0002',
	PowerPointMsoAnimEffectAnimationTypeFlip = '\001\002\0003',
	PowerPointMsoAnimEffectAnimationTypeShimmer = '\001\002\0004',
	PowerPointMsoAnimEffectAnimationTypeFold = '\001\002\0005',
	PowerPointMsoAnimEffectAnimationTypeChangeFillColor = '\001\002\0006',
	PowerPointMsoAnimEffectAnimationTypeChangeFont = '\001\002\0007',
	PowerPointMsoAnimEffectAnimationTypeChangeFontColor = '\001\002\0008',
	PowerPointMsoAnimEffectAnimationTypeChangeFontSize = '\001\002\0009',
	PowerPointMsoAnimEffectAnimationTypeChangeFontStyle = '\001\002\000:',
	PowerPointMsoAnimEffectAnimationTypeGrowShrink = '\001\002\000;',
	PowerPointMsoAnimEffectAnimationTypeChangeLineColor = '\001\002\000<',
	PowerPointMsoAnimEffectAnimationTypeSpin = '\001\002\000=',
	PowerPointMsoAnimEffectAnimationTypeTransparency = '\001\002\000>',
	PowerPointMsoAnimEffectAnimationTypeBoldFlash = '\001\002\000\?',
	PowerPointMsoAnimEffectAnimationTypeBlast = '\001\002\000@',
	PowerPointMsoAnimEffectAnimationTypeBoldReveal = '\001\002\000A',
	PowerPointMsoAnimEffectAnimationTypeBrushOnColor = '\001\002\000B',
	PowerPointMsoAnimEffectAnimationTypeBrushOnUnderline = '\001\002\000C',
	PowerPointMsoAnimEffectAnimationTypeColorBlend = '\001\002\000D',
	PowerPointMsoAnimEffectAnimationTypeColorWave = '\001\002\000E',
	PowerPointMsoAnimEffectAnimationTypeComplementaryColor = '\001\002\000F',
	PowerPointMsoAnimEffectAnimationTypeComplementaryColor2 = '\001\002\000G',
	PowerPointMsoAnimEffectAnimationTypeConstrastingColor = '\001\002\000H',
	PowerPointMsoAnimEffectAnimationTypeDarken = '\001\002\000I',
	PowerPointMsoAnimEffectAnimationTypeDesaturate = '\001\002\000J',
	PowerPointMsoAnimEffectAnimationTypeFlashBulb = '\001\002\000K',
	PowerPointMsoAnimEffectAnimationTypeFlicker = '\001\002\000L',
	PowerPointMsoAnimEffectAnimationTypeGrowWithColor = '\001\002\000M',
	PowerPointMsoAnimEffectAnimationTypeLighten = '\001\002\000N',
	PowerPointMsoAnimEffectAnimationTypeStyleEmphasis = '\001\002\000O',
	PowerPointMsoAnimEffectAnimationTypeTeeter = '\001\002\000P',
	PowerPointMsoAnimEffectAnimationTypeVerticalGrow = '\001\002\000Q',
	PowerPointMsoAnimEffectAnimationTypeWave = '\001\002\000R',
	PowerPointMsoAnimEffectAnimationTypeMediaPlay = '\001\002\000S',
	PowerPointMsoAnimEffectAnimationTypeMediaPause = '\001\002\000T',
	PowerPointMsoAnimEffectAnimationTypeMediaStop = '\001\002\000U',
	PowerPointMsoAnimEffectAnimationTypeCirclePath = '\001\002\000V',
	PowerPointMsoAnimEffectAnimationTypeRightTrianglePath = '\001\002\000W',
	PowerPointMsoAnimEffectAnimationTypeDiamondPath = '\001\002\000X',
	PowerPointMsoAnimEffectAnimationTypeHexagonPath = '\001\002\000Y',
	PowerPointMsoAnimEffectAnimationType5PointStarPath = '\001\002\000Z',
	PowerPointMsoAnimEffectAnimationTypeCrescentMoonPath = '\001\002\000[',
	PowerPointMsoAnimEffectAnimationTypeSquarePath = '\001\002\000\\',
	PowerPointMsoAnimEffectAnimationTypeTrapezoidPath = '\001\002\000]',
	PowerPointMsoAnimEffectAnimationTypeHeartPath = '\001\002\000^',
	PowerPointMsoAnimEffectAnimationTypeOctagonPath = '\001\002\000_',
	PowerPointMsoAnimEffectAnimationType6PointStarPath = '\001\002\000`',
	PowerPointMsoAnimEffectAnimationTypeFootballPath = '\001\002\000a',
	PowerPointMsoAnimEffectAnimationTypeEqualTrianglePath = '\001\002\000b',
	PowerPointMsoAnimEffectAnimationTypeParallelogramPath = '\001\002\000c',
	PowerPointMsoAnimEffectAnimationTypePentagonPath = '\001\002\000d',
	PowerPointMsoAnimEffectAnimationType4PointStarPath = '\001\002\000e',
	PowerPointMsoAnimEffectAnimationType8PointStarPath = '\001\002\000f',
	PowerPointMsoAnimEffectAnimationTypeTeardropPath = '\001\002\000g',
	PowerPointMsoAnimEffectAnimationTypePointyStarPath = '\001\002\000h',
	PowerPointMsoAnimEffectAnimationTypeCurvedSquarePath = '\001\002\000i',
	PowerPointMsoAnimEffectAnimationTypeCurvedXPath = '\001\002\000j',
	PowerPointMsoAnimEffectAnimationTypeVerticalFigure8Path = '\001\002\000k',
	PowerPointMsoAnimEffectAnimationTypeCurvyStarPath = '\001\002\000l',
	PowerPointMsoAnimEffectAnimationTypeLoopDeLoopPath = '\001\002\000m',
	PowerPointMsoAnimEffectAnimationTypeBuzzsawPath = '\001\002\000n',
	PowerPointMsoAnimEffectAnimationTypeHorizontalFigure8Path = '\001\002\000o',
	PowerPointMsoAnimEffectAnimationTypePeanutPath = '\001\002\000p',
	PowerPointMsoAnimEffectAnimationTypeFigure8FourPath = '\001\002\000q',
	PowerPointMsoAnimEffectAnimationTypeNeutronPath = '\001\002\000r',
	PowerPointMsoAnimEffectAnimationTypeSwooshPath = '\001\002\000s',
	PowerPointMsoAnimEffectAnimationTypeBeanPath = '\001\002\000t',
	PowerPointMsoAnimEffectAnimationTypePlusPath = '\001\002\000u',
	PowerPointMsoAnimEffectAnimationTypeInvertedTrianglePath = '\001\002\000v',
	PowerPointMsoAnimEffectAnimationTypeInvertedSquarePath = '\001\002\000w',
	PowerPointMsoAnimEffectAnimationTypeLeftPath = '\001\002\000x',
	PowerPointMsoAnimEffectAnimationTypeTurnRightPath = '\001\002\000y',
	PowerPointMsoAnimEffectAnimationTypeArcDownPath = '\001\002\000z',
	PowerPointMsoAnimEffectAnimationTypeZigzagPath = '\001\002\000{',
	PowerPointMsoAnimEffectAnimationTypeSCurve2Path = '\001\002\000|',
	PowerPointMsoAnimEffectAnimationTypeSineWavePath = '\001\002\000}',
	PowerPointMsoAnimEffectAnimationTypeBounceLeftPath = '\001\002\000~',
	PowerPointMsoAnimEffectAnimationTypeDownPath = '\001\002\000\177',
	PowerPointMsoAnimEffectAnimationTypeTurnUpPath = '\001\002\000\200',
	PowerPointMsoAnimEffectAnimationTypeArcUpPath = '\001\002\000\201',
	PowerPointMsoAnimEffectAnimationTypeHeartbeatPath = '\001\002\000\202',
	PowerPointMsoAnimEffectAnimationTypeSpiralRightPath = '\001\002\000\203',
	PowerPointMsoAnimEffectAnimationTypeWavePath = '\001\002\000\204',
	PowerPointMsoAnimEffectAnimationTypeCurvyLeftPath = '\001\002\000\205',
	PowerPointMsoAnimEffectAnimationTypeDiagonalDownRightPath = '\001\002\000\206',
	PowerPointMsoAnimEffectAnimationTypeTurnDownPath = '\001\002\000\207',
	PowerPointMsoAnimEffectAnimationTypeArcLeftPath = '\001\002\000\210',
	PowerPointMsoAnimEffectAnimationTypeFunnelPath = '\001\002\000\211',
	PowerPointMsoAnimEffectAnimationTypeSpringPath = '\001\002\000\212',
	PowerPointMsoAnimEffectAnimationTypeBounceRightPath = '\001\002\000\213',
	PowerPointMsoAnimEffectAnimationTypeSpiralLeftPath = '\001\002\000\214',
	PowerPointMsoAnimEffectAnimationTypeDiagonalUpRightPath = '\001\002\000\215',
	PowerPointMsoAnimEffectAnimationTypeTurnUpRightPath = '\001\002\000\216',
	PowerPointMsoAnimEffectAnimationTypeArcRightPath = '\001\002\000\217',
	PowerPointMsoAnimEffectAnimationTypeSCurve1Path = '\001\002\000\220',
	PowerPointMsoAnimEffectAnimationTypeDecayingWavePath = '\001\002\000\221',
	PowerPointMsoAnimEffectAnimationTypeCurvyRightPath = '\001\002\000\222',
	PowerPointMsoAnimEffectAnimationTypeStairsDownPath = '\001\002\000\223',
	PowerPointMsoAnimEffectAnimationTypeUpPath = '\001\002\000\224',
	PowerPointMsoAnimEffectAnimationTypeRightPath = '\001\002\000\225'
};
typedef enum PowerPointMsoAnimEffect PowerPointMsoAnimEffect;

enum PowerPointMsoAnimateByLevel {
	PowerPointMsoAnimateByLevelTextByNoLevels = '\001\001\000\000',
	PowerPointMsoAnimateByLevelTextByAllLevels = '\001\001\000\001',
	PowerPointMsoAnimateByLevelTextByFirstLevel = '\001\001\000\002',
	PowerPointMsoAnimateByLevelTextBySecondLevel = '\001\001\000\003',
	PowerPointMsoAnimateByLevelTextByThirdLevel = '\001\001\000\004',
	PowerPointMsoAnimateByLevelTextByFourthLevel = '\001\001\000\005',
	PowerPointMsoAnimateByLevelTextByFifthLevel = '\001\001\000\006',
	PowerPointMsoAnimateByLevelChartAllAtOnce = '\001\001\000\007',
	PowerPointMsoAnimateByLevelChartByCategory = '\001\001\000\010',
	PowerPointMsoAnimateByLevelChartByCtageoryElements = '\001\001\000\011',
	PowerPointMsoAnimateByLevelChartBySeries = '\001\001\000\012',
	PowerPointMsoAnimateByLevelChartBySeriesElements = '\001\001\000\013'
};
typedef enum PowerPointMsoAnimateByLevel PowerPointMsoAnimateByLevel;

enum PowerPointMsoAnimTriggerType {
	PowerPointMsoAnimTriggerTypeNoTrigger = '\001\000\000\000',
	PowerPointMsoAnimTriggerTypeOnPageClick = '\001\000\000\001',
	PowerPointMsoAnimTriggerTypeWithPrevious = '\001\000\000\002',
	PowerPointMsoAnimTriggerTypeAfterPrevious = '\001\000\000\003',
	PowerPointMsoAnimTriggerTypeOnShapeClick = '\001\000\000\004'
};
typedef enum PowerPointMsoAnimTriggerType PowerPointMsoAnimTriggerType;

enum PowerPointMsoAnimAfterEffect {
	PowerPointMsoAnimAfterEffectNoAfterEffect = '\000\377\000\000',
	PowerPointMsoAnimAfterEffectDim = '\000\377\000\001',
	PowerPointMsoAnimAfterEffectHide = '\000\377\000\002',
	PowerPointMsoAnimAfterEffectHideOnNextClick = '\000\377\000\003'
};
typedef enum PowerPointMsoAnimAfterEffect PowerPointMsoAnimAfterEffect;

enum PowerPointMsoAnimTextUnitEffect {
	PowerPointMsoAnimTextUnitEffectByParagraph = '\000\376\000\000',
	PowerPointMsoAnimTextUnitEffectByCharacter = '\000\376\000\001',
	PowerPointMsoAnimTextUnitEffectByWord = '\000\376\000\002'
};
typedef enum PowerPointMsoAnimTextUnitEffect PowerPointMsoAnimTextUnitEffect;

enum PowerPointMsoAnimEffectRestart {
	PowerPointMsoAnimEffectRestartRestartAlways = '\000\375\000\001',
	PowerPointMsoAnimEffectRestartRestartWhenOff = '\000\375\000\002',
	PowerPointMsoAnimEffectRestartNeverRestart = '\000\375\000\003'
};
typedef enum PowerPointMsoAnimEffectRestart PowerPointMsoAnimEffectRestart;

enum PowerPointMsoAnimEffectAfter {
	PowerPointMsoAnimEffectAfterAfterFreeze = '\000\374\000\001',
	PowerPointMsoAnimEffectAfterAfterRemove = '\000\374\000\002',
	PowerPointMsoAnimEffectAfterAfterHold = '\000\374\000\003',
	PowerPointMsoAnimEffectAfterAfterTransition = '\000\374\000\004'
};
typedef enum PowerPointMsoAnimEffectAfter PowerPointMsoAnimEffectAfter;

enum PowerPointMsoAnimDirection {
	PowerPointMsoAnimDirectionNoDirection = '\000\371\000\000',
	PowerPointMsoAnimDirectionUp = '\000\371\000\001',
	PowerPointMsoAnimDirectionRight = '\000\371\000\002',
	PowerPointMsoAnimDirectionDown = '\000\371\000\003',
	PowerPointMsoAnimDirectionLeft = '\000\371\000\004',
	PowerPointMsoAnimDirectionOrdinalMask = '\000\371\000\005',
	PowerPointMsoAnimDirectionUpLeft = '\000\371\000\006',
	PowerPointMsoAnimDirectionUpRight = '\000\371\000\007',
	PowerPointMsoAnimDirectionDownRight = '\000\371\000\010',
	PowerPointMsoAnimDirectionDownLeft = '\000\371\000\011',
	PowerPointMsoAnimDirectionTop = '\000\371\000\012',
	PowerPointMsoAnimDirectionBottom = '\000\371\000\013',
	PowerPointMsoAnimDirectionTopLeft = '\000\371\000\014',
	PowerPointMsoAnimDirectionTopRight = '\000\371\000\015',
	PowerPointMsoAnimDirectionBottomRight = '\000\371\000\016',
	PowerPointMsoAnimDirectionBottomLeft = '\000\371\000\017',
	PowerPointMsoAnimDirectionHorizontal = '\000\371\000\020',
	PowerPointMsoAnimDirectionVertical = '\000\371\000\021',
	PowerPointMsoAnimDirectionAcross = '\000\371\000\022',
	PowerPointMsoAnimDirectionInward = '\000\371\000\023',
	PowerPointMsoAnimDirectionOut = '\000\371\000\024',
	PowerPointMsoAnimDirectionClockwise = '\000\371\000\025',
	PowerPointMsoAnimDirectionCounterclockwise = '\000\371\000\026',
	PowerPointMsoAnimDirectionHorizontalIn = '\000\371\000\027',
	PowerPointMsoAnimDirectionHorizontalOut = '\000\371\000\030',
	PowerPointMsoAnimDirectionVerticalIn = '\000\371\000\031',
	PowerPointMsoAnimDirectionVerticalOut = '\000\371\000\032',
	PowerPointMsoAnimDirectionSlightly = '\000\371\000\033',
	PowerPointMsoAnimDirectionCenter = '\000\371\000\034',
	PowerPointMsoAnimDirectionInSlightly = '\000\371\000\035',
	PowerPointMsoAnimDirectionInCenter = '\000\371\000\036',
	PowerPointMsoAnimDirectionInBottom = '\000\371\000\037',
	PowerPointMsoAnimDirectionOutSlightly = '\000\371\000 ',
	PowerPointMsoAnimDirectionOutCenter = '\000\371\000!',
	PowerPointMsoAnimDirectionOutBottom = '\000\371\000\"',
	PowerPointMsoAnimDirectionFontBold = '\000\371\000#',
	PowerPointMsoAnimDirectionFontItalic = '\000\371\000$',
	PowerPointMsoAnimDirectionFontUnderline = '\000\371\000%',
	PowerPointMsoAnimDirectionFontStrikethrough = '\000\371\000&',
	PowerPointMsoAnimDirectionFontShadow = '\000\371\000\'',
	PowerPointMsoAnimDirectionFontAllCaps = '\000\371\000(',
	PowerPointMsoAnimDirectionInstant = '\000\371\000)',
	PowerPointMsoAnimDirectionGradual = '\000\371\000*',
	PowerPointMsoAnimDirectionCycleClockwise = '\000\371\000+',
	PowerPointMsoAnimDirectionCycleCounterclockwise = '\000\371\000,'
};
typedef enum PowerPointMsoAnimDirection PowerPointMsoAnimDirection;

enum PowerPointMsoAnimType {
	PowerPointMsoAnimTypeAnimationTypeNone = '\001\003\000\000',
	PowerPointMsoAnimTypeAnimationTypeMotion = '\001\003\000\001',
	PowerPointMsoAnimTypeAnimationTypeColor = '\001\003\000\002',
	PowerPointMsoAnimTypeAnimationTypeScale = '\001\003\000\003',
	PowerPointMsoAnimTypeAnimationTypeRotation = '\001\003\000\004',
	PowerPointMsoAnimTypeAnimationTypeProperty = '\001\003\000\005',
	PowerPointMsoAnimTypeAnimationTypeCommand = '\001\003\000\006',
	PowerPointMsoAnimTypeAnimationTypeFilter = '\001\003\000\007',
	PowerPointMsoAnimTypeAnimationTypeSet = '\001\003\000\010'
};
typedef enum PowerPointMsoAnimType PowerPointMsoAnimType;

enum PowerPointMsoAnimAdditive {
	PowerPointMsoAnimAdditiveNoAdditive = '\001\007\000\001',
	PowerPointMsoAnimAdditiveMotion = '\001\007\000\002'
};
typedef enum PowerPointMsoAnimAdditive PowerPointMsoAnimAdditive;

enum PowerPointMsoAnimAccumulate {
	PowerPointMsoAnimAccumulateNoAccumulate = '\001\004\000\001',
	PowerPointMsoAnimAccumulateAlways = '\001\004\000\002'
};
typedef enum PowerPointMsoAnimAccumulate PowerPointMsoAnimAccumulate;

enum PowerPointMsoAnimProperty {
	PowerPointMsoAnimPropertyNoProperty = '\001\005\000\000',
	PowerPointMsoAnimPropertyX = '\001\005\000\001',
	PowerPointMsoAnimPropertyY = '\001\005\000\002',
	PowerPointMsoAnimPropertyWidth = '\001\005\000\003',
	PowerPointMsoAnimPropertyHeight = '\001\005\000\004',
	PowerPointMsoAnimPropertyOpacity = '\001\005\000\005',
	PowerPointMsoAnimPropertyRotation = '\001\005\000\006',
	PowerPointMsoAnimPropertyColors = '\001\005\000\007',
	PowerPointMsoAnimPropertyVisibility = '\001\005\000\010',
	PowerPointMsoAnimPropertyTextFontBold = '\001\005\000d',
	PowerPointMsoAnimPropertyTextFontColor = '\001\005\000e',
	PowerPointMsoAnimPropertyTextFontEmboss = '\001\005\000f',
	PowerPointMsoAnimPropertyTextFontItalic = '\001\005\000g',
	PowerPointMsoAnimPropertyTextFontName = '\001\005\000h',
	PowerPointMsoAnimPropertyTextFontShadow = '\001\005\000i',
	PowerPointMsoAnimPropertyTextFontSize = '\001\005\000j',
	PowerPointMsoAnimPropertyTextFontSubscript = '\001\005\000k',
	PowerPointMsoAnimPropertyTextFontSuperscript = '\001\005\000l',
	PowerPointMsoAnimPropertyTextFontUnderline = '\001\005\000m',
	PowerPointMsoAnimPropertyTextFontStrikethrough = '\001\005\000n',
	PowerPointMsoAnimPropertyTextBulletCharacter = '\001\005\000o',
	PowerPointMsoAnimPropertyTextBulletFontName = '\001\005\000p',
	PowerPointMsoAnimPropertyTextBulletNumber = '\001\005\000q',
	PowerPointMsoAnimPropertyTextBulletColor = '\001\005\000r',
	PowerPointMsoAnimPropertyTextBulletRelativeSize = '\001\005\000s',
	PowerPointMsoAnimPropertyTextBulletStyle = '\001\005\000t',
	PowerPointMsoAnimPropertyTextBulletType = '\001\005\000u',
	PowerPointMsoAnimPropertyShapePictureContrast = '\001\005\003\350',
	PowerPointMsoAnimPropertyShapePictureBrightness = '\001\005\003\351',
	PowerPointMsoAnimPropertyShapePictureGamma = '\001\005\003\352',
	PowerPointMsoAnimPropertyShapePictureGrayscale = '\001\005\003\353',
	PowerPointMsoAnimPropertyShapeFillOn = '\001\005\003\354',
	PowerPointMsoAnimPropertyShapeFillColor = '\001\005\003\355',
	PowerPointMsoAnimPropertyShapeFillOpacity = '\001\005\003\356',
	PowerPointMsoAnimPropertyShapeFillBackColor = '\001\005\003\357',
	PowerPointMsoAnimPropertyShapeLineOn = '\001\005\003\360',
	PowerPointMsoAnimPropertyShapeLineColor = '\001\005\003\361',
	PowerPointMsoAnimPropertyShapeShadowOn = '\001\005\003\362',
	PowerPointMsoAnimPropertyShapeShadowType = '\001\005\003\363',
	PowerPointMsoAnimPropertyShapeShadowColor = '\001\005\003\364',
	PowerPointMsoAnimPropertyShapeShadowOpacity = '\001\005\003\365',
	PowerPointMsoAnimPropertyShapeShadowOffsetX = '\001\005\003\366',
	PowerPointMsoAnimPropertyShapeShadowOffsetY = '\001\005\003\367'
};
typedef enum PowerPointMsoAnimProperty PowerPointMsoAnimProperty;

enum PowerPointMsoAnimCommandType {
	PowerPointMsoAnimCommandTypeEvent = '\001\006\000\000',
	PowerPointMsoAnimCommandTypeCall = '\001\006\000\001',
	PowerPointMsoAnimCommandTypeVerb = '\001\006\000\002'
};
typedef enum PowerPointMsoAnimCommandType PowerPointMsoAnimCommandType;

enum PowerPointMsoAnimFilterEffectType {
	PowerPointMsoAnimFilterEffectTypeNoFilterEffectType = '\001\010\000\000',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeBarn = '\001\010\000\001',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeBlinds = '\001\010\000\002',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeBox = '\001\010\000\003',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeCheckerboard = '\001\010\000\004',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeCircle = '\001\010\000\005',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeDiamond = '\001\010\000\006',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeDissolve = '\001\010\000\007',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeFade = '\001\010\000\010',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeImage = '\001\010\000\011',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypePixelate = '\001\010\000\012',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypePlus = '\001\010\000\013',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeRandomBar = '\001\010\000\014',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeSlide = '\001\010\000\015',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeStretch = '\001\010\000\016',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeStrips = '\001\010\000\017',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeWedge = '\001\010\000\020',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeWheel = '\001\010\000\021',
	PowerPointMsoAnimFilterEffectTypeFilterEffectTypeWipe = '\001\010\000\022'
};
typedef enum PowerPointMsoAnimFilterEffectType PowerPointMsoAnimFilterEffectType;

enum PowerPointMsoAnimFilterEffectSubtype {
	PowerPointMsoAnimFilterEffectSubtypeNoEffectSubtype = '\001\011\000\000',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeInVertical = '\001\011\000\001',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeOutVertical = '\001\011\000\002',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeInHorizontal = '\001\011\000\003',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeOutHorizontal = '\001\011\000\004',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeHorizontal = '\001\011\000\005',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeVertical = '\001\011\000\006',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeInward = '\001\011\000\007',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeOut = '\001\011\000\010',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeAcross = '\001\011\000\011',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeFromLeft = '\001\011\000\012',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeFromRight = '\001\011\000\013',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeFromTop = '\001\011\000\014',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeFromBottom = '\001\011\000\015',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeDownLeft = '\001\011\000\016',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeUpLeft = '\001\011\000\017',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeDownRight = '\001\011\000\020',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeUpRight = '\001\011\000\021',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeSpoke1 = '\001\011\000\022',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeSpokes2 = '\001\011\000\023',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeSpokes3 = '\001\011\000\024',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeSpokes4 = '\001\011\000\025',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeSpokes8 = '\001\011\000\026',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeLeft = '\001\011\000\027',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeRight = '\001\011\000\030',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeDown = '\001\011\000\031',
	PowerPointMsoAnimFilterEffectSubtypeFilterEffectSubtypeUp = '\001\011\000\032'
};
typedef enum PowerPointMsoAnimFilterEffectSubtype PowerPointMsoAnimFilterEffectSubtype;

enum PowerPointGetPlayerFrom {
	PowerPointGetPlayerFromView = 'pVEW',
	PowerPointGetPlayerFromSlideShowView = 'PSSv'
};
typedef enum PowerPointGetPlayerFrom PowerPointGetPlayerFrom;

enum PowerPointPasteObject {
	PowerPointPasteObjectView = 'pVEW',
	PowerPointPasteObjectPresentation = 'pptP'
};
typedef enum PowerPointPasteObject PowerPointPasteObject;

enum PowerPointApplyTheme {
	PowerPointApplyThemeSlide = 'pSLD',
	PowerPointApplyThemeMaster = 'pMtr',
	PowerPointApplyThemePresentation = 'pptP'
};
typedef enum PowerPointApplyTheme PowerPointApplyTheme;

enum PowerPointOneColorGradient {
	PowerPointOneColorGradientShape = 'pShp',
	PowerPointOneColorGradientFillFormat = 'pFFm'
};
typedef enum PowerPointOneColorGradient PowerPointOneColorGradient;

enum PowerPointTwoColorGradient {
	PowerPointTwoColorGradientShape = 'pShp',
	PowerPointTwoColorGradientFillFormat = 'pFFm'
};
typedef enum PowerPointTwoColorGradient PowerPointTwoColorGradient;

enum PowerPointAutomaticLength {
	PowerPointAutomaticLengthCallout = 'cD00',
	PowerPointAutomaticLengthCalloutFormat = 'cCoF'
};
typedef enum PowerPointAutomaticLength PowerPointAutomaticLength;

enum PowerPointBeginConnect {
	PowerPointBeginConnectConnector = 'cD01',
	PowerPointBeginConnectConnectorFormat = 'pCxF'
};
typedef enum PowerPointBeginConnect PowerPointBeginConnect;

enum PowerPointBeginDisconnect {
	PowerPointBeginDisconnectConnector = 'cD01',
	PowerPointBeginDisconnectConnectorFormat = 'pCxF'
};
typedef enum PowerPointBeginDisconnect PowerPointBeginDisconnect;

enum PowerPointCustomLength {
	PowerPointCustomLengthCallout = 'cD00',
	PowerPointCustomLengthCalloutFormat = 'cCoF'
};
typedef enum PowerPointCustomLength PowerPointCustomLength;

enum PowerPointCustomDrop {
	PowerPointCustomDropCallout = 'cD00',
	PowerPointCustomDropCalloutFormat = 'cCoF'
};
typedef enum PowerPointCustomDrop PowerPointCustomDrop;

enum PowerPointEndConnect {
	PowerPointEndConnectConnector = 'cD01',
	PowerPointEndConnectConnectorFormat = 'pCxF'
};
typedef enum PowerPointEndConnect PowerPointEndConnect;

enum PowerPointEndDisconnect {
	PowerPointEndDisconnectConnector = 'cD01',
	PowerPointEndDisconnectConnectorFormat = 'pCxF'
};
typedef enum PowerPointEndDisconnect PowerPointEndDisconnect;

enum PowerPointPatterned {
	PowerPointPatternedShape = 'pShp',
	PowerPointPatternedFillFormat = 'pFFm'
};
typedef enum PowerPointPatterned PowerPointPatterned;

enum PowerPointGetActionSettingFor {
	PowerPointGetActionSettingForShape = 'pShp',
	PowerPointGetActionSettingForShapeRange = 'ShpR'
};
typedef enum PowerPointGetActionSettingFor PowerPointGetActionSettingFor;

enum PowerPointSolid {
	PowerPointSolidShape = 'pShp',
	PowerPointSolidFillFormat = 'pFFm'
};
typedef enum PowerPointSolid PowerPointSolid;

enum PowerPointResetRotation {
	PowerPointResetRotationShape = 'pShp',
	PowerPointResetRotationThreeDFormat = 'D3Df'
};
typedef enum PowerPointResetRotation PowerPointResetRotation;

enum PowerPointUserPicture {
	PowerPointUserPictureShape = 'pShp',
	PowerPointUserPictureFillFormat = 'pFFm'
};
typedef enum PowerPointUserPicture PowerPointUserPicture;

enum PowerPointUserTextured {
	PowerPointUserTexturedShape = 'pShp',
	PowerPointUserTexturedFillFormat = 'pFFm'
};
typedef enum PowerPointUserTextured PowerPointUserTextured;

enum PowerPointZOrder {
	PowerPointZOrderShape = 'pShp',
	PowerPointZOrderShapeRange = 'ShpR'
};
typedef enum PowerPointZOrder PowerPointZOrder;

enum PowerPointPresetTextured {
	PowerPointPresetTexturedShape = 'pShp',
	PowerPointPresetTexturedFillFormat = 'pFFm'
};
typedef enum PowerPointPresetTextured PowerPointPresetTextured;

enum PowerPointPresetGradient {
	PowerPointPresetGradientShape = 'pShp',
	PowerPointPresetGradientFillFormat = 'pFFm'
};
typedef enum PowerPointPresetGradient PowerPointPresetGradient;

enum PowerPointApply {
	PowerPointApplyShape = 'pShp',
	PowerPointApplyShapeRange = 'ShpR'
};
typedef enum PowerPointApply PowerPointApply;

enum PowerPointFlip {
	PowerPointFlipShape = 'pShp',
	PowerPointFlipShapeRange = 'ShpR'
};
typedef enum PowerPointFlip PowerPointFlip;

enum PowerPointPickUp {
	PowerPointPickUpShape = 'pShp',
	PowerPointPickUpShapeRange = 'ShpR'
};
typedef enum PowerPointPickUp PowerPointPickUp;

enum PowerPointRerouteConnections {
	PowerPointRerouteConnectionsShape = 'pShp',
	PowerPointRerouteConnectionsShapeRange = 'ShpR'
};
typedef enum PowerPointRerouteConnections PowerPointRerouteConnections;

enum PowerPointSetShapesDefaultProperties {
	PowerPointSetShapesDefaultPropertiesShape = 'pShp',
	PowerPointSetShapesDefaultPropertiesShapeRange = 'ShpR',
	PowerPointSetShapesDefaultPropertiesShapeRange = 'ShpR'
};
typedef enum PowerPointSetShapesDefaultProperties PowerPointSetShapesDefaultProperties;

@protocol PowerPointGenericMethods

- (void) closeSaving:(PowerPointSaveOptions)saving savingIn:(NSURL *)savingIn;  // Close a document.
- (void) saveIn:(NSURL *)in_ as:(id)as;  // Save a document.
- (void) printWithProperties:(NSDictionary *)withProperties printDialog:(BOOL)printDialog;  // Print a document.
- (void) delete;  // Delete an object.
- (void) duplicateTo:(SBObject *)to withProperties:(NSDictionary *)withProperties;  // Copy an object.
- (void) moveTo:(SBObject *)to;  // Move an object to a new location.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if PowerPoint can check out a specified presentation from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified presentation from a server to a local computer for editing. Returns a String that represents the local path and file name of the presentation checked out.
- (void) quit;

@end



/*
 * Standard Suite
 */

// The application's top-level scripting object.
@interface PowerPointApplication : SBApplication

- (SBElementArray<PowerPointDocument *> *) documents;
- (SBElementArray<PowerPointWindow *> *) windows;

@property (copy, readonly) NSString *name;  // The name of the application.
@property (readonly) BOOL frontmost;  // Is this the active application?
@property (copy, readonly) NSString *version;  // The version number of the application.

- (id) open:(id)x;  // Open a document.
- (void) print:(id)x withProperties:(NSDictionary *)withProperties printDialog:(BOOL)printDialog;  // Print a document.
- (void) quitSaving:(PowerPointSaveOptions)saving;  // Quit the application.
- (BOOL) exists:(id)x;  // Verify that an object exists.
- (void) reset:(id)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) applyTheme:(id)x fileName:(NSString *)fileName;  // Applies a theme or design template to the specified slide, master or presentation
- (void) arrangeWindows:(PowerPointEPPArrangeStyle)x;  // Arrange Document Windows
- (PowerPointPlayer *) getPlayerFrom:(id)x for:(PowerPointShape *)for_;  // get a player from a shape inside a slide show view
- (void) insertTheText:(NSString *)theText at:(SBObject *)at;
- (void) pasteObject:(id)x;
- (PowerPointAddIn *) registerAddIn:(NSString *)x;
- (NSInteger) runVBMacroMacroName:(NSString *)macroName listOfParameters:(NSArray<NSString *> *)listOfParameters;  // Runs a Visual Basic macro.
- (void) apply:(id)x;
- (void) automaticLength:(id)x;
- (void) beginConnect:(id)x connectedShape:(PowerPointShape *)connectedShape connectionSite:(NSInteger)connectionSite;
- (void) beginDisconnect:(id)x;
- (void) customDrop:(id)x dropAmount:(double)dropAmount;
- (void) customLength:(id)x length:(double)length;
- (void) endConnect:(id)x connectedShape:(PowerPointShape *)connectedShape connectionSite:(NSInteger)connectionSite;
- (void) endDisconnect:(id)x;
- (void) flip:(id)x direction:(PowerPointMsoFlipCmd)direction;
- (PowerPointActionSetting *) getActionSettingFor:(id)x event:(PowerPointEPPMouseActivation)event;
- (void) oneColorGradient:(id)x style:(PowerPointMsoGradientStyle)style variant:(NSInteger)variant degree:(double)degree;
- (void) patterned:(id)x pattern:(PowerPointMsoPatternType)pattern;
- (void) pickUp:(id)x;
- (void) presetGradient:(id)x style:(PowerPointMsoGradientStyle)style variant:(NSInteger)variant gradientType:(PowerPointMsoPresetGradientType)gradientType;
- (void) presetTextured:(id)x texture:(PowerPointMsoPresetTexture)texture;
- (void) rerouteConnections:(id)x;
- (void) resetRotation:(id)x;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- (void) setShapesDefaultProperties:(id)x;
- (void) solid:(id)x;
- (void) twoColorGradient:(id)x style:(PowerPointMsoGradientStyle)style variant:(NSInteger)variant;
- (void) userPicture:(id)x pictureFile:(NSString *)pictureFile;
- (void) userTextured:(id)x textureFile:(NSString *)textureFile;
- (void) zOrder:(id)x zOrderPosition:(PowerPointMsoZOrderCmd)zOrderPosition;

@end

// A document.
@interface PowerPointDocument : SBObject <PowerPointGenericMethods>

@property (copy, readonly) NSString *name;  // Its name.
@property (readonly) BOOL modified;  // Has it been modified since the last save?
@property (copy, readonly) NSURL *file;  // Its location on disk, if it has one.


@end

// A window.
@interface PowerPointWindow : SBObject <PowerPointGenericMethods>

@property (copy, readonly) NSString *name;  // The title of the window.
- (NSInteger) id;  // The unique identifier of the window.
@property NSInteger index;  // The index of the window, ordered front to back.
@property NSRect bounds;  // The bounding rectangle of the window.
@property (readonly) BOOL closeable;  // Does the window have a close button?
@property (readonly) BOOL miniaturizable;  // Does the window have a minimize button?
@property BOOL miniaturized;  // Is the window minimized right now?
@property (readonly) BOOL resizable;  // Can the window be resized?
@property BOOL visible;  // Is the window visible right now?
@property (readonly) BOOL zoomable;  // Does the window have a zoom button?
@property BOOL zoomed;  // Is the window zoomed right now?
@property (copy, readonly) PowerPointDocument *document;  // The document whose contents are displayed in the window.
@property (readonly) NSInteger entryIndex;  // the number of the window
@property NSPoint position;  // upper left coordinates of the window
@property (readonly) BOOL titled;  // Does the window have a title bar?
@property (readonly) BOOL floating;  // Does the window float?
@property (readonly) BOOL modal;  // Is the window modal?
@property (readonly) BOOL collapsable;  // Is the window collapasable?
@property BOOL collapsed;  // Is the window collapsed?
@property (readonly) BOOL sheet;  // Is this window a sheet window?


@end



/*
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface PowerPointCommandBarControl : SBObject <PowerPointGenericMethods>

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) PowerPointMsoControlType controlType;  // Returns the type of the command bar control.
@property (copy) NSString *descriptionText;  // Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control.
@property BOOL enabled;  // Returns or sets if the command bar control is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number for this command bar control.
@property NSInteger height;  // Returns or sets the height of a command bar control.
@property NSInteger helpContextID;  // Returns or sets the help context ID number for the Help topic attached to the command bar control.
@property (copy) NSString *helpFile;  // Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property.
- (NSInteger) id;  // Returns the id for a built-in command bar control.
@property (readonly) NSInteger leftPosition;  // Returns the left position of the command bar control.
@property (copy) NSString *name;  // Returns or sets the caption text for a command bar control.
@property (copy) NSString *parameter;  // Returns or sets a string that is used to execute a command.
@property NSInteger priority;  // Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7.
@property (copy) NSString *tag;  // Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control.
@property (copy) NSString *tooltipText;  // Returns or sets the text displayed in a command bar controls tooltip.
@property (readonly) NSInteger top;  // Returns the top position of a command bar control.
@property BOOL visible;  // Returns or sets if the command bar control is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar control.

- (void) execute;  // Runs the procedure or built-in command assigned to the specified command bar control.

@end

// A button control within a command bar.
@interface PowerPointCommandBarButton : PowerPointCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property PowerPointMsoButtonState buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property PowerPointMsoButtonStyle buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface PowerPointCommandBarCombobox : PowerPointCommandBarControl

@property PowerPointMsoComboStyle comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
@property (copy) NSString *comboboxText;  // Returns or sets the text in the display or edit portion of the command bar combobox control.
@property NSInteger dropDownLines;  // Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control.
@property NSInteger dropDownWidth;  // Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control.
@property NSInteger listIndex;

- (void) addItemToComboboxComboboxItem:(NSString *)comboboxItem entry_index:(NSInteger)entry_index;  // Add a new string to a custom combobox control.
- (void) clearCombobox;  // Clear all of the strings form a custom combobox.
- (NSString *) getComboboxItemEntry_index:(NSInteger)entry_index;  // Return the string at the given index within a combobox.
- (NSInteger) getCountOfComboboxItems;  // Return the number of strings within a combobox.
- (void) removeAnItemFromComboboxEntry_index:(NSInteger)entry_index;  // Remove a string from a custom combobox.
- (void) setComboboxItemEntry_index:(NSInteger)entry_index comboboxItem:(NSString *)comboboxItem;  // Set the string an a given index for a custom combobox.

@end

// A popup menu control within a command bar.
@interface PowerPointCommandBarPopup : PowerPointCommandBarControl

- (SBElementArray<PowerPointCommandBarControl *> *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface PowerPointCommandBar : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointCommandBarControl *> *) commandBarControls;

@property PowerPointMsoBarPosition barPosition;  // Returns or sets the position of the command bar.
@property (readonly) PowerPointMsoBarType barType;  // Returns the type of this command bar.
@property (readonly) BOOL builtIn;  // True if the command bar is built-in.
@property (copy, readonly) NSString *context;  // Returns or sets a string that determines where a command bar will be saved.
@property (readonly) BOOL embeddable;  // Returns if the command bar can be displayed inside the document window.
@property BOOL embedded;  // Returns or sets if the command bar will be displayed inside the document window.
@property BOOL enabled;  // Returns or set if the command bar is enabled.
@property (readonly) NSInteger entry_index;  // The index of the command bar.
@property NSInteger height;  // Returns or sets the height of the command bar.
@property NSInteger leftPosition;  // Returns or sets the left position of the command bar.
@property (copy) NSString *localName;  // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
@property (copy) NSString *name;  // Returns or sets the name of the command bar.
@property (copy) NSArray<NSAppleEventDescriptor *> *protection;  // Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock.
@property NSInteger rowIndex;  // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
@property NSInteger top;  // Returns or sets the top position of a command bar.
@property BOOL visible;  // Returns or sets if the command bar is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar.


@end

@interface PowerPointDocumentProperty : SBObject <PowerPointGenericMethods>

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface PowerPointCustomDocumentProperty : PowerPointDocumentProperty


@end

@interface PowerPointWebPageFont : SBObject <PowerPointGenericMethods>

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft PowerPoint Suite
 */

@interface PowerPointActionSetting : SBObject <PowerPointGenericMethods>

@property PowerPointEPPActionType action;
@property (copy) NSString *actionSettingToRun;
@property (copy, readonly) PowerPointSoundEffect *actionSoundEffect;
@property (copy) NSString *actionVerb;
@property BOOL animateAction;
@property (copy, readonly) PowerPointHyperlink *hyperlink;
@property (copy) NSString *slideShowName;


@end

@interface PowerPointAddIn : SBObject <PowerPointGenericMethods>

@property BOOL autoLoad;
@property (copy, readonly) NSString *fullName;
@property BOOL loaded;
@property (copy, readonly) NSString *name;
@property (copy, readonly) NSString *path;
@property BOOL registered;
@property (readonly) BOOL registeredInHKLM;


@end

@interface PowerPointAnimationBehavior : SBObject <PowerPointGenericMethods>

@property PowerPointMsoAnimAccumulate accumulate;
@property PowerPointMsoAnimAdditive additive;
@property PowerPointMsoAnimType animationBehaviorType;
@property (copy, readonly) PowerPointColorsEffect *colorsEffect;
@property (copy, readonly) PowerPointCommandEffect *commandEffect;
@property (copy, readonly) PowerPointFilterEffect *filterEffect;
@property (copy, readonly) PowerPointMotionEffect *motionEffect;
@property (copy, readonly) PowerPointPropertyEffect *propertyEffect;
@property (copy, readonly) PowerPointRotatingEffect *rotatingEffect;
@property (copy, readonly) PowerPointScaleEffect *scaleEffect;
@property (copy, readonly) PowerPointSetEffect *setEffect;
@property (copy, readonly) PowerPointTiming *timing;


@end

@interface PowerPointAnimationPoint : SBObject <PowerPointGenericMethods>

@property (copy) NSString *formula;
@property double time;
@property (copy) SBObject *value;


@end

@interface PowerPointAnimationSettings : SBObject <PowerPointGenericMethods>

@property double advanceTime;
@property PowerPointEPPAfterEffect afterEffect;
@property BOOL animate;
@property BOOL animateBackground;
@property BOOL animateTextInReverse;
@property NSInteger animationOrder;
@property (copy, readonly) PowerPointPlaySettings *animationPlaySettings;
@property (copy, readonly) PowerPointSoundEffect *animationSoundEffect;
@property PowerPointEPPChartUnitEffect chartUnitEffect;
@property (copy) NSArray<NSNumber *> *dimColor;
@property PowerPointMsoThemeColorIndex dimColorThemeIndex;
@property PowerPointEPPEntryEffect entryEffect;
@property PowerPointEPPTextLevelEffect textLevelEffect;
@property PowerPointEPPTextUnitEffect textUnitEffect;


@end

// An AppleScript object representing the Microsoft POWERPOINT application.
@interface PowerPointApplication (MicrosoftPowerPointSuite)

- (SBElementArray<PowerPointPresentation *> *) presentations;
- (SBElementArray<PowerPointDocumentWindow *> *) documentWindows;
- (SBElementArray<PowerPointSlideShowWindow *> *) slideShowWindows;
- (SBElementArray<PowerPointCommandBar *> *) commandBars;
- (SBElementArray<PowerPointAddIn *> *) addIns;

@property (copy, readonly) NSString *Version;
@property (copy, readonly) PowerPointPresentation *activePresentation;
@property (copy, readonly) NSString *activePrinter;
@property (copy, readonly) PowerPointDocumentWindow *activeWindow;
@property (copy, readonly) PowerPointAutocorrect *autocorrectObject;  // Returns the autocorrect object
@property PowerPointMsoAutomationSecurity automationSecurity;
@property (copy, readonly) NSString *build;
@property (copy, readonly) NSString *caption;
@property PowerPointEPPSaveAsFileType defaultSaveFormat;
@property (copy, readonly) PowerPointDefaultWebOptions *defaultWebOptionsObject;
@property (copy, readonly) NSString *name;
@property (copy, readonly) NSString *path;
@property (copy, readonly) PowerPointPreferences *preferenceObject;
@property (copy, readonly) PowerPointSaveAsMovieSettings *saveAsMovieSettingsObject;
@property BOOL startUpDialog;

@end

// Represents a single autocorrect entry.
@interface PowerPointAutocorrectEntry : SBObject <PowerPointGenericMethods>

@property (copy, readonly) NSString *autocorrectValue;  // Returns the value of the auto correct entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the auto correct entry.


@end

// Represents the autocorrect functionality in PowerPoint.
@interface PowerPointAutocorrect : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointAutocorrectEntry *> *) autocorrectEntries;
- (SBElementArray<PowerPointFirstLetterException *> *) firstLetterExceptions;
- (SBElementArray<PowerPointTwoInitialCapsException *> *) twoInitialCapsExceptions;

@property BOOL correctDays;  // Returns if PowerPoint automatically capitalizes the first letter of days of the week.
@property BOOL correctInitialCaps;  // Returns if PowerPoint automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, POwerPoint is corrected to PowerPoint.
@property BOOL correctSentenceCaps;  // Returns if PowerPoint automatically capitalizes the first letter in each sentence.
@property BOOL replaceText;  // Returns if Microsoft PowerPoint automatically replaces specified text with entries from the autocorrect list.


@end

// Represents an interface for broadcasting presentations
@interface PowerPointBroadcast : SBObject <PowerPointGenericMethods>

- (void) endSession;  // End a running broadcast session.
- (NSString *) getAttendeeURL;
- (BOOL) getIsBroadcasting;  // Returns true if the current presentation is being broadcast.
- (void) startServerUrl:(NSString *)serverUrl;  // Starts broadcasting to the given URL.

@end

@interface PowerPointBulletFormat : SBObject <PowerPointGenericMethods>

@property (copy) NSString *bulletCharacter;  // Returns or sets the Unicode character value that is used for bullets in the specified text.
@property (copy, readonly) PowerPointFont *bulletFont;  // Returns a font object that represents character formatting for a bullet format object.
@property (readonly) NSInteger bulletNumber;  // Returns the bullet number of a paragraph.
@property NSInteger bulletStartValue;
@property PowerPointMsoNumberedBulletStyle bulletStyle;  // Returns or sets a constant that represents the style of a bullet.
@property PowerPointMsoBulletType bulletType;  // Returns or sets a constant that represents the type of bullet.
@property double relativeSize;  // Returns or sets the bullet size relative to the size of the first text character in the paragraph.
@property BOOL useTextColor;  // Determines whether the specified bullets are set to the color of the first text character in the paragraph.
@property BOOL useTextFont;  // Determines whether the specified bullets are set to the font of the first text character in the paragraph.
@property BOOL visible;  // Returns or sets a value that specifies whether the bullet is visible.

- (void) setBulletPicturePictureFile:(NSString *)pictureFile;  // Sets the graphics file to be used for bullets in a bulleted list.

@end

@interface PowerPointColorScheme : SBObject <PowerPointGenericMethods>

- (NSArray<NSNumber *> *) getColorFromAt:(PowerPointEPPColorSchemeIndex)at;
- (void) setColorForAt:(PowerPointEPPColorSchemeIndex)at toColor:(NSArray<NSNumber *> *)toColor;

@end

@interface PowerPointColorsEffect : SBObject <PowerPointGenericMethods>

@property (copy) NSArray<NSNumber *> *by_color;
@property PowerPointMsoThemeColorIndex by_colorThemeIndex;  // Returns an object that represents a change to the color of the object by the specified number, expressed in RGB format.
@property (copy) NSArray<NSNumber *> *from_color;
@property PowerPointMsoThemeColorIndex from_colorThemeIndex;  // Returns or sets an object that represents the starting RGB color value of an animation behavior.
@property (copy) NSArray<NSNumber *> *to_color;
@property PowerPointMsoThemeColorIndex to_colorThemeIndex;  // Returns or sets an object that represents the RGB color value of an animation behavior.


@end

@interface PowerPointCommandEffect : SBObject <PowerPointGenericMethods>

@property (copy) NSString *command;
@property PowerPointMsoAnimCommandType type;


@end

@interface PowerPointCustomLayout : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointShape *> *) shapes;

@property (copy, readonly) PowerPointShape *background;
@property (copy, readonly) PowerPointDesign *design;
@property BOOL displayMasterShapes;
@property BOOL followMasterBackground;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property (readonly) double height;
@property (copy, readonly) PowerPointThemeColorScheme *themeColorScheme;  // Returns the color scheme of a custom layout.
@property (copy, readonly) PowerPointTimeline *timeline;
@property (readonly) double width;


@end

@interface PowerPointDefaultWebOptions : SBObject <PowerPointGenericMethods>

@property BOOL allowPNG;
@property BOOL alwaysSaveInDefaultEncoding;
@property BOOL checkIfOfficeIsHTMLEditor;
@property PowerPointMsoEncoding encoding;
@property BOOL updateLinksOnSave;


@end

@interface PowerPointDesign : SBObject <PowerPointGenericMethods>

@property (copy, readonly) PowerPointMaster *slideMaster;


@end

@interface PowerPointDocumentWindow : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointPane *> *) panes;

@property (readonly) BOOL active;
@property (copy, readonly) PowerPointPane *activePane;
@property BOOL blackAndWhite;
@property (copy, readonly) NSString *caption;
@property (readonly) NSInteger entry_index;
@property double height;
@property double leftPosition;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointSelection *selection;  // Represents the selection in the specified document window.
@property NSInteger splitHorizontal;
@property NSInteger splitVertical;
@property double top;
@property (copy, readonly) PowerPointView *view;
@property PowerPointEPPViewType viewType;
@property double width;

- (void) collapseSectionAtPosition:(NSInteger)atPosition;
- (void) expandSectionAtPosition:(NSInteger)atPosition;
- (BOOL) getIsExpandedOfSectionAtPosition:(NSInteger)atPosition;
- (void) launchSpellerOn;

@end

@interface PowerPointEffectInformation : SBObject <PowerPointGenericMethods>

@property (readonly) PowerPointMsoAnimAfterEffect afterEffectInformation;
@property (readonly) BOOL animateBackgroundInformation;
@property (readonly) BOOL animateTextInReverseInformation;
@property (readonly) PowerPointMsoAnimateByLevel buildByLevel;
@property (copy, readonly) NSArray<NSNumber *> *dim;
@property (copy, readonly) PowerPointPlaySettings *playSettingsInformation;
@property (copy, readonly) PowerPointSoundEffect *soundEffectInformation;
@property (readonly) PowerPointMsoAnimTextUnitEffect textUnitEffectInformation;


@end

@interface PowerPointEffectParameters : SBObject <PowerPointGenericMethods>

@property double amount;
@property (copy, readonly) NSArray<NSNumber *> *color2;
@property (readonly) PowerPointMsoThemeColorIndex color2ColorThemeIndex;  // Returns an object that represents the color on which to end a color-cycle animation.
@property PowerPointMsoAnimDirection direction;
@property (copy) NSString *font2;
@property BOOL relative;
@property double size2;


@end

@interface PowerPointEffect : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointAnimationBehavior *> *) animationBehaviors;

@property PowerPointMsoAnimEffect animationEffectType;
@property (copy, readonly) PowerPointEffectInformation *effectInformation;
@property (copy, readonly) PowerPointEffectParameters *effectParameters;
@property BOOL exitAnimation;
@property (copy, readonly) NSString *name;
@property (copy) PowerPointShape *shape;
@property NSInteger targetParagraph;
@property (readonly) NSInteger textRangeLength;
@property (readonly) NSInteger textRangeStart;
@property (copy, readonly) PowerPointTiming *timing;

- (PowerPointAnimationBehavior *) addBehaviorType:(PowerPointMsoAnimType)type;  // add an animation behavior

@end

@interface PowerPointFilterEffect : SBObject <PowerPointGenericMethods>

@property PowerPointMsoAnimFilterEffectType filterType;
@property BOOL reveal;
@property PowerPointMsoAnimFilterEffectSubtype subtype;


@end

// Represents an abbreviation excluded from automatic correction.
@interface PowerPointFirstLetterException : SBObject <PowerPointGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the FirstLetterException.


@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface PowerPointFont : SBObject <PowerPointGenericMethods>

@property (copy) NSString *ASCIIName;  // Returns or sets the font used for Latin text; characters with character codes from 0 through 127.
@property BOOL autoRotateNumbers;  // Returns or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.
@property double baseLineOffset;  // Returns or sets a value specifying the horizontaol offset of the selected font.
@property BOOL bold;  // Returns or sets a value specifying whether the font should be bold.
@property PowerPointMsoTextCaps capsType;  // Returns or sets a value specifying how the text should be capitalized.
@property (copy) NSString *eastAsianName;  // Returns or sets the font name used for Asian text.
@property (readonly) BOOL embedable;  // Returns a value indicating whether the font can be embedded in a page.
@property (readonly) BOOL embedded;  // Returns a value specifying whether the font is embedded in a page.
@property BOOL emboss;
@property BOOL equalizeCharacterHeight;  // Returns or sets a value specifying whether the text should have the same horizontal height.
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns a value specifying the fill formatting for text.
@property (copy) NSArray<NSNumber *> *fontColor;
@property PowerPointMsoThemeColorIndex fontColorThemeIndex;  // Returns or sets the color for the specified font.
@property (copy) NSString *fontName;  // Returns or sets a value specifying the font to use for a selection.
@property (copy) NSString *fontNameOther;  // Returns or sets the font used for characters whose character set numbers are greater than 127.
@property double fontSize;
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns a value specifying the glow formatting of the text.
@property (copy) NSArray<NSNumber *> *highlightColor;  // Returns or sets the text highlight color for object.
@property PowerPointMsoThemeColorIndex highlightColorThemeIndex;  // Returns or sets the specified text highlight color for object.
@property BOOL italic;
@property double kerning;  // Returns or sets a value specifying the amount of spacing between text characters.
@property (copy, readonly) PowerPointLineFormat *lineFormat;  // Returns a value specifiying the line formatting of the text.
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns a value specifying the reflection formatting of the text.
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;  // Returns the value specifying the type of shadow effect for the selection of text.
@property PowerPointMsoSoftEdgeType softEdgeFormat;  // Returns or sets a value specifying the soft edge fromatting of the text.
@property double spacing;  // Returns or sets a value specifying the spacing between characters in a selection of text.
@property PowerPointMsoTextStrike strikeType;  // Returns or sets a value specifying the strike format used for a selection of text.
@property BOOL subscript;  // Returns or sets a value specifying that the selected text should be displayed a subscript.
@property BOOL superscript;  // Returns or sets a value specifying that the selected text should be displayed a superscript.
@property double transparency;
@property BOOL underline;
@property (copy) NSArray<NSNumber *> *underlineColor;  // Returns or sets the 24-bit color of the underline for the specified font object.
@property PowerPointMsoThemeColorIndex underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.
@property PowerPointMsoTextUnderlineType underlineStyle;  // Returns or sets a value specifying the underline style for the selected text.
@property PowerPointMsoPresetTextEffect wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.


@end

@interface PowerPointHeaderOrFooter : SBObject <PowerPointGenericMethods>

@property PowerPointEPPDateTimeFormat dateFormat;
@property (copy) NSString *headerFooterText;
@property BOOL useDateFormat;
@property BOOL visible;


@end

@interface PowerPointHeadersAndFooters : SBObject <PowerPointGenericMethods>

@property (copy, readonly) PowerPointHeaderOrFooter *dateAndTime;
@property BOOL displayHeadersFootersOnTitleSlide;
@property (copy, readonly) PowerPointHeaderOrFooter *footer;
@property (copy, readonly) PowerPointHeaderOrFooter *header;
@property (copy, readonly) PowerPointHeaderOrFooter *slideNumber;


@end

@interface PowerPointHyperlink : SBObject <PowerPointGenericMethods>

@property (copy) NSString *hyperlinkAddress;
@property (copy) NSString *hyperlinkSubAddress;
@property (copy, readonly) id hyperlinkType;


@end

@interface PowerPointMaster : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointHyperlink *> *) hyperlinks;
- (SBElementArray<PowerPointCustomLayout *> *) customLayouts;

@property (copy, readonly) PowerPointShape *background;
@property (copy, readonly) PowerPointColorScheme *colorScheme;
@property (copy, readonly) PowerPointDesign *design;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property (readonly) double height;
@property (copy, readonly) PowerPointOfficeTheme *theme;
@property (copy, readonly) PowerPointTimeline *timeline;
@property (readonly) double width;

- (PowerPointTextStyle *) getTextStyleFromAt:(PowerPointPpTextStyleType)at;

@end

@interface PowerPointMotionEffect : SBObject <PowerPointGenericMethods>

@property double byX;
@property double byY;
@property double fromX;
@property double fromY;
@property (copy) NSString *path;
@property double toX;
@property double toY;


@end

@interface PowerPointNamedSlideShow : SBObject <PowerPointGenericMethods>

@property (copy, readonly) NSString *name;
@property (readonly) NSInteger numberOfSlides;
@property (copy, readonly) NSArray<NSNumber *> *slideIDs;


@end

@interface PowerPointPageSetup : SBObject <PowerPointGenericMethods>

@property NSInteger firstSlideNumber;
@property PowerPointMsoOrientation notesOrientation;
@property PowerPointMsoOrientation slideOrientation;
@property PowerPointEPPSlideSizeType slideSize;
@property double slideWidth;


@end

@interface PowerPointPane : SBObject <PowerPointGenericMethods>

@property (readonly) BOOL active;
@property (readonly) PowerPointEPPViewType paneViewType;


@end

@interface PowerPointParagraphFormat : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointTabStop *> *) tabStops;

@property PowerPointMsoParagraphAlignment alignment;
@property PowerPointMsoBaselineAlignment baselineAlignment;  // Returns or sets a constant that represents the vertical position of fonts in a paragraph.
@property (copy, readonly) PowerPointBulletFormat *bulletFormat;
@property BOOL eastAsianLineBreakControl;
@property double firstLineIndent;  // Returns or sets the value, in points, for a first line or hanging indent.
@property BOOL hangingPunctuation;  // Determines whether hanging punctuation is enabled for the specified paragraphs.
@property NSInteger indentLevel;  // Returns or sets a value representing the indent level assigned to text in the selected paragraph.
@property double leftIndent;  // Returns or sets a value that represents the left indent value, in points, for the specified paragraphs.
@property BOOL lineRuleAfter;  // Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleBefore;  // Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleWithin;  // Determines whether line spacing between base lines is set to a specific number of points or lines.
@property double rightIndent;  // Returns or sets the right indent, in points, for the specified paragraphs.
@property double spaceAfter;  // Returns or sets the spacing, in points, after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing, in points, before the specified paragraphs.
@property double spaceWithin;  // Returns or sets the amount of space between base lines in the specified paragraph, in points or lines.
@property PowerPointMsoTextDirection textDirection;  // Returns or sets the text direction for the specified paragraph.
@property BOOL wordWrap;  // Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.


@end

@interface PowerPointPlaySettings : SBObject <PowerPointGenericMethods>

@property BOOL hideWhileNotPlaying;
@property BOOL loopUntilStopped;
@property BOOL pauseAnimation;
@property BOOL playOnEntry;
@property BOOL rewindMove;
@property NSInteger stopAfterSlides;


@end

// Represents an interface for playing movies
@interface PowerPointPlayer : SBObject <PowerPointGenericMethods>

@property NSInteger currentPosition;
@property (readonly) PowerPointEPPPlayerState playerState;

- (void) goToNextBookmarkForPlayer;  // Advance the player to the next bookmark.
- (void) goToPreviousBookmarkForPlayer;  // Rewind the player to the previous bookmark.
- (void) pausePlayer;  // Pause playback.
- (void) playPlayer;  // Begin or resume playback.
- (void) stopPlayer;  // Stop playback.

@end

@interface PowerPointPreferences : SBObject <PowerPointGenericMethods>

@property NSInteger alwaysSuggestCorrections;
@property NSInteger appendDOSExtentions;
@property NSInteger autoFit;
@property NSInteger autoRecoveryFrequency;
@property NSInteger backgroundSpelling;
@property NSInteger compressFile;
@property NSInteger defaultView;
@property NSInteger dragDrop;
@property NSInteger endBlankSlide;
@property NSInteger filePropertiesPrompt;
@property NSInteger hideSpellingErrors;
@property NSInteger ignoreNumbersInWords;
@property NSInteger ignoreUppercase;
@property NSInteger optionBitmap;
@property NSInteger rulerUnits;
@property NSInteger saveAutoRecovery;
@property NSInteger showVerticalRuler;
@property NSInteger slideShowControl;
@property NSInteger slideShowRightMouse;
@property NSInteger smartCutPaste;
@property NSInteger smartQuotes;
@property NSInteger undoLevelCount;
@property (copy) NSString *userInitials;
@property (copy) NSString *userName;
@property NSInteger wordSelection;


@end

@interface PowerPointPresentation : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointSlide *> *) slides;
- (SBElementArray<PowerPointColorScheme *> *) colorSchemes;
- (SBElementArray<PowerPointFont *> *) fonts;
- (SBElementArray<PowerPointDocumentWindow *> *) documentWindows;
- (SBElementArray<PowerPointDocumentProperty *> *) documentProperties;
- (SBElementArray<PowerPointCustomDocumentProperty *> *) customDocumentProperties;
- (SBElementArray<PowerPointDesign *> *) designs;

@property (copy, readonly) PowerPointBroadcast *broadcast;
@property (copy, readonly) PowerPointShape *defaultShape;
@property PowerPointEPPFarEastLineBreakLevel eastAsianLineBreakLevel;  // Returns or sets the East Asian line break control level for the specified paragraph.
@property (copy, readonly) NSString *fullName;
@property (copy, readonly) PowerPointMaster *handoutMaster;
@property (readonly) BOOL hasTitleMaster;
@property (readonly) BOOL isProtected;  // Returns true if the presentation is protected by Information Rights Management.
@property PowerPointMsoTextDirection layoutDirection;
@property (copy, readonly) NSString *name;
@property (copy) NSString *noLineBreakAfter;
@property (copy) NSString *noLineBreakBefore;
@property (copy, readonly) PowerPointMaster *notesMaster;
@property (copy, readonly) PowerPointPageSetup *pageSetup;
@property (copy) NSString *password;  // The password for opening the presentation
@property (copy, readonly) NSString *path;
@property (copy, readonly) PowerPointPrintOptions *printOptions;
@property (readonly) BOOL readOnly;
@property (copy, readonly) PowerPointSaveAsMovieSettings *saveAsMovieSettings;
@property BOOL saved;
@property (copy, readonly) PowerPointSectionProperties *sectionProperties;
@property (copy, readonly) PowerPointMaster *slideMaster;
@property (copy, readonly) PowerPointSlideShowSettings *slideShowSettings;
@property (copy, readonly) PowerPointSlideShowWindow *slideShowWindow;
@property (copy, readonly) NSString *templateName;
@property (copy, readonly) PowerPointMaster *titleMaster;
@property (copy, readonly) PowerPointWebOptions *webOptions;
@property (copy) NSString *writePassword;  // The password for saving changes to the presentation

- (PowerPointDesign *) addDesignDesignName:(NSString *)DesignName index:(NSInteger)index;  // add a new design
- (void) applyTemplateFileName:(NSString *)fileName;  // Applies a theme or design template to the specified slide, master or presentation
- (BOOL) canCheckIn;  // Returns True if PowerPoint can check in a specified presentation to a server.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(PowerPointEPPCheckInVersionType)versionType;  // Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to printToFile:(NSString *)printToFile copies:(NSInteger)copies collate:(BOOL)collate showDialog:(BOOL)showDialog;
- (void) redoTimes:(NSInteger)times;
- (void) undoTimes:(NSInteger)times;
- (void) updateLinks;

@end

@interface PowerPointPresenterTool : SBObject <PowerPointGenericMethods>

@property (copy, readonly) PowerPointSlide *currentPresenterSlide;
@property (readonly) NSInteger currentSlideInShow;
@property (readonly) double elapsedTime;
@property (readonly) NSInteger maxSlidesInShow;
@property (copy) NSString *notesText;
@property NSInteger notesZoom;
@property (readonly) BOOL slideMiniature;

- (void) endShow;
- (void) next;
- (void) pauseTimer;
- (void) previous;
- (void) resetTimer;
- (void) startTimer;
- (void) swapDisplays;
- (void) toggleSlideMiniature;

@end

@interface PowerPointPresenterViewWindow : SBObject <PowerPointGenericMethods>

@property (readonly) BOOL active;
@property (readonly) double height;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointPresenterTool *presenterTool;
@property (readonly) double width;


@end

@interface PowerPointPrintOptions : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointPrintRange *> *) printRanges;

@property (copy, readonly) NSString *activePrinter;
@property BOOL collate;
@property BOOL fitToPage;
@property BOOL frameSlides;
@property NSInteger numberOfCopies;
@property PowerPointEPPPrintOutputType outputType;
@property PowerPointEPPPrintColorType printColorType;
@property BOOL printFontsAsGraphics;
@property BOOL printHiddenSlides;
@property PowerPointEPPPrintRangeType rangeType;
@property NSInteger sectionIndex;
@property (copy) NSString *slideShowName;


@end

@interface PowerPointPrintRange : SBObject <PowerPointGenericMethods>

@property (readonly) NSInteger rangeEnd;
@property (readonly) NSInteger rangeStart;


@end

@interface PowerPointPropertyEffect : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointAnimationPoint *> *) animationPoints;

@property (copy, readonly) id ending;
@property PowerPointMsoAnimProperty propertySetEffect;
@property (copy, readonly) id starting;


@end

@interface PowerPointRotatingEffect : SBObject <PowerPointGenericMethods>

@property double rotating;


@end

@interface PowerPointRulerLevel : SBObject <PowerPointGenericMethods>

@property double firstMargin;  // Returns or sets the first-line indent for the specified outline level, in points.
@property double leftMargin;  // Returns or sets the left indent for the specified outline level, in points.


@end

// Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.
@interface PowerPointRuler : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointTabStop *> *) tabStops;
- (SBElementArray<PowerPointRulerLevel *> *) rulerLevels;


@end

@interface PowerPointSaveAsMovieSettings : SBObject <PowerPointGenericMethods>

@property BOOL autoLoopEnabled;
@property (copy) NSString *backgroundSoundTrackFile;
@property NSInteger backgroundTrackSegmentEnd;
@property NSInteger backgroundTrackSegmentStart;
@property NSInteger backgroundTrackStart;
@property BOOL createMoviePreview;
@property BOOL forceAllInline;
@property BOOL includeNarrationAndSounds;
@property (copy) NSString *movieActors;
@property (copy) NSString *movieAuthor;
@property (copy) NSString *movieCopyright;
@property NSInteger movieFrameHeight;
@property NSInteger movieFrameWidth;
@property (copy) NSString *movieProducer;
@property PowerPointEPPMovieOptimization optimization;
@property BOOL showMovieController;
@property (copy) NSString *transitionDescription;
@property BOOL useSingleTransition;


@end

@interface PowerPointScaleEffect : SBObject <PowerPointGenericMethods>

@property double byX;
@property double byY;
@property double fromX;
@property double fromY;
@property double toX;
@property double toY;


@end

@interface PowerPointSectionProperties : SBObject <PowerPointGenericMethods>

- (void) deleteSectionAtPosition:(NSInteger)atPosition deletingSlides:(BOOL)deletingSlides;
- (NSInteger) getCountOfSections;
- (NSInteger) getFirstSlideOfSectionAtPosition:(NSInteger)atPosition;
- (NSString *) getIdOfSectionAtPosition:(NSInteger)atPosition;
- (NSString *) getNameOfSectionAtPosition:(NSInteger)atPosition;
- (NSInteger) getSlideCountOfSectionAtPosition:(NSInteger)atPosition;
- (NSInteger) insertSectionBeforeSection:(NSInteger)beforeSection beforeSlide:(NSInteger)beforeSlide titled:(NSString *)titled;
- (void) moveSectionAtPosition:(NSInteger)atPosition toPosition:(NSInteger)toPosition;
- (void) renameSectionAtPosition:(NSInteger)atPosition to:(NSString *)to;

@end

// Represents the selection in the specified document window
@interface PowerPointSelection : SBObject <PowerPointGenericMethods>

@property (copy, readonly) PowerPointShapeRange *childShapeRange;  // Returns the child shapes of a selection.
@property (readonly) BOOL hasChildShapeRange;  // Returns whether the selection contains child shapes.
@property (readonly) PowerPointEPPSelectionType selectionType;  // Represents the type of objects in a selection.
@property (copy, readonly) PowerPointShapeRange *shapeRange;  // Returns a collection of shapes selected on the specified slide.
@property (copy, readonly) PowerPointSlideRange *slideRange;  // Returns a collection of selected slides.
@property (copy, readonly) PowerPointTextRange *textRange;  // Returns the text and text properties of the selected text.

- (void) unselect;  // Cancels the current selection.

@end

@interface PowerPointSequence : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointEffect *> *) effects;

- (PowerPointEffect *) addEffectFor:(PowerPointShape *)for_ fx:(PowerPointMsoAnimEffect)fx level:(PowerPointMsoAnimateByLevel)level trigger:(PowerPointMsoAnimTriggerType)trigger index:(NSInteger)index;  // add an effect for a shape
- (PowerPointEffect *) convertToTextUnitEffectEffect:(PowerPointEffect *)Effect unit:(PowerPointMsoAnimTextUnitEffect)unit;  // convert an effect to a text unit effect

@end

@interface PowerPointSetEffect : SBObject <PowerPointGenericMethods>

@property (copy) id ending;
@property PowerPointMsoAnimProperty propertySetEffect;


@end

// A collection that represents a notes page or a slide range, which is a set of slides that can contain a single slide to as many as all slides in a presentation.
@interface PowerPointSlideRange : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointSlide *> *) slides;
- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointHyperlink *> *) hyperlinks;

@property (copy) PowerPointColorScheme *colorScheme;  // Returns or sets the color scheme for the specified slide, slide range, or slide master.
@property (copy) PowerPointCustomLayout *customLayout;  // Returns the custom layout associated with the specified range of slides.
@property (copy) PowerPointDesign *design;
@property BOOL displayMasterShapes;  // Determines whether the specified range of slides displays the background objects on the slide master.
@property BOOL followMasterBackground;  // Determines whether the range of slides follows the slide master background.
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;  // Returns a collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides.
@property PowerPointEPPSlideLayout layout;  // Returns or sets the slide layout.
@property (copy, readonly) PowerPointSlideRange *notesPage;  // Returns a slide range object that represents the notes pages for the specified slide or range of slides.
@property (readonly) NSInteger printSteps;
@property (readonly) NSInteger slideID;
@property (readonly) NSInteger slideIndex;
@property (copy, readonly) PowerPointMaster *slideMaster;  // Returns the slide master.
@property (readonly) NSInteger slideNumber;  // Returns the slide number.
@property (copy, readonly) PowerPointSlideShowTransition *slideShowTransitions;  // Returns the special effects for the specified slide transition.
@property (copy, readonly) PowerPointTimeline *timeline;  // Returns the animation timeline for the slide.


@end

@interface PowerPointSlideShowSettings : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointNamedSlideShow *> *) namedSlideShows;

@property PowerPointEPPSlideShowAdvanceMode advanceMode;
@property NSInteger endingSlide;
@property BOOL loopUntilStopped;
@property (copy) NSArray<NSNumber *> *pointerColor;
@property PowerPointMsoThemeColorIndex pointerColorThemeIndex;  // Returns or sets the color of  default pointer color for a presentation.
@property PowerPointEPPSlideShowRangeType rangeType;
@property PowerPointEPPSlideShowType showType;
@property (readonly) BOOL showWithAnimation;
@property BOOL showWithNarration;
@property BOOL showWithPresenter;
@property (copy) NSString *slideShowName;
@property NSInteger startingSlide;

- (PowerPointSlideShowWindow *) runSlideShow;

@end

@interface PowerPointSlideShowTransition : SBObject <PowerPointGenericMethods>

@property BOOL advanceOnClick;
@property BOOL advanceOnTime;
@property double advanceTime;
@property PowerPointEPPEntryEffect entryEffect;
@property BOOL hidden;
@property BOOL loopSoundUntilNext;
@property (copy, readonly) PowerPointSoundEffect *soundEffectTransition;
@property double transitionDuration;


@end

@interface PowerPointSlideShowView : SBObject <PowerPointGenericMethods>

@property BOOL accelerationsEnabled;
@property (readonly) NSInteger currentShowPosition;
@property (readonly) BOOL isNamedShow;
@property (copy, readonly) PowerPointSlide *lastSlideViewed;
@property (copy) NSArray<NSNumber *> *pointerColor;
@property PowerPointMsoThemeColorIndex pointerColorThemeIndex;  // Returns or sets the color of temporary pointer color for a view of a slide show.
@property PowerPointEPPSlideShowPointerType pointerType;
@property (readonly) double presentationElapsedTime;
@property (copy, readonly) PowerPointSlide *slide;
@property double slideElapsedTime;
@property (copy, readonly) NSString *slideShowName;
@property PowerPointEPPSlideShowState slideState;
@property (readonly) NSInteger zoom;

- (void) exitSlideShow;
- (void) goToFirstSlide;
- (void) goToLastSlide;
- (void) goToNextSlide;
- (void) goToPreviousSlide;
- (void) resetSlideTime;

@end

@interface PowerPointSlideShowWindow : SBObject <PowerPointGenericMethods>

@property (readonly) BOOL active;
@property NSRect bounds;
@property double height;
@property (readonly) BOOL isFullScreen;
@property double leftPosition;
@property (copy, readonly) PowerPointPresentation *presentation;
@property (copy, readonly) PowerPointSlideShowView *slideshowView;
@property double top;
@property double width;


@end

@interface PowerPointSlide : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointHyperlink *> *) hyperlinks;

@property (copy, readonly) PowerPointShape *background;
@property (copy) PowerPointColorScheme *colorScheme;
@property (copy) PowerPointCustomLayout *customLayout;
@property (copy) PowerPointDesign *design;
@property BOOL displayMasterShapes;
@property BOOL followMasterBackground;
@property (copy, readonly) PowerPointHeadersAndFooters *headersAndFooters;
@property PowerPointEPPSlideLayout layout;
@property (copy, readonly) PowerPointSlide *notesPage;
@property (readonly) NSInteger printSteps;
@property (readonly) NSInteger sectionIndex;
@property (readonly) NSInteger sectionNumber;
@property (readonly) NSInteger slideID;
@property (readonly) NSInteger slideIndex;
@property (copy, readonly) PowerPointMaster *slideMaster;
@property (readonly) NSInteger slideNumber;
@property (copy, readonly) PowerPointSlideShowTransition *slideShowTransition;
@property (copy, readonly) PowerPointTimeline *timeline;

- (void) applyThemeColorSchemeFileName:(NSString *)fileName;  // Applies a theme color to the specified slide.
- (void) copyObject NS_RETURNS_NOT_RETAINED;
- (void) cutObject;
- (void) moveToStartOfSectionAtPosition:(NSInteger)atPosition;

@end

@interface PowerPointSoundEffect : SBObject <PowerPointGenericMethods>

@property (copy, readonly) NSString *name;
@property PowerPointEPPSoundEffectType soundType;

- (void) importSoundFileSoundFileName:(NSString *)soundFileName;
- (void) playSoundEffect;

@end

// Represents a single tab stop.
@interface PowerPointTabStop : SBObject <PowerPointGenericMethods>

@property double tabPosition;  // Returns or sets the position of a tab stop relative to the left margin.
@property PowerPointMsoTabStopType tabStopType;  // Returns or sets the type of the tab stop object.


@end

@interface PowerPointTextStyleLevel : SBObject <PowerPointGenericMethods>

@property (copy, readonly) PowerPointFont *font;
@property (copy, readonly) PowerPointParagraphFormat *paragraphFormat;


@end

@interface PowerPointTextStyle : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointTextStyleLevel *> *) textStyleLevels;

@property (copy, readonly) PowerPointRuler *ruler;
@property (copy, readonly) PowerPointTextFrame *textFrame;


@end

@interface PowerPointTimeline : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointSequence *> *) sequences;

@property (copy, readonly) PowerPointSequence *mainSequence;

- (PowerPointSequence *) addSequenceIndex:(NSInteger)index;  // add an interactive sequence

@end

@interface PowerPointTiming : SBObject <PowerPointGenericMethods>

@property double acceleration;
@property BOOL autoreverse;
@property double deceleration;
@property double duration;
@property NSInteger repeatCount;
@property double repeatDuration;
@property PowerPointMsoAnimEffectRestart restart;
@property BOOL rewind;
@property BOOL smoothEnd;
@property BOOL smoothStart;
@property double speed;


@end

// Represents a single initial-capital autocorrect exception.
@interface PowerPointTwoInitialCapsException : SBObject <PowerPointGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the TwoInitialCapsException.


@end

@interface PowerPointView : SBObject <PowerPointGenericMethods>

@property BOOL displaySlideMiniature;
@property (copy) PowerPointSlide *slide;
@property (readonly) PowerPointEPPViewType viewType;
@property NSInteger zoom;
@property BOOL zoomToFit;

- (void) goToSlideNumber:(NSInteger)number;

@end

@interface PowerPointWebOptions : SBObject <PowerPointGenericMethods>

@property BOOL allowPNG;
@property PowerPointMsoEncoding encoding;


@end



/*
 * Drawing Suite
 */

@interface PowerPointAdjustment : SBObject <PowerPointGenericMethods>

@property double adjustment_value;


@end

@interface PowerPointCalloutFormat : SBObject <PowerPointGenericMethods>

@property BOOL accent;
@property PowerPointMsoCalloutAngleType angle;
@property BOOL autoAttach;
@property (readonly) BOOL autoLength;
@property BOOL border;
@property (readonly) double calloutFormatLength;
@property PowerPointMsoCalloutType calloutType;
@property (readonly) double drop;
@property (readonly) PowerPointMsoCalloutDropType dropType;
@property double gap;

- (void) presetDropDropType:(PowerPointMsoCalloutDropType)dropType;

@end

@interface PowerPointConnectorFormat : SBObject <PowerPointGenericMethods>

@property (readonly) BOOL beginConnected;
@property (copy, readonly) PowerPointShape *beginConnectedShape;
@property (readonly) NSInteger beginConnectionSite;
@property PowerPointMsoConnectorType connectorType;
@property (readonly) BOOL endConnected;
@property (copy, readonly) PowerPointShape *endConnectedShape;
@property (readonly) NSInteger endConnectionSite;


@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface PowerPointFillFormat : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointGradientStop *> *) gradientStops;

@property (copy) NSArray<NSNumber *> *backColor;
@property PowerPointMsoThemeColorIndex backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) PowerPointMsoFillType fillFormatType;
@property (copy) NSArray<NSNumber *> *foreColor;
@property PowerPointMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) PowerPointMsoGradientColorType gradientColorType;
@property (readonly) double gradientDegree;
@property (readonly) PowerPointMsoGradientStyle gradientStyle;
@property (readonly) NSInteger gradientVariant;
@property (readonly) PowerPointMsoPatternType pattern;
@property (readonly) PowerPointMsoPresetGradientType presetGradientType;
@property (readonly) PowerPointMsoPresetTexture presetTexture;
@property BOOL rotateWithObject;  // Returns or sets whether the fill rotates with the specified shape.
@property PowerPointMsoTextureAlignment textureAlignment;  // Returns or sets the texture alignment for the specified object.
@property double textureHorizontalScale;  // Returns or sets the texture alignment for the specified object.
@property (copy, readonly) NSString *textureName;
@property double textureOffsetX;  // Returns or sets the texture alignment for the specified object.
@property double textureOffsetY;  // Returns or sets the texture alignment for the specified object.
@property BOOL textureTile;  // Returns the texture tile style for the specified fill.
@property double textureVerticalScale;  // Returns or sets the texture alignment for the specified object.
@property double transparency;
@property BOOL visible;

- (void) deleteGradientStopStopIndex:(NSInteger)stopIndex;  // Removes a gradient stop.
- (void) insertGradientStopCustomColor:(NSArray<NSNumber *> *)customColor position:(double)position transparency:(double)transparency stopIndex:(NSInteger)stopIndex;  // Adds a stop to a gradient.

@end

// Represents the glow formatting for a shape or range of shapes
@interface PowerPointGlowFormat : SBObject <PowerPointGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the color for the specified glow format.
@property PowerPointMsoThemeColorIndex colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents one gradient stop.
@interface PowerPointGradientStop : SBObject <PowerPointGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the color for the specified the gradient stop.
@property PowerPointMsoThemeColorIndex colorThemeIndex;  // Returns or sets the color for the specified gradient stop.
@property double position;  // Returns or sets the position for the specified gradient stop expressed as a percent.
@property double transparency;  // Returns or sets a value representing the transparency of the gradient fill expressed as a percent.


@end

@interface PowerPointLineFormat : SBObject <PowerPointGenericMethods>

@property (copy) NSArray<NSNumber *> *backColor;
@property PowerPointMsoThemeColorIndex backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property PowerPointMsoArrowheadLength beginArrowHeadLength;
@property PowerPointMsoArrowheadStyle beginArrowheadStyle;
@property PowerPointMsoArrowheadWidth beginArrowheadWidth;
@property PowerPointMsoLineDashStyle dashStyle;
@property PowerPointMsoArrowheadLength endArrowheadLength;
@property PowerPointMsoArrowheadStyle endArrowheadStyle;
@property PowerPointMsoArrowheadWidth endArrowheadWidth;
@property (copy) NSArray<NSNumber *> *foreColor;
@property PowerPointMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property PowerPointMsoPatternType lineFormatPatterned;
@property PowerPointMsoLineStyle lineStyle;
@property double lineWeight;
@property double transparency;


@end

@interface PowerPointLinkFormat : SBObject <PowerPointGenericMethods>

@property PowerPointEPPUpdateOption autoUpdate;
@property (copy) NSString *sourceFullName;


@end

// Represents a Microsoft Office theme.
@interface PowerPointOfficeTheme : SBObject <PowerPointGenericMethods>

@property (copy, readonly) PowerPointThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) PowerPointThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) PowerPointThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

@interface PowerPointPictureFormat : SBObject <PowerPointGenericMethods>

@property double brightness;  // Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property PowerPointMsoPictureColorType colorType;  // Returns or sets the type of color transformation applied to the specified picture.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSArray<NSNumber *> *transparencyColor;  // Returns or sets the transparent color for the specified picture as aRGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

@interface PowerPointPlaceholderFormat : SBObject <PowerPointGenericMethods>

@property (readonly) PowerPointMsoShapeType containedType;
@property (copy) NSString *placeholderName;
@property (readonly) PowerPointEPPPlaceholderType placeholderType;


@end

// Represents the reflection effect in Office graphics.
@interface PowerPointReflectionFormat : SBObject <PowerPointGenericMethods>

@property PowerPointMsoReflectionType reflectionType;  // Returns or sets the type of the reflection format object.


@end

@interface PowerPointShadowFormat : SBObject <PowerPointGenericMethods>

@property double blur;
@property (copy) NSArray<NSNumber *> *foreColor;
@property PowerPointMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;
@property double offsetX;
@property double offsetY;
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property PowerPointMsoShadowStyle shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property PowerPointMsoShadowType shadowType;
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;
@property BOOL visible;


@end

// A collection that represents a set of shapes on a document.
@interface PowerPointShapeRange : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointAdjustment *> *) adjustments;
- (SBElementArray<PowerPointShape *> *) shapes;

@property (copy, readonly) PowerPointAnimationSettings *animationSettings;  // Returns all the special effects that you can apply to the animation of the specified shape.
@property PowerPointMsoAutoShapeType autoShapeType;  // Returns or sets the shape type for AutoShapes other than a line, freeform drawing, or connector.
@property PowerPointMsoBackgroundStyleIndex backgroundStyle;  // Returns or sets the background style for the specified object.
@property PowerPointMsoBlackWhiteMode blackAndWhiteMode;  // Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode.
@property (copy, readonly) PowerPointCalloutFormat *calloutFormat;  // Returns callout formatting properties for the specified line callout shapes.
@property (readonly) BOOL child;  // Returns whether all shapes in a shape range are child shapes of the same parent.
@property (readonly) NSInteger connectionSiteCount;  // Returns the number of connection sites on the specified shape.
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns the fill format properties for the specified object.
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns the glow format for the specified range of shapes.
@property (readonly) BOOL hasChild;
@property (readonly) BOOL hasTable;  // Returns whether the specified shape is a table.
@property (readonly) BOOL hasTextFrame;  // Returns whether the specified shape has a text frame.
@property double height;  // Returns or sets the height of the specified object.
@property (readonly) BOOL horizontalFlip;  // Returns whether the specified shape is flipped around the horizontal axis.
@property (readonly) BOOL isConnector;  // Determines whether the specified shape is a connector.
@property double leftPosition;  // Returns or sets a value equal to the distance from the left edge of the leftmost shape in the shape range to the left edge of the slide.
@property (copy, readonly) PowerPointLineFormat *lineFormat;  // Returns line format properties for the specified line or shape border.
@property (copy, readonly) PowerPointLinkFormat *linkFormat;  // Returns the properties for linked OLE objects.
@property BOOL lockAspectRatio;  // Determines whether the specified shape retains its original proportions when you resize it.
@property (readonly) PowerPointEPPMediaType mediaType;  // Returns the OLE media type.
@property (copy) NSString *name;  // Identifies the shape on a slide.
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns the reflection format for the specified range of shapes.
@property double rotation;  // Returns or sets the number of degrees that the specified shape is rotated around the z-axis.
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;  // Returns shadow formatting for specified shapes.
@property PowerPointMsoShapeStyleIndex shapeStyle;  // Returns or sets the shape style index for the specified object.
@property (readonly) PowerPointMsoShapeType shapeType;  // Represents the type of shape or shapes in a range of shapes.
@property (copy, readonly) PowerPointSoftEdgeFormat *softEdgeFormat;  // Returns the soft edge format for the specified range of shapes.
@property (copy, readonly) PowerPointTextFrame *textFrame;  // Returns the alignment and anchoring properties for the specified shape or master text style.
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns 3-D effect formatting properties for the specified shape.
@property double top;  // Returns or sets a value equal to the distance from the top edge of the topmost shape in the shape range to the top edge of the document.
@property (readonly) BOOL verticalFlip;  // Determines whether the specified shape is flipped around the vertical axis.
@property BOOL visible;  // Returns or sets the visibility of the specified object or the formatting applied to the specified object.
@property double width;  // Returns or sets the width of the specified object.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order.

- (void) alignAlignType:(PowerPointMsoAlignCmd)alignType relativeToSlide:(BOOL)relativeToSlide;  // Aligns the shapes in the specified range of shapes.
- (void) copyShapeRange NS_RETURNS_NOT_RETAINED;
- (void) cutShapeRange;
- (void) distributeDistributeType:(PowerPointMsoDistributeCmd)distributeType relativeToSlide:(BOOL)relativeToSlide;  // Evenly distributes the shapes in the specified range of shapes. You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire slide or over the space they originally occupy.
- (PowerPointShape *) group;  // Groups the shapes in the specified shape range.
- (PowerPointShape *) regroup;  // Regroups the group that the specified shape range belonged to previously.
- (PowerPointShapeRange *) ungroup;  // Ungroups any grouped shapes in the specified shape or range of shapes.

@end

@interface PowerPointShape : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointShape *> *) shapes;
- (SBElementArray<PowerPointCallout *> *) callouts;
- (SBElementArray<PowerPointConnector *> *) connectors;
- (SBElementArray<PowerPointPicture *> *) pictures;
- (SBElementArray<PowerPointLineShape *> *) lineShapes;
- (SBElementArray<PowerPointPlaceHolder *> *) placeHolders;
- (SBElementArray<PowerPointTextBox *> *) textBoxes;
- (SBElementArray<PowerPointComment *> *) comments;
- (SBElementArray<PowerPointShapeTable *> *) shapeTables;
- (SBElementArray<PowerPointAdjustment *> *) adjustments;

@property (copy, readonly) PowerPointAnimationSettings *animationSettings;
@property PowerPointMsoAutoShapeType autoShapeType;
@property PowerPointMsoBackgroundStyleIndex backgroundStyle;  // Returns or sets the background style.
@property PowerPointMsoBlackWhiteMode blackAndWhiteMode;
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger connectionSiteCount;
@property (copy, readonly) PowerPointFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified shape.
@property (copy, readonly) PowerPointGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (readonly) BOOL hasTable;
@property (readonly) BOOL hasTextFrame;
@property double height;
@property (readonly) BOOL horizontalFlip;
@property (readonly) BOOL isConnector;
@property double leftPosition;
@property (copy, readonly) PowerPointLineFormat *lineFormat;
@property (copy, readonly) PowerPointLinkFormat *linkFormat;
@property BOOL lockAspectRatio;
@property (readonly) PowerPointEPPMediaType mediaType;
@property (copy) NSString *name;
@property (copy, readonly) PowerPointReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property double rotation;
@property (copy, readonly) PowerPointShadowFormat *shadowFormat;
@property PowerPointMsoShapeStyleIndex shapeStyle;  // Returns or sets the shape style corresponding to the Quick Styles.
@property (readonly) PowerPointMsoShapeType shapeType;
@property (copy, readonly) PowerPointSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) PowerPointTextFrame *textFrame;
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape.
@property double top;
@property (readonly) BOOL verticalFlip;
@property BOOL visible;
@property double width;
@property (copy, readonly) PowerPointWordArtFormat *wordArtFormat;  // Returns the formatting properties for a word art effect.
@property (readonly) NSInteger zOrderPosition;

- (void) copyShape NS_RETURNS_NOT_RETAINED;
- (void) cutShape;
- (void) saveAsPicturePictureType:(PowerPointMsoPictureType)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format

@end

@interface PowerPointCallout : PowerPointShape

@property (readonly) PowerPointMsoCalloutType calloutType;
@property (copy, readonly) PowerPointCalloutFormat *calloutFormat;


@end

@interface PowerPointComment : PowerPointShape


@end

@interface PowerPointConnector : PowerPointShape

@property (copy, readonly) PowerPointConnectorFormat *connectorFormat;
@property (readonly) PowerPointMsoConnectorType connectorType;


@end

// The line shape uses begin line X, begin line Y, end line X, and end line Y when created
@interface PowerPointLineShape : PowerPointShape

@property double beginLineX;
@property double beginLineY;
@property double endLineX;
@property double endLineY;


@end

@interface PowerPointMediaObject : PowerPointShape

@property (copy, readonly) NSString *fileName;


@end

@interface PowerPointMedia2Object : PowerPointShape

@property (copy, readonly) NSString *fileName;
@property (readonly) BOOL linkToFile;
@property (readonly) BOOL saveWithDocument;


@end

@interface PowerPointPicture : PowerPointShape

@property (copy, readonly) NSString *fileName;
@property (readonly) BOOL linkToFile;
@property (copy, readonly) PowerPointPictureFormat *pictureFormat;
@property (readonly) BOOL saveWithDocument;

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(PowerPointMsoScaleFrom)scale;
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(PowerPointMsoScaleFrom)scale;

@end

@interface PowerPointPlaceHolder : PowerPointShape

@property (copy, readonly) PowerPointPlaceholderFormat *placeHolderFormat;
@property (readonly) PowerPointEPPPlaceholderType placeholderType;


@end

@interface PowerPointShapeTable : PowerPointShape

@property (readonly) NSInteger numberOfColumns;
@property (readonly) NSInteger numberOfRows;
@property (copy, readonly) PowerPointTable *tableObject;


@end

// Represents the soft edge formatting for a shape or range of shapes
@interface PowerPointSoftEdgeFormat : SBObject <PowerPointGenericMethods>

@property PowerPointMsoSoftEdgeType softEdgeType;  // Returns or sets the type soft edge format object.


@end

@interface PowerPointTextBox : PowerPointShape

@property (readonly) PowerPointMsoTextOrientation textOrientation;


@end

// Represents a single text column.
@interface PowerPointTextColumn : SBObject <PowerPointGenericMethods>

@property NSInteger columnNumber;  // Returns or sets the index of the text column object.
@property double spacing;  // Returns or sets the spacing between text columns in a text column object.
@property PowerPointMsoTextDirection textDirection;  // Returns or sets the direction of text in the text column object.


@end

@interface PowerPointTextFrame : SBObject <PowerPointGenericMethods>

@property PowerPointMsoAutoSize autoSize;
@property (readonly) BOOL hasText;  // Returns whether the specified text frame has text.
@property PowerPointMsoHorizontalAnchor horizontalAnchor;
@property double marginBottom;
@property double marginLeft;
@property double marginRight;
@property double marginTop;
@property (readonly) PowerPointMsoTextOrientation orientation;
@property PowerPointMsoPathFormat pathFormat;  // Returns or sets the path type for the specified text frame.
@property (copy, readonly) PowerPointRuler *ruler;
@property (copy, readonly) PowerPointTextColumn *textColumn;  // Returns the text column object that represents the columns within the text frame.
@property PowerPointMsoTextOrientation textOrientation;
@property (copy, readonly) PowerPointTextRange *textRange;
@property (copy, readonly) PowerPointThreeDFormat *threeDFormat;  // Returns the 3-D effect formatting properties for the specified text.
@property PowerPointMsoVerticalAnchor verticalAnchor;
@property PowerPointMsoWarpFormat warpFormat;  // Returns or sets the warp type for the specified text frame.
@property (readonly) PowerPointMsoPresetTextEffect wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property BOOL wordWrap;  // Returns or sets text break lines within or past the boundaries of the shape.
@property PowerPointMsoPresetTextEffect wordartFormat;  // Returns or sets the WordArt type for the specified text frame.


@end

// Represents the color scheme of an Office Theme
@interface PowerPointThemeColorScheme : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointThemeColor *> *) themeColors;

- (NSArray<NSNumber *> *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file.
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface PowerPointThemeColor : SBObject <PowerPointGenericMethods>

@property (copy) NSArray<NSNumber *> *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) PowerPointMsoThemeColorSchemeIndex themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface PowerPointThemeEffectScheme : SBObject <PowerPointGenericMethods>

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file.

@end

// Represents the font scheme of a Microsoft Office theme.
@interface PowerPointThemeFontScheme : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointMinorThemeFont *> *) minorThemeFonts;
- (SBElementArray<PowerPointMajorThemeFont *> *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface PowerPointThemeFont : SBObject <PowerPointGenericMethods>

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface PowerPointMajorThemeFont : PowerPointThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface PowerPointMinorThemeFont : PowerPointThemeFont


@end

// Represents a shape's three-dimensional formatting.
@interface PowerPointThreeDFormat : SBObject <PowerPointGenericMethods>

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property PowerPointMsoBevelType bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property PowerPointMsoBevelType bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy) NSArray<NSNumber *> *contourColor;  // Returns or sets the color of the contour of an object or text.
@property PowerPointMsoThemeColorIndex contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy) NSArray<NSNumber *> *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property PowerPointMsoThemeColorIndex extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property PowerPointMsoExtrusionColorType extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property PowerPointMsoPresetCamera presetCamera;  // Returns a constant that represents the camera preset.
@property PowerPointMsoPresetExtrusionDirection presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property PowerPointMsoPresetLightingDirection presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property PowerPointMsoLightRigType presetLightingRig;  // Returns a constant that represents the lighting preset.
@property PowerPointMsoPresetLightingSoftness presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property PowerPointMsoPresetMaterial presetMaterial;  // Returns or sets the extrusion surface material.
@property PowerPointMsoPresetThreeDFormat presetThreeDFormat;  // Returns or sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

@interface PowerPointWordArtFormat : SBObject <PowerPointGenericMethods>

@property BOOL fontBold;
@property BOOL fontItalic;
@property (copy) NSString *fontName;
@property double fontSize;
@property BOOL kernedPairs;
@property BOOL normalizedHeight;
@property PowerPointMsoPresetTextEffectShape presetShape;
@property BOOL rotatedChars;
@property PowerPointMsoTextEffectAlignment textAlignment;
@property double tracking;
@property PowerPointMsoPresetTextEffect wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.
@property (copy) NSString *wordArtText;

- (void) toggleVerticalText;

@end



/*
 * Text Suite
 */

@interface PowerPointTextRange : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointCharacter *> *) characters;
- (SBElementArray<PowerPointWord *> *) words;
- (SBElementArray<PowerPointSentence *> *) sentences;
- (SBElementArray<PowerPointLine *> *) lines;
- (SBElementArray<PowerPointParagraph *> *) paragraphs;
- (SBElementArray<PowerPointTextFlow *> *) textFlows;

@property (readonly) double boundsHeight;
@property (readonly) double boundsWidth;
@property (copy) NSString *content;
@property (copy, readonly) PowerPointFont *font;
@property NSInteger indentLevel;
@property (readonly) double leftBounds;
@property (readonly) NSInteger offset;
@property (copy, readonly) PowerPointParagraphFormat *paragraphFormat;
@property (readonly) NSInteger textLength;
@property (readonly) double topBounds;

- (void) addPeriodsTo;
- (void) changeCaseTo:(PowerPointMsoTextChangeCase)to;
- (void) copyTextRange NS_RETURNS_NOT_RETAINED;
- (void) cutTextRange;
- (NSArray<NSNumber *> *) getRotatedTextBounds;  // Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }
- (PowerPointActionSetting *) getTextActionSettingResult:(PowerPointEPPMouseActivation)result;
- (void) insertTextTextRangeInsertWhere:(PowerPointMsoTextRangeInsertPosition)insertWhere newText:(NSString *)newText;  // Adds text to existing text.
- (void) pasteTextRange;
- (void) removePeriodsFrom;

@end

@interface PowerPointCharacter : PowerPointTextRange


@end

@interface PowerPointLine : PowerPointTextRange


@end

@interface PowerPointParagraph : PowerPointTextRange


@end

@interface PowerPointSentence : PowerPointTextRange


@end

@interface PowerPointTextFlow : PowerPointTextRange


@end

@interface PowerPointWord : PowerPointTextRange


@end



/*
 * Table Suite
 */

@interface PowerPointCell : SBObject <PowerPointGenericMethods>

@property (readonly) BOOL selected;
@property (copy, readonly) PowerPointShape *shape;

- (PowerPointLineFormat *) getBorderEdge:(PowerPointPpBorderType)edge;
- (void) mergeMergeWith:(PowerPointCell *)mergeWith;
- (void) splitNumberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns;

@end

@interface PowerPointColumn : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointCell *> *) cells;

@property double width;


@end

@interface PowerPointRow : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointCell *> *) cells;

@property double height;


@end

@interface PowerPointTable : SBObject <PowerPointGenericMethods>

- (SBElementArray<PowerPointColumn *> *) columns;
- (SBElementArray<PowerPointRow *> *) rows;

@property PowerPointMsoTextDirection tableDirection;

- (PowerPointCell *) getCellFromRow:(NSInteger)row column:(NSInteger)column;

@end

