
@objc public protocol SBObjectProtocol: NSObjectProtocol {
    func get() -> AnyObject!
}

@objc public protocol SBApplicationProtocol: SBObjectProtocol {
    func activate()
    var delegate: SBApplicationDelegate! { get set }
}

// MARK: PowerPointSaveOptions
@objc public enum PowerPointSaveOptions : AEKeyword {
    case Yes = 0x79657320 /* 'yes ' */
    case No = 0x6e6f2020 /* 'no  ' */
    case Ask = 0x61736b20 /* 'ask ' */
}

// MARK: PowerPointPrintingErrorHandling
@objc public enum PowerPointPrintingErrorHandling : AEKeyword {
    case Standard = 0x6c777374 /* 'lwst' */
    case Detailed = 0x6c776474 /* 'lwdt' */
}

// MARK: PowerPointMsoLineDashStyle
@objc public enum PowerPointMsoLineDashStyle : AEKeyword {
    case LineDashStyleUnset = 0x92fffe /* '\x00\x92\xff\xfe' */
    case LineDashStyleSolid = 0x930001 /* '\x00\x93\x00\x01' */
    case LineDashStyleSquareDot = 0x930002 /* '\x00\x93\x00\x02' */
    case LineDashStyleRoundDot = 0x930003 /* '\x00\x93\x00\x03' */
    case LineDashStyleDash = 0x930004 /* '\x00\x93\x00\x04' */
    case LineDashStyleDashDot = 0x930005 /* '\x00\x93\x00\x05' */
    case LineDashStyleDashDotDot = 0x930006 /* '\x00\x93\x00\x06' */
    case LineDashStyleLongDash = 0x930007 /* '\x00\x93\x00\x07' */
    case LineDashStyleLongDashDot = 0x930008 /* '\x00\x93\x00\x08' */
    case LineDashStyleLongDashDotDot = 0x930009 /* '\x00\x93\x00\t' */
    case LineDashStyleSystemDash = 0x93000a /* '\x00\x93\x00\n' */
    case LineDashStyleSystemDot = 0x93000b /* '\x00\x93\x00\x0b' */
    case LineDashStyleSystemDashDot = 0x93000c /* '\x00\x93\x00\x0c' */
}

// MARK: PowerPointMsoLineStyle
@objc public enum PowerPointMsoLineStyle : AEKeyword {
    case LineStyleUnset = 0x94fffe /* '\x00\x94\xff\xfe' */
    case SingleLine = 0x950001 /* '\x00\x95\x00\x01' */
    case ThinThinLine = 0x950002 /* '\x00\x95\x00\x02' */
    case ThinThickLine = 0x950003 /* '\x00\x95\x00\x03' */
    case ThickThinLine = 0x950004 /* '\x00\x95\x00\x04' */
    case ThickBetweenThinLine = 0x950005 /* '\x00\x95\x00\x05' */
}

// MARK: PowerPointMsoArrowheadStyle
@objc public enum PowerPointMsoArrowheadStyle : AEKeyword {
    case ArrowheadStyleUnset = 0x91fffe /* '\x00\x91\xff\xfe' */
    case NoArrowhead = 0x920001 /* '\x00\x92\x00\x01' */
    case TriangleArrowhead = 0x920002 /* '\x00\x92\x00\x02' */
    case Open_arrowhead = 0x920003 /* '\x00\x92\x00\x03' */
    case StealthArrowhead = 0x920004 /* '\x00\x92\x00\x04' */
    case DiamondArrowhead = 0x920005 /* '\x00\x92\x00\x05' */
    case OvalArrowhead = 0x920006 /* '\x00\x92\x00\x06' */
}

// MARK: PowerPointMsoArrowheadWidth
@objc public enum PowerPointMsoArrowheadWidth : AEKeyword {
    case ArrowheadWidthUnset = 0x90fffe /* '\x00\x90\xff\xfe' */
    case NarrowWidthArrowhead = 0x910001 /* '\x00\x91\x00\x01' */
    case MediumWidthArrowhead = 0x910002 /* '\x00\x91\x00\x02' */
    case WideArrowhead = 0x910003 /* '\x00\x91\x00\x03' */
}

// MARK: PowerPointMsoArrowheadLength
@objc public enum PowerPointMsoArrowheadLength : AEKeyword {
    case ArrowheadLengthUnset = 0x93fffe /* '\x00\x93\xff\xfe' */
    case ShortArrowhead = 0x940001 /* '\x00\x94\x00\x01' */
    case MediumArrowhead = 0x940002 /* '\x00\x94\x00\x02' */
    case LongArrowhead = 0x940003 /* '\x00\x94\x00\x03' */
}

// MARK: PowerPointMsoFillType
@objc public enum PowerPointMsoFillType : AEKeyword {
    case FillUnset = 0x63fffe /* '\x00c\xff\xfe' */
    case FillSolid = 0x640001 /* '\x00d\x00\x01' */
    case FillPatterned = 0x640002 /* '\x00d\x00\x02' */
    case FillGradient = 0x640003 /* '\x00d\x00\x03' */
    case FillTextured = 0x640004 /* '\x00d\x00\x04' */
    case FillBackground = 0x640005 /* '\x00d\x00\x05' */
    case FillPicture = 0x640006 /* '\x00d\x00\x06' */
}

// MARK: PowerPointMsoGradientStyle
@objc public enum PowerPointMsoGradientStyle : AEKeyword {
    case GradientUnset = 0x64fffe /* '\x00d\xff\xfe' */
    case HorizontalGradient = 0x650001 /* '\x00e\x00\x01' */
    case VerticalGradient = 0x650002 /* '\x00e\x00\x02' */
    case DiagonalUpGradient = 0x650003 /* '\x00e\x00\x03' */
    case DiagonalDownGradient = 0x650004 /* '\x00e\x00\x04' */
    case FromCornerGradient = 0x650005 /* '\x00e\x00\x05' */
    case FromTitleGradient = 0x650006 /* '\x00e\x00\x06' */
    case FromCenterGradient = 0x650007 /* '\x00e\x00\x07' */
}

// MARK: PowerPointMsoGradientColorType
@objc public enum PowerPointMsoGradientColorType : AEKeyword {
    case GradientTypeUnset = 0x3effffe /* '\x03\xef\xff\xfe' */
    case SingleShadeGradientType = 0x3f00001 /* '\x03\xf0\x00\x01' */
    case TwoColorsGradientType = 0x3f00002 /* '\x03\xf0\x00\x02' */
    case PresetColorsGradientType = 0x3f00003 /* '\x03\xf0\x00\x03' */
    case MultiColorsGradientType = 0x3f00004 /* '\x03\xf0\x00\x04' */
}

// MARK: PowerPointMsoTextureType
@objc public enum PowerPointMsoTextureType : AEKeyword {
    case TextureTypeTextureTypeUnset = 0x3f0fffe /* '\x03\xf0\xff\xfe' */
    case TextureTypePresetTexture = 0x3f10001 /* '\x03\xf1\x00\x01' */
    case TextureTypeUserDefinedTexture = 0x3f10002 /* '\x03\xf1\x00\x02' */
}

// MARK: PowerPointMsoPresetTexture
@objc public enum PowerPointMsoPresetTexture : AEKeyword {
    case PresetTextureUnset = 0x65fffe /* '\x00e\xff\xfe' */
    case TexturePapyrus = 0x660001 /* '\x00f\x00\x01' */
    case TextureCanvas = 0x660002 /* '\x00f\x00\x02' */
    case TextureDenim = 0x660003 /* '\x00f\x00\x03' */
    case TextureWovenMat = 0x660004 /* '\x00f\x00\x04' */
    case TextureWaterDroplets = 0x660005 /* '\x00f\x00\x05' */
    case TexturePaperBag = 0x660006 /* '\x00f\x00\x06' */
    case TextureFishFossil = 0x660007 /* '\x00f\x00\x07' */
    case TextureSand = 0x660008 /* '\x00f\x00\x08' */
    case TextureGreenMarble = 0x660009 /* '\x00f\x00\t' */
    case TextureWhiteMarble = 0x66000a /* '\x00f\x00\n' */
    case TextureBrownMarble = 0x66000b /* '\x00f\x00\x0b' */
    case TextureGranite = 0x66000c /* '\x00f\x00\x0c' */
    case TextureNewsprint = 0x66000d /* '\x00f\x00\r' */
    case TextureRecycledPaper = 0x66000e /* '\x00f\x00\x0e' */
    case TextureParchment = 0x66000f /* '\x00f\x00\x0f' */
    case TextureStationery = 0x660010 /* '\x00f\x00\x10' */
    case TextureBlueTissuePaper = 0x660011 /* '\x00f\x00\x11' */
    case TexturePinkTissuePaper = 0x660012 /* '\x00f\x00\x12' */
    case TexturePurpleMesh = 0x660013 /* '\x00f\x00\x13' */
    case TextureBouquet = 0x660014 /* '\x00f\x00\x14' */
    case TextureCork = 0x660015 /* '\x00f\x00\x15' */
    case TextureWalnut = 0x660016 /* '\x00f\x00\x16' */
    case TextureOak = 0x660017 /* '\x00f\x00\x17' */
    case TextureMediumWood = 0x660018 /* '\x00f\x00\x18' */
}

// MARK: PowerPointMsoPatternType
@objc public enum PowerPointMsoPatternType : AEKeyword {
    case PatternUnset = 0x66fffe /* '\x00f\xff\xfe' */
    case FivePercentPattern = 0x670001 /* '\x00g\x00\x01' */
    case TenPercentPattern = 0x670002 /* '\x00g\x00\x02' */
    case TwentyPercentPattern = 0x670003 /* '\x00g\x00\x03' */
    case TwentyFivePercentPattern = 0x670004 /* '\x00g\x00\x04' */
    case ThirtyPercentPattern = 0x670005 /* '\x00g\x00\x05' */
    case FortyPercentPattern = 0x670006 /* '\x00g\x00\x06' */
    case FiftyPercentPattern = 0x670007 /* '\x00g\x00\x07' */
    case SixtyPercentPattern = 0x670008 /* '\x00g\x00\x08' */
    case SeventyPercentPattern = 0x670009 /* '\x00g\x00\t' */
    case SeventyFivePercentPattern = 0x67000a /* '\x00g\x00\n' */
    case EightyPercentPattern = 0x67000b /* '\x00g\x00\x0b' */
    case NinetyPercentPattern = 0x67000c /* '\x00g\x00\x0c' */
    case DarkHorizontalPattern = 0x67000d /* '\x00g\x00\r' */
    case DarkVerticalPattern = 0x67000e /* '\x00g\x00\x0e' */
    case DarkDownwardDiagonalPattern = 0x67000f /* '\x00g\x00\x0f' */
    case DarkUpwardDiagonalPattern = 0x670010 /* '\x00g\x00\x10' */
    case SmallCheckerBoardPattern = 0x670011 /* '\x00g\x00\x11' */
    case TrellisPattern = 0x670012 /* '\x00g\x00\x12' */
    case LightHorizontalPattern = 0x670013 /* '\x00g\x00\x13' */
    case LightVerticalPattern = 0x670014 /* '\x00g\x00\x14' */
    case LightDownwardDiagonalPattern = 0x670015 /* '\x00g\x00\x15' */
    case LightUpwardDiagonalPattern = 0x670016 /* '\x00g\x00\x16' */
    case SmallGridPattern = 0x670017 /* '\x00g\x00\x17' */
    case DottedDiamondPattern = 0x670018 /* '\x00g\x00\x18' */
    case WideDownwardDiagonal = 0x670019 /* '\x00g\x00\x19' */
    case WideUpwardDiagonalPattern = 0x67001a /* '\x00g\x00\x1a' */
    case DashedUpwardDiagonalPattern = 0x67001b /* '\x00g\x00\x1b' */
    case DashedDownwardDiagonalPattern = 0x67001c /* '\x00g\x00\x1c' */
    case NarrowVerticalPattern = 0x67001d /* '\x00g\x00\x1d' */
    case NarrowHorizontalPattern = 0x67001e /* '\x00g\x00\x1e' */
    case DashedVerticalPattern = 0x67001f /* '\x00g\x00\x1f' */
    case DashedHorizontalPattern = 0x670020 /* '\x00g\x00 ' */
    case LargeConfettiPattern = 0x670021 /* '\x00g\x00!' */
    case LargeGridPattern = 0x670022 /* '\x00g\x00"' */
    case HorizontalBrickPattern = 0x670023 /* '\x00g\x00#' */
    case LargeCheckerBoardPattern = 0x670024 /* '\x00g\x00$' */
    case SmallConfettiPattern = 0x670025 /* '\x00g\x00%' */
    case ZigZagPattern = 0x670026 /* '\x00g\x00&' */
    case SolidDiamondPattern = 0x670027 /* "\x00g\x00'" */
    case DiagonalBrickPattern = 0x670028 /* '\x00g\x00(' */
    case OutlinedDiamondPattern = 0x670029 /* '\x00g\x00)' */
    case PlaidPattern = 0x67002a /* '\x00g\x00*' */
    case SpherePattern = 0x67002b /* '\x00g\x00+' */
    case WeavePattern = 0x67002c /* '\x00g\x00,' */
    case DottedGridPattern = 0x67002d /* '\x00g\x00-' */
    case DivotPattern = 0x67002e /* '\x00g\x00.' */
    case ShinglePattern = 0x67002f /* '\x00g\x00/' */
    case WavePattern = 0x670030 /* '\x00g\x000' */
    case HorizontalPattern = 0x670031 /* '\x00g\x001' */
    case VerticalPattern = 0x670032 /* '\x00g\x002' */
    case CrossPattern = 0x670033 /* '\x00g\x003' */
    case DownwardDiagonalPattern = 0x670034 /* '\x00g\x004' */
    case UpwardDiagonalPattern = 0x670035 /* '\x00g\x005' */
    case DiagonalCrossPattern = 0x670035 /* '\x00g\x005' */
}

// MARK: PowerPointMsoPresetGradientType
@objc public enum PowerPointMsoPresetGradientType : AEKeyword {
    case PresetGradientUnset = 0x67fffe /* '\x00g\xff\xfe' */
    case GradientEarlySunset = 0x680001 /* '\x00h\x00\x01' */
    case GradientLateSunset = 0x680002 /* '\x00h\x00\x02' */
    case GradientNightfall = 0x680003 /* '\x00h\x00\x03' */
    case GradientDaybreak = 0x680004 /* '\x00h\x00\x04' */
    case GradientHorizon = 0x680005 /* '\x00h\x00\x05' */
    case GradientDesert = 0x680006 /* '\x00h\x00\x06' */
    case GradientOcean = 0x680007 /* '\x00h\x00\x07' */
    case GradientCalmWater = 0x680008 /* '\x00h\x00\x08' */
    case GradientFire = 0x680009 /* '\x00h\x00\t' */
    case GradientFog = 0x68000a /* '\x00h\x00\n' */
    case GradientMoss = 0x68000b /* '\x00h\x00\x0b' */
    case GradientPeacock = 0x68000c /* '\x00h\x00\x0c' */
    case GradientWheat = 0x68000d /* '\x00h\x00\r' */
    case GradientParchment = 0x68000e /* '\x00h\x00\x0e' */
    case GradientMahogany = 0x68000f /* '\x00h\x00\x0f' */
    case GradientRainbow = 0x680010 /* '\x00h\x00\x10' */
    case GradientRainbow2 = 0x680011 /* '\x00h\x00\x11' */
    case GradientGold = 0x680012 /* '\x00h\x00\x12' */
    case GradientGold2 = 0x680013 /* '\x00h\x00\x13' */
    case GradientBrass = 0x680014 /* '\x00h\x00\x14' */
    case GradientChrome = 0x680015 /* '\x00h\x00\x15' */
    case GradientChrome2 = 0x680016 /* '\x00h\x00\x16' */
    case GradientSilver = 0x680017 /* '\x00h\x00\x17' */
    case GradientSapphire = 0x680018 /* '\x00h\x00\x18' */
}

// MARK: PowerPointMsoShadowType
@objc public enum PowerPointMsoShadowType : AEKeyword {
    case ShadowUnset = 0x35ffffe /* '\x03_\xff\xfe' */
    case Shadow1 = 0x3600001 /* '\x03`\x00\x01' */
    case Shadow2 = 0x3600002 /* '\x03`\x00\x02' */
    case Shadow3 = 0x3600003 /* '\x03`\x00\x03' */
    case Shadow4 = 0x3600004 /* '\x03`\x00\x04' */
    case Shadow5 = 0x3600005 /* '\x03`\x00\x05' */
    case Shadow6 = 0x3600006 /* '\x03`\x00\x06' */
    case Shadow7 = 0x3600007 /* '\x03`\x00\x07' */
    case Shadow8 = 0x3600008 /* '\x03`\x00\x08' */
    case Shadow9 = 0x3600009 /* '\x03`\x00\t' */
    case Shadow10 = 0x360000a /* '\x03`\x00\n' */
    case Shadow11 = 0x360000b /* '\x03`\x00\x0b' */
    case Shadow12 = 0x360000c /* '\x03`\x00\x0c' */
    case Shadow13 = 0x360000d /* '\x03`\x00\r' */
    case Shadow14 = 0x360000e /* '\x03`\x00\x0e' */
    case Shadow15 = 0x360000f /* '\x03`\x00\x0f' */
    case Shadow16 = 0x3600010 /* '\x03`\x00\x10' */
    case Shadow17 = 0x3600011 /* '\x03`\x00\x11' */
    case Shadow18 = 0x3600012 /* '\x03`\x00\x12' */
    case Shadow19 = 0x3600013 /* '\x03`\x00\x13' */
    case Shadow20 = 0x3600014 /* '\x03`\x00\x14' */
    case Shadow21 = 0x3600015 /* '\x03`\x00\x15' */
    case Shadow22 = 0x3600016 /* '\x03`\x00\x16' */
    case Shadow23 = 0x3600017 /* '\x03`\x00\x17' */
    case Shadow24 = 0x3600018 /* '\x03`\x00\x18' */
    case Shadow25 = 0x3600019 /* '\x03`\x00\x19' */
    case Shadow26 = 0x360001a /* '\x03`\x00\x1a' */
    case Shadow27 = 0x360001b /* '\x03`\x00\x1b' */
    case Shadow28 = 0x360001c /* '\x03`\x00\x1c' */
    case Shadow29 = 0x360001d /* '\x03`\x00\x1d' */
    case Shadow30 = 0x360001e /* '\x03`\x00\x1e' */
    case Shadow31 = 0x360001f /* '\x03`\x00\x1f' */
    case Shadow32 = 0x3600020 /* '\x03`\x00 ' */
    case Shadow33 = 0x3600021 /* '\x03`\x00!' */
    case Shadow34 = 0x3600022 /* '\x03`\x00"' */
    case Shadow35 = 0x3600023 /* '\x03`\x00#' */
    case Shadow36 = 0x3600024 /* '\x03`\x00$' */
    case Shadow37 = 0x3600025 /* '\x03`\x00%' */
    case Shadow38 = 0x3600026 /* '\x03`\x00&' */
    case Shadow39 = 0x3600027 /* "\x03`\x00'" */
    case Shadow40 = 0x3600028 /* '\x03`\x00(' */
    case Shadow41 = 0x3600029 /* '\x03`\x00)' */
    case Shadow42 = 0x360002a /* '\x03`\x00*' */
    case Shadow43 = 0x360002b /* '\x03`\x00+' */
}

// MARK: PowerPointMsoPresetTextEffect
@objc public enum PowerPointMsoPresetTextEffect : AEKeyword {
    case WordartFormatUnset = 0x3f1fffe /* '\x03\xf1\xff\xfe' */
    case WordartFormat1 = 0x3f20000 /* '\x03\xf2\x00\x00' */
    case WordartFormat2 = 0x3f20001 /* '\x03\xf2\x00\x01' */
    case WordartFormat3 = 0x3f20002 /* '\x03\xf2\x00\x02' */
    case WordartFormat4 = 0x3f20003 /* '\x03\xf2\x00\x03' */
    case WordartFormat5 = 0x3f20004 /* '\x03\xf2\x00\x04' */
    case WordartFormat6 = 0x3f20005 /* '\x03\xf2\x00\x05' */
    case WordartFormat7 = 0x3f20006 /* '\x03\xf2\x00\x06' */
    case WordartFormat8 = 0x3f20007 /* '\x03\xf2\x00\x07' */
    case WordartFormat9 = 0x3f20008 /* '\x03\xf2\x00\x08' */
    case WordartFormat10 = 0x3f20009 /* '\x03\xf2\x00\t' */
    case WordartFormat11 = 0x3f2000a /* '\x03\xf2\x00\n' */
    case WordartFormat12 = 0x3f2000b /* '\x03\xf2\x00\x0b' */
    case WordartFormat13 = 0x3f2000c /* '\x03\xf2\x00\x0c' */
    case WordartFormat14 = 0x3f2000d /* '\x03\xf2\x00\r' */
    case WordartFormat15 = 0x3f2000e /* '\x03\xf2\x00\x0e' */
    case WordartFormat16 = 0x3f2000f /* '\x03\xf2\x00\x0f' */
    case WordartFormat17 = 0x3f20010 /* '\x03\xf2\x00\x10' */
    case WordartFormat18 = 0x3f20011 /* '\x03\xf2\x00\x11' */
    case WordartFormat19 = 0x3f20012 /* '\x03\xf2\x00\x12' */
    case WordartFormat20 = 0x3f20013 /* '\x03\xf2\x00\x13' */
    case WordartFormat21 = 0x3f20014 /* '\x03\xf2\x00\x14' */
    case WordartFormat22 = 0x3f20015 /* '\x03\xf2\x00\x15' */
    case WordartFormat23 = 0x3f20016 /* '\x03\xf2\x00\x16' */
    case WordartFormat24 = 0x3f20017 /* '\x03\xf2\x00\x17' */
    case WordartFormat25 = 0x3f20018 /* '\x03\xf2\x00\x18' */
    case WordartFormat26 = 0x3f20019 /* '\x03\xf2\x00\x19' */
    case WordartFormat27 = 0x3f2001a /* '\x03\xf2\x00\x1a' */
    case WordartFormat28 = 0x3f2001b /* '\x03\xf2\x00\x1b' */
    case WordartFormat29 = 0x3f2001c /* '\x03\xf2\x00\x1c' */
    case WordartFormat30 = 0x3f2001d /* '\x03\xf2\x00\x1d' */
    case WordartFormat31 = 0x3f2001e /* '\x03\xf2\x00\x1e' */
    case WordartFormat32 = 0x3f2001f /* '\x03\xf2\x00\x1f' */
    case WordartFormat33 = 0x3f20020 /* '\x03\xf2\x00 ' */
    case WordartFormat34 = 0x3f20021 /* '\x03\xf2\x00!' */
    case WordartFormat35 = 0x3f20022 /* '\x03\xf2\x00"' */
    case WordartFormat36 = 0x3f20023 /* '\x03\xf2\x00#' */
    case WordartFormat37 = 0x3f20024 /* '\x03\xf2\x00$' */
    case WordartFormat38 = 0x3f20025 /* '\x03\xf2\x00%' */
    case WordartFormat39 = 0x3f20026 /* '\x03\xf2\x00&' */
    case WordartFormat40 = 0x3f20027 /* "\x03\xf2\x00'" */
    case WordartFormat41 = 0x3f20028 /* '\x03\xf2\x00(' */
    case WordartFormat42 = 0x3f20029 /* '\x03\xf2\x00)' */
    case WordartFormat43 = 0x3f2002a /* '\x03\xf2\x00*' */
    case WordartFormat44 = 0x3f2002b /* '\x03\xf2\x00+' */
    case WordartFormat45 = 0x3f2002c /* '\x03\xf2\x00,' */
    case WordartFormat46 = 0x3f2002d /* '\x03\xf2\x00-' */
    case WordartFormat47 = 0x3f2002e /* '\x03\xf2\x00.' */
    case WordartFormat48 = 0x3f2002f /* '\x03\xf2\x00/' */
    case WordartFormat49 = 0x3f20030 /* '\x03\xf2\x000' */
    case WordartFormat50 = 0x3f20031 /* '\x03\xf2\x001' */
}

// MARK: PowerPointMsoPresetTextEffectShape
@objc public enum PowerPointMsoPresetTextEffectShape : AEKeyword {
    case TextEffectShapeUnset = 0x97fffe /* '\x00\x97\xff\xfe' */
    case PlainText = 0x980001 /* '\x00\x98\x00\x01' */
    case Stop = 0x980002 /* '\x00\x98\x00\x02' */
    case TriangleUp = 0x980003 /* '\x00\x98\x00\x03' */
    case TriangleDown = 0x980004 /* '\x00\x98\x00\x04' */
    case ChevronUp = 0x980005 /* '\x00\x98\x00\x05' */
    case ChevronDown = 0x980006 /* '\x00\x98\x00\x06' */
    case RingInside = 0x980007 /* '\x00\x98\x00\x07' */
    case RingOutside = 0x980008 /* '\x00\x98\x00\x08' */
    case ArchUpCurve = 0x980009 /* '\x00\x98\x00\t' */
    case ArchDownCurve = 0x98000a /* '\x00\x98\x00\n' */
    case CircleCurve = 0x98000b /* '\x00\x98\x00\x0b' */
    case ButtonCurve = 0x98000c /* '\x00\x98\x00\x0c' */
    case ArchUpPour = 0x98000d /* '\x00\x98\x00\r' */
    case ArchDownPour = 0x98000e /* '\x00\x98\x00\x0e' */
    case CirclePour = 0x98000f /* '\x00\x98\x00\x0f' */
    case ButtonPour = 0x980010 /* '\x00\x98\x00\x10' */
    case CurveUp = 0x980011 /* '\x00\x98\x00\x11' */
    case CurveDown = 0x980012 /* '\x00\x98\x00\x12' */
    case CanUp = 0x980013 /* '\x00\x98\x00\x13' */
    case CanDown = 0x980014 /* '\x00\x98\x00\x14' */
    case Wave1 = 0x980015 /* '\x00\x98\x00\x15' */
    case Wave2 = 0x980016 /* '\x00\x98\x00\x16' */
    case DoubleWave1 = 0x980017 /* '\x00\x98\x00\x17' */
    case DoubleWave2 = 0x980018 /* '\x00\x98\x00\x18' */
    case Inflate = 0x980019 /* '\x00\x98\x00\x19' */
    case Deflate = 0x98001a /* '\x00\x98\x00\x1a' */
    case InflateBottom = 0x98001b /* '\x00\x98\x00\x1b' */
    case DeflateBottom = 0x98001c /* '\x00\x98\x00\x1c' */
    case InflateTop = 0x98001d /* '\x00\x98\x00\x1d' */
    case DeflateTop = 0x98001e /* '\x00\x98\x00\x1e' */
    case DeflateInflate = 0x98001f /* '\x00\x98\x00\x1f' */
    case DeflateInflateDeflate = 0x980020 /* '\x00\x98\x00 ' */
    case FadeRight = 0x980021 /* '\x00\x98\x00!' */
    case FadeLeft = 0x980022 /* '\x00\x98\x00"' */
    case FadeUp = 0x980023 /* '\x00\x98\x00#' */
    case FadeDown = 0x980024 /* '\x00\x98\x00$' */
    case SlantUp = 0x980025 /* '\x00\x98\x00%' */
    case SlantDown = 0x980026 /* '\x00\x98\x00&' */
    case CascadeUp = 0x980027 /* "\x00\x98\x00'" */
    case CascadeDown = 0x980028 /* '\x00\x98\x00(' */
}

// MARK: PowerPointMsoTextEffectAlignment
@objc public enum PowerPointMsoTextEffectAlignment : AEKeyword {
    case TextEffectAlignmentUnset = 0x96fffe /* '\x00\x96\xff\xfe' */
    case LeftTextEffectAlignment = 0x970001 /* '\x00\x97\x00\x01' */
    case CenteredTextEffectAlignment = 0x970002 /* '\x00\x97\x00\x02' */
    case RightTextEffectAlignment = 0x970003 /* '\x00\x97\x00\x03' */
    case JustifyTextEffectAlignment = 0x970004 /* '\x00\x97\x00\x04' */
    case WordJustifyTextEffectAlignment = 0x970005 /* '\x00\x97\x00\x05' */
    case StretchJustifyTextEffectAlignment = 0x970006 /* '\x00\x97\x00\x06' */
}

// MARK: PowerPointMsoPresetLightingDirection
@objc public enum PowerPointMsoPresetLightingDirection : AEKeyword {
    case PresetLightingDirectionUnset = 0x9bfffe /* '\x00\x9b\xff\xfe' */
    case LightFromTopLeft = 0x9c0001 /* '\x00\x9c\x00\x01' */
    case LightFromTop = 0x9c0002 /* '\x00\x9c\x00\x02' */
    case LightFromTopRight = 0x9c0003 /* '\x00\x9c\x00\x03' */
    case LightFromLeft = 0x9c0004 /* '\x00\x9c\x00\x04' */
    case LightFromNone = 0x9c0005 /* '\x00\x9c\x00\x05' */
    case LightFromRight = 0x9c0006 /* '\x00\x9c\x00\x06' */
    case LightFromBottomLeft = 0x9c0007 /* '\x00\x9c\x00\x07' */
    case LightFromBottom = 0x9c0008 /* '\x00\x9c\x00\x08' */
    case LightFromBottomRight = 0x9c0009 /* '\x00\x9c\x00\t' */
}

// MARK: PowerPointMsoPresetLightingSoftness
@objc public enum PowerPointMsoPresetLightingSoftness : AEKeyword {
    case LightingSoftnessUnset = 0x9cfffe /* '\x00\x9c\xff\xfe' */
    case LightingDim = 0x9d0001 /* '\x00\x9d\x00\x01' */
    case LightingNormal = 0x9d0002 /* '\x00\x9d\x00\x02' */
    case LightingBright = 0x9d0003 /* '\x00\x9d\x00\x03' */
}

// MARK: PowerPointMsoPresetMaterial
@objc public enum PowerPointMsoPresetMaterial : AEKeyword {
    case PresetMaterialUnset = 0x9dfffe /* '\x00\x9d\xff\xfe' */
    case Matte = 0x9e0001 /* '\x00\x9e\x00\x01' */
    case Plastic = 0x9e0002 /* '\x00\x9e\x00\x02' */
    case Metal = 0x9e0003 /* '\x00\x9e\x00\x03' */
    case Wireframe = 0x9e0004 /* '\x00\x9e\x00\x04' */
    case Matte2 = 0x9e0005 /* '\x00\x9e\x00\x05' */
    case Plastic2 = 0x9e0006 /* '\x00\x9e\x00\x06' */
    case Metal2 = 0x9e0007 /* '\x00\x9e\x00\x07' */
    case WarmMatte = 0x9e0008 /* '\x00\x9e\x00\x08' */
    case TranslucentPowder = 0x9e0009 /* '\x00\x9e\x00\t' */
    case Powder = 0x9e000a /* '\x00\x9e\x00\n' */
    case DarkEdge = 0x9e000b /* '\x00\x9e\x00\x0b' */
    case SoftEdge = 0x9e000c /* '\x00\x9e\x00\x0c' */
    case MaterialClear = 0x9e000d /* '\x00\x9e\x00\r' */
    case Flat = 0x9e000e /* '\x00\x9e\x00\x0e' */
    case SoftMetal = 0x9e000f /* '\x00\x9e\x00\x0f' */
}

// MARK: PowerPointMsoPresetExtrusionDirection
@objc public enum PowerPointMsoPresetExtrusionDirection : AEKeyword {
    case PresetExtrusionDirectionUnset = 0x99fffe /* '\x00\x99\xff\xfe' */
    case ExtrudeBottomRight = 0x9a0001 /* '\x00\x9a\x00\x01' */
    case ExtrudeBottom = 0x9a0002 /* '\x00\x9a\x00\x02' */
    case ExtrudeBottomLeft = 0x9a0003 /* '\x00\x9a\x00\x03' */
    case ExtrudeRight = 0x9a0004 /* '\x00\x9a\x00\x04' */
    case ExtrudeNone = 0x9a0005 /* '\x00\x9a\x00\x05' */
    case ExtrudeLeft = 0x9a0006 /* '\x00\x9a\x00\x06' */
    case ExtrudeTopRight = 0x9a0007 /* '\x00\x9a\x00\x07' */
    case ExtrudeTop = 0x9a0008 /* '\x00\x9a\x00\x08' */
    case ExtrudeTopLeft = 0x9a0009 /* '\x00\x9a\x00\t' */
}

// MARK: PowerPointMsoPresetThreeDFormat
@objc public enum PowerPointMsoPresetThreeDFormat : AEKeyword {
    case PresetThreeDFormatUnset = 0x98fffe /* '\x00\x98\xff\xfe' */
    case Format1 = 0x990001 /* '\x00\x99\x00\x01' */
    case Format2 = 0x990002 /* '\x00\x99\x00\x02' */
    case Format3 = 0x990003 /* '\x00\x99\x00\x03' */
    case Format4 = 0x990004 /* '\x00\x99\x00\x04' */
    case Format5 = 0x990005 /* '\x00\x99\x00\x05' */
    case Format6 = 0x990006 /* '\x00\x99\x00\x06' */
    case Format7 = 0x990007 /* '\x00\x99\x00\x07' */
    case Format8 = 0x990008 /* '\x00\x99\x00\x08' */
    case Format9 = 0x990009 /* '\x00\x99\x00\t' */
    case Format10 = 0x99000a /* '\x00\x99\x00\n' */
    case Format11 = 0x99000b /* '\x00\x99\x00\x0b' */
    case Format12 = 0x99000c /* '\x00\x99\x00\x0c' */
    case Format13 = 0x99000d /* '\x00\x99\x00\r' */
    case Format14 = 0x99000e /* '\x00\x99\x00\x0e' */
    case Format15 = 0x99000f /* '\x00\x99\x00\x0f' */
    case Format16 = 0x990010 /* '\x00\x99\x00\x10' */
    case Format17 = 0x990011 /* '\x00\x99\x00\x11' */
    case Format18 = 0x990012 /* '\x00\x99\x00\x12' */
    case Format19 = 0x990013 /* '\x00\x99\x00\x13' */
    case Format20 = 0x990014 /* '\x00\x99\x00\x14' */
}

// MARK: PowerPointMsoExtrusionColorType
@objc public enum PowerPointMsoExtrusionColorType : AEKeyword {
    case ExtrusionColorUnset = 0x9afffe /* '\x00\x9a\xff\xfe' */
    case ExtrusionColorAutomatic = 0x9b0001 /* '\x00\x9b\x00\x01' */
    case ExtrusionColorCustom = 0x9b0002 /* '\x00\x9b\x00\x02' */
}

// MARK: PowerPointMsoConnectorType
@objc public enum PowerPointMsoConnectorType : AEKeyword {
    case ConnectorTypeUnset = 0x68fffe /* '\x00h\xff\xfe' */
    case Straight = 0x690001 /* '\x00i\x00\x01' */
    case Elbow = 0x690002 /* '\x00i\x00\x02' */
    case Curve = 0x690003 /* '\x00i\x00\x03' */
}

// MARK: PowerPointMsoHorizontalAnchor
@objc public enum PowerPointMsoHorizontalAnchor : AEKeyword {
    case HorizontalAnchorUnset = 0x9efffe /* '\x00\x9e\xff\xfe' */
    case HorizontalAnchorNone = 0x9f0001 /* '\x00\x9f\x00\x01' */
    case HorizontalAnchorCenter = 0x9f0002 /* '\x00\x9f\x00\x02' */
}

// MARK: PowerPointMsoVerticalAnchor
@objc public enum PowerPointMsoVerticalAnchor : AEKeyword {
    case VerticalAnchorUnset = 0x9ffffe /* '\x00\x9f\xff\xfe' */
    case AnchorTop = 0xa00001 /* '\x00\xa0\x00\x01' */
    case AnchorTopBaseline = 0xa00002 /* '\x00\xa0\x00\x02' */
    case AnchorMiddle = 0xa00003 /* '\x00\xa0\x00\x03' */
    case AnchorBottom = 0xa00004 /* '\x00\xa0\x00\x04' */
    case AnchorBottomBaseline = 0xa00005 /* '\x00\xa0\x00\x05' */
}

// MARK: PowerPointMsoAutoShapeType
@objc public enum PowerPointMsoAutoShapeType : AEKeyword {
    case AutoshapeShapeTypeUnset = 0x69fffe /* '\x00i\xff\xfe' */
    case AutoshapeRectangle = 0x6a0001 /* '\x00j\x00\x01' */
    case AutoshapeParallelogram = 0x6a0002 /* '\x00j\x00\x02' */
    case AutoshapeTrapezoid = 0x6a0003 /* '\x00j\x00\x03' */
    case AutoshapeDiamond = 0x6a0004 /* '\x00j\x00\x04' */
    case AutoshapeRoundedRectangle = 0x6a0005 /* '\x00j\x00\x05' */
    case AutoshapeOctagon = 0x6a0006 /* '\x00j\x00\x06' */
    case AutoshapeIsoscelesTriangle = 0x6a0007 /* '\x00j\x00\x07' */
    case AutoshapeRightTriangle = 0x6a0008 /* '\x00j\x00\x08' */
    case AutoshapeOval = 0x6a0009 /* '\x00j\x00\t' */
    case AutoshapeHexagon = 0x6a000a /* '\x00j\x00\n' */
    case AutoshapeCross = 0x6a000b /* '\x00j\x00\x0b' */
    case AutoshapeRegularPentagon = 0x6a000c /* '\x00j\x00\x0c' */
    case AutoshapeCan = 0x6a000d /* '\x00j\x00\r' */
    case AutoshapeCube = 0x6a000e /* '\x00j\x00\x0e' */
    case AutoshapeBevel = 0x6a000f /* '\x00j\x00\x0f' */
    case AutoshapeFoldedCorner = 0x6a0010 /* '\x00j\x00\x10' */
    case AutoshapeSmileyFace = 0x6a0011 /* '\x00j\x00\x11' */
    case AutoshapeDonut = 0x6a0012 /* '\x00j\x00\x12' */
    case AutoshapeNoSymbol = 0x6a0013 /* '\x00j\x00\x13' */
    case AutoshapeBlockArc = 0x6a0014 /* '\x00j\x00\x14' */
    case AutoshapeHeart = 0x6a0015 /* '\x00j\x00\x15' */
    case AutoshapeLightningBolt = 0x6a0016 /* '\x00j\x00\x16' */
    case AutoshapeSun = 0x6a0017 /* '\x00j\x00\x17' */
    case AutoshapeMoon = 0x6a0018 /* '\x00j\x00\x18' */
    case AutoshapeArc = 0x6a0019 /* '\x00j\x00\x19' */
    case AutoshapeDoubleBracket = 0x6a001a /* '\x00j\x00\x1a' */
    case AutoshapeDoubleBrace = 0x6a001b /* '\x00j\x00\x1b' */
    case AutoshapePlaque = 0x6a001c /* '\x00j\x00\x1c' */
    case AutoshapeLeftBracket = 0x6a001d /* '\x00j\x00\x1d' */
    case AutoshapeRightBracket = 0x6a001e /* '\x00j\x00\x1e' */
    case AutoshapeLeftBrace = 0x6a001f /* '\x00j\x00\x1f' */
    case AutoshapeRightBrace = 0x6a0020 /* '\x00j\x00 ' */
    case AutoshapeRightArrow = 0x6a0021 /* '\x00j\x00!' */
    case AutoshapeLeftArrow = 0x6a0022 /* '\x00j\x00"' */
    case AutoshapeUpArrow = 0x6a0023 /* '\x00j\x00#' */
    case AutoshapeDownArrow = 0x6a0024 /* '\x00j\x00$' */
    case AutoshapeLeftRightArrow = 0x6a0025 /* '\x00j\x00%' */
    case AutoshapeUpDownArrow = 0x6a0026 /* '\x00j\x00&' */
    case AutoshapeQuadArrow = 0x6a0027 /* "\x00j\x00'" */
    case AutoshapeLeftRightUpArrow = 0x6a0028 /* '\x00j\x00(' */
    case AutoshapeBentArrow = 0x6a0029 /* '\x00j\x00)' */
    case AutoshapeUTurnArrow = 0x6a002a /* '\x00j\x00*' */
    case AutoshapeLeftUpArrow = 0x6a002b /* '\x00j\x00+' */
    case AutoshapeBentUpArrow = 0x6a002c /* '\x00j\x00,' */
    case AutoshapeCurvedRightArrow = 0x6a002d /* '\x00j\x00-' */
    case AutoshapeCurvedLeftArrow = 0x6a002e /* '\x00j\x00.' */
    case AutoshapeCurvedUpArrow = 0x6a002f /* '\x00j\x00/' */
    case AutoshapeCurvedDownArrow = 0x6a0030 /* '\x00j\x000' */
    case AutoshapeStripedRightArrow = 0x6a0031 /* '\x00j\x001' */
    case AutoshapeNotchedRightArrow = 0x6a0032 /* '\x00j\x002' */
    case AutoshapePentagon = 0x6a0033 /* '\x00j\x003' */
    case AutoshapeChevron = 0x6a0034 /* '\x00j\x004' */
    case AutoshapeRightArrowCallout = 0x6a0035 /* '\x00j\x005' */
    case AutoshapeLeftArrowCallout = 0x6a0036 /* '\x00j\x006' */
    case AutoshapeUpArrowCallout = 0x6a0037 /* '\x00j\x007' */
    case AutoshapeDownArrowCallout = 0x6a0038 /* '\x00j\x008' */
    case AutoshapeLeftRightArrowCallout = 0x6a0039 /* '\x00j\x009' */
    case AutoshapeUpDownArrowCallout = 0x6a003a /* '\x00j\x00:' */
    case AutoshapeQuadArrowCallout = 0x6a003b /* '\x00j\x00;' */
    case AutoshapeCircularArrow = 0x6a003c /* '\x00j\x00<' */
    case AutoshapeFlowchartProcess = 0x6a003d /* '\x00j\x00=' */
    case AutoshapeFlowchartAlternateProcess = 0x6a003e /* '\x00j\x00>' */
    case AutoshapeFlowchartDecision = 0x6a003f /* '\x00j\x00?' */
    case AutoshapeFlowchartData = 0x6a0040 /* '\x00j\x00@' */
    case AutoshapeFlowchartPredefinedProcess = 0x6a0041 /* '\x00j\x00A' */
    case AutoshapeFlowchartInternalStorage = 0x6a0042 /* '\x00j\x00B' */
    case AutoshapeFlowchartDocument = 0x6a0043 /* '\x00j\x00C' */
    case AutoshapeFlowchartMultiDocument = 0x6a0044 /* '\x00j\x00D' */
    case AutoshapeFlowchartTerminator = 0x6a0045 /* '\x00j\x00E' */
    case AutoshapeFlowchartPreparation = 0x6a0046 /* '\x00j\x00F' */
    case AutoshapeFlowchartManualInput = 0x6a0047 /* '\x00j\x00G' */
    case AutoshapeFlowchartManualOperation = 0x6a0048 /* '\x00j\x00H' */
    case AutoshapeFlowchartConnector = 0x6a0049 /* '\x00j\x00I' */
    case AutoshapeFlowchartOffpageConnector = 0x6a004a /* '\x00j\x00J' */
    case AutoshapeFlowchartCard = 0x6a004b /* '\x00j\x00K' */
    case AutoshapeFlowchartPunchedTape = 0x6a004c /* '\x00j\x00L' */
    case AutoshapeFlowchartSummingJunction = 0x6a004d /* '\x00j\x00M' */
    case AutoshapeFlowchartOr = 0x6a004e /* '\x00j\x00N' */
    case AutoshapeFlowchartCollate = 0x6a004f /* '\x00j\x00O' */
    case AutoshapeFlowchartSort = 0x6a0050 /* '\x00j\x00P' */
    case AutoshapeFlowchartExtract = 0x6a0051 /* '\x00j\x00Q' */
    case AutoshapeFlowchartMerge = 0x6a0052 /* '\x00j\x00R' */
    case AutoshapeFlowchartStoredData = 0x6a0053 /* '\x00j\x00S' */
    case AutoshapeFlowchartDelay = 0x6a0054 /* '\x00j\x00T' */
    case AutoshapeFlowchartSequentialAccessStorage = 0x6a0055 /* '\x00j\x00U' */
    case AutoshapeFlowchartMagneticDisk = 0x6a0056 /* '\x00j\x00V' */
    case AutoshapeFlowchartDirectAccessStorage = 0x6a0057 /* '\x00j\x00W' */
    case AutoshapeFlowchartDisplay = 0x6a0058 /* '\x00j\x00X' */
    case AutoshapeExplosionOne = 0x6a0059 /* '\x00j\x00Y' */
    case AutoshapeExplosionTwo = 0x6a005a /* '\x00j\x00Z' */
    case AutoshapeFourPointStar = 0x6a005b /* '\x00j\x00[' */
    case AutoshapeFivePointStar = 0x6a005c /* '\x00j\x00\\' */
    case AutoshapeEightPointStar = 0x6a005d /* '\x00j\x00]' */
    case AutoshapeSixteenPointStar = 0x6a005e /* '\x00j\x00^' */
    case AutoshapeTwentyFourPointStar = 0x6a005f /* '\x00j\x00_' */
    case AutoshapeThirtyTwoPointStar = 0x6a0060 /* '\x00j\x00`' */
    case AutoshapeUpRibbon = 0x6a0061 /* '\x00j\x00a' */
    case AutoshapeDownRibbon = 0x6a0062 /* '\x00j\x00b' */
    case AutoshapeCurvedUpRibbon = 0x6a0063 /* '\x00j\x00c' */
    case AutoshapeCurvedDownRibbon = 0x6a0064 /* '\x00j\x00d' */
    case AutoshapeVerticalScroll = 0x6a0065 /* '\x00j\x00e' */
    case AutoshapeHorizontalScroll = 0x6a0066 /* '\x00j\x00f' */
    case AutoshapeWave = 0x6a0067 /* '\x00j\x00g' */
    case AutoshapeDoubleWave = 0x6a0068 /* '\x00j\x00h' */
    case AutoshapeRectangularCallout = 0x6a0069 /* '\x00j\x00i' */
    case AutoshapeRoundedRectangularCallout = 0x6a006a /* '\x00j\x00j' */
    case AutoshapeOvalCallout = 0x6a006b /* '\x00j\x00k' */
    case AutoshapeCloudCallout = 0x6a006c /* '\x00j\x00l' */
    case AutoshapeLineCalloutOne = 0x6a006d /* '\x00j\x00m' */
    case AutoshapeLineCalloutTwo = 0x6a006e /* '\x00j\x00n' */
    case AutoshapeLineCalloutThree = 0x6a006f /* '\x00j\x00o' */
    case AutoshapeLineCalloutFour = 0x6a0070 /* '\x00j\x00p' */
    case AutoshapeLineCalloutOneAccentBar = 0x6a0071 /* '\x00j\x00q' */
    case AutoshapeLineCalloutTwoAccentBar = 0x6a0072 /* '\x00j\x00r' */
    case AutoshapeLineCalloutThreeAccentBar = 0x6a0073 /* '\x00j\x00s' */
    case AutoshapeLineCalloutFourAccentBar = 0x6a0074 /* '\x00j\x00t' */
    case AutoshapeLineCalloutOneNoBorder = 0x6a0075 /* '\x00j\x00u' */
    case AutoshapeLineCalloutTwoNoBorder = 0x6a0076 /* '\x00j\x00v' */
    case AutoshapeLineCalloutThreeNoBorder = 0x6a0077 /* '\x00j\x00w' */
    case AutoshapeLineCalloutFourNoBorder = 0x6a0078 /* '\x00j\x00x' */
    case AutoshapeCalloutOneBorderAndAccentBar = 0x6a0079 /* '\x00j\x00y' */
    case AutoshapeCalloutTwoBorderAndAccentBar = 0x6a007a /* '\x00j\x00z' */
    case AutoshapeCalloutThreeBorderAndAccentBar = 0x6a007b /* '\x00j\x00{' */
    case AutoshapeCalloutFourBorderAndAccentBar = 0x6a007c /* '\x00j\x00|' */
    case AutoshapeActionButtonCustom = 0x6a007d /* '\x00j\x00}' */
    case AutoshapeActionButtonHome = 0x6a007e /* '\x00j\x00~' */
    case AutoshapeActionButtonHelp = 0x6a007f /* '\x00j\x00\x7f' */
    case AutoshapeActionButtonInformation = 0x6a0080 /* '\x00j\x00\x80' */
    case AutoshapeActionButtonBackOrPrevious = 0x6a0081 /* '\x00j\x00\x81' */
    case AutoshapeActionButtonForwardOrNext = 0x6a0082 /* '\x00j\x00\x82' */
    case AutoshapeActionButtonBeginning = 0x6a0083 /* '\x00j\x00\x83' */
    case AutoshapeActionButtonEnd = 0x6a0084 /* '\x00j\x00\x84' */
    case AutoshapeActionButtonReturn = 0x6a0085 /* '\x00j\x00\x85' */
    case AutoshapeActionButtonDocument = 0x6a0086 /* '\x00j\x00\x86' */
    case AutoshapeActionButtonSound = 0x6a0087 /* '\x00j\x00\x87' */
    case AutoshapeActionButtonMovie = 0x6a0088 /* '\x00j\x00\x88' */
    case AutoshapeBalloon = 0x6a0089 /* '\x00j\x00\x89' */
    case AutoshapeNotPrimitive = 0x6a008a /* '\x00j\x00\x8a' */
    case AutoshapeFlowchartOfflineStorage = 0x6a008b /* '\x00j\x00\x8b' */
    case AutoshapeLeftRightRibbon = 0x6a008c /* '\x00j\x00\x8c' */
    case AutoshapeDiagonalStripe = 0x6a008d /* '\x00j\x00\x8d' */
    case AutoshapePie = 0x6a008e /* '\x00j\x00\x8e' */
    case AutoshapeNonIsoscelesTrapezoid = 0x6a008f /* '\x00j\x00\x8f' */
    case AutoshapeDecagon = 0x6a0090 /* '\x00j\x00\x90' */
    case AutoshapeHeptagon = 0x6a0091 /* '\x00j\x00\x91' */
    case AutoshapeDodecagon = 0x6a0092 /* '\x00j\x00\x92' */
    case AutoshapeSixPointsStar = 0x6a0093 /* '\x00j\x00\x93' */
    case AutoshapeSevenPointsStar = 0x6a0094 /* '\x00j\x00\x94' */
    case AutoshapeTenPointsStar = 0x6a0095 /* '\x00j\x00\x95' */
    case AutoshapeTwelvePointsStar = 0x6a0096 /* '\x00j\x00\x96' */
    case AutoshapeRoundOneRectangle = 0x6a0097 /* '\x00j\x00\x97' */
    case AutoshapeRoundTwoSameRectangle = 0x6a0098 /* '\x00j\x00\x98' */
    case AutoshapeRoundTwoDiagonalRectangle = 0x6a0099 /* '\x00j\x00\x99' */
    case AutoshapeSnipRoundRectangle = 0x6a009a /* '\x00j\x00\x9a' */
    case AutoshapeSnipOneRectangle = 0x6a009b /* '\x00j\x00\x9b' */
    case AutoshapeSnipTwoSameRectangle = 0x6a009c /* '\x00j\x00\x9c' */
    case AutoshapeSnipTwoDiagonalRectangle = 0x6a009d /* '\x00j\x00\x9d' */
    case AutoshapeFrame = 0x6a009e /* '\x00j\x00\x9e' */
    case AutoshapeHalfFrame = 0x6a009f /* '\x00j\x00\x9f' */
    case AutoshapeTear = 0x6a00a0 /* '\x00j\x00\xa0' */
    case AutoshapeChord = 0x6a00a1 /* '\x00j\x00\xa1' */
    case AutoshapeCorner = 0x6a00a2 /* '\x00j\x00\xa2' */
    case AutoshapeMathPlus = 0x6a00a3 /* '\x00j\x00\xa3' */
    case AutoshapeMathMinus = 0x6a00a4 /* '\x00j\x00\xa4' */
    case AutoshapeMathMultiply = 0x6a00a5 /* '\x00j\x00\xa5' */
    case AutoshapeMathDivide = 0x6a00a6 /* '\x00j\x00\xa6' */
    case AutoshapeMathEqual = 0x6a00a7 /* '\x00j\x00\xa7' */
    case AutoshapeMathNotEqual = 0x6a00a8 /* '\x00j\x00\xa8' */
    case AutoshapeCornerTabs = 0x6a00a9 /* '\x00j\x00\xa9' */
    case AutoshapeSquareTabs = 0x6a00aa /* '\x00j\x00\xaa' */
    case AutoshapePlaqueTabs = 0x6a00ab /* '\x00j\x00\xab' */
    case AutoshapeGearSix = 0x6a00ac /* '\x00j\x00\xac' */
    case AutoshapeGearNine = 0x6a00ad /* '\x00j\x00\xad' */
    case AutoshapeFunnel = 0x6a00ae /* '\x00j\x00\xae' */
    case AutoshapePieWedge = 0x6a00af /* '\x00j\x00\xaf' */
    case AutoshapeLeftCircularArrow = 0x6a00b0 /* '\x00j\x00\xb0' */
    case AutoshapeLeftRightCircularArrow = 0x6a00b1 /* '\x00j\x00\xb1' */
    case AutoshapeSwooshArrow = 0x6a00b2 /* '\x00j\x00\xb2' */
    case AutoshapeCloud = 0x6a00b3 /* '\x00j\x00\xb3' */
    case AutoshapeChartX = 0x6a00b4 /* '\x00j\x00\xb4' */
    case AutoshapeChartStar = 0x6a00b5 /* '\x00j\x00\xb5' */
    case AutoshapeChartPlus = 0x6a00b6 /* '\x00j\x00\xb6' */
    case AutoshapeLineInverse = 0x6a00b7 /* '\x00j\x00\xb7' */
}

// MARK: PowerPointMsoShapeType
@objc public enum PowerPointMsoShapeType : AEKeyword {
    case ShapeTypeUnset = 0x8bfffe /* '\x00\x8b\xff\xfe' */
    case ShapeTypeAuto = 0x8c0001 /* '\x00\x8c\x00\x01' */
    case ShapeTypeCallout = 0x8c0002 /* '\x00\x8c\x00\x02' */
    case ShapeTypeChart = 0x8c0003 /* '\x00\x8c\x00\x03' */
    case ShapeTypeComment = 0x8c0004 /* '\x00\x8c\x00\x04' */
    case ShapeTypeFreeForm = 0x8c0005 /* '\x00\x8c\x00\x05' */
    case ShapeTypeGroup = 0x8c0006 /* '\x00\x8c\x00\x06' */
    case ShapeTypeEmbeddedOLEControl = 0x8c0007 /* '\x00\x8c\x00\x07' */
    case ShapeTypeFormControl = 0x8c0008 /* '\x00\x8c\x00\x08' */
    case ShapeTypeLine = 0x8c0009 /* '\x00\x8c\x00\t' */
    case ShapeTypeLinkedOLEObject = 0x8c000a /* '\x00\x8c\x00\n' */
    case ShapeTypeLinkedPicture = 0x8c000b /* '\x00\x8c\x00\x0b' */
    case ShapeTypeOLEControl = 0x8c000c /* '\x00\x8c\x00\x0c' */
    case ShapeTypePicture = 0x8c000d /* '\x00\x8c\x00\r' */
    case ShapeTypePlaceHolder = 0x8c000e /* '\x00\x8c\x00\x0e' */
    case ShapeTypeWordArt = 0x8c000f /* '\x00\x8c\x00\x0f' */
    case ShapeTypeMedia = 0x8c0010 /* '\x00\x8c\x00\x10' */
    case ShapeTypeTextBox = 0x8c0011 /* '\x00\x8c\x00\x11' */
    case ShapeTypeTable = 0x8c0012 /* '\x00\x8c\x00\x12' */
    case ShapeTypeCanvas = 0x8c0013 /* '\x00\x8c\x00\x13' */
    case ShapeTypeDiagram = 0x8c0014 /* '\x00\x8c\x00\x14' */
    case ShapeTypeInk = 0x8c0015 /* '\x00\x8c\x00\x15' */
    case ShapeTypeInkComment = 0x8c0016 /* '\x00\x8c\x00\x16' */
    case ShapeTypeSmartartGraphic = 0x8c0017 /* '\x00\x8c\x00\x17' */
    case ShapeTypeSlicer = 0x8c0018 /* '\x00\x8c\x00\x18' */
    case ShapeTypeWebVideo = 0x8c0019 /* '\x00\x8c\x00\x19' */
    case ShapeTypeContentApplication = 0x8c001a /* '\x00\x8c\x00\x1a' */
}

// MARK: PowerPointMsoColorType
@objc public enum PowerPointMsoColorType : AEKeyword {
    case ColorTypeUnset = 0x6afffe /* '\x00j\xff\xfe' */
    case RGB = 0x6b0001 /* '\x00k\x00\x01' */
    case Scheme = 0x6b0002 /* '\x00k\x00\x02' */
}

// MARK: PowerPointMsoPictureColorType
@objc public enum PowerPointMsoPictureColorType : AEKeyword {
    case PictureColorTypeUnset = 0xb5fffe /* '\x00\xb5\xff\xfe' */
    case PictureColorAutomatic = 0xb60001 /* '\x00\xb6\x00\x01' */
    case PictureColorGrayScale = 0xb60002 /* '\x00\xb6\x00\x02' */
    case PictureColorBlackAndWhite = 0xb60003 /* '\x00\xb6\x00\x03' */
    case PictureColorWatermark = 0xb60004 /* '\x00\xb6\x00\x04' */
}

// MARK: PowerPointMsoCalloutAngleType
@objc public enum PowerPointMsoCalloutAngleType : AEKeyword {
    case AngleUnset = 0x6bfffe /* '\x00k\xff\xfe' */
    case AngleAutomatic = 0x6c0001 /* '\x00l\x00\x01' */
    case Angle30 = 0x6c0002 /* '\x00l\x00\x02' */
    case Angle45 = 0x6c0003 /* '\x00l\x00\x03' */
    case Angle60 = 0x6c0004 /* '\x00l\x00\x04' */
    case Angle90 = 0x6c0005 /* '\x00l\x00\x05' */
}

// MARK: PowerPointMsoCalloutDropType
@objc public enum PowerPointMsoCalloutDropType : AEKeyword {
    case DropUnset = 0x6cfffe /* '\x00l\xff\xfe' */
    case DropCustom = 0x6d0001 /* '\x00m\x00\x01' */
    case DropTop = 0x6d0002 /* '\x00m\x00\x02' */
    case DropCenter = 0x6d0003 /* '\x00m\x00\x03' */
    case DropBottom = 0x6d0004 /* '\x00m\x00\x04' */
}

// MARK: PowerPointMsoCalloutType
@objc public enum PowerPointMsoCalloutType : AEKeyword {
    case CalloutUnset = 0x6dfffe /* '\x00m\xff\xfe' */
    case CalloutOne = 0x6e0001 /* '\x00n\x00\x01' */
    case CalloutTwo = 0x6e0002 /* '\x00n\x00\x02' */
    case CalloutThree = 0x6e0003 /* '\x00n\x00\x03' */
    case CalloutFour = 0x6e0004 /* '\x00n\x00\x04' */
}

// MARK: PowerPointMsoTextOrientation
@objc public enum PowerPointMsoTextOrientation : AEKeyword {
    case TextOrientationUnset = 0x8dfffe /* '\x00\x8d\xff\xfe' */
    case Horizontal = 0x8e0001 /* '\x00\x8e\x00\x01' */
    case Upward = 0x8e0002 /* '\x00\x8e\x00\x02' */
    case Downward = 0x8e0003 /* '\x00\x8e\x00\x03' */
    case VerticalEastAsian = 0x8e0004 /* '\x00\x8e\x00\x04' */
    case Vertical = 0x8e0005 /* '\x00\x8e\x00\x05' */
    case HorizontalRotatedEastAsian = 0x8e0006 /* '\x00\x8e\x00\x06' */
}

// MARK: PowerPointMsoScaleFrom
@objc public enum PowerPointMsoScaleFrom : AEKeyword {
    case ScaleFromTopLeft = 0x6f0000 /* '\x00o\x00\x00' */
    case ScaleFromMiddle = 0x6f0001 /* '\x00o\x00\x01' */
    case ScaleFromBottomRight = 0x6f0002 /* '\x00o\x00\x02' */
}

// MARK: PowerPointMsoPresetCamera
@objc public enum PowerPointMsoPresetCamera : AEKeyword {
    case PresetCameraUnset = 0xaefffe /* '\x00\xae\xff\xfe' */
    case CameraLegacyObliqueFromTopLeft = 0xaf0001 /* '\x00\xaf\x00\x01' */
    case CameraLegacyObliqueFromTop = 0xaf0002 /* '\x00\xaf\x00\x02' */
    case CameraLegacyObliqueFromTopright = 0xaf0003 /* '\x00\xaf\x00\x03' */
    case CameraLegacyObliqueFromLeft = 0xaf0004 /* '\x00\xaf\x00\x04' */
    case CameraLegacyObliqueFromFront = 0xaf0005 /* '\x00\xaf\x00\x05' */
    case CameraLegacyObliqueFromRight = 0xaf0006 /* '\x00\xaf\x00\x06' */
    case CameraLegacyObliqueFromBottomLeft = 0xaf0007 /* '\x00\xaf\x00\x07' */
    case CameraLegacyObliqueFromBottom = 0xaf0008 /* '\x00\xaf\x00\x08' */
    case CameraLegacyObliqueFromBottomRight = 0xaf0009 /* '\x00\xaf\x00\t' */
    case CameraLegacyPerspectiveFromTopLeft = 0xaf000a /* '\x00\xaf\x00\n' */
    case CameraLegacyPerspectiveFromTop = 0xaf000b /* '\x00\xaf\x00\x0b' */
    case CameraLegacyPerspectiveFromTopRight = 0xaf000c /* '\x00\xaf\x00\x0c' */
    case CameraLegacyPerspectiveFromLeft = 0xaf000d /* '\x00\xaf\x00\r' */
    case CameraLegacyPerspectiveFromFront = 0xaf000e /* '\x00\xaf\x00\x0e' */
    case CameraLegacyPerspectiveFromRight = 0xaf000f /* '\x00\xaf\x00\x0f' */
    case CameraLegacyPerspectiveFromBottomLeft = 0xaf0010 /* '\x00\xaf\x00\x10' */
    case CameraLegacyPerspectiveFromBottom = 0xaf0011 /* '\x00\xaf\x00\x11' */
    case CameraLegacyPerspectiveFromBottomRight = 0xaf0012 /* '\x00\xaf\x00\x12' */
    case CameraOrthographic = 0xaf0013 /* '\x00\xaf\x00\x13' */
    case CameraIsometricFromTopUp = 0xaf0014 /* '\x00\xaf\x00\x14' */
    case CameraIsometricFromTopDown = 0xaf0015 /* '\x00\xaf\x00\x15' */
    case CameraIsometricFromBottomUp = 0xaf0016 /* '\x00\xaf\x00\x16' */
    case CameraIsometricFromBottomDown = 0xaf0017 /* '\x00\xaf\x00\x17' */
    case CameraIsometricFromLeftUp = 0xaf0018 /* '\x00\xaf\x00\x18' */
    case CameraIsometricFromLeftDown = 0xaf0019 /* '\x00\xaf\x00\x19' */
    case CameraIsometricFromRightUp = 0xaf001a /* '\x00\xaf\x00\x1a' */
    case CameraIsometricFromRightDown = 0xaf001b /* '\x00\xaf\x00\x1b' */
    case CameraIsometricOffAxis1FromLeft = 0xaf001c /* '\x00\xaf\x00\x1c' */
    case CameraIsometricOffAxis1FromRight = 0xaf001d /* '\x00\xaf\x00\x1d' */
    case CameraIsometricOffAxis1FromTop = 0xaf001e /* '\x00\xaf\x00\x1e' */
    case CameraIsometricOffAxis2FromLeft = 0xaf001f /* '\x00\xaf\x00\x1f' */
    case CameraIsometricOffAxis2FromRight = 0xaf0020 /* '\x00\xaf\x00 ' */
    case CameraIsometricOffAxis2FromTop = 0xaf0021 /* '\x00\xaf\x00!' */
    case CameraIsometricOffAxis3FromLeft = 0xaf0022 /* '\x00\xaf\x00"' */
    case CameraIsometricOffAxis3FromRight = 0xaf0023 /* '\x00\xaf\x00#' */
    case CameraIsometricOffAxis3FromBottom = 0xaf0024 /* '\x00\xaf\x00$' */
    case CameraIsometricOffAxis4FromLeft = 0xaf0025 /* '\x00\xaf\x00%' */
    case CameraIsometricOffAxis4FromRight = 0xaf0026 /* '\x00\xaf\x00&' */
    case CameraIsometricOffAxis4FromBottom = 0xaf0027 /* "\x00\xaf\x00'" */
    case CameraObliqueFromTopLeft = 0xaf0028 /* '\x00\xaf\x00(' */
    case CameraObliqueFromTop = 0xaf0029 /* '\x00\xaf\x00)' */
    case CameraObliqueFromTopRight = 0xaf002a /* '\x00\xaf\x00*' */
    case CameraObliqueFromLeft = 0xaf002b /* '\x00\xaf\x00+' */
    case CameraObliqueFromRight = 0xaf002c /* '\x00\xaf\x00,' */
    case CameraObliqueFromBottomLeft = 0xaf002d /* '\x00\xaf\x00-' */
    case CameraObliqueFromBottom = 0xaf002e /* '\x00\xaf\x00.' */
    case CameraObliqueFromBottomRight = 0xaf002f /* '\x00\xaf\x00/' */
    case CameraPerspectiveFromFront = 0xaf0030 /* '\x00\xaf\x000' */
    case CameraPerspectiveFromLeft = 0xaf0031 /* '\x00\xaf\x001' */
    case CameraPerspectiveFromRight = 0xaf0032 /* '\x00\xaf\x002' */
    case CameraPerspectiveFromAbove = 0xaf0033 /* '\x00\xaf\x003' */
    case CameraPerspectiveFromBelow = 0xaf0034 /* '\x00\xaf\x004' */
    case CameraPerspectiveFromAboveFacingLeft = 0xaf0035 /* '\x00\xaf\x005' */
    case CameraPerspectiveFromAboveFacingRight = 0xaf0036 /* '\x00\xaf\x006' */
    case CameraPerspectiveContrastingFacingLeft = 0xaf0037 /* '\x00\xaf\x007' */
    case CameraPerspectiveContrastingFacingRight = 0xaf0038 /* '\x00\xaf\x008' */
    case CameraPerspectiveHeroicFacingLeft = 0xaf0039 /* '\x00\xaf\x009' */
    case CameraPerspectiveHeroicFacingRight = 0xaf003a /* '\x00\xaf\x00:' */
    case CameraPerspectiveHeroicExtremeFacingLeft = 0xaf003b /* '\x00\xaf\x00;' */
    case CameraPerspectiveHeroicExtremeFacingRight = 0xaf003c /* '\x00\xaf\x00<' */
    case CameraPerspectiveRelaxed = 0xaf003d /* '\x00\xaf\x00=' */
    case CameraPerspectiveRelaxedModerately = 0xaf003e /* '\x00\xaf\x00>' */
}

// MARK: PowerPointMsoLightRigType
@objc public enum PowerPointMsoLightRigType : AEKeyword {
    case LightRigUnset = 0xaffffe /* '\x00\xaf\xff\xfe' */
    case LightRigFlat1 = 0xb00001 /* '\x00\xb0\x00\x01' */
    case LightRigFlat2 = 0xb00002 /* '\x00\xb0\x00\x02' */
    case LightRigFlat3 = 0xb00003 /* '\x00\xb0\x00\x03' */
    case LightRigFlat4 = 0xb00004 /* '\x00\xb0\x00\x04' */
    case LightRigNormal1 = 0xb00005 /* '\x00\xb0\x00\x05' */
    case LightRigNormal2 = 0xb00006 /* '\x00\xb0\x00\x06' */
    case LightRigNormal3 = 0xb00007 /* '\x00\xb0\x00\x07' */
    case LightRigNormal4 = 0xb00008 /* '\x00\xb0\x00\x08' */
    case LightRigHarsh1 = 0xb00009 /* '\x00\xb0\x00\t' */
    case LightRigHarsh2 = 0xb0000a /* '\x00\xb0\x00\n' */
    case LightRigHarsh3 = 0xb0000b /* '\x00\xb0\x00\x0b' */
    case LightRigHarsh4 = 0xb0000c /* '\x00\xb0\x00\x0c' */
    case LightRigThreePoint = 0xb0000d /* '\x00\xb0\x00\r' */
    case LightRigBalanced = 0xb0000e /* '\x00\xb0\x00\x0e' */
    case LightRigSoft = 0xb0000f /* '\x00\xb0\x00\x0f' */
    case LightRigHarsh = 0xb00010 /* '\x00\xb0\x00\x10' */
    case LightRigFlood = 0xb00011 /* '\x00\xb0\x00\x11' */
    case LightRigContrasting = 0xb00012 /* '\x00\xb0\x00\x12' */
    case LightRigMorning = 0xb00013 /* '\x00\xb0\x00\x13' */
    case LightRigSunrise = 0xb00014 /* '\x00\xb0\x00\x14' */
    case LightRigSunset = 0xb00015 /* '\x00\xb0\x00\x15' */
    case LightRigChilly = 0xb00016 /* '\x00\xb0\x00\x16' */
    case LightRigFreezing = 0xb00017 /* '\x00\xb0\x00\x17' */
    case LightRigFlat = 0xb00018 /* '\x00\xb0\x00\x18' */
    case LightRigTwoPoint = 0xb00019 /* '\x00\xb0\x00\x19' */
    case LightRigGlow = 0xb0001a /* '\x00\xb0\x00\x1a' */
    case LightRigBrightRoom = 0xb0001b /* '\x00\xb0\x00\x1b' */
}

// MARK: PowerPointMsoBevelType
@objc public enum PowerPointMsoBevelType : AEKeyword {
    case BevelTypeUnset = 0xb0fffe /* '\x00\xb0\xff\xfe' */
    case BevelNone = 0xb10001 /* '\x00\xb1\x00\x01' */
    case BevelRelaxedInset = 0xb10002 /* '\x00\xb1\x00\x02' */
    case BevelCircle = 0xb10003 /* '\x00\xb1\x00\x03' */
    case BevelSlope = 0xb10004 /* '\x00\xb1\x00\x04' */
    case BevelCross = 0xb10005 /* '\x00\xb1\x00\x05' */
    case BevelAngle = 0xb10006 /* '\x00\xb1\x00\x06' */
    case BevelSoftRound = 0xb10007 /* '\x00\xb1\x00\x07' */
    case BevelConvex = 0xb10008 /* '\x00\xb1\x00\x08' */
    case BevelCoolSlant = 0xb10009 /* '\x00\xb1\x00\t' */
    case BevelDivot = 0xb1000a /* '\x00\xb1\x00\n' */
    case BevelRiblet = 0xb1000b /* '\x00\xb1\x00\x0b' */
    case BevelHardEdge = 0xb1000c /* '\x00\xb1\x00\x0c' */
    case BevelArtDeco = 0xb1000d /* '\x00\xb1\x00\r' */
}

// MARK: PowerPointMsoShadowStyle
@objc public enum PowerPointMsoShadowStyle : AEKeyword {
    case ShadowStyleUnset = 0xb1fffe /* '\x00\xb1\xff\xfe' */
    case ShadowStyleInner = 0xb20001 /* '\x00\xb2\x00\x01' */
    case ShadowStyleOuter = 0xb20002 /* '\x00\xb2\x00\x02' */
}

// MARK: PowerPointMsoParagraphAlignment
@objc public enum PowerPointMsoParagraphAlignment : AEKeyword {
    case ParagraphAlignmentUnset = 0xe6fffe /* '\x00\xe6\xff\xfe' */
    case ParagraphAlignLeft = 0xe70000 /* '\x00\xe7\x00\x00' */
    case ParagraphAlignCenter = 0xe70001 /* '\x00\xe7\x00\x01' */
    case ParagraphAlignRight = 0xe70002 /* '\x00\xe7\x00\x02' */
    case ParagraphAlignJustify = 0xe70003 /* '\x00\xe7\x00\x03' */
    case ParagraphAlignDistribute = 0xe70004 /* '\x00\xe7\x00\x04' */
    case ParagraphAlignThai = 0xe70005 /* '\x00\xe7\x00\x05' */
    case ParagraphAlignJustifyLow = 0xe70006 /* '\x00\xe7\x00\x06' */
}

// MARK: PowerPointMsoTextStrike
@objc public enum PowerPointMsoTextStrike : AEKeyword {
    case StrikeUnset = 0xb3fffe /* '\x00\xb3\xff\xfe' */
    case NoStrike = 0xb40000 /* '\x00\xb4\x00\x00' */
    case SingleStrike = 0xb40001 /* '\x00\xb4\x00\x01' */
    case DoubleStrike = 0xb40002 /* '\x00\xb4\x00\x02' */
}

// MARK: PowerPointMsoTextCaps
@objc public enum PowerPointMsoTextCaps : AEKeyword {
    case CapsUnset = 0xb4fffe /* '\x00\xb4\xff\xfe' */
    case NoCaps = 0xb50000 /* '\x00\xb5\x00\x00' */
    case SmallCaps = 0xb50001 /* '\x00\xb5\x00\x01' */
    case AllCaps = 0xb50002 /* '\x00\xb5\x00\x02' */
}

// MARK: PowerPointMsoTextUnderlineType
@objc public enum PowerPointMsoTextUnderlineType : AEKeyword {
    case UnderlineUnset = 0x3eefffe /* '\x03\xee\xff\xfe' */
    case NoUnderline = 0x3ef0000 /* '\x03\xef\x00\x00' */
    case UnderlineWordsOnly = 0x3ef0001 /* '\x03\xef\x00\x01' */
    case UnderlineSingleLine = 0x3ef0002 /* '\x03\xef\x00\x02' */
    case UnderlineDoubleLine = 0x3ef0003 /* '\x03\xef\x00\x03' */
    case UnderlineHeavyLine = 0x3ef0004 /* '\x03\xef\x00\x04' */
    case UnderlineDottedLine = 0x3ef0005 /* '\x03\xef\x00\x05' */
    case UnderlineHeavyDottedLine = 0x3ef0006 /* '\x03\xef\x00\x06' */
    case UnderlineDashLine = 0x3ef0007 /* '\x03\xef\x00\x07' */
    case UnderlineHeavyDashLine = 0x3ef0008 /* '\x03\xef\x00\x08' */
    case UnderlineLongDashLine = 0x3ef0009 /* '\x03\xef\x00\t' */
    case UnderlineHeavyLongDashLine = 0x3ef000a /* '\x03\xef\x00\n' */
    case UnderlineDotDashLine = 0x3ef000b /* '\x03\xef\x00\x0b' */
    case UnderlineHeavyDotDashLine = 0x3ef000c /* '\x03\xef\x00\x0c' */
    case UnderlineDotDotDashLine = 0x3ef000d /* '\x03\xef\x00\r' */
    case UnderlineHeavyDotDotDashLine = 0x3ef000e /* '\x03\xef\x00\x0e' */
    case UnderlineWavyLine = 0x3ef000f /* '\x03\xef\x00\x0f' */
    case UnderlineHeavyWavyLine = 0x3ef0010 /* '\x03\xef\x00\x10' */
    case UnderlineWavyDoubleLine = 0x3ef0011 /* '\x03\xef\x00\x11' */
}

// MARK: PowerPointMsoTextTabAlign
@objc public enum PowerPointMsoTextTabAlign : AEKeyword {
    case TabUnset = 0xb6fffe /* '\x00\xb6\xff\xfe' */
    case LeftTab = 0xb70000 /* '\x00\xb7\x00\x00' */
    case CenterTab = 0xb70001 /* '\x00\xb7\x00\x01' */
    case RightTab = 0xb70002 /* '\x00\xb7\x00\x02' */
    case DecimalTab = 0xb70003 /* '\x00\xb7\x00\x03' */
}

// MARK: PowerPointMsoTextCharWrap
@objc public enum PowerPointMsoTextCharWrap : AEKeyword {
    case CharacterWrapUnset = 0xb7fffe /* '\x00\xb7\xff\xfe' */
    case NoCharacterWrap = 0xb80000 /* '\x00\xb8\x00\x00' */
    case StandardCharacterWrap = 0xb80001 /* '\x00\xb8\x00\x01' */
    case StrictCharacterWrap = 0xb80002 /* '\x00\xb8\x00\x02' */
    case CustomCharacterWrap = 0xb80003 /* '\x00\xb8\x00\x03' */
}

// MARK: PowerPointMsoTextFontAlign
@objc public enum PowerPointMsoTextFontAlign : AEKeyword {
    case FontAlignmentUnset = 0xb8fffe /* '\x00\xb8\xff\xfe' */
    case AutomaticAlignment = 0xb90000 /* '\x00\xb9\x00\x00' */
    case TopAlignment = 0xb90001 /* '\x00\xb9\x00\x01' */
    case CenterAlignment = 0xb90002 /* '\x00\xb9\x00\x02' */
    case BaselineAlignment = 0xb90003 /* '\x00\xb9\x00\x03' */
    case BottomAlignment = 0xb90004 /* '\x00\xb9\x00\x04' */
}

// MARK: PowerPointMsoAutoSize
@objc public enum PowerPointMsoAutoSize : AEKeyword {
    case AutoSizeUnset = 0xe4fffe /* '\x00\xe4\xff\xfe' */
    case AutoSizeNone = 0xe50000 /* '\x00\xe5\x00\x00' */
    case ShapeToFitText = 0xe50001 /* '\x00\xe5\x00\x01' */
    case TextToFitShape = 0xe50002 /* '\x00\xe5\x00\x02' */
}

// MARK: PowerPointMsoPathFormat
@objc public enum PowerPointMsoPathFormat : AEKeyword {
    case PathTypeUnset = 0xbafffe /* '\x00\xba\xff\xfe' */
    case NoPathType = 0xbb0000 /* '\x00\xbb\x00\x00' */
    case PathType1 = 0xbb0001 /* '\x00\xbb\x00\x01' */
    case PathType2 = 0xbb0002 /* '\x00\xbb\x00\x02' */
    case PathType3 = 0xbb0003 /* '\x00\xbb\x00\x03' */
    case PathType4 = 0xbb0004 /* '\x00\xbb\x00\x04' */
}

// MARK: PowerPointMsoWarpFormat
@objc public enum PowerPointMsoWarpFormat : AEKeyword {
    case WarpFormatUnset = 0xbbfffe /* '\x00\xbb\xff\xfe' */
    case WarpFormat1 = 0xbc0000 /* '\x00\xbc\x00\x00' */
    case WarpFormat2 = 0xbc0001 /* '\x00\xbc\x00\x01' */
    case WarpFormat3 = 0xbc0002 /* '\x00\xbc\x00\x02' */
    case WarpFormat4 = 0xbc0003 /* '\x00\xbc\x00\x03' */
    case WarpFormat5 = 0xbc0004 /* '\x00\xbc\x00\x04' */
    case WarpFormat6 = 0xbc0005 /* '\x00\xbc\x00\x05' */
    case WarpFormat7 = 0xbc0006 /* '\x00\xbc\x00\x06' */
    case WarpFormat8 = 0xbc0007 /* '\x00\xbc\x00\x07' */
    case WarpFormat9 = 0xbc0008 /* '\x00\xbc\x00\x08' */
    case WarpFormat10 = 0xbc0009 /* '\x00\xbc\x00\t' */
    case WarpFormat11 = 0xbc000a /* '\x00\xbc\x00\n' */
    case WarpFormat12 = 0xbc000b /* '\x00\xbc\x00\x0b' */
    case WarpFormat13 = 0xbc000c /* '\x00\xbc\x00\x0c' */
    case WarpFormat14 = 0xbc000d /* '\x00\xbc\x00\r' */
    case WarpFormat15 = 0xbc000e /* '\x00\xbc\x00\x0e' */
    case WarpFormat16 = 0xbc000f /* '\x00\xbc\x00\x0f' */
    case WarpFormat17 = 0xbc0010 /* '\x00\xbc\x00\x10' */
    case WarpFormat18 = 0xbc0011 /* '\x00\xbc\x00\x11' */
    case WarpFormat19 = 0xbc0012 /* '\x00\xbc\x00\x12' */
    case WarpFormat20 = 0xbc0013 /* '\x00\xbc\x00\x13' */
    case WarpFormat21 = 0xbc0014 /* '\x00\xbc\x00\x14' */
    case WarpFormat22 = 0xbc0015 /* '\x00\xbc\x00\x15' */
    case WarpFormat23 = 0xbc0016 /* '\x00\xbc\x00\x16' */
    case WarpFormat24 = 0xbc0017 /* '\x00\xbc\x00\x17' */
    case WarpFormat25 = 0xbc0018 /* '\x00\xbc\x00\x18' */
    case WarpFormat26 = 0xbc0019 /* '\x00\xbc\x00\x19' */
    case WarpFormat27 = 0xbc001a /* '\x00\xbc\x00\x1a' */
    case WarpFormat28 = 0xbc001b /* '\x00\xbc\x00\x1b' */
    case WarpFormat29 = 0xbc001c /* '\x00\xbc\x00\x1c' */
    case WarpFormat30 = 0xbc001d /* '\x00\xbc\x00\x1d' */
    case WarpFormat31 = 0xbc001e /* '\x00\xbc\x00\x1e' */
    case WarpFormat32 = 0xbc001f /* '\x00\xbc\x00\x1f' */
    case WarpFormat33 = 0xbc0020 /* '\x00\xbc\x00 ' */
    case WarpFormat34 = 0xbc0021 /* '\x00\xbc\x00!' */
    case WarpFormat35 = 0xbc0022 /* '\x00\xbc\x00"' */
    case WarpFormat36 = 0xbc0023 /* '\x00\xbc\x00#' */
    case WarpFormat37 = 0xbc0024 /* '\x00\xbc\x00$' */
}

// MARK: PowerPointMsoTextChangeCase
@objc public enum PowerPointMsoTextChangeCase : AEKeyword {
    case CaseSentence = 0xe40001 /* '\x00\xe4\x00\x01' */
    case CaseLower = 0xe40002 /* '\x00\xe4\x00\x02' */
    case CaseUpper = 0xe40003 /* '\x00\xe4\x00\x03' */
    case CaseTitle = 0xe40004 /* '\x00\xe4\x00\x04' */
    case CaseToggle = 0xe40005 /* '\x00\xe4\x00\x05' */
}

// MARK: PowerPointMsoDateTimeFormat
@objc public enum PowerPointMsoDateTimeFormat : AEKeyword {
    case DateTimeFormatUnset = 0xbdfffe /* '\x00\xbd\xff\xfe' */
    case DateTimeFormatMdyy = 0xbe0001 /* '\x00\xbe\x00\x01' */
    case DateTimeFormatDdddMMMMddyyyy = 0xbe0002 /* '\x00\xbe\x00\x02' */
    case DateTimeFormatDMMMMyyyy = 0xbe0003 /* '\x00\xbe\x00\x03' */
    case DateTimeFormatMMMMdyyyy = 0xbe0004 /* '\x00\xbe\x00\x04' */
    case DateTimeFormatDMMMyy = 0xbe0005 /* '\x00\xbe\x00\x05' */
    case DateTimeFormatMMMMyy = 0xbe0006 /* '\x00\xbe\x00\x06' */
    case DateTimeFormatMMyy = 0xbe0007 /* '\x00\xbe\x00\x07' */
    case DateTimeFormatMMddyyHmm = 0xbe0008 /* '\x00\xbe\x00\x08' */
    case DateTimeFormatMMddyyhmmAMPM = 0xbe0009 /* '\x00\xbe\x00\t' */
    case DateTimeFormatHmm = 0xbe000a /* '\x00\xbe\x00\n' */
    case DateTimeFormatHmmss = 0xbe000b /* '\x00\xbe\x00\x0b' */
    case DateTimeFormatHmmAMPM = 0xbe000c /* '\x00\xbe\x00\x0c' */
    case DateTimeFormatHmmssAMPM = 0xbe000d /* '\x00\xbe\x00\r' */
    case DateTimeFormatFigureOut = 0xbe000e /* '\x00\xbe\x00\x0e' */
}

// MARK: PowerPointMsoSoftEdgeType
@objc public enum PowerPointMsoSoftEdgeType : AEKeyword {
    case SoftEdgeUnset = 0xbffffe /* '\x00\xbf\xff\xfe' */
    case NoSoftEdge = 0xc00000 /* '\x00\xc0\x00\x00' */
    case SoftEdgeType1 = 0xc00001 /* '\x00\xc0\x00\x01' */
    case SoftEdgeType2 = 0xc00002 /* '\x00\xc0\x00\x02' */
    case SoftEdgeType3 = 0xc00003 /* '\x00\xc0\x00\x03' */
    case SoftEdgeType4 = 0xc00004 /* '\x00\xc0\x00\x04' */
    case SoftEdgeType5 = 0xc00005 /* '\x00\xc0\x00\x05' */
    case SoftEdgeType6 = 0xc00006 /* '\x00\xc0\x00\x06' */
}

// MARK: PowerPointMsoThemeColorSchemeIndex
@objc public enum PowerPointMsoThemeColorSchemeIndex : AEKeyword {
    case FirstDarkSchemeColor = 0xc10001 /* '\x00\xc1\x00\x01' */
    case FirstLightSchemeColor = 0xc10002 /* '\x00\xc1\x00\x02' */
    case SecondDarkSchemeColor = 0xc10003 /* '\x00\xc1\x00\x03' */
    case SecondLightSchemeColor = 0xc10004 /* '\x00\xc1\x00\x04' */
    case FirstAccentSchemeColor = 0xc10005 /* '\x00\xc1\x00\x05' */
    case SecondAccentSchemeColor = 0xc10006 /* '\x00\xc1\x00\x06' */
    case ThirdAccentSchemeColor = 0xc10007 /* '\x00\xc1\x00\x07' */
    case FourthAccentSchemeColor = 0xc10008 /* '\x00\xc1\x00\x08' */
    case FifthAccentSchemeColor = 0xc10009 /* '\x00\xc1\x00\t' */
    case SixthAccentSchemeColor = 0xc1000a /* '\x00\xc1\x00\n' */
    case HyperlinkSchemeColor = 0xc1000b /* '\x00\xc1\x00\x0b' */
    case FollowedHyperlinkSchemeColor = 0xc1000c /* '\x00\xc1\x00\x0c' */
}

// MARK: PowerPointMsoThemeColorIndex
@objc public enum PowerPointMsoThemeColorIndex : AEKeyword {
    case ThemeColorUnset = 0xc1fffe /* '\x00\xc1\xff\xfe' */
    case NoThemeColor = 0xc20000 /* '\x00\xc2\x00\x00' */
    case FirstDarkThemeColor = 0xc20001 /* '\x00\xc2\x00\x01' */
    case FirstLightThemeColor = 0xc20002 /* '\x00\xc2\x00\x02' */
    case SecondDarkThemeColor = 0xc20003 /* '\x00\xc2\x00\x03' */
    case SecondLightThemeColor = 0xc20004 /* '\x00\xc2\x00\x04' */
    case FirstAccentThemeColor = 0xc20005 /* '\x00\xc2\x00\x05' */
    case SecondAccentThemeColor = 0xc20006 /* '\x00\xc2\x00\x06' */
    case ThirdAccentThemeColor = 0xc20007 /* '\x00\xc2\x00\x07' */
    case FourthAccentThemeColor = 0xc20008 /* '\x00\xc2\x00\x08' */
    case FifthAccentThemeColor = 0xc20009 /* '\x00\xc2\x00\t' */
    case SixthAccentThemeColor = 0xc2000a /* '\x00\xc2\x00\n' */
    case HyperlinkThemeColor = 0xc2000b /* '\x00\xc2\x00\x0b' */
    case FollowedHyperlinkThemeColor = 0xc2000c /* '\x00\xc2\x00\x0c' */
    case FirstTextThemeColor = 0xc2000d /* '\x00\xc2\x00\r' */
    case FirstBackgroundThemeColor = 0xc2000e /* '\x00\xc2\x00\x0e' */
    case SecondTextThemeColor = 0xc2000f /* '\x00\xc2\x00\x0f' */
    case SecondBackgroundThemeColor = 0xc20010 /* '\x00\xc2\x00\x10' */
}

// MARK: PowerPointMsoFontLanguageIndex
@objc public enum PowerPointMsoFontLanguageIndex : AEKeyword {
    case ThemeFontLatin = 0xc30001 /* '\x00\xc3\x00\x01' */
    case ThemeFontComplexScript = 0xc30002 /* '\x00\xc3\x00\x02' */
    case ThemeFontHighAnsi = 0xc30003 /* '\x00\xc3\x00\x03' */
    case ThemeFontEastAsian = 0xc30004 /* '\x00\xc3\x00\x04' */
}

// MARK: PowerPointMsoShapeStyleIndex
@objc public enum PowerPointMsoShapeStyleIndex : AEKeyword {
    case ShapeStyleUnset = 0xc3fffe /* '\x00\xc3\xff\xfe' */
    case ShapeNotAPreset = 0xc40000 /* '\x00\xc4\x00\x00' */
    case ShapePreset1 = 0xc40001 /* '\x00\xc4\x00\x01' */
    case ShapePreset2 = 0xc40002 /* '\x00\xc4\x00\x02' */
    case ShapePreset3 = 0xc40003 /* '\x00\xc4\x00\x03' */
    case ShapePreset4 = 0xc40004 /* '\x00\xc4\x00\x04' */
    case ShapePreset5 = 0xc40005 /* '\x00\xc4\x00\x05' */
    case ShapePreset6 = 0xc40006 /* '\x00\xc4\x00\x06' */
    case ShapePreset7 = 0xc40007 /* '\x00\xc4\x00\x07' */
    case ShapePreset8 = 0xc40008 /* '\x00\xc4\x00\x08' */
    case ShapePreset9 = 0xc40009 /* '\x00\xc4\x00\t' */
    case ShapePreset10 = 0xc4000a /* '\x00\xc4\x00\n' */
    case ShapePreset11 = 0xc4000b /* '\x00\xc4\x00\x0b' */
    case ShapePreset12 = 0xc4000c /* '\x00\xc4\x00\x0c' */
    case ShapePreset13 = 0xc4000d /* '\x00\xc4\x00\r' */
    case ShapePreset14 = 0xc4000e /* '\x00\xc4\x00\x0e' */
    case ShapePreset15 = 0xc4000f /* '\x00\xc4\x00\x0f' */
    case ShapePreset16 = 0xc40010 /* '\x00\xc4\x00\x10' */
    case ShapePreset17 = 0xc40011 /* '\x00\xc4\x00\x11' */
    case ShapePreset18 = 0xc40012 /* '\x00\xc4\x00\x12' */
    case ShapePreset19 = 0xc40013 /* '\x00\xc4\x00\x13' */
    case ShapePreset20 = 0xc40014 /* '\x00\xc4\x00\x14' */
    case ShapePreset21 = 0xc40015 /* '\x00\xc4\x00\x15' */
    case ShapePreset22 = 0xc40016 /* '\x00\xc4\x00\x16' */
    case ShapePreset23 = 0xc40017 /* '\x00\xc4\x00\x17' */
    case ShapePreset24 = 0xc40018 /* '\x00\xc4\x00\x18' */
    case ShapePreset25 = 0xc40019 /* '\x00\xc4\x00\x19' */
    case ShapePreset26 = 0xc4001a /* '\x00\xc4\x00\x1a' */
    case ShapePreset27 = 0xc4001b /* '\x00\xc4\x00\x1b' */
    case ShapePreset28 = 0xc4001c /* '\x00\xc4\x00\x1c' */
    case ShapePreset29 = 0xc4001d /* '\x00\xc4\x00\x1d' */
    case ShapePreset30 = 0xc4001e /* '\x00\xc4\x00\x1e' */
    case ShapePreset31 = 0xc4001f /* '\x00\xc4\x00\x1f' */
    case ShapePreset32 = 0xc40020 /* '\x00\xc4\x00 ' */
    case ShapePreset33 = 0xc40021 /* '\x00\xc4\x00!' */
    case ShapePreset34 = 0xc40022 /* '\x00\xc4\x00"' */
    case ShapePreset35 = 0xc40023 /* '\x00\xc4\x00#' */
    case ShapePreset36 = 0xc40024 /* '\x00\xc4\x00$' */
    case ShapePreset37 = 0xc40025 /* '\x00\xc4\x00%' */
    case ShapePreset38 = 0xc40026 /* '\x00\xc4\x00&' */
    case ShapePreset39 = 0xc40027 /* "\x00\xc4\x00'" */
    case ShapePreset40 = 0xc40028 /* '\x00\xc4\x00(' */
    case ShapePreset41 = 0xc40029 /* '\x00\xc4\x00)' */
    case ShapePreset42 = 0xc4002a /* '\x00\xc4\x00*' */
    case LinePreset1 = 0xc42711 /* "\x00\xc4'\x11" */
    case LinePreset2 = 0xc42712 /* "\x00\xc4'\x12" */
    case LinePreset3 = 0xc42713 /* "\x00\xc4'\x13" */
    case LinePreset4 = 0xc42714 /* "\x00\xc4'\x14" */
    case LinePreset5 = 0xc42715 /* "\x00\xc4'\x15" */
    case LinePreset6 = 0xc42716 /* "\x00\xc4'\x16" */
    case LinePreset7 = 0xc42717 /* "\x00\xc4'\x17" */
    case LinePreset8 = 0xc42718 /* "\x00\xc4'\x18" */
    case LinePreset9 = 0xc42719 /* "\x00\xc4'\x19" */
    case LinePreset10 = 0xc4271a /* "\x00\xc4'\x1a" */
    case LinePreset11 = 0xc4271b /* "\x00\xc4'\x1b" */
    case LinePreset12 = 0xc4271c /* "\x00\xc4'\x1c" */
    case LinePreset13 = 0xc4271d /* "\x00\xc4'\x1d" */
    case LinePreset14 = 0xc4271e /* "\x00\xc4'\x1e" */
    case LinePreset15 = 0xc4271f /* "\x00\xc4'\x1f" */
    case LinePreset16 = 0xc42720 /* "\x00\xc4' " */
    case LinePreset17 = 0xc42721 /* "\x00\xc4'!" */
    case LinePreset18 = 0xc42722 /* '\x00\xc4\'"' */
    case LinePreset19 = 0xc42723 /* "\x00\xc4'#" */
    case LinePreset20 = 0xc42724 /* "\x00\xc4'$" */
    case LinePreset21 = 0xc42725 /* "\x00\xc4'%" */
}

// MARK: PowerPointMsoBackgroundStyleIndex
@objc public enum PowerPointMsoBackgroundStyleIndex : AEKeyword {
    case BackgroundUnset = 0xc4fffe /* '\x00\xc4\xff\xfe' */
    case BackgroundNotAPreset = 0xc50000 /* '\x00\xc5\x00\x00' */
    case BackgroundPreset1 = 0xc50001 /* '\x00\xc5\x00\x01' */
    case BackgroundPreset2 = 0xc50002 /* '\x00\xc5\x00\x02' */
    case BackgroundPreset3 = 0xc50003 /* '\x00\xc5\x00\x03' */
    case BackgroundPreset4 = 0xc50004 /* '\x00\xc5\x00\x04' */
    case BackgroundPreset5 = 0xc50005 /* '\x00\xc5\x00\x05' */
    case BackgroundPreset6 = 0xc50006 /* '\x00\xc5\x00\x06' */
    case BackgroundPreset7 = 0xc50007 /* '\x00\xc5\x00\x07' */
    case BackgroundPreset8 = 0xc50008 /* '\x00\xc5\x00\x08' */
    case BackgroundPreset9 = 0xc50009 /* '\x00\xc5\x00\t' */
    case BackgroundPreset10 = 0xc5000a /* '\x00\xc5\x00\n' */
    case BackgroundPreset11 = 0xc5000b /* '\x00\xc5\x00\x0b' */
    case BackgroundPreset12 = 0xc5000c /* '\x00\xc5\x00\x0c' */
}

// MARK: PowerPointMsoTextDirection
@objc public enum PowerPointMsoTextDirection : AEKeyword {
    case TextDirectionUnset = 0xeafffe /* '\x00\xea\xff\xfe' */
    case LeftToRight = 0xeb0001 /* '\x00\xeb\x00\x01' */
    case RightToLeft = 0xeb0002 /* '\x00\xeb\x00\x02' */
}

// MARK: PowerPointMsoBulletType
@objc public enum PowerPointMsoBulletType : AEKeyword {
    case BulletTypeUnset = 0xe7fffe /* '\x00\xe7\xff\xfe' */
    case BulletTypeNone = 0xe80000 /* '\x00\xe8\x00\x00' */
    case BulletTypeUnnumbered = 0xe80001 /* '\x00\xe8\x00\x01' */
    case BulletTypeNumbered = 0xe80002 /* '\x00\xe8\x00\x02' */
    case PictureBulletType = 0xe80003 /* '\x00\xe8\x00\x03' */
}

// MARK: PowerPointMsoNumberedBulletStyle
@objc public enum PowerPointMsoNumberedBulletStyle : AEKeyword {
    case NumberedBulletStyleUnset = 0xe8fffe /* '\x00\xe8\xff\xfe' */
    case NumberedBulletStyleAlphaLowercasePeriod = 0xe90000 /* '\x00\xe9\x00\x00' */
    case NumberedBulletStyleAlphaUppercasePeriod = 0xe90001 /* '\x00\xe9\x00\x01' */
    case NumberedBulletStyleArabicRightParen = 0xe90002 /* '\x00\xe9\x00\x02' */
    case NumberedBulletStyleArabicPeriod = 0xe90003 /* '\x00\xe9\x00\x03' */
    case NumberedBulletStyleRomanLowercaseParenBoth = 0xe90004 /* '\x00\xe9\x00\x04' */
    case NumberedBulletStyleRomanLowercaseParenRight = 0xe90005 /* '\x00\xe9\x00\x05' */
    case NumberedBulletStyleRomanLowercasePeriod = 0xe90006 /* '\x00\xe9\x00\x06' */
    case NumberedBulletStyleRomanUppercasePeriod = 0xe90007 /* '\x00\xe9\x00\x07' */
    case NumberedBulletStyleAlphaLowercaseParenBoth = 0xe90008 /* '\x00\xe9\x00\x08' */
    case NumberedBulletStyleAlphaLowercaseParenRight = 0xe90009 /* '\x00\xe9\x00\t' */
    case NumberedBulletStyleAlphaUppercaseParenBoth = 0xe9000a /* '\x00\xe9\x00\n' */
    case NumberedBulletStyleAlphaUppercaseParenRight = 0xe9000b /* '\x00\xe9\x00\x0b' */
    case NumberedBulletStyleArabicParenBoth = 0xe9000c /* '\x00\xe9\x00\x0c' */
    case NumberedBulletStyleArabicPlain = 0xe9000d /* '\x00\xe9\x00\r' */
    case NumberedBulletStyleRomanUppercaseParenBoth = 0xe9000e /* '\x00\xe9\x00\x0e' */
    case NumberedBulletStyleRomanUppercaseParenRight = 0xe9000f /* '\x00\xe9\x00\x0f' */
    case NumberedBulletStyleSimplifiedChinesePlain = 0xe90010 /* '\x00\xe9\x00\x10' */
    case NumberedBulletStyleSimplifiedChinesePeriod = 0xe90011 /* '\x00\xe9\x00\x11' */
    case NumberedBulletStyleCircleNumberPlain = 0xe90012 /* '\x00\xe9\x00\x12' */
    case NumberedBulletStyleCircleNumberWhitePlain = 0xe90013 /* '\x00\xe9\x00\x13' */
    case NumberedBulletStyleCircleNumberBlackPlain = 0xe90014 /* '\x00\xe9\x00\x14' */
    case NumberedBulletStyleTraditionalChinesePlain = 0xe90015 /* '\x00\xe9\x00\x15' */
    case NumberedBulletStyleTraditionalChinesePeriod = 0xe90016 /* '\x00\xe9\x00\x16' */
    case NumberedBulletStyleArabicAlphaDash = 0xe90017 /* '\x00\xe9\x00\x17' */
    case NumberedBulletStyleArabicAbjadDash = 0xe90018 /* '\x00\xe9\x00\x18' */
    case NumberedBulletStyleHebrewAlphaDash = 0xe90019 /* '\x00\xe9\x00\x19' */
    case NumberedBulletStyleKanjiKoreanPlain = 0xe9001a /* '\x00\xe9\x00\x1a' */
    case NumberedBulletStyleKanjiKoreanPeriod = 0xe9001b /* '\x00\xe9\x00\x1b' */
    case NumberedBulletStyleArabicDBPlain = 0xe9001c /* '\x00\xe9\x00\x1c' */
    case NumberedBulletStyleArabicDBPeriod = 0xe9001d /* '\x00\xe9\x00\x1d' */
    case NumberedBulletStyleThaiAlphaPeriod = 0xe9001e /* '\x00\xe9\x00\x1e' */
    case NumberedBulletStyleThaiAlphaParenRight = 0xe9001f /* '\x00\xe9\x00\x1f' */
    case NumberedBulletStyleThaiAlphaParenBoth = 0xe90020 /* '\x00\xe9\x00 ' */
    case NumberedBulletStyleThaiNumberPeriod = 0xe90021 /* '\x00\xe9\x00!' */
    case NumberedBulletStyleThaiNumberParenRight = 0xe90022 /* '\x00\xe9\x00"' */
    case NumberedBulletStyleThaiParenBoth = 0xe90023 /* '\x00\xe9\x00#' */
    case NumberedBulletStyleHindiAlphaPeriod = 0xe90024 /* '\x00\xe9\x00$' */
    case NumberedBulletStyleHindiNumberPeriod = 0xe90025 /* '\x00\xe9\x00%' */
    case NumberedBulletStyleKanjiSimpifiedChineseDBPeriod = 0xe90026 /* '\x00\xe9\x00&' */
    case NumberedBulletStyleHindiNumberParenRight = 0xe90027 /* "\x00\xe9\x00'" */
    case NumberedBulletStyleHindiAlpha1Period = 0xe90028 /* '\x00\xe9\x00(' */
}

// MARK: PowerPointMsoTabStopType
@objc public enum PowerPointMsoTabStopType : AEKeyword {
    case TabstopUnset = 0xf4fffe /* '\x00\xf4\xff\xfe' */
    case TabstopLeft = 0xf50001 /* '\x00\xf5\x00\x01' */
    case TabstopCenter = 0xf50002 /* '\x00\xf5\x00\x02' */
    case TabstopRight = 0xf50003 /* '\x00\xf5\x00\x03' */
    case TabstopDecimal = 0xf50004 /* '\x00\xf5\x00\x04' */
}

// MARK: PowerPointMsoReflectionType
@objc public enum PowerPointMsoReflectionType : AEKeyword {
    case ReflectionUnset = 0x3e9fffe /* '\x03\xe9\xff\xfe' */
    case ReflectionTypeNone = 0x3ea0000 /* '\x03\xea\x00\x00' */
    case ReflectionType1 = 0x3ea0001 /* '\x03\xea\x00\x01' */
    case ReflectionType2 = 0x3ea0002 /* '\x03\xea\x00\x02' */
    case ReflectionType3 = 0x3ea0003 /* '\x03\xea\x00\x03' */
    case ReflectionType4 = 0x3ea0004 /* '\x03\xea\x00\x04' */
    case ReflectionType5 = 0x3ea0005 /* '\x03\xea\x00\x05' */
    case ReflectionType6 = 0x3ea0006 /* '\x03\xea\x00\x06' */
    case ReflectionType7 = 0x3ea0007 /* '\x03\xea\x00\x07' */
    case ReflectionType8 = 0x3ea0008 /* '\x03\xea\x00\x08' */
    case ReflectionType9 = 0x3ea0009 /* '\x03\xea\x00\t' */
}

// MARK: PowerPointMsoTextureAlignment
@objc public enum PowerPointMsoTextureAlignment : AEKeyword {
    case TextureUnset = 0x3eafffe /* '\x03\xea\xff\xfe' */
    case TextureTopLeft = 0x3eb0000 /* '\x03\xeb\x00\x00' */
    case TextureTop = 0x3eb0001 /* '\x03\xeb\x00\x01' */
    case TextureTopRight = 0x3eb0002 /* '\x03\xeb\x00\x02' */
    case TextureLeft = 0x3eb0003 /* '\x03\xeb\x00\x03' */
    case TextureCenter = 0x3eb0004 /* '\x03\xeb\x00\x04' */
    case TextureRight = 0x3eb0005 /* '\x03\xeb\x00\x05' */
    case TextureBottomLeft = 0x3eb0006 /* '\x03\xeb\x00\x06' */
    case TextureBotton = 0x3eb0007 /* '\x03\xeb\x00\x07' */
    case TextureBottomRight = 0x3eb0008 /* '\x03\xeb\x00\x08' */
}

// MARK: PowerPointMsoBaselineAlignment
@objc public enum PowerPointMsoBaselineAlignment : AEKeyword {
    case TextBaselineAlignmentUnset = 0x3ebfffe /* '\x03\xeb\xff\xfe' */
    case TextBaselineAlignBaseline = 0x3ec0001 /* '\x03\xec\x00\x01' */
    case TextBaselineAlignTop = 0x3ec0002 /* '\x03\xec\x00\x02' */
    case TextBaselineAlignCenter = 0x3ec0003 /* '\x03\xec\x00\x03' */
    case TextBaselineAlignEastAsian50 = 0x3ec0004 /* '\x03\xec\x00\x04' */
    case TextBaselineAlignAutomatic = 0x3ec0005 /* '\x03\xec\x00\x05' */
}

// MARK: PowerPointMsoClipboardFormat
@objc public enum PowerPointMsoClipboardFormat : AEKeyword {
    case ClipboardFormatUnset = 0x3ecfffe /* '\x03\xec\xff\xfe' */
    case NativeClipboardFormat = 0x3ed0001 /* '\x03\xed\x00\x01' */
    case HTMlClipboardFormat = 0x3ed0002 /* '\x03\xed\x00\x02' */
    case RTFClipboardFormat = 0x3ed0003 /* '\x03\xed\x00\x03' */
    case PlainTextClipboardFormat = 0x3ed0004 /* '\x03\xed\x00\x04' */
}

// MARK: PowerPointMsoTextRangeInsertPosition
@objc public enum PowerPointMsoTextRangeInsertPosition : AEKeyword {
    case InsertBefore = 0x3ee0000 /* '\x03\xee\x00\x00' */
    case InsertAfter = 0x3ee0001 /* '\x03\xee\x00\x01' */
}

// MARK: PowerPointMsoPictureType
@objc public enum PowerPointMsoPictureType : AEKeyword {
    case SaveAsDefault = 0x3f2fffe /* '\x03\xf2\xff\xfe' */
    case SaveAsPNGFile = 0x3f30000 /* '\x03\xf3\x00\x00' */
    case SaveAsBMPFile = 0x3f30001 /* '\x03\xf3\x00\x01' */
    case SaveAsGIFFile = 0x3f30002 /* '\x03\xf3\x00\x02' */
    case SaveAsJPGFile = 0x3f30003 /* '\x03\xf3\x00\x03' */
    case SaveAsPDFFile = 0x3f30004 /* '\x03\xf3\x00\x04' */
}

// MARK: PowerPointMsoPictureEffectType
@objc public enum PowerPointMsoPictureEffectType : AEKeyword {
    case NoEffect = 0x3f40000 /* '\x03\xf4\x00\x00' */
    case EffectBackgroundRemoval = 0x3f40001 /* '\x03\xf4\x00\x01' */
    case EffectBlur = 0x3f40002 /* '\x03\xf4\x00\x02' */
    case EffectBrightnessContrast = 0x3f40003 /* '\x03\xf4\x00\x03' */
    case EffectCement = 0x3f40004 /* '\x03\xf4\x00\x04' */
    case EffectCrisscrossEtching = 0x3f40005 /* '\x03\xf4\x00\x05' */
    case EffectChalkSketch = 0x3f40006 /* '\x03\xf4\x00\x06' */
    case EffectColorTemperature = 0x3f40007 /* '\x03\xf4\x00\x07' */
    case EffectCutout = 0x3f40008 /* '\x03\xf4\x00\x08' */
    case EffectFilmGrain = 0x3f40009 /* '\x03\xf4\x00\t' */
    case EffectGlass = 0x3f4000a /* '\x03\xf4\x00\n' */
    case EffectGlowDiffused = 0x3f4000b /* '\x03\xf4\x00\x0b' */
    case EffectGlowEdges = 0x3f4000c /* '\x03\xf4\x00\x0c' */
    case EffectLightScreen = 0x3f4000d /* '\x03\xf4\x00\r' */
    case EffectLineDrawing = 0x3f4000e /* '\x03\xf4\x00\x0e' */
    case EffectMarker = 0x3f4000f /* '\x03\xf4\x00\x0f' */
    case EffectMosiaicBubbles = 0x3f40010 /* '\x03\xf4\x00\x10' */
    case EffectPaintBrush = 0x3f40011 /* '\x03\xf4\x00\x11' */
    case EffectPaintStrokes = 0x3f40012 /* '\x03\xf4\x00\x12' */
    case EffectPastelsSmooth = 0x3f40013 /* '\x03\xf4\x00\x13' */
    case EffectPencilGrayscale = 0x3f40014 /* '\x03\xf4\x00\x14' */
    case EffectPencilSketch = 0x3f40015 /* '\x03\xf4\x00\x15' */
    case EffectPhotocopy = 0x3f40016 /* '\x03\xf4\x00\x16' */
    case EffectPlasticWrap = 0x3f40017 /* '\x03\xf4\x00\x17' */
    case EffectSaturation = 0x3f40018 /* '\x03\xf4\x00\x18' */
    case EffectSharpenSoften = 0x3f40019 /* '\x03\xf4\x00\x19' */
    case EffectTexturizer = 0x3f4001a /* '\x03\xf4\x00\x1a' */
    case EffectWatercolorSponge = 0x3f4001b /* '\x03\xf4\x00\x1b' */
}

// MARK: PowerPointMsoSegmentType
@objc public enum PowerPointMsoSegmentType : AEKeyword {
    case Line = 0x8f0000 /* '\x00\x8f\x00\x00' */
    case Curve = 0x8f0001 /* '\x00\x8f\x00\x01' */
}

// MARK: PowerPointMsoEditingType
@objc public enum PowerPointMsoEditingType : AEKeyword {
    case Auto = 0x900000 /* '\x00\x90\x00\x00' */
    case Corner = 0x900001 /* '\x00\x90\x00\x01' */
    case Smooth = 0x900002 /* '\x00\x90\x00\x02' */
    case Symmetric = 0x900003 /* '\x00\x90\x00\x03' */
}

// MARK: PowerPointMsoSmartArtNodePosition
@objc public enum PowerPointMsoSmartArtNodePosition : AEKeyword {
    case DefaultNodePosition = 0x3f50001 /* '\x03\xf5\x00\x01' */
    case AfterNode = 0x3f50002 /* '\x03\xf5\x00\x02' */
    case BeforeNode = 0x3f50003 /* '\x03\xf5\x00\x03' */
    case AboveNode = 0x3f50004 /* '\x03\xf5\x00\x04' */
    case BelowNode = 0x3f50005 /* '\x03\xf5\x00\x05' */
}

// MARK: PowerPointMsoSmartArtNodeType
@objc public enum PowerPointMsoSmartArtNodeType : AEKeyword {
    case DefaultNode = 0x3f60001 /* '\x03\xf6\x00\x01' */
    case AssistantNode = 0x3f60002 /* '\x03\xf6\x00\x02' */
}

// MARK: PowerPointMsoOrgChartLayoutType
@objc public enum PowerPointMsoOrgChartLayoutType : AEKeyword {
    case OrgChartLayoutUnset = 0x3f6fffe /* '\x03\xf6\xff\xfe' */
    case OrgChartLayoutStandard = 0x3f70001 /* '\x03\xf7\x00\x01' */
    case OrgChartLayoutBothHanging = 0x3f70002 /* '\x03\xf7\x00\x02' */
    case OrgChartLayoutLeftHanging = 0x3f70003 /* '\x03\xf7\x00\x03' */
    case OrgChartLayoutRightHanging = 0x3f70004 /* '\x03\xf7\x00\x04' */
    case OrgChartLayoutDefault = 0x3f70005 /* '\x03\xf7\x00\x05' */
}

// MARK: PowerPointMsoAlignCmd
@objc public enum PowerPointMsoAlignCmd : AEKeyword {
    case AlignLefts = 0x0 /* '\x00\x00\x00\x00' */
    case AlignCenters = 0x1 /* '\x00\x00\x00\x01' */
    case AlignRights = 0x2 /* '\x00\x00\x00\x02' */
    case AlignTops = 0x3 /* '\x00\x00\x00\x03' */
    case AlignMiddles = 0x4 /* '\x00\x00\x00\x04' */
    case AlignBottoms = 0x5 /* '\x00\x00\x00\x05' */
}

// MARK: PowerPointMsoDistributeCmd
@objc public enum PowerPointMsoDistributeCmd : AEKeyword {
    case DistributeHorizontally = 0x0 /* '\x00\x00\x00\x00' */
    case DistributeVertically = 0x1 /* '\x00\x00\x00\x01' */
}

// MARK: PowerPointMsoOrientation
@objc public enum PowerPointMsoOrientation : AEKeyword {
    case OrientationUnset = 0x8cfffe /* '\x00\x8c\xff\xfe' */
    case HorizontalOrientation = 0x8d0001 /* '\x00\x8d\x00\x01' */
    case VerticalOrientation = 0x8d0002 /* '\x00\x8d\x00\x02' */
}

// MARK: PowerPointMsoZOrderCmd
@objc public enum PowerPointMsoZOrderCmd : AEKeyword {
    case BringShapeToFront = 0x700000 /* '\x00p\x00\x00' */
    case SendShapeToBack = 0x700001 /* '\x00p\x00\x01' */
    case BringShapeForward = 0x700002 /* '\x00p\x00\x02' */
    case SendShapeBackward = 0x700003 /* '\x00p\x00\x03' */
    case BringShapeInFrontOfText = 0x700004 /* '\x00p\x00\x04' */
    case SendShapeBehindText = 0x700005 /* '\x00p\x00\x05' */
}

// MARK: PowerPointMsoScriptLanguage
@objc public enum PowerPointMsoScriptLanguage : AEKeyword {
    case BringShapeToFront = 0x6f0001 /* '\x00o\x00\x01' */
    case SendShapeToBack = 0x6f0002 /* '\x00o\x00\x02' */
    case BringShapeForward = 0x6f0003 /* '\x00o\x00\x03' */
    case SendShapeBackward = 0x6f0004 /* '\x00o\x00\x04' */
}

// MARK: PowerPointMsoFlipCmd
@objc public enum PowerPointMsoFlipCmd : AEKeyword {
    case FlipHorizontal = 0x710000 /* '\x00q\x00\x00' */
    case FlipVertical = 0x710001 /* '\x00q\x00\x01' */
}

// MARK: PowerPointMsoTriState
@objc public enum PowerPointMsoTriState : AEKeyword {
    case True = 0xa0ffff /* '\x00\xa0\xff\xff' */
    case False = 0xa10000 /* '\x00\xa1\x00\x00' */
    case CTrue = 0xa10001 /* '\x00\xa1\x00\x01' */
    case Toggle = 0xa0fffd /* '\x00\xa0\xff\xfd' */
    case TriStateUnset = 0xa0fffe /* '\x00\xa0\xff\xfe' */
}

// MARK: PowerPointMsoBlackWhiteMode
@objc public enum PowerPointMsoBlackWhiteMode : AEKeyword {
    case BlackAndWhiteUnset = 0xacfffe /* '\x00\xac\xff\xfe' */
    case BlackAndWhiteModeAutomatic = 0xad0001 /* '\x00\xad\x00\x01' */
    case BlackAndWhiteModeGrayScale = 0xad0002 /* '\x00\xad\x00\x02' */
    case BlackAndWhiteModeLightGrayScale = 0xad0003 /* '\x00\xad\x00\x03' */
    case BlackAndWhiteModeInverseGrayScale = 0xad0004 /* '\x00\xad\x00\x04' */
    case BlackAndWhiteModeGrayOutline = 0xad0005 /* '\x00\xad\x00\x05' */
    case BlackAndWhiteModeBlackTextAndLine = 0xad0006 /* '\x00\xad\x00\x06' */
    case BlackAndWhiteModeHighContrast = 0xad0007 /* '\x00\xad\x00\x07' */
    case BlackAndWhiteModeBlack = 0xad0008 /* '\x00\xad\x00\x08' */
    case BlackAndWhiteModeWhite = 0xad0009 /* '\x00\xad\x00\t' */
    case BlackAndWhiteModeDontShow = 0xad000a /* '\x00\xad\x00\n' */
}

// MARK: PowerPointMsoBarPosition
@objc public enum PowerPointMsoBarPosition : AEKeyword {
    case BarLeft = 0x720000 /* '\x00r\x00\x00' */
    case BarTop = 0x720001 /* '\x00r\x00\x01' */
    case BarRight = 0x720002 /* '\x00r\x00\x02' */
    case BarBottom = 0x720003 /* '\x00r\x00\x03' */
    case BarFloating = 0x720004 /* '\x00r\x00\x04' */
    case BarPopUp = 0x720005 /* '\x00r\x00\x05' */
    case BarMenu = 0x720006 /* '\x00r\x00\x06' */
}

// MARK: PowerPointMsoBarProtection
@objc public enum PowerPointMsoBarProtection : AEKeyword {
    case NoProtection = 0x730000 /* '\x00s\x00\x00' */
    case NoCustomize = 0x730001 /* '\x00s\x00\x01' */
    case NoResize = 0x730002 /* '\x00s\x00\x02' */
    case NoMove = 0x730004 /* '\x00s\x00\x04' */
    case NoChangeVisible = 0x730008 /* '\x00s\x00\x08' */
    case NoChangeDock = 0x730010 /* '\x00s\x00\x10' */
    case NoVerticalDock = 0x730020 /* '\x00s\x00 ' */
    case NoHorizontalDock = 0x730040 /* '\x00s\x00@' */
}

// MARK: PowerPointMsoBarType
@objc public enum PowerPointMsoBarType : AEKeyword {
    case NormalCommandBar = 0x740000 /* '\x00t\x00\x00' */
    case MenubarCommandBar = 0x740001 /* '\x00t\x00\x01' */
    case PopupCommandBar = 0x740002 /* '\x00t\x00\x02' */
}

// MARK: PowerPointMsoControlType
@objc public enum PowerPointMsoControlType : AEKeyword {
    case ControlCustom = 0x750000 /* '\x00u\x00\x00' */
    case ControlButton = 0x750001 /* '\x00u\x00\x01' */
    case ControlEdit = 0x750002 /* '\x00u\x00\x02' */
    case ControlDropDown = 0x750003 /* '\x00u\x00\x03' */
    case ControlCombobox = 0x750004 /* '\x00u\x00\x04' */
    case ButtonDropDown = 0x750005 /* '\x00u\x00\x05' */
    case SplitDropDown = 0x750006 /* '\x00u\x00\x06' */
    case OCXDropDown = 0x750007 /* '\x00u\x00\x07' */
    case GenericDropDown = 0x750008 /* '\x00u\x00\x08' */
    case GraphicDropDown = 0x750009 /* '\x00u\x00\t' */
    case ControlPopup = 0x75000a /* '\x00u\x00\n' */
    case GraphicPopup = 0x75000b /* '\x00u\x00\x0b' */
    case ButtonPopup = 0x75000c /* '\x00u\x00\x0c' */
    case SplitButtonPopup = 0x75000d /* '\x00u\x00\r' */
    case SplitButtonMRUPopup = 0x75000e /* '\x00u\x00\x0e' */
    case ControlLabel = 0x75000f /* '\x00u\x00\x0f' */
    case ExpandingGrid = 0x750010 /* '\x00u\x00\x10' */
    case SplitExpandingGrid = 0x750011 /* '\x00u\x00\x11' */
    case ControlGrid = 0x750012 /* '\x00u\x00\x12' */
    case ControlGauge = 0x750013 /* '\x00u\x00\x13' */
    case GraphicCombobox = 0x750014 /* '\x00u\x00\x14' */
    case ControlPane = 0x750015 /* '\x00u\x00\x15' */
    case ActiveX = 0x750016 /* '\x00u\x00\x16' */
    case ControlGroup = 0x750017 /* '\x00u\x00\x17' */
    case ControlTab = 0x750018 /* '\x00u\x00\x18' */
    case ControlSpinner = 0x750019 /* '\x00u\x00\x19' */
}

// MARK: PowerPointMsoButtonState
@objc public enum PowerPointMsoButtonState : AEKeyword {
    case ButtonStateUp = 0x760000 /* '\x00v\x00\x00' */
    case ButtonStateDown = 0x75ffff /* '\x00u\xff\xff' */
    case ButtonStateUnset = 0x760002 /* '\x00v\x00\x02' */
}

// MARK: PowerPointMsoControlOLEUsage
@objc public enum PowerPointMsoControlOLEUsage : AEKeyword {
    case Neither = 0x770000 /* '\x00w\x00\x00' */
    case Server = 0x770001 /* '\x00w\x00\x01' */
    case Client = 0x770002 /* '\x00w\x00\x02' */
    case Both = 0x770003 /* '\x00w\x00\x03' */
}

// MARK: PowerPointMsoButtonStyle
@objc public enum PowerPointMsoButtonStyle : AEKeyword {
    case ButtonAutomatic = 0x780000 /* '\x00x\x00\x00' */
    case ButtonIcon = 0x780001 /* '\x00x\x00\x01' */
    case ButtonCaption = 0x780002 /* '\x00x\x00\x02' */
    case ButtonIconAndCaption = 0x780003 /* '\x00x\x00\x03' */
}

// MARK: PowerPointMsoComboStyle
@objc public enum PowerPointMsoComboStyle : AEKeyword {
    case ComboboxStyleNormal = 0x790000 /* '\x00y\x00\x00' */
    case ComboboxStyleLabel = 0x790001 /* '\x00y\x00\x01' */
}

// MARK: PowerPointMsoMenuAnimation
@objc public enum PowerPointMsoMenuAnimation : AEKeyword {
    case None = 0x7b0000 /* '\x00{\x00\x00' */
    case Random = 0x7b0001 /* '\x00{\x00\x01' */
    case Unfold = 0x7b0002 /* '\x00{\x00\x02' */
    case Slide = 0x7b0003 /* '\x00{\x00\x03' */
}

// MARK: PowerPointMsoHyperlinkType
@objc public enum PowerPointMsoHyperlinkType : AEKeyword {
    case HyperlinkTypeTextRange = 0x960000 /* '\x00\x96\x00\x00' */
    case HyperlinkTypeShape = 0x960001 /* '\x00\x96\x00\x01' */
    case HyperlinkTypeInlineShape = 0x960002 /* '\x00\x96\x00\x02' */
}

// MARK: PowerPointMsoExtraInfoMethod
@objc public enum PowerPointMsoExtraInfoMethod : AEKeyword {
    case AppendString = 0xae0000 /* '\x00\xae\x00\x00' */
    case PostString = 0xae0001 /* '\x00\xae\x00\x01' */
}

// MARK: PowerPointMsoAnimationType
@objc public enum PowerPointMsoAnimationType : AEKeyword {
    case Idle = 0x7c0001 /* '\x00|\x00\x01' */
    case Greeting = 0x7c0002 /* '\x00|\x00\x02' */
    case Goodbye = 0x7c0003 /* '\x00|\x00\x03' */
    case BeginSpeaking = 0x7c0004 /* '\x00|\x00\x04' */
    case CharacterSuccessMajor = 0x7c0006 /* '\x00|\x00\x06' */
    case GetAttentionMajor = 0x7c000b /* '\x00|\x00\x0b' */
    case GetAttentionMinor = 0x7c000c /* '\x00|\x00\x0c' */
    case Searching = 0x7c000d /* '\x00|\x00\r' */
    case Printing = 0x7c0012 /* '\x00|\x00\x12' */
    case GestureRight = 0x7c0013 /* '\x00|\x00\x13' */
    case WritingNotingSomething = 0x7c0016 /* '\x00|\x00\x16' */
    case WorkingAtSomething = 0x7c0017 /* '\x00|\x00\x17' */
    case Thinking = 0x7c0018 /* '\x00|\x00\x18' */
    case SendingMail = 0x7c0019 /* '\x00|\x00\x19' */
    case ListensToComputer = 0x7c001a /* '\x00|\x00\x1a' */
    case Disappear = 0x7c001f /* '\x00|\x00\x1f' */
    case Appear = 0x7c0020 /* '\x00|\x00 ' */
    case GetArtsy = 0x7c0064 /* '\x00|\x00d' */
    case GetTechy = 0x7c0065 /* '\x00|\x00e' */
    case GetWizardy = 0x7c0066 /* '\x00|\x00f' */
    case CheckingSomething = 0x7c0067 /* '\x00|\x00g' */
    case LookDown = 0x7c0068 /* '\x00|\x00h' */
    case LookDownLeft = 0x7c0069 /* '\x00|\x00i' */
    case LookDownRight = 0x7c006a /* '\x00|\x00j' */
    case LookLeft = 0x7c006b /* '\x00|\x00k' */
    case LookRight = 0x7c006c /* '\x00|\x00l' */
    case LookUp = 0x7c006d /* '\x00|\x00m' */
    case LookUpLeft = 0x7c006e /* '\x00|\x00n' */
    case LookUpRight = 0x7c006f /* '\x00|\x00o' */
    case Saving = 0x7c0070 /* '\x00|\x00p' */
    case GestureDown = 0x7c0071 /* '\x00|\x00q' */
    case GestureLeft = 0x7c0072 /* '\x00|\x00r' */
    case GestureUp = 0x7c0073 /* '\x00|\x00s' */
    case EmptyTrash = 0x7c0074 /* '\x00|\x00t' */
}

// MARK: PowerPointMsoButtonSetType
@objc public enum PowerPointMsoButtonSetType : AEKeyword {
    case ButtonNone = 0x7d0000 /* '\x00}\x00\x00' */
    case ButtonOk = 0x7d0001 /* '\x00}\x00\x01' */
    case ButtonCancel = 0x7d0002 /* '\x00}\x00\x02' */
    case ButtonsOkCancel = 0x7d0003 /* '\x00}\x00\x03' */
    case ButtonsYesNo = 0x7d0004 /* '\x00}\x00\x04' */
    case ButtonsYesNoCancel = 0x7d0005 /* '\x00}\x00\x05' */
    case ButtonsBackClose = 0x7d0006 /* '\x00}\x00\x06' */
    case ButtonsNextClose = 0x7d0007 /* '\x00}\x00\x07' */
    case ButtonsBackNextClose = 0x7d0008 /* '\x00}\x00\x08' */
    case ButtonsRetryCancel = 0x7d0009 /* '\x00}\x00\t' */
    case ButtonsAbortRetryIgnore = 0x7d000a /* '\x00}\x00\n' */
    case ButtonsSearchClose = 0x7d000b /* '\x00}\x00\x0b' */
    case ButtonsBackNextSnooze = 0x7d000c /* '\x00}\x00\x0c' */
    case ButtonsTipsOptionsClose = 0x7d000d /* '\x00}\x00\r' */
    case ButtonsYesAllNoCancel = 0x7d000e /* '\x00}\x00\x0e' */
}

// MARK: PowerPointMsoIconType
@objc public enum PowerPointMsoIconType : AEKeyword {
    case IconNone = 0x7e0000 /* '\x00~\x00\x00' */
    case IconApplication = 0x7e0001 /* '\x00~\x00\x01' */
    case IconAlert = 0x7e0002 /* '\x00~\x00\x02' */
    case IconTip = 0x7e0003 /* '\x00~\x00\x03' */
    case IconAlertCritical = 0x7e0065 /* '\x00~\x00e' */
    case IconAlertWarning = 0x7e0067 /* '\x00~\x00g' */
    case IconAlertInfo = 0x7e0068 /* '\x00~\x00h' */
}

// MARK: PowerPointMsoWizardActType
@objc public enum PowerPointMsoWizardActType : AEKeyword {
    case Inactive = 0x820000 /* '\x00\x82\x00\x00' */
    case Active = 0x820001 /* '\x00\x82\x00\x01' */
    case Suspend = 0x820002 /* '\x00\x82\x00\x02' */
    case Resume = 0x820003 /* '\x00\x82\x00\x03' */
}

// MARK: PowerPointMsoDocProperties
@objc public enum PowerPointMsoDocProperties : AEKeyword {
    case PropertyTypeNumber = 0xa20001 /* '\x00\xa2\x00\x01' */
    case PropertyTypeBoolean = 0xa20002 /* '\x00\xa2\x00\x02' */
    case PropertyTypeDate = 0xa20003 /* '\x00\xa2\x00\x03' */
    case PropertyTypeString = 0xa20004 /* '\x00\xa2\x00\x04' */
    case PropertyTypeFloat = 0xa20005 /* '\x00\xa2\x00\x05' */
}

// MARK: PowerPointMsoAutomationSecurity
@objc public enum PowerPointMsoAutomationSecurity : AEKeyword {
    case MsoAutomationSecurityLow = 0xa30001 /* '\x00\xa3\x00\x01' */
    case MsoAutomationSecurityByUI = 0xa30002 /* '\x00\xa3\x00\x02' */
    case MsoAutomationSecurityForceDisable = 0xa30003 /* '\x00\xa3\x00\x03' */
}

// MARK: PowerPointMsoScreenSize
@objc public enum PowerPointMsoScreenSize : AEKeyword {
    case Resolution544x376 = 0x840000 /* '\x00\x84\x00\x00' */
    case Resolution640x480 = 0x840001 /* '\x00\x84\x00\x01' */
    case Resolution720x512 = 0x840002 /* '\x00\x84\x00\x02' */
    case Resolution800x600 = 0x840003 /* '\x00\x84\x00\x03' */
    case Resolution1024x768 = 0x840004 /* '\x00\x84\x00\x04' */
    case Resolution1152x882 = 0x840005 /* '\x00\x84\x00\x05' */
    case Resolution1152x900 = 0x840006 /* '\x00\x84\x00\x06' */
    case Resolution1280x1024 = 0x840007 /* '\x00\x84\x00\x07' */
    case Resolution1600x1200 = 0x840008 /* '\x00\x84\x00\x08' */
    case Resolution1800x1440 = 0x840009 /* '\x00\x84\x00\t' */
    case Resolution1920x1200 = 0x84000a /* '\x00\x84\x00\n' */
}

// MARK: PowerPointMsoCharacterSet
@objc public enum PowerPointMsoCharacterSet : AEKeyword {
    case ArabicCharacterSet = 0x850001 /* '\x00\x85\x00\x01' */
    case CyrillicCharacterSet = 0x850002 /* '\x00\x85\x00\x02' */
    case EnglishCharacterSet = 0x850003 /* '\x00\x85\x00\x03' */
    case GreekCharacterSet = 0x850004 /* '\x00\x85\x00\x04' */
    case HebrewCharacterSet = 0x850005 /* '\x00\x85\x00\x05' */
    case JapaneseCharacterSet = 0x850006 /* '\x00\x85\x00\x06' */
    case KoreanCharacterSet = 0x850007 /* '\x00\x85\x00\x07' */
    case MultilingualUnicodeCharacterSet = 0x850008 /* '\x00\x85\x00\x08' */
    case SimplifiedChineseCharacterSet = 0x850009 /* '\x00\x85\x00\t' */
    case ThaiCharacterSet = 0x85000a /* '\x00\x85\x00\n' */
    case TraditionalChineseCharacterSet = 0x85000b /* '\x00\x85\x00\x0b' */
    case VietnameseCharacterSet = 0x85000c /* '\x00\x85\x00\x0c' */
}

// MARK: PowerPointMsoEncoding
@objc public enum PowerPointMsoEncoding : AEKeyword {
    case EncodingThai = 0x8b036a /* '\x00\x8b\x03j' */
    case EncodingJapaneseShiftJIS = 0x8b03a4 /* '\x00\x8b\x03\xa4' */
    case EncodingSimplifiedChinese = 0x8b03a8 /* '\x00\x8b\x03\xa8' */
    case EncodingKorean = 0x8b03b5 /* '\x00\x8b\x03\xb5' */
    case EncodingBig5TraditionalChinese = 0x8b03b6 /* '\x00\x8b\x03\xb6' */
    case EncodingLittleEndian = 0x8b04b0 /* '\x00\x8b\x04\xb0' */
    case EncodingBigEndian = 0x8b04b1 /* '\x00\x8b\x04\xb1' */
    case EncodingCentralEuropean = 0x8b04e2 /* '\x00\x8b\x04\xe2' */
    case EncodingCyrillic = 0x8b04e3 /* '\x00\x8b\x04\xe3' */
    case EncodingWestern = 0x8b04e4 /* '\x00\x8b\x04\xe4' */
    case EncodingGreek = 0x8b04e5 /* '\x00\x8b\x04\xe5' */
    case EncodingTurkish = 0x8b04e6 /* '\x00\x8b\x04\xe6' */
    case EncodingHebrew = 0x8b04e7 /* '\x00\x8b\x04\xe7' */
    case EncodingArabic = 0x8b04e8 /* '\x00\x8b\x04\xe8' */
    case EncodingBaltic = 0x8b04e9 /* '\x00\x8b\x04\xe9' */
    case EncodingVietnamese = 0x8b04ea /* '\x00\x8b\x04\xea' */
    case EncodingISO88591Latin1 = 0x8b6faf /* '\x00\x8bo\xaf' */
    case EncodingISO88592CentralEurope = 0x8b6fb0 /* '\x00\x8bo\xb0' */
    case EncodingISO88593Latin3 = 0x8b6fb1 /* '\x00\x8bo\xb1' */
    case EncodingISO88594Baltic = 0x8b6fb2 /* '\x00\x8bo\xb2' */
    case EncodingISO88595Cyrillic = 0x8b6fb3 /* '\x00\x8bo\xb3' */
    case EncodingISO88596Arabic = 0x8b6fb4 /* '\x00\x8bo\xb4' */
    case EncodingISO88597Greek = 0x8b6fb5 /* '\x00\x8bo\xb5' */
    case EncodingISO88598Hebrew = 0x8b6fb6 /* '\x00\x8bo\xb6' */
    case EncodingISO88599Turkish = 0x8b6fb7 /* '\x00\x8bo\xb7' */
    case EncodingISO885915Latin9 = 0x8b6fbd /* '\x00\x8bo\xbd' */
    case EncodingISO2022JapaneseNoHalfWidthKatakana = 0x8bc42c /* '\x00\x8b\xc4,' */
    case EncodingISO2022JapaneseJISX02021984 = 0x8bc42d /* '\x00\x8b\xc4-' */
    case EncodingISO2022JapaneseJISX02011989 = 0x8bc42e /* '\x00\x8b\xc4.' */
    case EncodingISO2022KR = 0x8bc431 /* '\x00\x8b\xc41' */
    case EncodingISO2022CNTraditionalChinese = 0x8bc433 /* '\x00\x8b\xc43' */
    case EncodingISO2022CNSimplifiedChinese = 0x8bc435 /* '\x00\x8b\xc45' */
    case EncodingMacRoman = 0x8b2710 /* "\x00\x8b'\x10" */
    case EncodingMacJapanese = 0x8b2711 /* "\x00\x8b'\x11" */
    case EncodingMacTraditionalChinese = 0x8b2712 /* "\x00\x8b'\x12" */
    case EncodingMacKorean = 0x8b2713 /* "\x00\x8b'\x13" */
    case EncodingMacArabic = 0x8b2714 /* "\x00\x8b'\x14" */
    case EncodingMacHebrew = 0x8b2715 /* "\x00\x8b'\x15" */
    case EncodingMacGreek1 = 0x8b2716 /* "\x00\x8b'\x16" */
    case EncodingMacCyrillic = 0x8b2717 /* "\x00\x8b'\x17" */
    case EncodingMacSimplifiedChineseGB2312 = 0x8b2718 /* "\x00\x8b'\x18" */
    case EncodingMacRomania = 0x8b271a /* "\x00\x8b'\x1a" */
    case EncodingMacUkraine = 0x8b2721 /* "\x00\x8b'!" */
    case EncodingMacLatin2 = 0x8b272d /* "\x00\x8b'-" */
    case EncodingMacIcelandic = 0x8b275f /* "\x00\x8b'_" */
    case EncodingMacTurkish = 0x8b2761 /* "\x00\x8b'a" */
    case EncodingMacCroatia = 0x8b2762 /* "\x00\x8b'b" */
    case EncodingEBCDICUSCanada = 0x8b0025 /* '\x00\x8b\x00%' */
    case EncodingEBCDICInternational = 0x8b01f4 /* '\x00\x8b\x01\xf4' */
    case EncodingEBCDICMultilingualROECELatin2 = 0x8b0366 /* '\x00\x8b\x03f' */
    case EncodingEBCDICGreekModern = 0x8b036b /* '\x00\x8b\x03k' */
    case EncodingEBCDICTurkishLatin5 = 0x8b0402 /* '\x00\x8b\x04\x02' */
    case EncodingEBCDICGermany = 0x8b4f31 /* '\x00\x8bO1' */
    case EncodingEBCDICDenmarkNorway = 0x8b4f35 /* '\x00\x8bO5' */
    case EncodingEBCDICFinlandSweden = 0x8b4f36 /* '\x00\x8bO6' */
    case EncodingEBCDICItaly = 0x8b4f38 /* '\x00\x8bO8' */
    case EncodingEBCDICLatinAmericaSpain = 0x8b4f3c /* '\x00\x8bO<' */
    case EncodingEBCDICUnitedKingdom = 0x8b4f3d /* '\x00\x8bO=' */
    case EncodingEBCDICJapaneseKatakanaExtended = 0x8b4f42 /* '\x00\x8bOB' */
    case EncodingEBCDICFrance = 0x8b4f49 /* '\x00\x8bOI' */
    case EncodingEBCDICArabic = 0x8b4fc4 /* '\x00\x8bO\xc4' */
    case EncodingEBCDICGreek = 0x8b4fc7 /* '\x00\x8bO\xc7' */
    case EncodingEBCDICHebrew = 0x8b4fc8 /* '\x00\x8bO\xc8' */
    case EncodingEBCDICKoreanExtended = 0x8b5161 /* '\x00\x8bQa' */
    case EncodingEBCDICThai = 0x8b5166 /* '\x00\x8bQf' */
    case EncodingEBCDICIcelandic = 0x8b5187 /* '\x00\x8bQ\x87' */
    case EncodingEBCDICTurkish = 0x8b51a9 /* '\x00\x8bQ\xa9' */
    case EncodingEBCDICRussian = 0x8b5190 /* '\x00\x8bQ\x90' */
    case EncodingEBCDICSerbianBulgarian = 0x8b5221 /* '\x00\x8bR!' */
    case EncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 0x8bc6f2 /* '\x00\x8b\xc6\xf2' */
    case EncodingEBCDICUSCanadaAndJapanese = 0x8bc6f3 /* '\x00\x8b\xc6\xf3' */
    case EncodingEBCDICExtendedAndKorean = 0x8bc6f5 /* '\x00\x8b\xc6\xf5' */
    case EncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 0x8bc6f7 /* '\x00\x8b\xc6\xf7' */
    case EncodingEBCDICUSCanadaAndTraditionalChinese = 0x8bc6f9 /* '\x00\x8b\xc6\xf9' */
    case EncodingEBCDICJapaneseLatinExtendedAndJapanese = 0x8bc6fb /* '\x00\x8b\xc6\xfb' */
    case EncodingOEMUnitedStates = 0x8b01b5 /* '\x00\x8b\x01\xb5' */
    case EncodingOEMGreek = 0x8b02e1 /* '\x00\x8b\x02\xe1' */
    case EncodingOEMBaltic = 0x8b0307 /* '\x00\x8b\x03\x07' */
    case EncodingOEMMultilingualLatinI = 0x8b0352 /* '\x00\x8b\x03R' */
    case EncodingOEMMultilingualLatinII = 0x8b0354 /* '\x00\x8b\x03T' */
    case EncodingOEMCyrillic = 0x8b0357 /* '\x00\x8b\x03W' */
    case EncodingOEMTurkish = 0x8b0359 /* '\x00\x8b\x03Y' */
    case EncodingOEMPortuguese = 0x8b035c /* '\x00\x8b\x03\\' */
    case EncodingOEMIcelandic = 0x8b035d /* '\x00\x8b\x03]' */
    case EncodingOEMHebrew = 0x8b035e /* '\x00\x8b\x03^' */
    case EncodingOEMCanadianFrench = 0x8b035f /* '\x00\x8b\x03_' */
    case EncodingOEMArabic = 0x8b0360 /* '\x00\x8b\x03`' */
    case EncodingOEMNordic = 0x8b0361 /* '\x00\x8b\x03a' */
    case EncodingOEMCyrillicII = 0x8b0362 /* '\x00\x8b\x03b' */
    case EncodingOEMModernGreek = 0x8b0365 /* '\x00\x8b\x03e' */
    case EncodingEUCJapanese = 0x8bcadc /* '\x00\x8b\xca\xdc' */
    case EncodingEUCChineseSimplifiedChinese = 0x8bcae0 /* '\x00\x8b\xca\xe0' */
    case EncodingEUCKorean = 0x8bcaed /* '\x00\x8b\xca\xed' */
    case EncodingEUCTaiwaneseTraditionalChinese = 0x8bcaee /* '\x00\x8b\xca\xee' */
    case EncodingDevanagari = 0x8bdeaa /* '\x00\x8b\xde\xaa' */
    case EncodingBengali = 0x8bdeab /* '\x00\x8b\xde\xab' */
    case EncodingTamil = 0x8bdeac /* '\x00\x8b\xde\xac' */
    case EncodingTelugu = 0x8bdead /* '\x00\x8b\xde\xad' */
    case EncodingAssamese = 0x8bdeae /* '\x00\x8b\xde\xae' */
    case EncodingOriya = 0x8bdeaf /* '\x00\x8b\xde\xaf' */
    case EncodingKannada = 0x8bdeb0 /* '\x00\x8b\xde\xb0' */
    case EncodingMalayalam = 0x8bdeb1 /* '\x00\x8b\xde\xb1' */
    case EncodingGujarati = 0x8bdeb2 /* '\x00\x8b\xde\xb2' */
    case EncodingPunjabi = 0x8bdeb3 /* '\x00\x8b\xde\xb3' */
    case EncodingArabicASMO = 0x8b02c4 /* '\x00\x8b\x02\xc4' */
    case EncodingArabicTransparentASMO = 0x8b02d0 /* '\x00\x8b\x02\xd0' */
    case EncodingKoreanJohab = 0x8b0551 /* '\x00\x8b\x05Q' */
    case EncodingTaiwanCNS = 0x8b4e20 /* '\x00\x8bN ' */
    case EncodingTaiwanTCA = 0x8b4e21 /* '\x00\x8bN!' */
    case EncodingTaiwanEten = 0x8b4e22 /* '\x00\x8bN"' */
    case EncodingTaiwanIBM5550 = 0x8b4e23 /* '\x00\x8bN#' */
    case EncodingTaiwanTeletext = 0x8b4e24 /* '\x00\x8bN$' */
    case EncodingTaiwanWang = 0x8b4e25 /* '\x00\x8bN%' */
    case EncodingIA5IRV = 0x8b4e89 /* '\x00\x8bN\x89' */
    case EncodingIA5German = 0x8b4e8a /* '\x00\x8bN\x8a' */
    case EncodingIA5Swedish = 0x8b4e8b /* '\x00\x8bN\x8b' */
    case EncodingIA5Norwegian = 0x8b4e8c /* '\x00\x8bN\x8c' */
    case EncodingUSASCII = 0x8b4e9f /* '\x00\x8bN\x9f' */
    case EncodingT61 = 0x8b4f25 /* '\x00\x8bO%' */
    case EncodingISO6937NonspacingAccent = 0x8b4f2d /* '\x00\x8bO-' */
    case EncodingKOI8R = 0x8b5182 /* '\x00\x8bQ\x82' */
    case EncodingExtAlphaLowercase = 0x8b5223 /* '\x00\x8bR#' */
    case EncodingKOI8U = 0x8b556a /* '\x00\x8bUj' */
    case EncodingEuropa3 = 0x8b7149 /* '\x00\x8bqI' */
    case EncodingHZGBSimplifiedChinese = 0x8bcec8 /* '\x00\x8b\xce\xc8' */
    case EncodingUTF7 = 0x8bfde8 /* '\x00\x8b\xfd\xe8' */
    case EncodingUTF8 = 0x8bfde9 /* '\x00\x8b\xfd\xe9' */
}

// MARK: PowerPointReset
@objc public enum PowerPointReset : AEKeyword {
    case CommandBar = 0x6d734342 /* 'msCB' */
    case CommandBarControl = 0x6d434243 /* 'mCBC' */
}

// MARK: PowerPointEPPWindowState
@objc public enum PowerPointEPPWindowState : AEKeyword {
    case WindowNormal = 0xc90001 /* '\x00\xc9\x00\x01' */
    case WindowMinimized = 0xc90002 /* '\x00\xc9\x00\x02' */
}

// MARK: PowerPointEPPArrangeStyle
@objc public enum PowerPointEPPArrangeStyle : AEKeyword {
    case ArrangeTiled = 0xd10001 /* '\x00\xd1\x00\x01' */
    case ArrangeCascade = 0xd10002 /* '\x00\xd1\x00\x02' */
}

// MARK: PowerPointEPPViewType
@objc public enum PowerPointEPPViewType : AEKeyword {
    case SlideView = 0xca0001 /* '\x00\xca\x00\x01' */
    case MasterView = 0xca0002 /* '\x00\xca\x00\x02' */
    case PageView = 0xca0003 /* '\x00\xca\x00\x03' */
    case HandoutMasterView = 0xca0004 /* '\x00\xca\x00\x04' */
    case NotesMasterView = 0xca0005 /* '\x00\xca\x00\x05' */
    case OutlineView = 0xca0006 /* '\x00\xca\x00\x06' */
    case SlideSorterView = 0xca0007 /* '\x00\xca\x00\x07' */
    case TitleMasterView = 0xca0008 /* '\x00\xca\x00\x08' */
    case NormalView = 0xca0009 /* '\x00\xca\x00\t' */
    case PrintPreview = 0xca000a /* '\x00\xca\x00\n' */
    case ThumbnailView = 0xca000b /* '\x00\xca\x00\x0b' */
    case ThumbnailMasterView = 0xca000c /* '\x00\xca\x00\x0c' */
}

// MARK: PowerPointEPPColorSchemeIndex
@objc public enum PowerPointEPPColorSchemeIndex : AEKeyword {
    case SchemeColorUnset = 0xf2fffe /* '\x00\xf2\xff\xfe' */
    case NotASchemeColor = 0xf30000 /* '\x00\xf3\x00\x00' */
    case BackgroundScheme = 0xf30001 /* '\x00\xf3\x00\x01' */
    case ForegroundScheme = 0xf30002 /* '\x00\xf3\x00\x02' */
    case ShadowScheme = 0xf30003 /* '\x00\xf3\x00\x03' */
    case TitleScheme = 0xf30004 /* '\x00\xf3\x00\x04' */
    case FillScheme = 0xf30005 /* '\x00\xf3\x00\x05' */
    case Accent1Scheme = 0xf30006 /* '\x00\xf3\x00\x06' */
    case Accent2Scheme = 0xf30007 /* '\x00\xf3\x00\x07' */
    case Accent3Scheme = 0xf30008 /* '\x00\xf3\x00\x08' */
}

// MARK: PowerPointEPPSlideSizeType
@objc public enum PowerPointEPPSlideSizeType : AEKeyword {
    case SlideSizeOnScreen = 0xcb0001 /* '\x00\xcb\x00\x01' */
    case SlideSizeLetterPaper = 0xcb0002 /* '\x00\xcb\x00\x02' */
    case SlideSizeA4Paper = 0xcb0003 /* '\x00\xcb\x00\x03' */
    case SlideSize35MM = 0xcb0004 /* '\x00\xcb\x00\x04' */
    case SlideSizeOverhead = 0xcb0005 /* '\x00\xcb\x00\x05' */
    case SlideSizeBanner = 0xcb0006 /* '\x00\xcb\x00\x06' */
    case SlideSizeCustom = 0xcb0007 /* '\x00\xcb\x00\x07' */
    case SlideSizeLedgerPaper = 0xcb0008 /* '\x00\xcb\x00\x08' */
    case SlideSizeA3Paper = 0xcb0009 /* '\x00\xcb\x00\t' */
    case SlideSizeB4ISOPaper = 0xcb000a /* '\x00\xcb\x00\n' */
    case SlideSizeB5ISOPaper = 0xcb000b /* '\x00\xcb\x00\x0b' */
    case SlideSizeB4JISPaper = 0xcb000c /* '\x00\xcb\x00\x0c' */
    case SlideSizeB5JISPaper = 0xcb000d /* '\x00\xcb\x00\r' */
    case SlideSizeHagakiCard = 0xcb000e /* '\x00\xcb\x00\x0e' */
    case SlideSizeOnScreen16x9 = 0xcb000f /* '\x00\xcb\x00\x0f' */
    case SlideSizeOnScreen16x10 = 0xcb0010 /* '\x00\xcb\x00\x10' */
}

// MARK: PowerPointEPPSaveAsFileType
@objc public enum PowerPointEPPSaveAsFileType : AEKeyword {
    case SaveAsPresentation = 0xcc0001 /* '\x00\xcc\x00\x01' */
    case SaveAsTemplate = 0xcc0005 /* '\x00\xcc\x00\x05' */
    case SaveAsRTF = 0xcc0006 /* '\x00\xcc\x00\x06' */
    case SaveAsShow = 0xcc0007 /* '\x00\xcc\x00\x07' */
    case SaveAsAddIn = 0xcc0008 /* '\x00\xcc\x00\x08' */
    case SaveAsDefault = 0xcc000a /* '\x00\xcc\x00\n' */
    case SaveAsHTML = 0xcc000b /* '\x00\xcc\x00\x0b' */
    case SaveAsMovie = 0xcc000c /* '\x00\xcc\x00\x0c' */
    case SaveAsPackage = 0xcc000d /* '\x00\xcc\x00\r' */
    case SaveAsPDF = 0xcc000e /* '\x00\xcc\x00\x0e' */
    case SaveAsOpenXMLPresentation = 0xcc000f /* '\x00\xcc\x00\x0f' */
    case SaveAsOpenXMLPresentationMacroEnabled = 0xcc0010 /* '\x00\xcc\x00\x10' */
    case SaveAsOpenXMLShow = 0xcc0011 /* '\x00\xcc\x00\x11' */
    case SaveAsOpenXMLShowMacroEnabled = 0xcc0012 /* '\x00\xcc\x00\x12' */
    case SaveAsOpenXMLTemplate = 0xcc0013 /* '\x00\xcc\x00\x13' */
    case SaveAsOpenXMLTemplateMacroEnabled = 0xcc0014 /* '\x00\xcc\x00\x14' */
    case SaveAsOpenXMLTheme = 0xcc0015 /* '\x00\xcc\x00\x15' */
    case SaveAsGIF = 0xcc0016 /* '\x00\xcc\x00\x16' */
    case SaveAsJPG = 0xcc0017 /* '\x00\xcc\x00\x17' */
    case SaveAsPNG = 0xcc0018 /* '\x00\xcc\x00\x18' */
    case SaveAsBMP = 0xcc0019 /* '\x00\xcc\x00\x19' */
    case SaveAsTIF = 0xcc001a /* '\x00\xcc\x00\x1a' */
}

// MARK: PowerPointPpTextStyleType
@objc public enum PowerPointPpTextStyleType : AEKeyword {
    case TextStyleDefault = 0x12a0001 /* '\x01*\x00\x01' */
    case TextStyleTitle = 0x12a0002 /* '\x01*\x00\x02' */
    case TextStyleBody = 0x12a0003 /* '\x01*\x00\x03' */
}

// MARK: PowerPointEPPSlideLayout
@objc public enum PowerPointEPPSlideLayout : AEKeyword {
    case SlideLayoutUnset = 0xccfffe /* '\x00\xcc\xff\xfe' */
    case SlideLayoutTitleSlide = 0xcd0001 /* '\x00\xcd\x00\x01' */
    case SlideLayoutTextSlide = 0xcd0002 /* '\x00\xcd\x00\x02' */
    case SlideLayoutTwoColumnText = 0xcd0003 /* '\x00\xcd\x00\x03' */
    case SlideLayoutTable = 0xcd0004 /* '\x00\xcd\x00\x04' */
    case SlideLayoutTextAndChart = 0xcd0005 /* '\x00\xcd\x00\x05' */
    case SlideLayoutChartAndText = 0xcd0006 /* '\x00\xcd\x00\x06' */
    case SlideLayoutOrgchart = 0xcd0007 /* '\x00\xcd\x00\x07' */
    case SlideLayoutChart = 0xcd0008 /* '\x00\xcd\x00\x08' */
    case SlideLayoutTextAndClipart = 0xcd0009 /* '\x00\xcd\x00\t' */
    case SlideLayoutClipartAndText = 0xcd000a /* '\x00\xcd\x00\n' */
    case SlideLayoutTitleOnly = 0xcd000b /* '\x00\xcd\x00\x0b' */
    case SlideLayoutBlank = 0xcd000c /* '\x00\xcd\x00\x0c' */
    case SlideLayoutTextAndObject = 0xcd000d /* '\x00\xcd\x00\r' */
    case SlideLayoutObjectAndText = 0xcd000e /* '\x00\xcd\x00\x0e' */
    case SlideLayoutLargeObject = 0xcd000f /* '\x00\xcd\x00\x0f' */
    case SlideLayoutObject = 0xcd0010 /* '\x00\xcd\x00\x10' */
    case SlideLayoutMediaClip = 0xcd0011 /* '\x00\xcd\x00\x11' */
    case SlideLayoutMediaClipAndText = 0xcd0012 /* '\x00\xcd\x00\x12' */
    case SlideLayoutObjectOverText = 0xcd0013 /* '\x00\xcd\x00\x13' */
    case SlideLayoutTextOverObject = 0xcd0014 /* '\x00\xcd\x00\x14' */
    case SlideLayoutTextAndTwoObjects = 0xcd0015 /* '\x00\xcd\x00\x15' */
    case SlideLayoutTwoObjectsAndText = 0xcd0016 /* '\x00\xcd\x00\x16' */
    case SlideLayoutTwoObjectsOverText = 0xcd0017 /* '\x00\xcd\x00\x17' */
    case SlideLayoutFourObjects = 0xcd0018 /* '\x00\xcd\x00\x18' */
    case SlideLayoutVerticalText = 0xcd0019 /* '\x00\xcd\x00\x19' */
    case SlideLayoutClipartAndVerticalText = 0xcd001a /* '\x00\xcd\x00\x1a' */
    case SlideLayoutVerticalTitleAndText = 0xcd001b /* '\x00\xcd\x00\x1b' */
    case SlideLayoutVerticalTitleAndTextOverChart = 0xcd001c /* '\x00\xcd\x00\x1c' */
    case SlideLayoutTwoObjects = 0xcd001d /* '\x00\xcd\x00\x1d' */
    case SlideLayoutObjectAndTwoObjects = 0xcd001e /* '\x00\xcd\x00\x1e' */
    case SlideLayoutTwoObjectsAndObject = 0xcd001f /* '\x00\xcd\x00\x1f' */
    case SlideLayoutCustom = 0xcd0020 /* '\x00\xcd\x00 ' */
    case SlideLayoutSectionHeader = 0xcd0021 /* '\x00\xcd\x00!' */
    case SlideLayoutComparison = 0xcd0022 /* '\x00\xcd\x00"' */
    case SlideLayoutContentWithCaption = 0xcd0023 /* '\x00\xcd\x00#' */
    case SlideLayoutPictureWithCaption = 0xcd0024 /* '\x00\xcd\x00$' */
}

// MARK: PowerPointEPPEntryEffect
@objc public enum PowerPointEPPEntryEffect : AEKeyword {
    case EntryEffectAppear = 0xf60f04 /* '\x00\xf6\x0f\x04' */
    case EntryEffectBlindsHorizontal = 0xf60301 /* '\x00\xf6\x03\x01' */
    case EntryEffectBlindsVertical = 0xf60302 /* '\x00\xf6\x03\x02' */
    case EntryEffectBoxDown = 0xf60f55 /* '\x00\xf6\x0fU' */
    case EntryEffectBoxIn = 0xf60c02 /* '\x00\xf6\x0c\x02' */
    case EntryEffectBoxLeft = 0xf60f52 /* '\x00\xf6\x0fR' */
    case EntryEffectBoxOut = 0xf60c01 /* '\x00\xf6\x0c\x01' */
    case EntryEffectBoxRight = 0xf60f54 /* '\x00\xf6\x0fT' */
    case EntryEffectBoxUp = 0xf60f53 /* '\x00\xf6\x0fS' */
    case EntryEffectCheckerboardAcross = 0xf60401 /* '\x00\xf6\x04\x01' */
    case EntryEffectCheckerboardDown = 0xf60402 /* '\x00\xf6\x04\x02' */
    case EntryEffectCircle = 0xf60f05 /* '\x00\xf6\x0f\x05' */
    case EntryEffectCollapseAcross = 0xf60d17 /* '\x00\xf6\r\x17' */
    case EntryEffectCollapseBottom = 0xf60d1b /* '\x00\xf6\r\x1b' */
    case EntryEffectCollapseLeft = 0xf60d18 /* '\x00\xf6\r\x18' */
    case EntryEffectCollapseRight = 0xf60d1a /* '\x00\xf6\r\x1a' */
    case EntryEffectCollapseUp = 0xf60d19 /* '\x00\xf6\r\x19' */
    case EntryEffectCombHorizontal = 0xf60f07 /* '\x00\xf6\x0f\x07' */
    case EntryEffectCombVertical = 0xf60f08 /* '\x00\xf6\x0f\x08' */
    case EntryEffectConveyorLeft = 0xf60f2a /* '\x00\xf6\x0f*' */
    case EntryEffectConveyorRight = 0xf60f2b /* '\x00\xf6\x0f+' */
    case EntryEffectCoverDown = 0xf60504 /* '\x00\xf6\x05\x04' */
    case EntryEffectCoverLeftDown = 0xf60507 /* '\x00\xf6\x05\x07' */
    case EntryEffectCoverLeftUp = 0xf60505 /* '\x00\xf6\x05\x05' */
    case EntryEffectCoverLeft = 0xf60501 /* '\x00\xf6\x05\x01' */
    case EntryEffectCoverRightDown = 0xf60508 /* '\x00\xf6\x05\x08' */
    case EntryEffectCoverRightUp = 0xf60506 /* '\x00\xf6\x05\x06' */
    case EntryEffectCoverRight = 0xf60503 /* '\x00\xf6\x05\x03' */
    case EntryEffectCoverUp = 0xf60502 /* '\x00\xf6\x05\x02' */
    case EntryEffectCrawlFromDown = 0xf60d10 /* '\x00\xf6\r\x10' */
    case EntryEffectCrawlFromLeft = 0xf60d0d /* '\x00\xf6\r\r' */
    case EntryEffectCrawlFromRight = 0xf60d0f /* '\x00\xf6\r\x0f' */
    case EntryEffectCrawlFromUp = 0xf60d0e /* '\x00\xf6\r\x0e' */
    case EntryEffectCubeDown = 0xf60f4d /* '\x00\xf6\x0fM' */
    case EntryEffectCubeLeft = 0xf60f4a /* '\x00\xf6\x0fJ' */
    case EntryEffectCubeRight = 0xf60f4b /* '\x00\xf6\x0fK' */
    case EntryEffectCubeUp = 0xf60f4c /* '\x00\xf6\x0fL' */
    case EntryEffectCutBlack = 0xf60102 /* '\x00\xf6\x01\x02' */
    case EntryEffectCut = 0xf60101 /* '\x00\xf6\x01\x01' */
    case EntryEffectDiamond = 0xf60f06 /* '\x00\xf6\x0f\x06' */
    case EntryEffectDissolve = 0xf60601 /* '\x00\xf6\x06\x01' */
    case EntryEffectDoorsHorizontal = 0xf60f2d /* '\x00\xf6\x0f-' */
    case EntryEffectDoorsVertical = 0xf60f2c /* '\x00\xf6\x0f,' */
    case EntryEffectFadeBlack = 0xf60701 /* '\x00\xf6\x07\x01' */
    case EntryEffectFadeFlyFromBottomLeft = 0xf60d24 /* '\x00\xf6\r$' */
    case EntryEffectFadeFlyFromBottomRight = 0xf60d25 /* '\x00\xf6\r%' */
    case EntryEffectFadeFlyFromBottom = 0xf60d21 /* '\x00\xf6\r!' */
    case EntryEffectFadeFlyFromLeft = 0xf60d1e /* '\x00\xf6\r\x1e' */
    case EntryEffectFadeFlyFromRight = 0xf60d20 /* '\x00\xf6\r ' */
    case EntryEffectFadeFlyFromTopLeft = 0xf60d22 /* '\x00\xf6\r"' */
    case EntryEffectFadeFlyFromTopRight = 0xf60d23 /* '\x00\xf6\r#' */
    case EntryEffectFadeFlyFromTop = 0xf60d1f /* '\x00\xf6\r\x1f' */
    case EntryEffectFadeSmoothly = 0xf60f09 /* '\x00\xf6\x0f\t' */
    case EntryEffectFade = 0xf60f09 /* '\x00\xf6\x0f\t' */
    case EntryEffectFerrisWheelLeft = 0xf60f3b /* '\x00\xf6\x0f;' */
    case EntryEffectFerrisWheelRight = 0xf60f3c /* '\x00\xf6\x0f<' */
    case EntryEffectFlashOnceFast = 0xf60f01 /* '\x00\xf6\x0f\x01' */
    case EntryEffectFlashOnceMedium = 0xf60f02 /* '\x00\xf6\x0f\x02' */
    case EntryEffectFlashOnceSlow = 0xf60f03 /* '\x00\xf6\x0f\x03' */
    case EntryEffectFlashbulb = 0xf60f45 /* '\x00\xf6\x0fE' */
    case EntryEffectFlipDown = 0xf60f44 /* '\x00\xf6\x0fD' */
    case EntryEffectFlipLeft = 0xf60f41 /* '\x00\xf6\x0fA' */
    case EntryEffectFlipRight = 0xf60f42 /* '\x00\xf6\x0fB' */
    case EntryEffectFlipUp = 0xf60f43 /* '\x00\xf6\x0fC' */
    case EntryEffectFlyFromBottomLeft = 0xf60d07 /* '\x00\xf6\r\x07' */
    case EntryEffectFlyFromBottomRight = 0xf60d08 /* '\x00\xf6\r\x08' */
    case EntryEffectFlyFromBottom = 0xf60d04 /* '\x00\xf6\r\x04' */
    case EntryEffectFlyFromLeft = 0xf60d01 /* '\x00\xf6\r\x01' */
    case EntryEffectFlyFromRight = 0xf60d03 /* '\x00\xf6\r\x03' */
    case EntryEffectFlyFromTopLeft = 0xf60d05 /* '\x00\xf6\r\x05' */
    case EntryEffectFlyFromTopRight = 0xf60d06 /* '\x00\xf6\r\x06' */
    case EntryEffectFlyFromTop = 0xf60d02 /* '\x00\xf6\r\x02' */
    case EntryEffectFlyThroughInBounce = 0xf60f34 /* '\x00\xf6\x0f4' */
    case EntryEffectFlyThroughIn = 0xf60f32 /* '\x00\xf6\x0f2' */
    case EntryEffectFlyThroughOutBounce = 0xf60f35 /* '\x00\xf6\x0f5' */
    case EntryEffectFlyThroughOut = 0xf60f33 /* '\x00\xf6\x0f3' */
    case EntryEffectGalleryLeft = 0xf60f28 /* '\x00\xf6\x0f(' */
    case EntryEffectGalleryRight = 0xf60f29 /* '\x00\xf6\x0f)' */
    case EntryEffectGlitterDiamondDown = 0xf60f23 /* '\x00\xf6\x0f#' */
    case EntryEffectGlitterDiamondLeft = 0xf60f20 /* '\x00\xf6\x0f ' */
    case EntryEffectGlitterDiamondRight = 0xf60f22 /* '\x00\xf6\x0f"' */
    case EntryEffectGlitterDiamondUp = 0xf60f21 /* '\x00\xf6\x0f!' */
    case EntryEffectGlitterHexagonDown = 0xf60f27 /* "\x00\xf6\x0f'" */
    case EntryEffectGlitterHexagonLeft = 0xf60f24 /* '\x00\xf6\x0f$' */
    case EntryEffectGlitterHexagonRight = 0xf60f26 /* '\x00\xf6\x0f&' */
    case EntryEffectGlitterHexagonUp = 0xf60f25 /* '\x00\xf6\x0f%' */
    case EntryEffectHoneycomb = 0xf60f3a /* '\x00\xf6\x0f:' */
    case EntryEffectNewsFlash = 0xf60f0a /* '\x00\xf6\x0f\n' */
    case EntryEffectNone = 0xf60000 /* '\x00\xf6\x00\x00' */
    case EntryEffectOrbitDown = 0xf60f59 /* '\x00\xf6\x0fY' */
    case EntryEffectOrbitLeft = 0xf60f56 /* '\x00\xf6\x0fV' */
    case EntryEffectOrbitRight = 0xf60f58 /* '\x00\xf6\x0fX' */
    case EntryEffectOrbitUp = 0xf60f57 /* '\x00\xf6\x0fW' */
    case EntryEffectPanDown = 0xf60f5d /* '\x00\xf6\x0f]' */
    case EntryEffectPanLeft = 0xf60f5a /* '\x00\xf6\x0fZ' */
    case EntryEffectPanRight = 0xf60f5c /* '\x00\xf6\x0f\\' */
    case EntryEffectPanUp = 0xf60f5b /* '\x00\xf6\x0f[' */
    case EntryEffectPeekFromDown = 0xf60d0a /* '\x00\xf6\r\n' */
    case EntryEffectPeekFromLeft = 0xf60d09 /* '\x00\xf6\r\t' */
    case EntryEffectPeekFromRight = 0xf60d0b /* '\x00\xf6\r\x0b' */
    case EntryEffectPeekFromUp = 0xf60d0c /* '\x00\xf6\r\x0c' */
    case EntryEffectPlus = 0xf60f0b /* '\x00\xf6\x0f\x0b' */
    case EntryEffectPushDown = 0xf60f0c /* '\x00\xf6\x0f\x0c' */
    case EntryEffectPushLeft = 0xf60f0d /* '\x00\xf6\x0f\r' */
    case EntryEffectPushRight = 0xf60f0e /* '\x00\xf6\x0f\x0e' */
    case EntryEffectPushUp = 0xf60f0f /* '\x00\xf6\x0f\x0f' */
    case EntryEffectRandomBarsHorizontal = 0xf60901 /* '\x00\xf6\t\x01' */
    case EntryEffectRandomBarsVertical = 0xf60902 /* '\x00\xf6\t\x02' */
    case EntryEffectRandom = 0xf60201 /* '\x00\xf6\x02\x01' */
    case EntryEffectRevealBlackLeft = 0xf60f38 /* '\x00\xf6\x0f8' */
    case EntryEffectRevealBlackRight = 0xf60f39 /* '\x00\xf6\x0f9' */
    case EntryEffectRevealSmoothLeft = 0xf60f36 /* '\x00\xf6\x0f6' */
    case EntryEffectRevealSmoothRight = 0xf60f37 /* '\x00\xf6\x0f7' */
    case EntryEffectRippleCenter = 0xf60f1b /* '\x00\xf6\x0f\x1b' */
    case EntryEffectRippleLeftDown = 0xf60f1e /* '\x00\xf6\x0f\x1e' */
    case EntryEffectRippleLeftUp = 0xf60f1d /* '\x00\xf6\x0f\x1d' */
    case EntryEffectRippleRightDown = 0xf60f1f /* '\x00\xf6\x0f\x1f' */
    case EntryEffectRippleRightUp = 0xf60f1c /* '\x00\xf6\x0f\x1c' */
    case EntryEffectRotateDown = 0xf60f51 /* '\x00\xf6\x0fQ' */
    case EntryEffectRotateLeft = 0xf60f4e /* '\x00\xf6\x0fN' */
    case EntryEffectRotateRight = 0xf60f50 /* '\x00\xf6\x0fP' */
    case EntryEffectRotateUp = 0xf60f4f /* '\x00\xf6\x0fO' */
    case EntryEffectShredRectangleIn = 0xf60f48 /* '\x00\xf6\x0fH' */
    case EntryEffectShredRectangleOut = 0xf60f49 /* '\x00\xf6\x0fI' */
    case EntryEffectShredStripsIn = 0xf60f46 /* '\x00\xf6\x0fF' */
    case EntryEffectShredStripsOut = 0xf60f47 /* '\x00\xf6\x0fG' */
    case EntryEffectSpinner = 0xf60f0a /* '\x00\xf6\x0f\n' */
    case EntryEffectSpiral = 0xf60d1d /* '\x00\xf6\r\x1d' */
    case EntryEffectSplitHorizontalIn = 0xf60e02 /* '\x00\xf6\x0e\x02' */
    case EntryEffectSplitHorizontalOut = 0xf60e01 /* '\x00\xf6\x0e\x01' */
    case EntryEffectSplitVerticalIn = 0xf60e04 /* '\x00\xf6\x0e\x04' */
    case EntryEffectSplitVerticalOut = 0xf60e03 /* '\x00\xf6\x0e\x03' */
    case EntryEffectStretchAcross = 0xf60d17 /* '\x00\xf6\r\x17' */
    case EntryEffectStretchDown = 0xf60d1b /* '\x00\xf6\r\x1b' */
    case EntryEffectStretchLeft = 0xf60d18 /* '\x00\xf6\r\x18' */
    case EntryEffectStretchRight = 0xf60d1a /* '\x00\xf6\r\x1a' */
    case EntryEffectStretchUp = 0xf60d19 /* '\x00\xf6\r\x19' */
    case EntryEffectStripsDownLeft = 0xf60a03 /* '\x00\xf6\n\x03' */
    case EntryEffectStripsDownRight = 0xf60a04 /* '\x00\xf6\n\x04' */
    case EntryEffectStripsLeftDown = 0xf60a07 /* '\x00\xf6\n\x07' */
    case EntryEffectStripsLeftUp = 0xf60a05 /* '\x00\xf6\n\x05' */
    case EntryEffectStripsRightDown = 0xf60a08 /* '\x00\xf6\n\x08' */
    case EntryEffectStripsRightUp = 0xf60a06 /* '\x00\xf6\n\x06' */
    case EntryEffectStripsUpLeft = 0xf60a01 /* '\x00\xf6\n\x01' */
    case EntryEffectStripsUpRight = 0xf60a02 /* '\x00\xf6\n\x02' */
    case EntryEffectSwitchDown = 0xf60f40 /* '\x00\xf6\x0f@' */
    case EntryEffectSwitchLeft = 0xf60f3d /* '\x00\xf6\x0f=' */
    case EntryEffectSwitchRight = 0xf60f3f /* '\x00\xf6\x0f?' */
    case EntryEffectSwitchUp = 0xf60f3e /* '\x00\xf6\x0f>' */
    case EntryEffectSwivel = 0xf60d1c /* '\x00\xf6\r\x1c' */
    case EntryEffectUncoverDown = 0xf60804 /* '\x00\xf6\x08\x04' */
    case EntryEffectUncoverLeftDown = 0xf60807 /* '\x00\xf6\x08\x07' */
    case EntryEffectUncoverLeftUp = 0xf60805 /* '\x00\xf6\x08\x05' */
    case EntryEffectUncoverLeft = 0xf60801 /* '\x00\xf6\x08\x01' */
    case EntryEffectUncoverRightDown = 0xf60808 /* '\x00\xf6\x08\x08' */
    case EntryEffectUncoverRightUp = 0xf60806 /* '\x00\xf6\x08\x06' */
    case EntryEffectUncoverRight = 0xf60803 /* '\x00\xf6\x08\x03' */
    case EntryEffectUncoverUp = 0xf60802 /* '\x00\xf6\x08\x02' */
    case EntryEffectUnset = 0xf5fffe /* '\x00\xf5\xff\xfe' */
    case EntryEffectVortexDown = 0xf60f1a /* '\x00\xf6\x0f\x1a' */
    case EntryEffectVortexLeft = 0xf60f17 /* '\x00\xf6\x0f\x17' */
    case EntryEffectVortexRight = 0xf60f19 /* '\x00\xf6\x0f\x19' */
    case EntryEffectVortexUp = 0xf60f18 /* '\x00\xf6\x0f\x18' */
    case EntryEffectWarpIn = 0xf60f30 /* '\x00\xf6\x0f0' */
    case EntryEffectWarpOut = 0xf60f31 /* '\x00\xf6\x0f1' */
    case EntryEffectWedge = 0xf60f10 /* '\x00\xf6\x0f\x10' */
    case EntryEffectWheel1Spoke = 0xf60f11 /* '\x00\xf6\x0f\x11' */
    case EntryEffectWheel2Spokes = 0xf60f12 /* '\x00\xf6\x0f\x12' */
    case EntryEffectWheel3Spokes = 0xf60f13 /* '\x00\xf6\x0f\x13' */
    case EntryEffectWheel4Spokes = 0xf60f14 /* '\x00\xf6\x0f\x14' */
    case EntryEffectWheel8Spokes = 0xf60f15 /* '\x00\xf6\x0f\x15' */
    case EntryEffectWheelReverse1Spoke = 0xf60f16 /* '\x00\xf6\x0f\x16' */
    case EntryEffectWindowHorizontal = 0xf60f2f /* '\x00\xf6\x0f/' */
    case EntryEffectWindowVertical = 0xf60f2e /* '\x00\xf6\x0f.' */
    case EntryEffectWipeDown = 0xf60b04 /* '\x00\xf6\x0b\x04' */
    case EntryEffectWipeLeft = 0xf60b01 /* '\x00\xf6\x0b\x01' */
    case EntryEffectWipeRight = 0xf60b03 /* '\x00\xf6\x0b\x03' */
    case EntryEffectWipeUp = 0xf60b02 /* '\x00\xf6\x0b\x02' */
    case EntryEffectZoomBottom = 0xf60d16 /* '\x00\xf6\r\x16' */
    case EntryEffectZoomCenter = 0xf60d15 /* '\x00\xf6\r\x15' */
    case EntryEffectZoomInSlightly = 0xf60d12 /* '\x00\xf6\r\x12' */
    case EntryEffectZoomIn = 0xf60d11 /* '\x00\xf6\r\x11' */
    case EntryEffectZoomOutSlightly = 0xf60d14 /* '\x00\xf6\r\x14' */
    case EntryEffectZoomOut = 0xf60d13 /* '\x00\xf6\r\x13' */
}

// MARK: PowerPointEPPTextLevelEffect
@objc public enum PowerPointEPPTextLevelEffect : AEKeyword {
    case AnimationLevelUnset = 0xdffffe /* '\x00\xdf\xff\xfe' */
    case AnimateLevelNone = 0xe00000 /* '\x00\xe0\x00\x00' */
    case AnimateLevelFirstLevel = 0xe00001 /* '\x00\xe0\x00\x01' */
    case AnimateLevelSecondLevel = 0xe00002 /* '\x00\xe0\x00\x02' */
    case AnimateLevelThirdLevel = 0xe00003 /* '\x00\xe0\x00\x03' */
    case AnimateLevelFourthLevel = 0xe00004 /* '\x00\xe0\x00\x04' */
    case AnimateLevelFifthLevel = 0xe00005 /* '\x00\xe0\x00\x05' */
    case AnimateLevelAllLevels = 0xe00010 /* '\x00\xe0\x00\x10' */
}

// MARK: PowerPointEPPTextUnitEffect
@objc public enum PowerPointEPPTextUnitEffect : AEKeyword {
    case AnimationUnitUnset = 0xe0fffe /* '\x00\xe0\xff\xfe' */
    case TextUnitEffectByParagraph = 0xe10000 /* '\x00\xe1\x00\x00' */
    case TextUnitEffectByWord = 0xe10001 /* '\x00\xe1\x00\x01' */
    case TextUnitEffectByCharacter = 0xe10002 /* '\x00\xe1\x00\x02' */
}

// MARK: PowerPointEPPChartUnitEffect
@objc public enum PowerPointEPPChartUnitEffect : AEKeyword {
    case AnimationChartUnset = 0xe1fffe /* '\x00\xe1\xff\xfe' */
    case ChartUnitEffectBySeries = 0xe20001 /* '\x00\xe2\x00\x01' */
    case ChartUnitEffectByCategory = 0xe20002 /* '\x00\xe2\x00\x02' */
    case ChartUnitEffectBySeriesElement = 0xe20003 /* '\x00\xe2\x00\x03' */
}

// MARK: PowerPointEPPAfterEffect
@objc public enum PowerPointEPPAfterEffect : AEKeyword {
    case AfterEffectUnset = 0xf3fffe /* '\x00\xf3\xff\xfe' */
    case AfterEffectNone = 0xf40000 /* '\x00\xf4\x00\x00' */
    case AfterEffectHide = 0xf40001 /* '\x00\xf4\x00\x01' */
    case AfterEffectDim = 0xf40002 /* '\x00\xf4\x00\x02' */
}

// MARK: PowerPointEPPAdvanceMode
@objc public enum PowerPointEPPAdvanceMode : AEKeyword {
    case AdvanceModeUnset = 0xf1fffe /* '\x00\xf1\xff\xfe' */
    case AdvanceModeOnClick = 0xf20001 /* '\x00\xf2\x00\x01' */
}

// MARK: PowerPointEPPSoundEffectType
@objc public enum PowerPointEPPSoundEffectType : AEKeyword {
    case SoundEffectUnset = 0xd9fffe /* '\x00\xd9\xff\xfe' */
    case SoundEffectNone = 0xda0000 /* '\x00\xda\x00\x00' */
    case SoundEffectStopPrevious = 0xda0001 /* '\x00\xda\x00\x01' */
    case SoundEffectFile = 0xda0002 /* '\x00\xda\x00\x02' */
}

// MARK: PowerPointEPPUpdateOption
@objc public enum PowerPointEPPUpdateOption : AEKeyword {
    case UpdateOptionUnset = 0xdefffe /* '\x00\xde\xff\xfe' */
    case UpdateOptionManual = 0xdf0001 /* '\x00\xdf\x00\x01' */
}

// MARK: PowerPointEPPChangeCase
@objc public enum PowerPointEPPChangeCase : AEKeyword {
    case PpCaseSentence = 0xe40001 /* '\x00\xe4\x00\x01' */
    case PpCaseLower = 0xe40002 /* '\x00\xe4\x00\x02' */
    case PpCaseUpper = 0xe40003 /* '\x00\xe4\x00\x03' */
    case PpCaseTitle = 0xe40004 /* '\x00\xe4\x00\x04' */
}

// MARK: PowerPointEPPDialogMode
@objc public enum PowerPointEPPDialogMode : AEKeyword {
    case DialogModeUnset = 0xeffffe /* '\x00\xef\xff\xfe' */
    case DialogModeModless = 0xf00000 /* '\x00\xf0\x00\x00' */
    case DialogModeModal = 0xf00001 /* '\x00\xf0\x00\x01' */
}

// MARK: PowerPointEPPDialogStyle
@objc public enum PowerPointEPPDialogStyle : AEKeyword {
    case DialogStyleUnset = 0xf0fffe /* '\x00\xf0\xff\xfe' */
    case DialogStyleStandard = 0xf10001 /* '\x00\xf1\x00\x01' */
}

// MARK: PowerPointEPPSlideShowPointerType
@objc public enum PowerPointEPPSlideShowPointerType : AEKeyword {
    case SlideShowPointerNone = 0xd20000 /* '\x00\xd2\x00\x00' */
    case SlideShowPointerArrow = 0xd20001 /* '\x00\xd2\x00\x01' */
    case SlideShowPointerPen = 0xd20002 /* '\x00\xd2\x00\x02' */
    case SlideShowPointerAlwaysHidden = 0xd20003 /* '\x00\xd2\x00\x03' */
}

// MARK: PowerPointEPPSlideShowState
@objc public enum PowerPointEPPSlideShowState : AEKeyword {
    case SlideShowStateRunning = 0xd30001 /* '\x00\xd3\x00\x01' */
    case SlideShowStatePaused = 0xd30002 /* '\x00\xd3\x00\x02' */
    case SlideShowStateBlackScreen = 0xd30003 /* '\x00\xd3\x00\x03' */
    case SlideShowStateWhiteScreen = 0xd30004 /* '\x00\xd3\x00\x04' */
    case SlideShowStateDone = 0xd30005 /* '\x00\xd3\x00\x05' */
}

// MARK: PowerPointEPPPlayerState
@objc public enum PowerPointEPPPlayerState : AEKeyword {
    case PlayerStatePlaying = 0xf50000 /* '\x00\xf5\x00\x00' */
    case PlayerStatePaused = 0xf50001 /* '\x00\xf5\x00\x01' */
    case PlayerStateStopped = 0xf50002 /* '\x00\xf5\x00\x02' */
    case PlayerStateNotReady = 0xf50003 /* '\x00\xf5\x00\x03' */
}

// MARK: PowerPointEPPSlideShowAdvanceMode
@objc public enum PowerPointEPPSlideShowAdvanceMode : AEKeyword {
    case SlideShowAdvanceManualAdvance = 0xd40001 /* '\x00\xd4\x00\x01' */
    case SlideShowAdvanceUseSlideTimings = 0xd40002 /* '\x00\xd4\x00\x02' */
}

// MARK: PowerPointEPPPrintOutputType
@objc public enum PowerPointEPPPrintOutputType : AEKeyword {
    case PrintSlides = 0xd80001 /* '\x00\xd8\x00\x01' */
    case PrintTwoSlideHandouts = 0xd80002 /* '\x00\xd8\x00\x02' */
    case PrintThreeSlideHandouts = 0xd80003 /* '\x00\xd8\x00\x03' */
    case PrintSixSlideHandouts = 0xd80004 /* '\x00\xd8\x00\x04' */
    case PrintNotesPages = 0xd80005 /* '\x00\xd8\x00\x05' */
    case PrintOutline = 0xd80006 /* '\x00\xd8\x00\x06' */
    case PrintFourSlideHandouts = 0xd80007 /* '\x00\xd8\x00\x07' */
    case PrintNineSlideHandouts = 0xd80008 /* '\x00\xd8\x00\x08' */
}

// MARK: PowerPointEPPPrintColorType
@objc public enum PowerPointEPPPrintColorType : AEKeyword {
    case PrintColor = 0xd70001 /* '\x00\xd7\x00\x01' */
    case PrintBlackAndWhite = 0xd70002 /* '\x00\xd7\x00\x02' */
}

// MARK: PowerPointEPPSelectionType
@objc public enum PowerPointEPPSelectionType : AEKeyword {
    case SelectionTypeNone = 0xce0000 /* '\x00\xce\x00\x00' */
    case SelectionTypeSlides = 0xce0001 /* '\x00\xce\x00\x01' */
    case SelectionTypeShapes = 0xce0002 /* '\x00\xce\x00\x02' */
    case SelectionTypeText = 0xce0003 /* '\x00\xce\x00\x03' */
}

// MARK: PowerPointEPPDirection
@objc public enum PowerPointEPPDirection : AEKeyword {
    case DirectionUnset = 0xeafffe /* '\x00\xea\xff\xfe' */
    case LeftToRight = 0xeb0001 /* '\x00\xeb\x00\x01' */
}

// MARK: PowerPointEPPDateTimeFormat
@objc public enum PowerPointEPPDateTimeFormat : AEKeyword {
    case Unset = 0xe2fffe /* '\x00\xe2\xff\xfe' */
    case Mdyy = 0xe30001 /* '\x00\xe3\x00\x01' */
    case DdddMMMMddyyyy = 0xe30002 /* '\x00\xe3\x00\x02' */
    case MMMMyyyy = 0xe30003 /* '\x00\xe3\x00\x03' */
    case MMMMdyyyy = 0xe30004 /* '\x00\xe3\x00\x04' */
    case MMMyy = 0xe30005 /* '\x00\xe3\x00\x05' */
    case MMMMyy = 0xe30006 /* '\x00\xe3\x00\x06' */
    case MMyy = 0xe30007 /* '\x00\xe3\x00\x07' */
    case MMddyyHmm = 0xe30008 /* '\x00\xe3\x00\x08' */
    case MddyyhmmAMPM = 0xe30009 /* '\x00\xe3\x00\t' */
    case Hmm = 0xe3000a /* '\x00\xe3\x00\n' */
    case Hmmss = 0xe3000b /* '\x00\xe3\x00\x0b' */
    case HmmAMPM = 0xe3000c /* '\x00\xe3\x00\x0c' */
    case HmmssAMPM = 0xe3000d /* '\x00\xe3\x00\r' */
}

// MARK: PowerPointEPPTransitionSpeed
@objc public enum PowerPointEPPTransitionSpeed : AEKeyword {
    case TransitionSpeedUnset = 0xd8fffe /* '\x00\xd8\xff\xfe' */
    case TransistionSpeedSlow = 0xd90001 /* '\x00\xd9\x00\x01' */
    case TransistionSpeedMedium = 0xd90002 /* '\x00\xd9\x00\x02' */
}

// MARK: PowerPointEPPMouseActivation
@objc public enum PowerPointEPPMouseActivation : AEKeyword {
    case MouseActivationMouseClick = 0xfa0001 /* '\x00\xfa\x00\x01' */
    case MouseActivationMouseOver = 0xfa0002 /* '\x00\xfa\x00\x02' */
}

// MARK: PowerPointEPPActionType
@objc public enum PowerPointEPPActionType : AEKeyword {
    case ActionTypeUnset = 0xe5fffe /* '\x00\xe5\xff\xfe' */
    case ActionTypeNone = 0xe60000 /* '\x00\xe6\x00\x00' */
    case ActionTypeNextSlide = 0xe60001 /* '\x00\xe6\x00\x01' */
    case ActionTypePreviousSlide = 0xe60002 /* '\x00\xe6\x00\x02' */
    case ActionTypeFirstSlide = 0xe60003 /* '\x00\xe6\x00\x03' */
    case ActionTypeLastSlide = 0xe60004 /* '\x00\xe6\x00\x04' */
    case ActionTypeLastSlideViewed = 0xe60005 /* '\x00\xe6\x00\x05' */
    case ActionTypeEndShow = 0xe60006 /* '\x00\xe6\x00\x06' */
    case ActionTypeHyperlinkAction = 0xe60007 /* '\x00\xe6\x00\x07' */
    case ActionTypeRunMacro = 0xe60008 /* '\x00\xe6\x00\x08' */
    case ActionTypeRunProgram = 0xe60009 /* '\x00\xe6\x00\t' */
    case ActionTypeNamedSlideshowAction = 0xe6000a /* '\x00\xe6\x00\n' */
    case ActionTypeOLEVerb = 0xe6000b /* '\x00\xe6\x00\x0b' */
}

// MARK: PowerPointEPPPlaceholderType
@objc public enum PowerPointEPPPlaceholderType : AEKeyword {
    case PlaceholderTypeUnset = 0xdafffe /* '\x00\xda\xff\xfe' */
    case PlaceholderTypeTitlePlaceholder = 0xdb0001 /* '\x00\xdb\x00\x01' */
    case PlaceholderTypeBodyPlaceholder = 0xdb0002 /* '\x00\xdb\x00\x02' */
    case PlaceholderTypeCenterTitlePlaceholder = 0xdb0003 /* '\x00\xdb\x00\x03' */
    case PlaceholderTypeSubtitlePlaceholder = 0xdb0004 /* '\x00\xdb\x00\x04' */
    case PlaceholderTypeVerticalTitlePlaceholder = 0xdb0005 /* '\x00\xdb\x00\x05' */
    case PlaceholderTypeVerticalBodyPlaceholder = 0xdb0006 /* '\x00\xdb\x00\x06' */
    case PlaceholderTypeObjectPlaceholder = 0xdb0007 /* '\x00\xdb\x00\x07' */
    case PlaceholderTypeChartPlaceholder = 0xdb0008 /* '\x00\xdb\x00\x08' */
    case PlaceholderTypeBitmapPlaceholder = 0xdb0009 /* '\x00\xdb\x00\t' */
    case PlaceholderTypeMediaClipPlaceholder = 0xdb000a /* '\x00\xdb\x00\n' */
    case PlaceholderTypeOrgChartPlaceholder = 0xdb000b /* '\x00\xdb\x00\x0b' */
    case PlaceholderTypeTablePlaceholder = 0xdb000c /* '\x00\xdb\x00\x0c' */
    case PlaceholderTypeSlideNumberPlaceholder = 0xdb000d /* '\x00\xdb\x00\r' */
    case PlaceholderTypeHeaderPlaceholder = 0xdb000e /* '\x00\xdb\x00\x0e' */
    case PlaceholderTypeFooterPlaceholder = 0xdb000f /* '\x00\xdb\x00\x0f' */
    case PlaceholderTypeDatePlaceholder = 0xdb0010 /* '\x00\xdb\x00\x10' */
    case PlaceholderTypeVerticalObjectPlaceholder = 0xdb0011 /* '\x00\xdb\x00\x11' */
    case PlaceholderTypePicturePlaceholder = 0xdb0012 /* '\x00\xdb\x00\x12' */
}

// MARK: PowerPointEPPSlideShowType
@objc public enum PowerPointEPPSlideShowType : AEKeyword {
    case SlideShowTypeSpeaker = 0xd50001 /* '\x00\xd5\x00\x01' */
    case SlideShowTypeWindow = 0xd50002 /* '\x00\xd5\x00\x02' */
    case SlideShowTypeKiosk = 0xd50003 /* '\x00\xd5\x00\x03' */
    case SlideShowTypePresenter = 0xd50005 /* '\x00\xd5\x00\x05' */
}

// MARK: PowerPointEPPPrintRangeType
@objc public enum PowerPointEPPPrintRangeType : AEKeyword {
    case PrintRangeAll = 0xf70001 /* '\x00\xf7\x00\x01' */
    case PrintRangeSelection = 0xf70002 /* '\x00\xf7\x00\x02' */
    case PrintRangeCurrent = 0xf70003 /* '\x00\xf7\x00\x03' */
    case PrintRangeSlideRange = 0xf70004 /* '\x00\xf7\x00\x04' */
    case PrintSection = 0xf70005 /* '\x00\xf7\x00\x05' */
}

// MARK: PowerPointEPPAutoSize
@objc public enum PowerPointEPPAutoSize : AEKeyword {
    case PpAutoSizemixed = 0xe4fffe /* '\x00\xe4\xff\xfe' */
    case PpAutoSizeNone = 0xe50000 /* '\x00\xe5\x00\x00' */
    case PpAutoSizeShapeToFitText = 0xe50001 /* '\x00\xe5\x00\x01' */
    case PpAutoSizeTextToFitShape = 0xe50002 /* '\x00\xe5\x00\x02' */
}

// MARK: PowerPointEPPMediaType
@objc public enum PowerPointEPPMediaType : AEKeyword {
    case MediaTypeUnset = 0xdbfffe /* '\x00\xdb\xff\xfe' */
    case MediaTypeOther = 0xdc0001 /* '\x00\xdc\x00\x01' */
    case MediaTypeSound = 0xdc0002 /* '\x00\xdc\x00\x02' */
    case MediaTypeMovie = 0xdc0003 /* '\x00\xdc\x00\x03' */
}

// MARK: PowerPointEPPSoundFormatType
@objc public enum PowerPointEPPSoundFormatType : AEKeyword {
    case SoundFormatUnset = 0xf7fffe /* '\x00\xf7\xff\xfe' */
    case SoundFormatNone = 0xf80000 /* '\x00\xf8\x00\x00' */
    case SoundFormatWAV = 0xf80001 /* '\x00\xf8\x00\x01' */
    case SoundFormatMIDI = 0xf80002 /* '\x00\xf8\x00\x02' */
}

// MARK: PowerPointEPPFarEastLineBreakLevel
@objc public enum PowerPointEPPFarEastLineBreakLevel : AEKeyword {
    case EastAsianLineBreakNormal = 0xec0001 /* '\x00\xec\x00\x01' */
    case EastAsianLineBreakStrict = 0xec0002 /* '\x00\xec\x00\x02' */
    case EastAsianLineBreakCustom = 0xec0003 /* '\x00\xec\x00\x03' */
}

// MARK: PowerPointEPPSlideShowRangeType
@objc public enum PowerPointEPPSlideShowRangeType : AEKeyword {
    case SlideShowRangeShowAll = 0xd60001 /* '\x00\xd6\x00\x01' */
    case SlideShowRange = 0xd60002 /* '\x00\xd6\x00\x02' */
    case SlideShowRangeNamedSlideshow = 0xd60003 /* '\x00\xd6\x00\x03' */
}

// MARK: PowerPointEPPFrameColors
@objc public enum PowerPointEPPFrameColors : AEKeyword {
    case FrameColorsBrowserColors = 0xcf0001 /* '\x00\xcf\x00\x01' */
    case FrameColorsPresentationSchemeTextColor = 0xcf0002 /* '\x00\xcf\x00\x02' */
    case FrameColorsPresentationSchemeAccentColor = 0xcf0003 /* '\x00\xcf\x00\x03' */
    case FrameColorsWhiteTextOnBlack = 0xcf0004 /* '\x00\xcf\x00\x04' */
    case FrameColorsBlackTextOnWhite = 0xcf0005 /* '\x00\xcf\x00\x05' */
}

// MARK: PowerPointEPPMovieOptimization
@objc public enum PowerPointEPPMovieOptimization : AEKeyword {
    case MovieOptimizationNormal = 0xcffffe /* '\x00\xcf\xff\xfe' */
    case MovieOptimizationSize = 0xd00001 /* '\x00\xd0\x00\x01' */
    case MovieOptimizationSpeed = 0xd00002 /* '\x00\xd0\x00\x02' */
    case MovieOptimizationQuality = 0xd00003 /* '\x00\xd0\x00\x03' */
}

// MARK: PowerPointEPPBulletType
@objc public enum PowerPointEPPBulletType : AEKeyword {
    case PpBulletmixed = 0xe7fffe /* '\x00\xe7\xff\xfe' */
    case PpBulletNone = 0xe80000 /* '\x00\xe8\x00\x00' */
    case PpBulletUnnumbered = 0xe80001 /* '\x00\xe8\x00\x01' */
    case PpBulletNumbered = 0xe80002 /* '\x00\xe8\x00\x02' */
    case PpBulletPicture = 0xe80003 /* '\x00\xe8\x00\x03' */
}

// MARK: PowerPointEPPNumberedBulletStyle
@objc public enum PowerPointEPPNumberedBulletStyle : AEKeyword {
    case PpBulletStylemixed = 0xe8fffe /* '\x00\xe8\xff\xfe' */
    case PpBulletAlphaLCPeriod = 0xe90000 /* '\x00\xe9\x00\x00' */
    case PpBulletAlphaUCPeriod = 0xe90001 /* '\x00\xe9\x00\x01' */
    case PpBulletArabicParenRight = 0xe90002 /* '\x00\xe9\x00\x02' */
    case PpBulletArabicPeriod = 0xe90003 /* '\x00\xe9\x00\x03' */
    case PpBulletRomanLCParenBoth = 0xe90004 /* '\x00\xe9\x00\x04' */
    case PpBulletRomanLCParenRight = 0xe90005 /* '\x00\xe9\x00\x05' */
    case PpBulletRomanLCPeriod = 0xe90006 /* '\x00\xe9\x00\x06' */
    case PpBulletRomanUCPeriod = 0xe90007 /* '\x00\xe9\x00\x07' */
    case PpBulletAlphaLCParenBoth = 0xe90008 /* '\x00\xe9\x00\x08' */
    case PpBulletAlphaLCParenRight = 0xe90009 /* '\x00\xe9\x00\t' */
    case PpBulletAlphaUCParenBoth = 0xe9000a /* '\x00\xe9\x00\n' */
    case PpBulletAlphaUCParenRight = 0xe9000b /* '\x00\xe9\x00\x0b' */
    case PpBulletArabicParenBoth = 0xe9000c /* '\x00\xe9\x00\x0c' */
    case PpBulletArabicPlain = 0xe9000d /* '\x00\xe9\x00\r' */
    case PpBulletRomanUCParenBoth = 0xe9000e /* '\x00\xe9\x00\x0e' */
    case PpBulletRomanUCParenRight = 0xe9000f /* '\x00\xe9\x00\x0f' */
    case PpBulletSimpChinPlain = 0xe90010 /* '\x00\xe9\x00\x10' */
    case PpBulletSimpChinPeriod = 0xe90011 /* '\x00\xe9\x00\x11' */
    case PpBulletCircleNumDBPlain = 0xe90012 /* '\x00\xe9\x00\x12' */
    case PpBulletCircleNumWDWhitePlain = 0xe90013 /* '\x00\xe9\x00\x13' */
    case PpBulletCircleNumWDBlackPlain = 0xe90014 /* '\x00\xe9\x00\x14' */
    case PpBulletTradChinPlain = 0xe90015 /* '\x00\xe9\x00\x15' */
    case PpBulletTradChinPeriod = 0xe90016 /* '\x00\xe9\x00\x16' */
    case PpBulletArabicAlphaDash = 0xe90017 /* '\x00\xe9\x00\x17' */
    case PpBulletArabicAbjadDash = 0xe90018 /* '\x00\xe9\x00\x18' */
    case PpBulletHebrewAlphaDash = 0xe90019 /* '\x00\xe9\x00\x19' */
    case PpBulletKanjiKoreanPlain = 0xe9001a /* '\x00\xe9\x00\x1a' */
    case PpBulletKanjiKoreanPeriod = 0xe9001b /* '\x00\xe9\x00\x1b' */
    case PpBulletArabicDBPlain = 0xe9001c /* '\x00\xe9\x00\x1c' */
}

// MARK: PowerPointEPPShapeFormat
@objc public enum PowerPointEPPShapeFormat : AEKeyword {
    case ShapeFormatGIF = 0xdd0000 /* '\x00\xdd\x00\x00' */
    case ShapeFormatJPEG = 0xdd0001 /* '\x00\xdd\x00\x01' */
    case ShapeFormatPNG = 0xdd0002 /* '\x00\xdd\x00\x02' */
    case ShapeFormatPICT = 0xdd0003 /* '\x00\xdd\x00\x03' */
}

// MARK: PowerPointEPPExportMode
@objc public enum PowerPointEPPExportMode : AEKeyword {
    case ExportModeRelativeToSlide = 0xde0001 /* '\x00\xde\x00\x01' */
    case ExportModeClipRelativeToSlide = 0xde0001 /* '\x00\xde\x00\x01' */
    case ExportModeScaleToFit = 0xde0002 /* '\x00\xde\x00\x02' */
    case ExportModeScaleXY = 0xde0003 /* '\x00\xde\x00\x03' */
}

// MARK: PowerPointPpBorderType
@objc public enum PowerPointPpBorderType : AEKeyword {
    case TopBorder = 0x11a0001 /* '\x01\x1a\x00\x01' */
    case LeftBorder = 0x11a0002 /* '\x01\x1a\x00\x02' */
    case BottomBorder = 0x11a0003 /* '\x01\x1a\x00\x03' */
    case RightBorder = 0x11a0004 /* '\x01\x1a\x00\x04' */
    case DiagonalDownBorder = 0x11a0005 /* '\x01\x1a\x00\x05' */
    case DiagonalUpBorder = 0x11a0006 /* '\x01\x1a\x00\x06' */
}

// MARK: PowerPointEPPCheckInVersionType
@objc public enum PowerPointEPPCheckInVersionType : AEKeyword {
    case MinorVersion = 0x2c40000 /* '\x02\xc4\x00\x00' */
    case MajorVersion = 0x2c40001 /* '\x02\xc4\x00\x01' */
    case OverwriteCurrentVersion = 0x2c40002 /* '\x02\xc4\x00\x02' */
}

// MARK: PowerPointIppPageLayout
@objc public enum PowerPointIppPageLayout : AEKeyword {
    case PageLayoutNormal = 0xed0000 /* '\x00\xed\x00\x00' */
    case PageLayoutFullScreen = 0xed0001 /* '\x00\xed\x00\x01' */
}

// MARK: PowerPointIppButtonsType
@objc public enum PowerPointIppButtonsType : AEKeyword {
    case Regular = 0xee0001 /* '\x00\xee\x00\x01' */
    case Fancy = 0xee0002 /* '\x00\xee\x00\x02' */
    case TextOnly = 0xee0003 /* '\x00\xee\x00\x03' */
}

// MARK: PowerPointIppNavBarPlacement
@objc public enum PowerPointIppNavBarPlacement : AEKeyword {
    case BarPlacementTop = 0xef0001 /* '\x00\xef\x00\x01' */
    case BarPlacementBottom = 0xef0002 /* '\x00\xef\x00\x02' */
}

// MARK: PowerPointMsoAnimEffect
@objc public enum PowerPointMsoAnimEffect : AEKeyword {
    case AnimationTypeCustom = 0x1020000 /* '\x01\x02\x00\x00' */
    case AnimationTypeAppear = 0x1020001 /* '\x01\x02\x00\x01' */
    case AnimationTypeFly = 0x1020002 /* '\x01\x02\x00\x02' */
    case AnimationTypeBlinds = 0x1020003 /* '\x01\x02\x00\x03' */
    case AnimationTypeBox = 0x1020004 /* '\x01\x02\x00\x04' */
    case AnimationTypeCheckerboard = 0x1020005 /* '\x01\x02\x00\x05' */
    case AnimationTypeCircle = 0x1020006 /* '\x01\x02\x00\x06' */
    case AnimationTypeCrawl = 0x1020007 /* '\x01\x02\x00\x07' */
    case AnimationTypeDiamond = 0x1020008 /* '\x01\x02\x00\x08' */
    case AnimationTypeDissolve = 0x1020009 /* '\x01\x02\x00\t' */
    case AnimationTypeFade = 0x102000a /* '\x01\x02\x00\n' */
    case AnimationTypeFlashOnce = 0x102000b /* '\x01\x02\x00\x0b' */
    case AnimationTypePeek = 0x102000c /* '\x01\x02\x00\x0c' */
    case AnimationTypePlus = 0x102000d /* '\x01\x02\x00\r' */
    case AnimationTypeRandomBars = 0x102000e /* '\x01\x02\x00\x0e' */
    case AnimationTypeSpiral = 0x102000f /* '\x01\x02\x00\x0f' */
    case AnimationTypeSplit = 0x1020010 /* '\x01\x02\x00\x10' */
    case AnimationTypeStretch = 0x1020011 /* '\x01\x02\x00\x11' */
    case AnimationTypeStrips = 0x1020012 /* '\x01\x02\x00\x12' */
    case AnimationTypeSwivel = 0x1020013 /* '\x01\x02\x00\x13' */
    case AnimationTypeWedge = 0x1020014 /* '\x01\x02\x00\x14' */
    case AnimationTypeWheel = 0x1020015 /* '\x01\x02\x00\x15' */
    case AnimationTypeWipe = 0x1020016 /* '\x01\x02\x00\x16' */
    case AnimationTypeZoom = 0x1020017 /* '\x01\x02\x00\x17' */
    case AnimationTypeRandomEffect = 0x1020018 /* '\x01\x02\x00\x18' */
    case AnimationTypeBoomerang = 0x1020019 /* '\x01\x02\x00\x19' */
    case AnimationTypeBounce = 0x102001a /* '\x01\x02\x00\x1a' */
    case AnimationTypeColorReveal = 0x102001b /* '\x01\x02\x00\x1b' */
    case AnimationTypeCredits = 0x102001c /* '\x01\x02\x00\x1c' */
    case AnimationTypeEaseIn = 0x102001d /* '\x01\x02\x00\x1d' */
    case AnimationTypeFloat = 0x102001e /* '\x01\x02\x00\x1e' */
    case AnimationTypeGrowAndTurn = 0x102001f /* '\x01\x02\x00\x1f' */
    case AnimationTypeLightSpeed = 0x1020020 /* '\x01\x02\x00 ' */
    case AnimationTypePinwheel = 0x1020021 /* '\x01\x02\x00!' */
    case AnimationTypeRiseUp = 0x1020022 /* '\x01\x02\x00"' */
    case AnimationTypeSwish = 0x1020023 /* '\x01\x02\x00#' */
    case AnimationTypeThinLine = 0x1020024 /* '\x01\x02\x00$' */
    case AnimationTypeUnfold = 0x1020025 /* '\x01\x02\x00%' */
    case AnimationTypeWhip = 0x1020026 /* '\x01\x02\x00&' */
    case AnimationTypeAscend = 0x1020027 /* "\x01\x02\x00'" */
    case AnimationTypeCenterRevolve = 0x1020028 /* '\x01\x02\x00(' */
    case AnimationTypeFadedSwivel = 0x1020029 /* '\x01\x02\x00)' */
    case AnimationTypeDescend = 0x102002a /* '\x01\x02\x00*' */
    case AnimationTypeSling = 0x102002b /* '\x01\x02\x00+' */
    case AnimationTypeSpinner = 0x102002c /* '\x01\x02\x00,' */
    case AnimationTypeStretchy = 0x102002d /* '\x01\x02\x00-' */
    case AnimationTypeZip = 0x102002e /* '\x01\x02\x00.' */
    case AnimationTypeArcUp = 0x102002f /* '\x01\x02\x00/' */
    case AnimationTypeFadeZoom = 0x1020030 /* '\x01\x02\x000' */
    case AnimationTypeGlide = 0x1020031 /* '\x01\x02\x001' */
    case AnimationTypeExpand = 0x1020032 /* '\x01\x02\x002' */
    case AnimationTypeFlip = 0x1020033 /* '\x01\x02\x003' */
    case AnimationTypeShimmer = 0x1020034 /* '\x01\x02\x004' */
    case AnimationTypeFold = 0x1020035 /* '\x01\x02\x005' */
    case AnimationTypeChangeFillColor = 0x1020036 /* '\x01\x02\x006' */
    case AnimationTypeChangeFont = 0x1020037 /* '\x01\x02\x007' */
    case AnimationTypeChangeFontColor = 0x1020038 /* '\x01\x02\x008' */
    case AnimationTypeChangeFontSize = 0x1020039 /* '\x01\x02\x009' */
    case AnimationTypeChangeFontStyle = 0x102003a /* '\x01\x02\x00:' */
    case AnimationTypeGrowShrink = 0x102003b /* '\x01\x02\x00;' */
    case AnimationTypeChangeLineColor = 0x102003c /* '\x01\x02\x00<' */
    case AnimationTypeSpin = 0x102003d /* '\x01\x02\x00=' */
    case AnimationTypeTransparency = 0x102003e /* '\x01\x02\x00>' */
    case AnimationTypeBoldFlash = 0x102003f /* '\x01\x02\x00?' */
    case AnimationTypeBlast = 0x1020040 /* '\x01\x02\x00@' */
    case AnimationTypeBoldReveal = 0x1020041 /* '\x01\x02\x00A' */
    case AnimationTypeBrushOnColor = 0x1020042 /* '\x01\x02\x00B' */
    case AnimationTypeBrushOnUnderline = 0x1020043 /* '\x01\x02\x00C' */
    case AnimationTypeColorBlend = 0x1020044 /* '\x01\x02\x00D' */
    case AnimationTypeColorWave = 0x1020045 /* '\x01\x02\x00E' */
    case AnimationTypeComplementaryColor = 0x1020046 /* '\x01\x02\x00F' */
    case AnimationTypeComplementaryColor2 = 0x1020047 /* '\x01\x02\x00G' */
    case AnimationTypeConstrastingColor = 0x1020048 /* '\x01\x02\x00H' */
    case AnimationTypeDarken = 0x1020049 /* '\x01\x02\x00I' */
    case AnimationTypeDesaturate = 0x102004a /* '\x01\x02\x00J' */
    case AnimationTypeFlashBulb = 0x102004b /* '\x01\x02\x00K' */
    case AnimationTypeFlicker = 0x102004c /* '\x01\x02\x00L' */
    case AnimationTypeGrowWithColor = 0x102004d /* '\x01\x02\x00M' */
    case AnimationTypeLighten = 0x102004e /* '\x01\x02\x00N' */
    case AnimationTypeStyleEmphasis = 0x102004f /* '\x01\x02\x00O' */
    case AnimationTypeTeeter = 0x1020050 /* '\x01\x02\x00P' */
    case AnimationTypeVerticalGrow = 0x1020051 /* '\x01\x02\x00Q' */
    case AnimationTypeWave = 0x1020052 /* '\x01\x02\x00R' */
    case AnimationTypeMediaPlay = 0x1020053 /* '\x01\x02\x00S' */
    case AnimationTypeMediaPause = 0x1020054 /* '\x01\x02\x00T' */
    case AnimationTypeMediaStop = 0x1020055 /* '\x01\x02\x00U' */
    case AnimationTypeCirclePath = 0x1020056 /* '\x01\x02\x00V' */
    case AnimationTypeRightTrianglePath = 0x1020057 /* '\x01\x02\x00W' */
    case AnimationTypeDiamondPath = 0x1020058 /* '\x01\x02\x00X' */
    case AnimationTypeHexagonPath = 0x1020059 /* '\x01\x02\x00Y' */
    case AnimationType5PointStarPath = 0x102005a /* '\x01\x02\x00Z' */
    case AnimationTypeCrescentMoonPath = 0x102005b /* '\x01\x02\x00[' */
    case AnimationTypeSquarePath = 0x102005c /* '\x01\x02\x00\\' */
    case AnimationTypeTrapezoidPath = 0x102005d /* '\x01\x02\x00]' */
    case AnimationTypeHeartPath = 0x102005e /* '\x01\x02\x00^' */
    case AnimationTypeOctagonPath = 0x102005f /* '\x01\x02\x00_' */
    case AnimationType6PointStarPath = 0x1020060 /* '\x01\x02\x00`' */
    case AnimationTypeFootballPath = 0x1020061 /* '\x01\x02\x00a' */
    case AnimationTypeEqualTrianglePath = 0x1020062 /* '\x01\x02\x00b' */
    case AnimationTypeParallelogramPath = 0x1020063 /* '\x01\x02\x00c' */
    case AnimationTypePentagonPath = 0x1020064 /* '\x01\x02\x00d' */
    case AnimationType4PointStarPath = 0x1020065 /* '\x01\x02\x00e' */
    case AnimationType8PointStarPath = 0x1020066 /* '\x01\x02\x00f' */
    case AnimationTypeTeardropPath = 0x1020067 /* '\x01\x02\x00g' */
    case AnimationTypePointyStarPath = 0x1020068 /* '\x01\x02\x00h' */
    case AnimationTypeCurvedSquarePath = 0x1020069 /* '\x01\x02\x00i' */
    case AnimationTypeCurvedXPath = 0x102006a /* '\x01\x02\x00j' */
    case AnimationTypeVerticalFigure8Path = 0x102006b /* '\x01\x02\x00k' */
    case AnimationTypeCurvyStarPath = 0x102006c /* '\x01\x02\x00l' */
    case AnimationTypeLoopDeLoopPath = 0x102006d /* '\x01\x02\x00m' */
    case AnimationTypeBuzzsawPath = 0x102006e /* '\x01\x02\x00n' */
    case AnimationTypeHorizontalFigure8Path = 0x102006f /* '\x01\x02\x00o' */
    case AnimationTypePeanutPath = 0x1020070 /* '\x01\x02\x00p' */
    case AnimationTypeFigure8FourPath = 0x1020071 /* '\x01\x02\x00q' */
    case AnimationTypeNeutronPath = 0x1020072 /* '\x01\x02\x00r' */
    case AnimationTypeSwooshPath = 0x1020073 /* '\x01\x02\x00s' */
    case AnimationTypeBeanPath = 0x1020074 /* '\x01\x02\x00t' */
    case AnimationTypePlusPath = 0x1020075 /* '\x01\x02\x00u' */
    case AnimationTypeInvertedTrianglePath = 0x1020076 /* '\x01\x02\x00v' */
    case AnimationTypeInvertedSquarePath = 0x1020077 /* '\x01\x02\x00w' */
    case AnimationTypeLeftPath = 0x1020078 /* '\x01\x02\x00x' */
    case AnimationTypeTurnRightPath = 0x1020079 /* '\x01\x02\x00y' */
    case AnimationTypeArcDownPath = 0x102007a /* '\x01\x02\x00z' */
    case AnimationTypeZigzagPath = 0x102007b /* '\x01\x02\x00{' */
    case AnimationTypeSCurve2Path = 0x102007c /* '\x01\x02\x00|' */
    case AnimationTypeSineWavePath = 0x102007d /* '\x01\x02\x00}' */
    case AnimationTypeBounceLeftPath = 0x102007e /* '\x01\x02\x00~' */
    case AnimationTypeDownPath = 0x102007f /* '\x01\x02\x00\x7f' */
    case AnimationTypeTurnUpPath = 0x1020080 /* '\x01\x02\x00\x80' */
    case AnimationTypeArcUpPath = 0x1020081 /* '\x01\x02\x00\x81' */
    case AnimationTypeHeartbeatPath = 0x1020082 /* '\x01\x02\x00\x82' */
    case AnimationTypeSpiralRightPath = 0x1020083 /* '\x01\x02\x00\x83' */
    case AnimationTypeWavePath = 0x1020084 /* '\x01\x02\x00\x84' */
    case AnimationTypeCurvyLeftPath = 0x1020085 /* '\x01\x02\x00\x85' */
    case AnimationTypeDiagonalDownRightPath = 0x1020086 /* '\x01\x02\x00\x86' */
    case AnimationTypeTurnDownPath = 0x1020087 /* '\x01\x02\x00\x87' */
    case AnimationTypeArcLeftPath = 0x1020088 /* '\x01\x02\x00\x88' */
    case AnimationTypeFunnelPath = 0x1020089 /* '\x01\x02\x00\x89' */
    case AnimationTypeSpringPath = 0x102008a /* '\x01\x02\x00\x8a' */
    case AnimationTypeBounceRightPath = 0x102008b /* '\x01\x02\x00\x8b' */
    case AnimationTypeSpiralLeftPath = 0x102008c /* '\x01\x02\x00\x8c' */
    case AnimationTypeDiagonalUpRightPath = 0x102008d /* '\x01\x02\x00\x8d' */
    case AnimationTypeTurnUpRightPath = 0x102008e /* '\x01\x02\x00\x8e' */
    case AnimationTypeArcRightPath = 0x102008f /* '\x01\x02\x00\x8f' */
    case AnimationTypeSCurve1Path = 0x1020090 /* '\x01\x02\x00\x90' */
    case AnimationTypeDecayingWavePath = 0x1020091 /* '\x01\x02\x00\x91' */
    case AnimationTypeCurvyRightPath = 0x1020092 /* '\x01\x02\x00\x92' */
    case AnimationTypeStairsDownPath = 0x1020093 /* '\x01\x02\x00\x93' */
    case AnimationTypeUpPath = 0x1020094 /* '\x01\x02\x00\x94' */
    case AnimationTypeRightPath = 0x1020095 /* '\x01\x02\x00\x95' */
}

// MARK: PowerPointMsoAnimateByLevel
@objc public enum PowerPointMsoAnimateByLevel : AEKeyword {
    case TextByNoLevels = 0x1010000 /* '\x01\x01\x00\x00' */
    case TextByAllLevels = 0x1010001 /* '\x01\x01\x00\x01' */
    case TextByFirstLevel = 0x1010002 /* '\x01\x01\x00\x02' */
    case TextBySecondLevel = 0x1010003 /* '\x01\x01\x00\x03' */
    case TextByThirdLevel = 0x1010004 /* '\x01\x01\x00\x04' */
    case TextByFourthLevel = 0x1010005 /* '\x01\x01\x00\x05' */
    case TextByFifthLevel = 0x1010006 /* '\x01\x01\x00\x06' */
    case ChartAllAtOnce = 0x1010007 /* '\x01\x01\x00\x07' */
    case ChartByCategory = 0x1010008 /* '\x01\x01\x00\x08' */
    case ChartByCtageoryElements = 0x1010009 /* '\x01\x01\x00\t' */
    case ChartBySeries = 0x101000a /* '\x01\x01\x00\n' */
    case ChartBySeriesElements = 0x101000b /* '\x01\x01\x00\x0b' */
}

// MARK: PowerPointMsoAnimTriggerType
@objc public enum PowerPointMsoAnimTriggerType : AEKeyword {
    case NoTrigger = 0x1000000 /* '\x01\x00\x00\x00' */
    case OnPageClick = 0x1000001 /* '\x01\x00\x00\x01' */
    case WithPrevious = 0x1000002 /* '\x01\x00\x00\x02' */
    case AfterPrevious = 0x1000003 /* '\x01\x00\x00\x03' */
    case OnShapeClick = 0x1000004 /* '\x01\x00\x00\x04' */
}

// MARK: PowerPointMsoAnimAfterEffect
@objc public enum PowerPointMsoAnimAfterEffect : AEKeyword {
    case NoAfterEffect = 0xff0000 /* '\x00\xff\x00\x00' */
    case Dim = 0xff0001 /* '\x00\xff\x00\x01' */
    case Hide = 0xff0002 /* '\x00\xff\x00\x02' */
    case HideOnNextClick = 0xff0003 /* '\x00\xff\x00\x03' */
}

// MARK: PowerPointMsoAnimTextUnitEffect
@objc public enum PowerPointMsoAnimTextUnitEffect : AEKeyword {
    case ByParagraph = 0xfe0000 /* '\x00\xfe\x00\x00' */
    case ByCharacter = 0xfe0001 /* '\x00\xfe\x00\x01' */
    case ByWord = 0xfe0002 /* '\x00\xfe\x00\x02' */
}

// MARK: PowerPointMsoAnimEffectRestart
@objc public enum PowerPointMsoAnimEffectRestart : AEKeyword {
    case RestartAlways = 0xfd0001 /* '\x00\xfd\x00\x01' */
    case RestartWhenOff = 0xfd0002 /* '\x00\xfd\x00\x02' */
    case NeverRestart = 0xfd0003 /* '\x00\xfd\x00\x03' */
}

// MARK: PowerPointMsoAnimEffectAfter
@objc public enum PowerPointMsoAnimEffectAfter : AEKeyword {
    case AfterFreeze = 0xfc0001 /* '\x00\xfc\x00\x01' */
    case AfterRemove = 0xfc0002 /* '\x00\xfc\x00\x02' */
    case AfterHold = 0xfc0003 /* '\x00\xfc\x00\x03' */
    case AfterTransition = 0xfc0004 /* '\x00\xfc\x00\x04' */
}

// MARK: PowerPointMsoAnimDirection
@objc public enum PowerPointMsoAnimDirection : AEKeyword {
    case NoDirection = 0xf90000 /* '\x00\xf9\x00\x00' */
    case Up = 0xf90001 /* '\x00\xf9\x00\x01' */
    case Right = 0xf90002 /* '\x00\xf9\x00\x02' */
    case Down = 0xf90003 /* '\x00\xf9\x00\x03' */
    case Left = 0xf90004 /* '\x00\xf9\x00\x04' */
    case OrdinalMask = 0xf90005 /* '\x00\xf9\x00\x05' */
    case UpLeft = 0xf90006 /* '\x00\xf9\x00\x06' */
    case UpRight = 0xf90007 /* '\x00\xf9\x00\x07' */
    case DownRight = 0xf90008 /* '\x00\xf9\x00\x08' */
    case DownLeft = 0xf90009 /* '\x00\xf9\x00\t' */
    case Top = 0xf9000a /* '\x00\xf9\x00\n' */
    case Bottom = 0xf9000b /* '\x00\xf9\x00\x0b' */
    case TopLeft = 0xf9000c /* '\x00\xf9\x00\x0c' */
    case TopRight = 0xf9000d /* '\x00\xf9\x00\r' */
    case BottomRight = 0xf9000e /* '\x00\xf9\x00\x0e' */
    case BottomLeft = 0xf9000f /* '\x00\xf9\x00\x0f' */
    case Horizontal = 0xf90010 /* '\x00\xf9\x00\x10' */
    case Vertical = 0xf90011 /* '\x00\xf9\x00\x11' */
    case Across = 0xf90012 /* '\x00\xf9\x00\x12' */
    case Inward = 0xf90013 /* '\x00\xf9\x00\x13' */
    case Out = 0xf90014 /* '\x00\xf9\x00\x14' */
    case Clockwise = 0xf90015 /* '\x00\xf9\x00\x15' */
    case Counterclockwise = 0xf90016 /* '\x00\xf9\x00\x16' */
    case HorizontalIn = 0xf90017 /* '\x00\xf9\x00\x17' */
    case HorizontalOut = 0xf90018 /* '\x00\xf9\x00\x18' */
    case VerticalIn = 0xf90019 /* '\x00\xf9\x00\x19' */
    case VerticalOut = 0xf9001a /* '\x00\xf9\x00\x1a' */
    case Slightly = 0xf9001b /* '\x00\xf9\x00\x1b' */
    case Center = 0xf9001c /* '\x00\xf9\x00\x1c' */
    case InSlightly = 0xf9001d /* '\x00\xf9\x00\x1d' */
    case InCenter = 0xf9001e /* '\x00\xf9\x00\x1e' */
    case InBottom = 0xf9001f /* '\x00\xf9\x00\x1f' */
    case OutSlightly = 0xf90020 /* '\x00\xf9\x00 ' */
    case OutCenter = 0xf90021 /* '\x00\xf9\x00!' */
    case OutBottom = 0xf90022 /* '\x00\xf9\x00"' */
    case FontBold = 0xf90023 /* '\x00\xf9\x00#' */
    case FontItalic = 0xf90024 /* '\x00\xf9\x00$' */
    case FontUnderline = 0xf90025 /* '\x00\xf9\x00%' */
    case FontStrikethrough = 0xf90026 /* '\x00\xf9\x00&' */
    case FontShadow = 0xf90027 /* "\x00\xf9\x00'" */
    case FontAllCaps = 0xf90028 /* '\x00\xf9\x00(' */
    case Instant = 0xf90029 /* '\x00\xf9\x00)' */
    case Gradual = 0xf9002a /* '\x00\xf9\x00*' */
    case CycleClockwise = 0xf9002b /* '\x00\xf9\x00+' */
    case CycleCounterclockwise = 0xf9002c /* '\x00\xf9\x00,' */
}

// MARK: PowerPointMsoAnimType
@objc public enum PowerPointMsoAnimType : AEKeyword {
    case AnimationTypeNone = 0x1030000 /* '\x01\x03\x00\x00' */
    case AnimationTypeMotion = 0x1030001 /* '\x01\x03\x00\x01' */
    case AnimationTypeColor = 0x1030002 /* '\x01\x03\x00\x02' */
    case AnimationTypeScale = 0x1030003 /* '\x01\x03\x00\x03' */
    case AnimationTypeRotation = 0x1030004 /* '\x01\x03\x00\x04' */
    case AnimationTypeProperty = 0x1030005 /* '\x01\x03\x00\x05' */
    case AnimationTypeCommand = 0x1030006 /* '\x01\x03\x00\x06' */
    case AnimationTypeFilter = 0x1030007 /* '\x01\x03\x00\x07' */
    case AnimationTypeSet = 0x1030008 /* '\x01\x03\x00\x08' */
}

// MARK: PowerPointMsoAnimAdditive
@objc public enum PowerPointMsoAnimAdditive : AEKeyword {
    case NoAdditive = 0x1070001 /* '\x01\x07\x00\x01' */
    case Motion = 0x1070002 /* '\x01\x07\x00\x02' */
}

// MARK: PowerPointMsoAnimAccumulate
@objc public enum PowerPointMsoAnimAccumulate : AEKeyword {
    case NoAccumulate = 0x1040001 /* '\x01\x04\x00\x01' */
    case Always = 0x1040002 /* '\x01\x04\x00\x02' */
}

// MARK: PowerPointMsoAnimProperty
@objc public enum PowerPointMsoAnimProperty : AEKeyword {
    case NoProperty = 0x1050000 /* '\x01\x05\x00\x00' */
    case X = 0x1050001 /* '\x01\x05\x00\x01' */
    case Y = 0x1050002 /* '\x01\x05\x00\x02' */
    case Width = 0x1050003 /* '\x01\x05\x00\x03' */
    case Height = 0x1050004 /* '\x01\x05\x00\x04' */
    case Opacity = 0x1050005 /* '\x01\x05\x00\x05' */
    case Rotation = 0x1050006 /* '\x01\x05\x00\x06' */
    case Colors = 0x1050007 /* '\x01\x05\x00\x07' */
    case Visibility = 0x1050008 /* '\x01\x05\x00\x08' */
    case TextFontBold = 0x1050064 /* '\x01\x05\x00d' */
    case TextFontColor = 0x1050065 /* '\x01\x05\x00e' */
    case TextFontEmboss = 0x1050066 /* '\x01\x05\x00f' */
    case TextFontItalic = 0x1050067 /* '\x01\x05\x00g' */
    case TextFontName = 0x1050068 /* '\x01\x05\x00h' */
    case TextFontShadow = 0x1050069 /* '\x01\x05\x00i' */
    case TextFontSize = 0x105006a /* '\x01\x05\x00j' */
    case TextFontSubscript = 0x105006b /* '\x01\x05\x00k' */
    case TextFontSuperscript = 0x105006c /* '\x01\x05\x00l' */
    case TextFontUnderline = 0x105006d /* '\x01\x05\x00m' */
    case TextFontStrikethrough = 0x105006e /* '\x01\x05\x00n' */
    case TextBulletCharacter = 0x105006f /* '\x01\x05\x00o' */
    case TextBulletFontName = 0x1050070 /* '\x01\x05\x00p' */
    case TextBulletNumber = 0x1050071 /* '\x01\x05\x00q' */
    case TextBulletColor = 0x1050072 /* '\x01\x05\x00r' */
    case TextBulletRelativeSize = 0x1050073 /* '\x01\x05\x00s' */
    case TextBulletStyle = 0x1050074 /* '\x01\x05\x00t' */
    case TextBulletType = 0x1050075 /* '\x01\x05\x00u' */
    case ShapePictureContrast = 0x10503e8 /* '\x01\x05\x03\xe8' */
    case ShapePictureBrightness = 0x10503e9 /* '\x01\x05\x03\xe9' */
    case ShapePictureGamma = 0x10503ea /* '\x01\x05\x03\xea' */
    case ShapePictureGrayscale = 0x10503eb /* '\x01\x05\x03\xeb' */
    case ShapeFillOn = 0x10503ec /* '\x01\x05\x03\xec' */
    case ShapeFillColor = 0x10503ed /* '\x01\x05\x03\xed' */
    case ShapeFillOpacity = 0x10503ee /* '\x01\x05\x03\xee' */
    case ShapeFillBackColor = 0x10503ef /* '\x01\x05\x03\xef' */
    case ShapeLineOn = 0x10503f0 /* '\x01\x05\x03\xf0' */
    case ShapeLineColor = 0x10503f1 /* '\x01\x05\x03\xf1' */
    case ShapeShadowOn = 0x10503f2 /* '\x01\x05\x03\xf2' */
    case ShapeShadowType = 0x10503f3 /* '\x01\x05\x03\xf3' */
    case ShapeShadowColor = 0x10503f4 /* '\x01\x05\x03\xf4' */
    case ShapeShadowOpacity = 0x10503f5 /* '\x01\x05\x03\xf5' */
    case ShapeShadowOffsetX = 0x10503f6 /* '\x01\x05\x03\xf6' */
    case ShapeShadowOffsetY = 0x10503f7 /* '\x01\x05\x03\xf7' */
}

// MARK: PowerPointMsoAnimCommandType
@objc public enum PowerPointMsoAnimCommandType : AEKeyword {
    case Event = 0x1060000 /* '\x01\x06\x00\x00' */
    case Call = 0x1060001 /* '\x01\x06\x00\x01' */
    case Verb = 0x1060002 /* '\x01\x06\x00\x02' */
}

// MARK: PowerPointMsoAnimFilterEffectType
@objc public enum PowerPointMsoAnimFilterEffectType : AEKeyword {
    case NoFilterEffectType = 0x1080000 /* '\x01\x08\x00\x00' */
    case FilterEffectTypeBarn = 0x1080001 /* '\x01\x08\x00\x01' */
    case FilterEffectTypeBlinds = 0x1080002 /* '\x01\x08\x00\x02' */
    case FilterEffectTypeBox = 0x1080003 /* '\x01\x08\x00\x03' */
    case FilterEffectTypeCheckerboard = 0x1080004 /* '\x01\x08\x00\x04' */
    case FilterEffectTypeCircle = 0x1080005 /* '\x01\x08\x00\x05' */
    case FilterEffectTypeDiamond = 0x1080006 /* '\x01\x08\x00\x06' */
    case FilterEffectTypeDissolve = 0x1080007 /* '\x01\x08\x00\x07' */
    case FilterEffectTypeFade = 0x1080008 /* '\x01\x08\x00\x08' */
    case FilterEffectTypeImage = 0x1080009 /* '\x01\x08\x00\t' */
    case FilterEffectTypePixelate = 0x108000a /* '\x01\x08\x00\n' */
    case FilterEffectTypePlus = 0x108000b /* '\x01\x08\x00\x0b' */
    case FilterEffectTypeRandomBar = 0x108000c /* '\x01\x08\x00\x0c' */
    case FilterEffectTypeSlide = 0x108000d /* '\x01\x08\x00\r' */
    case FilterEffectTypeStretch = 0x108000e /* '\x01\x08\x00\x0e' */
    case FilterEffectTypeStrips = 0x108000f /* '\x01\x08\x00\x0f' */
    case FilterEffectTypeWedge = 0x1080010 /* '\x01\x08\x00\x10' */
    case FilterEffectTypeWheel = 0x1080011 /* '\x01\x08\x00\x11' */
    case FilterEffectTypeWipe = 0x1080012 /* '\x01\x08\x00\x12' */
}

// MARK: PowerPointMsoAnimFilterEffectSubtype
@objc public enum PowerPointMsoAnimFilterEffectSubtype : AEKeyword {
    case NoEffectSubtype = 0x1090000 /* '\x01\t\x00\x00' */
    case FilterEffectSubtypeInVertical = 0x1090001 /* '\x01\t\x00\x01' */
    case FilterEffectSubtypeOutVertical = 0x1090002 /* '\x01\t\x00\x02' */
    case FilterEffectSubtypeInHorizontal = 0x1090003 /* '\x01\t\x00\x03' */
    case FilterEffectSubtypeOutHorizontal = 0x1090004 /* '\x01\t\x00\x04' */
    case FilterEffectSubtypeHorizontal = 0x1090005 /* '\x01\t\x00\x05' */
    case FilterEffectSubtypeVertical = 0x1090006 /* '\x01\t\x00\x06' */
    case FilterEffectSubtypeInward = 0x1090007 /* '\x01\t\x00\x07' */
    case FilterEffectSubtypeOut = 0x1090008 /* '\x01\t\x00\x08' */
    case FilterEffectSubtypeAcross = 0x1090009 /* '\x01\t\x00\t' */
    case FilterEffectSubtypeFromLeft = 0x109000a /* '\x01\t\x00\n' */
    case FilterEffectSubtypeFromRight = 0x109000b /* '\x01\t\x00\x0b' */
    case FilterEffectSubtypeFromTop = 0x109000c /* '\x01\t\x00\x0c' */
    case FilterEffectSubtypeFromBottom = 0x109000d /* '\x01\t\x00\r' */
    case FilterEffectSubtypeDownLeft = 0x109000e /* '\x01\t\x00\x0e' */
    case FilterEffectSubtypeUpLeft = 0x109000f /* '\x01\t\x00\x0f' */
    case FilterEffectSubtypeDownRight = 0x1090010 /* '\x01\t\x00\x10' */
    case FilterEffectSubtypeUpRight = 0x1090011 /* '\x01\t\x00\x11' */
    case FilterEffectSubtypeSpoke1 = 0x1090012 /* '\x01\t\x00\x12' */
    case FilterEffectSubtypeSpokes2 = 0x1090013 /* '\x01\t\x00\x13' */
    case FilterEffectSubtypeSpokes3 = 0x1090014 /* '\x01\t\x00\x14' */
    case FilterEffectSubtypeSpokes4 = 0x1090015 /* '\x01\t\x00\x15' */
    case FilterEffectSubtypeSpokes8 = 0x1090016 /* '\x01\t\x00\x16' */
    case FilterEffectSubtypeLeft = 0x1090017 /* '\x01\t\x00\x17' */
    case FilterEffectSubtypeRight = 0x1090018 /* '\x01\t\x00\x18' */
    case FilterEffectSubtypeDown = 0x1090019 /* '\x01\t\x00\x19' */
    case FilterEffectSubtypeUp = 0x109001a /* '\x01\t\x00\x1a' */
}

// MARK: PowerPointGetPlayerFrom
@objc public enum PowerPointGetPlayerFrom : AEKeyword {
    case View = 0x70564557 /* 'pVEW' */
    case SlideShowView = 0x50535376 /* 'PSSv' */
}

// MARK: PowerPointPasteObject
@objc public enum PowerPointPasteObject : AEKeyword {
    case View = 0x70564557 /* 'pVEW' */
    case Presentation = 0x70707450 /* 'pptP' */
}

// MARK: PowerPointApplyTheme
@objc public enum PowerPointApplyTheme : AEKeyword {
    case Slide = 0x70534c44 /* 'pSLD' */
    case Master = 0x704d7472 /* 'pMtr' */
    case Presentation = 0x70707450 /* 'pptP' */
}

// MARK: PowerPointOneColorGradient
@objc public enum PowerPointOneColorGradient : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointTwoColorGradient
@objc public enum PowerPointTwoColorGradient : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointAutomaticLength
@objc public enum PowerPointAutomaticLength : AEKeyword {
    case Callout = 0x63443030 /* 'cD00' */
    case CalloutFormat = 0x63436f46 /* 'cCoF' */
}

// MARK: PowerPointBeginConnect
@objc public enum PowerPointBeginConnect : AEKeyword {
    case Connector = 0x63443031 /* 'cD01' */
    case ConnectorFormat = 0x70437846 /* 'pCxF' */
}

// MARK: PowerPointBeginDisconnect
@objc public enum PowerPointBeginDisconnect : AEKeyword {
    case Connector = 0x63443031 /* 'cD01' */
    case ConnectorFormat = 0x70437846 /* 'pCxF' */
}

// MARK: PowerPointCustomLength
@objc public enum PowerPointCustomLength : AEKeyword {
    case Callout = 0x63443030 /* 'cD00' */
    case CalloutFormat = 0x63436f46 /* 'cCoF' */
}

// MARK: PowerPointCustomDrop
@objc public enum PowerPointCustomDrop : AEKeyword {
    case Callout = 0x63443030 /* 'cD00' */
    case CalloutFormat = 0x63436f46 /* 'cCoF' */
}

// MARK: PowerPointEndConnect
@objc public enum PowerPointEndConnect : AEKeyword {
    case Connector = 0x63443031 /* 'cD01' */
    case ConnectorFormat = 0x70437846 /* 'pCxF' */
}

// MARK: PowerPointEndDisconnect
@objc public enum PowerPointEndDisconnect : AEKeyword {
    case Connector = 0x63443031 /* 'cD01' */
    case ConnectorFormat = 0x70437846 /* 'pCxF' */
}

// MARK: PowerPointPatterned
@objc public enum PowerPointPatterned : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointGetActionSettingFor
@objc public enum PowerPointGetActionSettingFor : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ShapeRange = 0x53687052 /* 'ShpR' */
}

// MARK: PowerPointSolid
@objc public enum PowerPointSolid : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointResetRotation
@objc public enum PowerPointResetRotation : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ThreeDFormat = 0x44334466 /* 'D3Df' */
}

// MARK: PowerPointUserPicture
@objc public enum PowerPointUserPicture : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointUserTextured
@objc public enum PowerPointUserTextured : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointZOrder
@objc public enum PowerPointZOrder : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ShapeRange = 0x53687052 /* 'ShpR' */
}

// MARK: PowerPointPresetTextured
@objc public enum PowerPointPresetTextured : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointPresetGradient
@objc public enum PowerPointPresetGradient : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case FillFormat = 0x7046466d /* 'pFFm' */
}

// MARK: PowerPointApply
@objc public enum PowerPointApply : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ShapeRange = 0x53687052 /* 'ShpR' */
}

// MARK: PowerPointFlip
@objc public enum PowerPointFlip : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ShapeRange = 0x53687052 /* 'ShpR' */
}

// MARK: PowerPointPickUp
@objc public enum PowerPointPickUp : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ShapeRange = 0x53687052 /* 'ShpR' */
}

// MARK: PowerPointRerouteConnections
@objc public enum PowerPointRerouteConnections : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ShapeRange = 0x53687052 /* 'ShpR' */
}

// MARK: PowerPointSetShapesDefaultProperties
@objc public enum PowerPointSetShapesDefaultProperties : AEKeyword {
    case Shape = 0x70536870 /* 'pShp' */
    case ShapeRange = 0x53687052 /* 'ShpR' */
}

// MARK: PowerPointGenericMethods
@objc public protocol PowerPointGenericMethods {
    optional func closeSaving(saving: PowerPointSaveOptions, savingIn: AnyObject!) // Close a document.
    optional func saveIn(in_: AnyObject!, `as`: AnyObject!) // Save a document.
    optional func printWithProperties(withProperties: AnyObject!, printDialog: AnyObject!) // Print a document.
    optional func delete() // Delete an object.
    optional func duplicateTo(to: AnyObject!, withProperties: AnyObject!) // Copy an object.
    optional func moveTo(to: AnyObject!) // Move an object to a new location.
    optional func canCheckOutFileName(fileName: AnyObject!) // Returns True if PowerPoint can check out a specified presentation from a server.
    optional func checkOutFileName(fileName: AnyObject!) // Copies a specified presentation from a server to a local computer for editing. Returns a String that represents the local path and file name of the presentation checked out.
    optional func quit()
}

// MARK: PowerPointApplication
@objc public protocol PowerPointApplication: SBApplicationProtocol {
    optional func documents()
    optional func windows()
    optional var name: Int { get } // The name of the application.
    optional var frontmost: Int { get } // Is this the active application?
    optional var version: Int { get } // The version number of the application.
    optional func open(x: AnyObject!) -> AnyObject // Open a document.
    optional func print(x: AnyObject!, withProperties: AnyObject!, printDialog: AnyObject!) // Print a document.
    optional func quitSaving(saving: PowerPointSaveOptions) // Quit the application.
    optional func exists(x: AnyObject!) // Verify that an object exists.
    optional func reset(x: AnyObject!) // Resets a built-in command bar or command bar control to its default configuration.
    optional func applyTheme(x: AnyObject!, fileName: AnyObject!) // Applies a theme or design template to the specified slide, master or presentation
    optional func arrangeWindows(x: PowerPointEPPArrangeStyle) // Arrange Document Windows
    optional func getPlayerFrom(x: AnyObject!, `for` for_: PowerPointShape!) -> PowerPointPlayer // get a player from a shape inside a slide show view
    optional func insertTheText(theText: AnyObject!, at: AnyObject!)
    optional func pasteObject(x: AnyObject!)
    optional func registerAddIn(x: AnyObject!) -> PowerPointAddIn
    optional func runVBMacroMacroName(macroName: AnyObject!, listOfParameters: AnyObject!) // Runs a Visual Basic macro.
    optional func apply(x: AnyObject!)
    optional func automaticLength(x: AnyObject!)
    optional func beginConnect(x: AnyObject!, connectedShape: PowerPointShape!, connectionSite: AnyObject!)
    optional func beginDisconnect(x: AnyObject!)
    optional func customDrop(x: AnyObject!, dropAmount: Double)
    optional func customLength(x: AnyObject!, length: Double)
    optional func endConnect(x: AnyObject!, connectedShape: PowerPointShape!, connectionSite: AnyObject!)
    optional func endDisconnect(x: AnyObject!)
    optional func flip(x: AnyObject!, direction: PowerPointMsoFlipCmd)
    optional func getActionSettingFor(x: AnyObject!, event: PowerPointEPPMouseActivation) -> PowerPointActionSetting
    optional func oneColorGradient(x: AnyObject!, style: PowerPointMsoGradientStyle, variant: AnyObject!, degree: Double)
    optional func patterned(x: AnyObject!, pattern: PowerPointMsoPatternType)
    optional func pickUp(x: AnyObject!)
    optional func presetGradient(x: AnyObject!, style: PowerPointMsoGradientStyle, variant: AnyObject!, gradientType: PowerPointMsoPresetGradientType)
    optional func presetTextured(x: AnyObject!, texture: PowerPointMsoPresetTexture)
    optional func rerouteConnections(x: AnyObject!)
    optional func resetRotation(x: AnyObject!) // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
    optional func setShapesDefaultProperties(x: AnyObject!)
    optional func solid(x: AnyObject!)
    optional func twoColorGradient(x: AnyObject!, style: PowerPointMsoGradientStyle, variant: AnyObject!)
    optional func userPicture(x: AnyObject!, pictureFile: AnyObject!)
    optional func userTextured(x: AnyObject!, textureFile: AnyObject!)
    optional func zOrder(x: AnyObject!, zOrderPosition: PowerPointMsoZOrderCmd)
    optional func presentations()
    optional func documentWindows()
    optional func slideShowWindows()
    optional func commandBars()
    optional func addIns()
    optional var Version: Int { get }
    optional var activePresentation: PowerPointPresentation { get }
    optional var activePrinter: Int { get }
    optional var activeWindow: PowerPointDocumentWindow { get }
    optional var autocorrectObject: PowerPointAutocorrect { get } // Returns the autocorrect object
    optional var automationSecurity: PowerPointMsoAutomationSecurity { get set }
    optional var build: Int { get }
    optional var caption: Int { get }
    optional var defaultSaveFormat: PowerPointEPPSaveAsFileType { get set }
    optional var defaultWebOptionsObject: PowerPointDefaultWebOptions { get }
    optional var path: Int { get }
    optional var preferenceObject: PowerPointPreferences { get }
    optional var saveAsMovieSettingsObject: PowerPointSaveAsMovieSettings { get }
    optional var startUpDialog: Int { get set }
}
extension SBApplication: PowerPointApplication {}

// MARK: PowerPointDocument
@objc public protocol PowerPointDocument: SBObjectProtocol, PowerPointGenericMethods {
    optional var name: Int { get } // Its name.
    optional var modified: Int { get } // Has it been modified since the last save?
    optional var file: Int { get } // Its location on disk, if it has one.
}
extension SBObject: PowerPointDocument {}

// MARK: PowerPointWindow
@objc public protocol PowerPointWindow: SBObjectProtocol, PowerPointGenericMethods {
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
    optional var document: PowerPointDocument { get } // The document whose contents are displayed in the window.
    optional var entryIndex: Int { get } // the number of the window
    optional var position: Int { get set } // upper left coordinates of the window
    optional var titled: Int { get } // Does the window have a title bar?
    optional var floating: Int { get } // Does the window float?
    optional var modal: Int { get } // Is the window modal?
    optional var collapsable: Int { get } // Is the window collapasable?
    optional var collapsed: Int { get set } // Is the window collapsed?
    optional var sheet: Int { get } // Is this window a sheet window?
}
extension SBObject: PowerPointWindow {}

// MARK: PowerPointCommandBarControl
@objc public protocol PowerPointCommandBarControl: SBObjectProtocol, PowerPointGenericMethods {
    optional var beginGroup: Int { get set } // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
    optional var builtIn: Int { get } // Returns true if the command bar control is a built-in command bar control.
    optional var controlType: PowerPointMsoControlType { get } // Returns the type of the command bar control.
    optional var descriptionText: Int { get set } // Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control.
    optional var enabled: Int { get set } // Returns or sets if the command bar control is enabled.
    optional var entry_index: Int { get } // Returns the index number for this command bar control.
    optional var height: Int { get set } // Returns or sets the height of a command bar control.
    optional var helpContextID: Int { get set } // Returns or sets the help context ID number for the Help topic attached to the command bar control.
    optional var helpFile: Int { get set } // Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property.
    optional func id() // Returns the id for a built-in command bar control.
    optional var leftPosition: Int { get } // Returns the left position of the command bar control.
    optional var name: Int { get set } // Returns or sets the caption text for a command bar control.
    optional var parameter: Int { get set } // Returns or sets a string that is used to execute a command.
    optional var priority: Int { get set } // Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7.
    optional var tag: Int { get set } // Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control.
    optional var tooltipText: Int { get set } // Returns or sets the text displayed in a command bar controls tooltip.
    optional var top: Int { get } // Returns the top position of a command bar control.
    optional var visible: Int { get set } // Returns or sets if the command bar control is visible.
    optional var width: Int { get set } // Returns or sets the width in pixels of the command bar control.
    optional func execute() // Runs the procedure or built-in command assigned to the specified command bar control.
}
extension SBObject: PowerPointCommandBarControl {}

// MARK: PowerPointCommandBarButton
@objc public protocol PowerPointCommandBarButton: PowerPointCommandBarControl {
    optional var buttonFaceIsDefault: Int { get } // Returns if the face of a command bar button control is the original built-in face.
    optional var buttonState: PowerPointMsoButtonState { get set } // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
    optional var buttonStyle: PowerPointMsoButtonStyle { get set } // Returns or sets the way a command button control is displayed.
    optional var faceId: Int { get set } // Returns or sets the Id number for the face of the command bar button control.
}
extension SBObject: PowerPointCommandBarButton {}

// MARK: PowerPointCommandBarCombobox
@objc public protocol PowerPointCommandBarCombobox: PowerPointCommandBarControl {
    optional var comboboxStyle: PowerPointMsoComboStyle { get set } // Returns or sets the way a command bar combobox control is displayed.
    optional var comboboxText: Int { get set } // Returns or sets the text in the display or edit portion of the command bar combobox control.
    optional var dropDownLines: Int { get set } // Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control.
    optional var dropDownWidth: Int { get set } // Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control.
    optional var listIndex: Int { get set }
    optional func addItemToComboboxComboboxItem(comboboxItem: AnyObject!, entry_index: AnyObject!) // Add a new string to a custom combobox control.
    optional func clearCombobox() // Clear all of the strings form a custom combobox.
    optional func getComboboxItemEntry_index(entry_index: AnyObject!) // Return the string at the given index within a combobox.
    optional func getCountOfComboboxItems() // Return the number of strings within a combobox.
    optional func removeAnItemFromComboboxEntry_index(entry_index: AnyObject!) // Remove a string from a custom combobox.
    optional func setComboboxItemEntry_index(entry_index: AnyObject!, comboboxItem: AnyObject!) // Set the string an a given index for a custom combobox.
}
extension SBObject: PowerPointCommandBarCombobox {}

// MARK: PowerPointCommandBarPopup
@objc public protocol PowerPointCommandBarPopup: PowerPointCommandBarControl {
    optional func commandBarControls()
}
extension SBObject: PowerPointCommandBarPopup {}

// MARK: PowerPointCommandBar
@objc public protocol PowerPointCommandBar: SBObjectProtocol, PowerPointGenericMethods {
    optional func commandBarControls()
    optional var barPosition: PowerPointMsoBarPosition { get set } // Returns or sets the position of the command bar.
    optional var barType: PowerPointMsoBarType { get } // Returns the type of this command bar.
    optional var builtIn: Int { get } // True if the command bar is built-in.
    optional var context: Int { get } // Returns or sets a string that determines where a command bar will be saved.
    optional var embeddable: Int { get } // Returns if the command bar can be displayed inside the document window.
    optional var embedded: Int { get set } // Returns or sets if the command bar will be displayed inside the document window.
    optional var enabled: Int { get set } // Returns or set if the command bar is enabled.
    optional var entry_index: Int { get } // The index of the command bar.
    optional var height: Int { get set } // Returns or sets the height of the command bar.
    optional var leftPosition: Int { get set } // Returns or sets the left position of the command bar.
    optional var localName: Int { get set } // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
    optional var name: Int { get set } // Returns or sets the name of the command bar.
    optional var rowIndex: Int { get set } // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
    optional var top: Int { get set } // Returns or sets the top position of a command bar.
    optional var visible: Int { get set } // Returns or sets if the command bar is visible.
    optional var width: Int { get set } // Returns or sets the width in pixels of the command bar.
}
extension SBObject: PowerPointCommandBar {}

// MARK: PowerPointDocumentProperty
@objc public protocol PowerPointDocumentProperty: SBObjectProtocol, PowerPointGenericMethods {
    optional var documentPropertyType: Int { get set } // Returns or sets the document property type.
    optional var linkSource: Int { get set } // Returns or sets the source of a lined custom document property.
    optional var linkToContent: Int { get set } // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
    optional var name: Int { get set } // Returns or sets the name of the document property.
    optional var value: Int { get set } // Returns or sets the value of the document property.
}
extension SBObject: PowerPointDocumentProperty {}

// MARK: PowerPointCustomDocumentProperty
@objc public protocol PowerPointCustomDocumentProperty: PowerPointDocumentProperty {
}
extension SBObject: PowerPointCustomDocumentProperty {}

// MARK: PowerPointWebPageFont
@objc public protocol PowerPointWebPageFont: SBObjectProtocol, PowerPointGenericMethods {
    optional var fixedWidthFont: Int { get set } // Returns or sets the fixed-width font setting.
    optional var fixedWidthFontSize: Double { get set } // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
    optional var proportionalFont: Int { get set } // Returns or sets the proportional font setting.
    optional var proportionalFontSize: Double { get set } // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
}
extension SBObject: PowerPointWebPageFont {}

// MARK: PowerPointActionSetting
@objc public protocol PowerPointActionSetting: SBObjectProtocol, PowerPointGenericMethods {
    optional var action: PowerPointEPPActionType { get set }
    optional var actionSettingToRun: Int { get set }
    optional var actionSoundEffect: PowerPointSoundEffect { get }
    optional var actionVerb: Int { get set }
    optional var animateAction: Int { get set }
    optional var hyperlink: PowerPointHyperlink { get }
    optional var slideShowName: Int { get set }
}
extension SBObject: PowerPointActionSetting {}

// MARK: PowerPointAddIn
@objc public protocol PowerPointAddIn: SBObjectProtocol, PowerPointGenericMethods {
    optional var autoLoad: Int { get set }
    optional var fullName: Int { get }
    optional var loaded: Int { get set }
    optional var name: Int { get }
    optional var path: Int { get }
    optional var registered: Int { get set }
    optional var registeredInHKLM: Int { get }
}
extension SBObject: PowerPointAddIn {}

// MARK: PowerPointAnimationBehavior
@objc public protocol PowerPointAnimationBehavior: SBObjectProtocol, PowerPointGenericMethods {
    optional var accumulate: PowerPointMsoAnimAccumulate { get set }
    optional var additive: PowerPointMsoAnimAdditive { get set }
    optional var animationBehaviorType: PowerPointMsoAnimType { get set }
    optional var colorsEffect: PowerPointColorsEffect { get }
    optional var commandEffect: PowerPointCommandEffect { get }
    optional var filterEffect: PowerPointFilterEffect { get }
    optional var motionEffect: PowerPointMotionEffect { get }
    optional var propertyEffect: PowerPointPropertyEffect { get }
    optional var rotatingEffect: PowerPointRotatingEffect { get }
    optional var scaleEffect: PowerPointScaleEffect { get }
    optional var setEffect: PowerPointSetEffect { get }
    optional var timing: PowerPointTiming { get }
}
extension SBObject: PowerPointAnimationBehavior {}

// MARK: PowerPointAnimationPoint
@objc public protocol PowerPointAnimationPoint: SBObjectProtocol, PowerPointGenericMethods {
    optional var formula: Int { get set }
    optional var time: Double { get set }
    optional var value: Int { get set }
}
extension SBObject: PowerPointAnimationPoint {}

// MARK: PowerPointAnimationSettings
@objc public protocol PowerPointAnimationSettings: SBObjectProtocol, PowerPointGenericMethods {
    optional var advanceTime: Double { get set }
    optional var afterEffect: PowerPointEPPAfterEffect { get set }
    optional var animate: Int { get set }
    optional var animateBackground: Int { get set }
    optional var animateTextInReverse: Int { get set }
    optional var animationOrder: Int { get set }
    optional var animationPlaySettings: PowerPointPlaySettings { get }
    optional var animationSoundEffect: PowerPointSoundEffect { get }
    optional var chartUnitEffect: PowerPointEPPChartUnitEffect { get set }
    optional var dimColorThemeIndex: PowerPointMsoThemeColorIndex { get set }
    optional var entryEffect: PowerPointEPPEntryEffect { get set }
    optional var textLevelEffect: PowerPointEPPTextLevelEffect { get set }
    optional var textUnitEffect: PowerPointEPPTextUnitEffect { get set }
}
extension SBObject: PowerPointAnimationSettings {}

// MARK: PowerPointAutocorrectEntry
@objc public protocol PowerPointAutocorrectEntry: SBObjectProtocol, PowerPointGenericMethods {
    optional var autocorrectValue: Int { get } // Returns the value of the auto correct entry.
    optional var entry_index: Int { get } // Returns the index for the position of the object in its container element list.
    optional var name: Int { get } // Returns the name of the auto correct entry.
}
extension SBObject: PowerPointAutocorrectEntry {}

// MARK: PowerPointAutocorrect
@objc public protocol PowerPointAutocorrect: SBObjectProtocol, PowerPointGenericMethods {
    optional func autocorrectEntries()
    optional func firstLetterExceptions()
    optional func twoInitialCapsExceptions()
    optional var correctDays: Int { get set } // Returns if PowerPoint automatically capitalizes the first letter of days of the week.
    optional var correctInitialCaps: Int { get set } // Returns if PowerPoint automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, POwerPoint is corrected to PowerPoint.
    optional var correctSentenceCaps: Int { get set } // Returns if PowerPoint automatically capitalizes the first letter in each sentence.
    optional var replaceText: Int { get set } // Returns if Microsoft PowerPoint automatically replaces specified text with entries from the autocorrect list.
}
extension SBObject: PowerPointAutocorrect {}

// MARK: PowerPointBroadcast
@objc public protocol PowerPointBroadcast: SBObjectProtocol, PowerPointGenericMethods {
    optional func endSession() // End a running broadcast session.
    optional func getAttendeeURL()
    optional func getIsBroadcasting() // Returns true if the current presentation is being broadcast.
    optional func startServerUrl(serverUrl: AnyObject!) // Starts broadcasting to the given URL.
}
extension SBObject: PowerPointBroadcast {}

// MARK: PowerPointBulletFormat
@objc public protocol PowerPointBulletFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var bulletCharacter: Int { get set } // Returns or sets the Unicode character value that is used for bullets in the specified text.
    optional var bulletFont: PowerPointFont { get } // Returns a font object that represents character formatting for a bullet format object.
    optional var bulletNumber: Int { get } // Returns the bullet number of a paragraph.
    optional var bulletStartValue: Int { get set }
    optional var bulletStyle: PowerPointMsoNumberedBulletStyle { get set } // Returns or sets a constant that represents the style of a bullet.
    optional var bulletType: PowerPointMsoBulletType { get set } // Returns or sets a constant that represents the type of bullet.
    optional var relativeSize: Double { get set } // Returns or sets the bullet size relative to the size of the first text character in the paragraph.
    optional var useTextColor: Int { get set } // Determines whether the specified bullets are set to the color of the first text character in the paragraph.
    optional var useTextFont: Int { get set } // Determines whether the specified bullets are set to the font of the first text character in the paragraph.
    optional var visible: Int { get set } // Returns or sets a value that specifies whether the bullet is visible.
    optional func setBulletPicturePictureFile(pictureFile: AnyObject!) // Sets the graphics file to be used for bullets in a bulleted list.
}
extension SBObject: PowerPointBulletFormat {}

// MARK: PowerPointColorScheme
@objc public protocol PowerPointColorScheme: SBObjectProtocol, PowerPointGenericMethods {
    optional func getColorFromAt(at: PowerPointEPPColorSchemeIndex)
    optional func setColorForAt(at: PowerPointEPPColorSchemeIndex, toColor: AnyObject!)
}
extension SBObject: PowerPointColorScheme {}

// MARK: PowerPointColorsEffect
@objc public protocol PowerPointColorsEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var by_colorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns an object that represents a change to the color of the object by the specified number, expressed in RGB format.
    optional var from_colorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets an object that represents the starting RGB color value of an animation behavior.
    optional var to_colorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets an object that represents the RGB color value of an animation behavior.
}
extension SBObject: PowerPointColorsEffect {}

// MARK: PowerPointCommandEffect
@objc public protocol PowerPointCommandEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var command: Int { get set }
    optional var type: PowerPointMsoAnimCommandType { get set }
}
extension SBObject: PowerPointCommandEffect {}

// MARK: PowerPointCustomLayout
@objc public protocol PowerPointCustomLayout: SBObjectProtocol, PowerPointGenericMethods {
    optional func shapes()
    optional var background: PowerPointShape { get }
    optional var design: PowerPointDesign { get }
    optional var displayMasterShapes: Int { get set }
    optional var followMasterBackground: Int { get set }
    optional var headersAndFooters: PowerPointHeadersAndFooters { get }
    optional var height: Double { get }
    optional var themeColorScheme: PowerPointThemeColorScheme { get } // Returns the color scheme of a custom layout.
    optional var timeline: PowerPointTimeline { get }
    optional var width: Double { get }
}
extension SBObject: PowerPointCustomLayout {}

// MARK: PowerPointDefaultWebOptions
@objc public protocol PowerPointDefaultWebOptions: SBObjectProtocol, PowerPointGenericMethods {
    optional var allowPNG: Int { get set }
    optional var alwaysSaveInDefaultEncoding: Int { get set }
    optional var checkIfOfficeIsHTMLEditor: Int { get set }
    optional var encoding: PowerPointMsoEncoding { get set }
    optional var updateLinksOnSave: Int { get set }
}
extension SBObject: PowerPointDefaultWebOptions {}

// MARK: PowerPointDesign
@objc public protocol PowerPointDesign: SBObjectProtocol, PowerPointGenericMethods {
    optional var slideMaster: PowerPointMaster { get }
}
extension SBObject: PowerPointDesign {}

// MARK: PowerPointDocumentWindow
@objc public protocol PowerPointDocumentWindow: SBObjectProtocol, PowerPointGenericMethods {
    optional func panes()
    optional var active: Int { get }
    optional var activePane: PowerPointPane { get }
    optional var blackAndWhite: Int { get set }
    optional var caption: Int { get }
    optional var entry_index: Int { get }
    optional var height: Double { get set }
    optional var leftPosition: Double { get set }
    optional var presentation: PowerPointPresentation { get }
    optional var selection: PowerPointSelection { get } // Represents the selection in the specified document window.
    optional var splitHorizontal: Int { get set }
    optional var splitVertical: Int { get set }
    optional var top: Double { get set }
    optional var view: PowerPointView { get }
    optional var viewType: PowerPointEPPViewType { get set }
    optional var width: Double { get set }
    optional func collapseSectionAtPosition(atPosition: AnyObject!)
    optional func expandSectionAtPosition(atPosition: AnyObject!)
    optional func getIsExpandedOfSectionAtPosition(atPosition: AnyObject!)
    optional func launchSpellerOn()
}
extension SBObject: PowerPointDocumentWindow {}

// MARK: PowerPointEffectInformation
@objc public protocol PowerPointEffectInformation: SBObjectProtocol, PowerPointGenericMethods {
    optional var afterEffectInformation: PowerPointMsoAnimAfterEffect { get }
    optional var animateBackgroundInformation: Int { get }
    optional var animateTextInReverseInformation: Int { get }
    optional var buildByLevel: PowerPointMsoAnimateByLevel { get }
    optional var playSettingsInformation: PowerPointPlaySettings { get }
    optional var soundEffectInformation: PowerPointSoundEffect { get }
    optional var textUnitEffectInformation: PowerPointMsoAnimTextUnitEffect { get }
}
extension SBObject: PowerPointEffectInformation {}

// MARK: PowerPointEffectParameters
@objc public protocol PowerPointEffectParameters: SBObjectProtocol, PowerPointGenericMethods {
    optional var amount: Double { get set }
    optional var color2ColorThemeIndex: PowerPointMsoThemeColorIndex { get } // Returns an object that represents the color on which to end a color-cycle animation.
    optional var direction: PowerPointMsoAnimDirection { get set }
    optional var font2: Int { get set }
    optional var relative: Int { get set }
    optional var size2: Double { get set }
}
extension SBObject: PowerPointEffectParameters {}

// MARK: PowerPointEffect
@objc public protocol PowerPointEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional func animationBehaviors()
    optional var animationEffectType: PowerPointMsoAnimEffect { get set }
    optional var effectInformation: PowerPointEffectInformation { get }
    optional var effectParameters: PowerPointEffectParameters { get }
    optional var exitAnimation: Int { get set }
    optional var name: Int { get }
    optional var shape: PowerPointShape { get set }
    optional var targetParagraph: Int { get set }
    optional var textRangeLength: Int { get }
    optional var textRangeStart: Int { get }
    optional var timing: PowerPointTiming { get }
    optional func addBehaviorType(type: PowerPointMsoAnimType) -> PowerPointAnimationBehavior // add an animation behavior
}
extension SBObject: PowerPointEffect {}

// MARK: PowerPointFilterEffect
@objc public protocol PowerPointFilterEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var filterType: PowerPointMsoAnimFilterEffectType { get set }
    optional var reveal: Int { get set }
    optional var subtype: PowerPointMsoAnimFilterEffectSubtype { get set }
}
extension SBObject: PowerPointFilterEffect {}

// MARK: PowerPointFirstLetterException
@objc public protocol PowerPointFirstLetterException: SBObjectProtocol, PowerPointGenericMethods {
    optional var entry_index: Int { get } // Returns the index for the position of the object in its container element list.
    optional var name: Int { get } // Returns the name of the FirstLetterException.
}
extension SBObject: PowerPointFirstLetterException {}

// MARK: PowerPointFont
@objc public protocol PowerPointFont: SBObjectProtocol, PowerPointGenericMethods {
    optional var ASCIIName: Int { get set } // Returns or sets the font used for Latin text; characters with character codes from 0 through 127.
    optional var autoRotateNumbers: Int { get set } // Returns or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.
    optional var baseLineOffset: Double { get set } // Returns or sets a value specifying the horizontaol offset of the selected font.
    optional var bold: Int { get set } // Returns or sets a value specifying whether the font should be bold.
    optional var capsType: PowerPointMsoTextCaps { get set } // Returns or sets a value specifying how the text should be capitalized.
    optional var eastAsianName: Int { get set } // Returns or sets the font name used for Asian text.
    optional var embedable: Int { get } // Returns a value indicating whether the font can be embedded in a page.
    optional var embedded: Int { get } // Returns a value specifying whether the font is embedded in a page.
    optional var emboss: Int { get set }
    optional var equalizeCharacterHeight: Int { get set } // Returns or sets a value specifying whether the text should have the same horizontal height.
    optional var fillFormat: PowerPointFillFormat { get } // Returns a value specifying the fill formatting for text.
    optional var fontColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the color for the specified font.
    optional var fontName: Int { get set } // Returns or sets a value specifying the font to use for a selection.
    optional var fontNameOther: Int { get set } // Returns or sets the font used for characters whose character set numbers are greater than 127.
    optional var fontSize: Double { get set }
    optional var glowFormat: PowerPointGlowFormat { get } // Returns a value specifying the glow formatting of the text.
    optional var highlightColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the specified text highlight color for object.
    optional var italic: Int { get set }
    optional var kerning: Double { get set } // Returns or sets a value specifying the amount of spacing between text characters.
    optional var lineFormat: PowerPointLineFormat { get } // Returns a value specifiying the line formatting of the text.
    optional var reflectionFormat: PowerPointReflectionFormat { get } // Returns a value specifying the reflection formatting of the text.
    optional var shadowFormat: PowerPointShadowFormat { get } // Returns the value specifying the type of shadow effect for the selection of text.
    optional var softEdgeFormat: PowerPointMsoSoftEdgeType { get set } // Returns or sets a value specifying the soft edge fromatting of the text.
    optional var spacing: Double { get set } // Returns or sets a value specifying the spacing between characters in a selection of text.
    optional var strikeType: PowerPointMsoTextStrike { get set } // Returns or sets a value specifying the strike format used for a selection of text.
    optional var subscript: Int { get set } // Returns or sets a value specifying that the selected text should be displayed a subscript.
    optional var superscript: Int { get set } // Returns or sets a value specifying that the selected text should be displayed a superscript.
    optional var transparency: Double { get set }
    optional var underline: Int { get set }
    optional var underlineColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns a value specifying the color of the underline for the selected text.
    optional var underlineStyle: PowerPointMsoTextUnderlineType { get set } // Returns or sets a value specifying the underline style for the selected text.
    optional var wordArtStylesFormat: PowerPointMsoPresetTextEffect { get set } // Returns or sets a value specifying the text effect for the selected text.
}
extension SBObject: PowerPointFont {}

// MARK: PowerPointHeaderOrFooter
@objc public protocol PowerPointHeaderOrFooter: SBObjectProtocol, PowerPointGenericMethods {
    optional var dateFormat: PowerPointEPPDateTimeFormat { get set }
    optional var headerFooterText: Int { get set }
    optional var useDateFormat: Int { get set }
    optional var visible: Int { get set }
}
extension SBObject: PowerPointHeaderOrFooter {}

// MARK: PowerPointHeadersAndFooters
@objc public protocol PowerPointHeadersAndFooters: SBObjectProtocol, PowerPointGenericMethods {
    optional var dateAndTime: PowerPointHeaderOrFooter { get }
    optional var displayHeadersFootersOnTitleSlide: Int { get set }
    optional var footer: PowerPointHeaderOrFooter { get }
    optional var header: PowerPointHeaderOrFooter { get }
    optional var slideNumber: PowerPointHeaderOrFooter { get }
}
extension SBObject: PowerPointHeadersAndFooters {}

// MARK: PowerPointHyperlink
@objc public protocol PowerPointHyperlink: SBObjectProtocol, PowerPointGenericMethods {
    optional var hyperlinkAddress: Int { get set }
    optional var hyperlinkSubAddress: Int { get set }
    optional var hyperlinkType: AnyObject { get }
}
extension SBObject: PowerPointHyperlink {}

// MARK: PowerPointMaster
@objc public protocol PowerPointMaster: SBObjectProtocol, PowerPointGenericMethods {
    optional func shapes()
    optional func hyperlinks()
    optional func customLayouts()
    optional var background: PowerPointShape { get }
    optional var colorScheme: PowerPointColorScheme { get }
    optional var design: PowerPointDesign { get }
    optional var headersAndFooters: PowerPointHeadersAndFooters { get }
    optional var height: Double { get }
    optional var theme: PowerPointOfficeTheme { get }
    optional var timeline: PowerPointTimeline { get }
    optional var width: Double { get }
    optional func getTextStyleFromAt(at: PowerPointPpTextStyleType) -> PowerPointTextStyle
}
extension SBObject: PowerPointMaster {}

// MARK: PowerPointMotionEffect
@objc public protocol PowerPointMotionEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var byX: Double { get set }
    optional var byY: Double { get set }
    optional var fromX: Double { get set }
    optional var fromY: Double { get set }
    optional var path: Int { get set }
    optional var toX: Double { get set }
    optional var toY: Double { get set }
}
extension SBObject: PowerPointMotionEffect {}

// MARK: PowerPointNamedSlideShow
@objc public protocol PowerPointNamedSlideShow: SBObjectProtocol, PowerPointGenericMethods {
    optional var name: Int { get }
    optional var numberOfSlides: Int { get }
}
extension SBObject: PowerPointNamedSlideShow {}

// MARK: PowerPointPageSetup
@objc public protocol PowerPointPageSetup: SBObjectProtocol, PowerPointGenericMethods {
    optional var firstSlideNumber: Int { get set }
    optional var notesOrientation: PowerPointMsoOrientation { get set }
    optional var slideOrientation: PowerPointMsoOrientation { get set }
    optional var slideSize: PowerPointEPPSlideSizeType { get set }
    optional var slideWidth: Double { get set }
}
extension SBObject: PowerPointPageSetup {}

// MARK: PowerPointPane
@objc public protocol PowerPointPane: SBObjectProtocol, PowerPointGenericMethods {
    optional var active: Int { get }
    optional var paneViewType: PowerPointEPPViewType { get }
}
extension SBObject: PowerPointPane {}

// MARK: PowerPointParagraphFormat
@objc public protocol PowerPointParagraphFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional func tabStops()
    optional var alignment: PowerPointMsoParagraphAlignment { get set }
    optional var baselineAlignment: PowerPointMsoBaselineAlignment { get set } // Returns or sets a constant that represents the vertical position of fonts in a paragraph.
    optional var bulletFormat: PowerPointBulletFormat { get }
    optional var eastAsianLineBreakControl: Int { get set }
    optional var firstLineIndent: Double { get set } // Returns or sets the value, in points, for a first line or hanging indent.
    optional var hangingPunctuation: Int { get set } // Determines whether hanging punctuation is enabled for the specified paragraphs.
    optional var indentLevel: Int { get set } // Returns or sets a value representing the indent level assigned to text in the selected paragraph.
    optional var leftIndent: Double { get set } // Returns or sets a value that represents the left indent value, in points, for the specified paragraphs.
    optional var lineRuleAfter: Int { get set } // Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.
    optional var lineRuleBefore: Int { get set } // Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.
    optional var lineRuleWithin: Int { get set } // Determines whether line spacing between base lines is set to a specific number of points or lines.
    optional var rightIndent: Double { get set } // Returns or sets the right indent, in points, for the specified paragraphs.
    optional var spaceAfter: Double { get set } // Returns or sets the spacing, in points, after the specified paragraphs.
    optional var spaceBefore: Double { get set } // Returns or sets the spacing, in points, before the specified paragraphs.
    optional var spaceWithin: Double { get set } // Returns or sets the amount of space between base lines in the specified paragraph, in points or lines.
    optional var textDirection: PowerPointMsoTextDirection { get set } // Returns or sets the text direction for the specified paragraph.
    optional var wordWrap: Int { get set } // Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.
}
extension SBObject: PowerPointParagraphFormat {}

// MARK: PowerPointPlaySettings
@objc public protocol PowerPointPlaySettings: SBObjectProtocol, PowerPointGenericMethods {
    optional var hideWhileNotPlaying: Int { get set }
    optional var loopUntilStopped: Int { get set }
    optional var pauseAnimation: Int { get set }
    optional var playOnEntry: Int { get set }
    optional var rewindMove: Int { get set }
    optional var stopAfterSlides: Int { get set }
}
extension SBObject: PowerPointPlaySettings {}

// MARK: PowerPointPlayer
@objc public protocol PowerPointPlayer: SBObjectProtocol, PowerPointGenericMethods {
    optional var currentPosition: Int { get set }
    optional var playerState: PowerPointEPPPlayerState { get }
    optional func goToNextBookmarkForPlayer() // Advance the player to the next bookmark.
    optional func goToPreviousBookmarkForPlayer() // Rewind the player to the previous bookmark.
    optional func pausePlayer() // Pause playback.
    optional func playPlayer() // Begin or resume playback.
    optional func stopPlayer() // Stop playback.
}
extension SBObject: PowerPointPlayer {}

// MARK: PowerPointPreferences
@objc public protocol PowerPointPreferences: SBObjectProtocol, PowerPointGenericMethods {
    optional var alwaysSuggestCorrections: Int { get set }
    optional var appendDOSExtentions: Int { get set }
    optional var autoFit: Int { get set }
    optional var autoRecoveryFrequency: Int { get set }
    optional var backgroundSpelling: Int { get set }
    optional var compressFile: Int { get set }
    optional var defaultView: Int { get set }
    optional var dragDrop: Int { get set }
    optional var endBlankSlide: Int { get set }
    optional var filePropertiesPrompt: Int { get set }
    optional var hideSpellingErrors: Int { get set }
    optional var ignoreNumbersInWords: Int { get set }
    optional var ignoreUppercase: Int { get set }
    optional var optionBitmap: Int { get set }
    optional var rulerUnits: Int { get set }
    optional var saveAutoRecovery: Int { get set }
    optional var showVerticalRuler: Int { get set }
    optional var slideShowControl: Int { get set }
    optional var slideShowRightMouse: Int { get set }
    optional var smartCutPaste: Int { get set }
    optional var smartQuotes: Int { get set }
    optional var undoLevelCount: Int { get set }
    optional var userInitials: Int { get set }
    optional var userName: Int { get set }
    optional var wordSelection: Int { get set }
}
extension SBObject: PowerPointPreferences {}

// MARK: PowerPointPresentation
@objc public protocol PowerPointPresentation: SBObjectProtocol, PowerPointGenericMethods {
    optional func slides()
    optional func colorSchemes()
    optional func fonts()
    optional func documentWindows()
    optional func documentProperties()
    optional func customDocumentProperties()
    optional func designs()
    optional var broadcast: PowerPointBroadcast { get }
    optional var defaultShape: PowerPointShape { get }
    optional var eastAsianLineBreakLevel: PowerPointEPPFarEastLineBreakLevel { get set } // Returns or sets the East Asian line break control level for the specified paragraph.
    optional var fullName: Int { get }
    optional var handoutMaster: PowerPointMaster { get }
    optional var hasTitleMaster: Int { get }
    optional var isProtected: Int { get } // Returns true if the presentation is protected by Information Rights Management.
    optional var layoutDirection: PowerPointMsoTextDirection { get set }
    optional var name: Int { get }
    optional var noLineBreakAfter: Int { get set }
    optional var noLineBreakBefore: Int { get set }
    optional var notesMaster: PowerPointMaster { get }
    optional var pageSetup: PowerPointPageSetup { get }
    optional var password: Int { get set } // The password for opening the presentation
    optional var path: Int { get }
    optional var printOptions: PowerPointPrintOptions { get }
    optional var readOnly: Int { get }
    optional var saveAsMovieSettings: PowerPointSaveAsMovieSettings { get }
    optional var saved: Int { get set }
    optional var sectionProperties: PowerPointSectionProperties { get }
    optional var slideMaster: PowerPointMaster { get }
    optional var slideShowSettings: PowerPointSlideShowSettings { get }
    optional var slideShowWindow: PowerPointSlideShowWindow { get }
    optional var templateName: Int { get }
    optional var titleMaster: PowerPointMaster { get }
    optional var webOptions: PowerPointWebOptions { get }
    optional var writePassword: Int { get set } // The password for saving changes to the presentation
    optional func addDesignDesignName(DesignName: AnyObject!, index: AnyObject!) -> PowerPointDesign // add a new design
    optional func applyTemplateFileName(fileName: AnyObject!) // Applies a theme or design template to the specified slide, master or presentation
    optional func canCheckIn() // Returns True if PowerPoint can check in a specified presentation to a server.
    optional func checkInSaveChanges(saveChanges: AnyObject!, comments: AnyObject!, makePublic: AnyObject!) // Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
    optional func checkInWithVersionSaveChanges(saveChanges: AnyObject!, comments: AnyObject!, makePublic: AnyObject!, versionType: PowerPointEPPCheckInVersionType) // Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
    optional func printOutFrom(from: AnyObject!, to: AnyObject!, printToFile: AnyObject!, copies: AnyObject!, collate: AnyObject!, showDialog: AnyObject!)
    optional func redoTimes(times: AnyObject!)
    optional func undoTimes(times: AnyObject!)
    optional func updateLinks()
}
extension SBObject: PowerPointPresentation {}

// MARK: PowerPointPresenterTool
@objc public protocol PowerPointPresenterTool: SBObjectProtocol, PowerPointGenericMethods {
    optional var currentPresenterSlide: PowerPointSlide { get }
    optional var currentSlideInShow: Int { get }
    optional var elapsedTime: Double { get }
    optional var maxSlidesInShow: Int { get }
    optional var notesText: Int { get set }
    optional var notesZoom: Int { get set }
    optional var slideMiniature: Int { get }
    optional func endShow()
    optional func next()
    optional func pauseTimer()
    optional func previous()
    optional func resetTimer()
    optional func startTimer()
    optional func swapDisplays()
    optional func toggleSlideMiniature()
}
extension SBObject: PowerPointPresenterTool {}

// MARK: PowerPointPresenterViewWindow
@objc public protocol PowerPointPresenterViewWindow: SBObjectProtocol, PowerPointGenericMethods {
    optional var active: Int { get }
    optional var height: Double { get }
    optional var presentation: PowerPointPresentation { get }
    optional var presenterTool: PowerPointPresenterTool { get }
    optional var width: Double { get }
}
extension SBObject: PowerPointPresenterViewWindow {}

// MARK: PowerPointPrintOptions
@objc public protocol PowerPointPrintOptions: SBObjectProtocol, PowerPointGenericMethods {
    optional func printRanges()
    optional var activePrinter: Int { get }
    optional var collate: Int { get set }
    optional var fitToPage: Int { get set }
    optional var frameSlides: Int { get set }
    optional var numberOfCopies: Int { get set }
    optional var outputType: PowerPointEPPPrintOutputType { get set }
    optional var printColorType: PowerPointEPPPrintColorType { get set }
    optional var printFontsAsGraphics: Int { get set }
    optional var printHiddenSlides: Int { get set }
    optional var rangeType: PowerPointEPPPrintRangeType { get set }
    optional var sectionIndex: Int { get set }
    optional var slideShowName: Int { get set }
}
extension SBObject: PowerPointPrintOptions {}

// MARK: PowerPointPrintRange
@objc public protocol PowerPointPrintRange: SBObjectProtocol, PowerPointGenericMethods {
    optional var rangeEnd: Int { get }
    optional var rangeStart: Int { get }
}
extension SBObject: PowerPointPrintRange {}

// MARK: PowerPointPropertyEffect
@objc public protocol PowerPointPropertyEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional func animationPoints()
    optional var ending: AnyObject { get }
    optional var propertySetEffect: PowerPointMsoAnimProperty { get set }
    optional var starting: AnyObject { get }
}
extension SBObject: PowerPointPropertyEffect {}

// MARK: PowerPointRotatingEffect
@objc public protocol PowerPointRotatingEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var rotating: Double { get set }
}
extension SBObject: PowerPointRotatingEffect {}

// MARK: PowerPointRulerLevel
@objc public protocol PowerPointRulerLevel: SBObjectProtocol, PowerPointGenericMethods {
    optional var firstMargin: Double { get set } // Returns or sets the first-line indent for the specified outline level, in points.
    optional var leftMargin: Double { get set } // Returns or sets the left indent for the specified outline level, in points.
}
extension SBObject: PowerPointRulerLevel {}

// MARK: PowerPointRuler
@objc public protocol PowerPointRuler: SBObjectProtocol, PowerPointGenericMethods {
    optional func tabStops()
    optional func rulerLevels()
}
extension SBObject: PowerPointRuler {}

// MARK: PowerPointSaveAsMovieSettings
@objc public protocol PowerPointSaveAsMovieSettings: SBObjectProtocol, PowerPointGenericMethods {
    optional var autoLoopEnabled: Int { get set }
    optional var backgroundSoundTrackFile: Int { get set }
    optional var backgroundTrackSegmentEnd: Int { get set }
    optional var backgroundTrackSegmentStart: Int { get set }
    optional var backgroundTrackStart: Int { get set }
    optional var createMoviePreview: Int { get set }
    optional var forceAllInline: Int { get set }
    optional var includeNarrationAndSounds: Int { get set }
    optional var movieActors: Int { get set }
    optional var movieAuthor: Int { get set }
    optional var movieCopyright: Int { get set }
    optional var movieFrameHeight: Int { get set }
    optional var movieFrameWidth: Int { get set }
    optional var movieProducer: Int { get set }
    optional var optimization: PowerPointEPPMovieOptimization { get set }
    optional var showMovieController: Int { get set }
    optional var transitionDescription: Int { get set }
    optional var useSingleTransition: Int { get set }
}
extension SBObject: PowerPointSaveAsMovieSettings {}

// MARK: PowerPointScaleEffect
@objc public protocol PowerPointScaleEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var byX: Double { get set }
    optional var byY: Double { get set }
    optional var fromX: Double { get set }
    optional var fromY: Double { get set }
    optional var toX: Double { get set }
    optional var toY: Double { get set }
}
extension SBObject: PowerPointScaleEffect {}

// MARK: PowerPointSectionProperties
@objc public protocol PowerPointSectionProperties: SBObjectProtocol, PowerPointGenericMethods {
    optional func deleteSectionAtPosition(atPosition: AnyObject!, deletingSlides: AnyObject!)
    optional func getCountOfSections()
    optional func getFirstSlideOfSectionAtPosition(atPosition: AnyObject!)
    optional func getIdOfSectionAtPosition(atPosition: AnyObject!)
    optional func getNameOfSectionAtPosition(atPosition: AnyObject!)
    optional func getSlideCountOfSectionAtPosition(atPosition: AnyObject!)
    optional func insertSectionBeforeSection(beforeSection: AnyObject!, beforeSlide: AnyObject!, titled: AnyObject!)
    optional func moveSectionAtPosition(atPosition: AnyObject!, toPosition: AnyObject!)
    optional func renameSectionAtPosition(atPosition: AnyObject!, to: AnyObject!)
}
extension SBObject: PowerPointSectionProperties {}

// MARK: PowerPointSelection
@objc public protocol PowerPointSelection: SBObjectProtocol, PowerPointGenericMethods {
    optional var childShapeRange: PowerPointShapeRange { get } // Returns the child shapes of a selection.
    optional var hasChildShapeRange: Int { get } // Returns whether the selection contains child shapes.
    optional var selectionType: PowerPointEPPSelectionType { get } // Represents the type of objects in a selection.
    optional var shapeRange: PowerPointShapeRange { get } // Returns a collection of shapes selected on the specified slide.
    optional var slideRange: PowerPointSlideRange { get } // Returns a collection of selected slides.
    optional var textRange: PowerPointTextRange { get } // Returns the text and text properties of the selected text.
    optional func unselect() // Cancels the current selection.
}
extension SBObject: PowerPointSelection {}

// MARK: PowerPointSequence
@objc public protocol PowerPointSequence: SBObjectProtocol, PowerPointGenericMethods {
    optional func effects()
    optional func addEffectFor(for_: PowerPointShape!, fx: PowerPointMsoAnimEffect, level: PowerPointMsoAnimateByLevel, trigger: PowerPointMsoAnimTriggerType, index: AnyObject!) -> PowerPointEffect // add an effect for a shape
    optional func convertToTextUnitEffectEffect(Effect: PowerPointEffect!, unit: PowerPointMsoAnimTextUnitEffect) -> PowerPointEffect // convert an effect to a text unit effect
}
extension SBObject: PowerPointSequence {}

// MARK: PowerPointSetEffect
@objc public protocol PowerPointSetEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var ending: AnyObject { get set }
    optional var propertySetEffect: PowerPointMsoAnimProperty { get set }
}
extension SBObject: PowerPointSetEffect {}

// MARK: PowerPointSlideRange
@objc public protocol PowerPointSlideRange: SBObjectProtocol, PowerPointGenericMethods {
    optional func slides()
    optional func shapes()
    optional func hyperlinks()
    optional var colorScheme: PowerPointColorScheme { get set } // Returns or sets the color scheme for the specified slide, slide range, or slide master.
    optional var customLayout: PowerPointCustomLayout { get set } // Returns the custom layout associated with the specified range of slides.
    optional var design: PowerPointDesign { get set }
    optional var displayMasterShapes: Int { get set } // Determines whether the specified range of slides displays the background objects on the slide master.
    optional var followMasterBackground: Int { get set } // Determines whether the range of slides follows the slide master background.
    optional var headersAndFooters: PowerPointHeadersAndFooters { get } // Returns a collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides.
    optional var layout: PowerPointEPPSlideLayout { get set } // Returns or sets the slide layout.
    optional var notesPage: PowerPointSlideRange { get } // Returns a slide range object that represents the notes pages for the specified slide or range of slides.
    optional var printSteps: Int { get }
    optional var slideID: Int { get }
    optional var slideIndex: Int { get }
    optional var slideMaster: PowerPointMaster { get } // Returns the slide master.
    optional var slideNumber: Int { get } // Returns the slide number.
    optional var slideShowTransitions: PowerPointSlideShowTransition { get } // Returns the special effects for the specified slide transition.
    optional var timeline: PowerPointTimeline { get } // Returns the animation timeline for the slide.
}
extension SBObject: PowerPointSlideRange {}

// MARK: PowerPointSlideShowSettings
@objc public protocol PowerPointSlideShowSettings: SBObjectProtocol, PowerPointGenericMethods {
    optional func namedSlideShows()
    optional var advanceMode: PowerPointEPPSlideShowAdvanceMode { get set }
    optional var endingSlide: Int { get set }
    optional var loopUntilStopped: Int { get set }
    optional var pointerColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the color of  default pointer color for a presentation.
    optional var rangeType: PowerPointEPPSlideShowRangeType { get set }
    optional var showType: PowerPointEPPSlideShowType { get set }
    optional var showWithAnimation: Int { get }
    optional var showWithNarration: Int { get set }
    optional var showWithPresenter: Int { get set }
    optional var slideShowName: Int { get set }
    optional var startingSlide: Int { get set }
    optional func runSlideShow() -> PowerPointSlideShowWindow
}
extension SBObject: PowerPointSlideShowSettings {}

// MARK: PowerPointSlideShowTransition
@objc public protocol PowerPointSlideShowTransition: SBObjectProtocol, PowerPointGenericMethods {
    optional var advanceOnClick: Int { get set }
    optional var advanceOnTime: Int { get set }
    optional var advanceTime: Double { get set }
    optional var entryEffect: PowerPointEPPEntryEffect { get set }
    optional var hidden: Int { get set }
    optional var loopSoundUntilNext: Int { get set }
    optional var soundEffectTransition: PowerPointSoundEffect { get }
    optional var transitionDuration: Double { get set }
}
extension SBObject: PowerPointSlideShowTransition {}

// MARK: PowerPointSlideShowView
@objc public protocol PowerPointSlideShowView: SBObjectProtocol, PowerPointGenericMethods {
    optional var accelerationsEnabled: Int { get set }
    optional var currentShowPosition: Int { get }
    optional var isNamedShow: Int { get }
    optional var lastSlideViewed: PowerPointSlide { get }
    optional var pointerColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the color of temporary pointer color for a view of a slide show.
    optional var pointerType: PowerPointEPPSlideShowPointerType { get set }
    optional var presentationElapsedTime: Double { get }
    optional var slide: PowerPointSlide { get }
    optional var slideElapsedTime: Double { get set }
    optional var slideShowName: Int { get }
    optional var slideState: PowerPointEPPSlideShowState { get set }
    optional var zoom: Int { get }
    optional func exitSlideShow()
    optional func goToFirstSlide()
    optional func goToLastSlide()
    optional func goToNextSlide()
    optional func goToPreviousSlide()
    optional func resetSlideTime()
}
extension SBObject: PowerPointSlideShowView {}

// MARK: PowerPointSlideShowWindow
@objc public protocol PowerPointSlideShowWindow: SBObjectProtocol, PowerPointGenericMethods {
    optional var active: Int { get }
    optional var bounds: Int { get set }
    optional var height: Double { get set }
    optional var isFullScreen: Int { get }
    optional var leftPosition: Double { get set }
    optional var presentation: PowerPointPresentation { get }
    optional var slideshowView: PowerPointSlideShowView { get }
    optional var top: Double { get set }
    optional var width: Double { get set }
}
extension SBObject: PowerPointSlideShowWindow {}

// MARK: PowerPointSlide
@objc public protocol PowerPointSlide: SBObjectProtocol, PowerPointGenericMethods {
    optional func shapes()
    optional func hyperlinks()
    optional var background: PowerPointShape { get }
    optional var colorScheme: PowerPointColorScheme { get set }
    optional var customLayout: PowerPointCustomLayout { get set }
    optional var design: PowerPointDesign { get set }
    optional var displayMasterShapes: Int { get set }
    optional var followMasterBackground: Int { get set }
    optional var headersAndFooters: PowerPointHeadersAndFooters { get }
    optional var layout: PowerPointEPPSlideLayout { get set }
    optional var notesPage: PowerPointSlide { get }
    optional var printSteps: Int { get }
    optional var sectionIndex: Int { get }
    optional var sectionNumber: Int { get }
    optional var slideID: Int { get }
    optional var slideIndex: Int { get }
    optional var slideMaster: PowerPointMaster { get }
    optional var slideNumber: Int { get }
    optional var slideShowTransition: PowerPointSlideShowTransition { get }
    optional var timeline: PowerPointTimeline { get }
    optional func applyThemeColorSchemeFileName(fileName: AnyObject!) // Applies a theme color to the specified slide.
    optional func copyObject()
    optional func cutObject()
    optional func moveToStartOfSectionAtPosition(atPosition: AnyObject!)
}
extension SBObject: PowerPointSlide {}

// MARK: PowerPointSoundEffect
@objc public protocol PowerPointSoundEffect: SBObjectProtocol, PowerPointGenericMethods {
    optional var name: Int { get }
    optional var soundType: PowerPointEPPSoundEffectType { get set }
    optional func importSoundFileSoundFileName(soundFileName: AnyObject!)
    optional func playSoundEffect()
}
extension SBObject: PowerPointSoundEffect {}

// MARK: PowerPointTabStop
@objc public protocol PowerPointTabStop: SBObjectProtocol, PowerPointGenericMethods {
    optional var tabPosition: Double { get set } // Returns or sets the position of a tab stop relative to the left margin.
    optional var tabStopType: PowerPointMsoTabStopType { get set } // Returns or sets the type of the tab stop object.
}
extension SBObject: PowerPointTabStop {}

// MARK: PowerPointTextStyleLevel
@objc public protocol PowerPointTextStyleLevel: SBObjectProtocol, PowerPointGenericMethods {
    optional var font: PowerPointFont { get }
    optional var paragraphFormat: PowerPointParagraphFormat { get }
}
extension SBObject: PowerPointTextStyleLevel {}

// MARK: PowerPointTextStyle
@objc public protocol PowerPointTextStyle: SBObjectProtocol, PowerPointGenericMethods {
    optional func textStyleLevels()
    optional var ruler: PowerPointRuler { get }
    optional var textFrame: PowerPointTextFrame { get }
}
extension SBObject: PowerPointTextStyle {}

// MARK: PowerPointTimeline
@objc public protocol PowerPointTimeline: SBObjectProtocol, PowerPointGenericMethods {
    optional func sequences()
    optional var mainSequence: PowerPointSequence { get }
    optional func addSequenceIndex(index: AnyObject!) -> PowerPointSequence // add an interactive sequence
}
extension SBObject: PowerPointTimeline {}

// MARK: PowerPointTiming
@objc public protocol PowerPointTiming: SBObjectProtocol, PowerPointGenericMethods {
    optional var acceleration: Double { get set }
    optional var autoreverse: Int { get set }
    optional var deceleration: Double { get set }
    optional var duration: Double { get set }
    optional var repeatCount: Int { get set }
    optional var repeatDuration: Double { get set }
    optional var restart: PowerPointMsoAnimEffectRestart { get set }
    optional var rewind: Int { get set }
    optional var smoothEnd: Int { get set }
    optional var smoothStart: Int { get set }
    optional var speed: Double { get set }
}
extension SBObject: PowerPointTiming {}

// MARK: PowerPointTwoInitialCapsException
@objc public protocol PowerPointTwoInitialCapsException: SBObjectProtocol, PowerPointGenericMethods {
    optional var entry_index: Int { get } // Returns the index for the position of the object in its container element list.
    optional var name: Int { get } // Returns the name of the TwoInitialCapsException.
}
extension SBObject: PowerPointTwoInitialCapsException {}

// MARK: PowerPointView
@objc public protocol PowerPointView: SBObjectProtocol, PowerPointGenericMethods {
    optional var displaySlideMiniature: Int { get set }
    optional var slide: PowerPointSlide { get set }
    optional var viewType: PowerPointEPPViewType { get }
    optional var zoom: Int { get set }
    optional var zoomToFit: Int { get set }
    optional func goToSlideNumber(number: AnyObject!)
}
extension SBObject: PowerPointView {}

// MARK: PowerPointWebOptions
@objc public protocol PowerPointWebOptions: SBObjectProtocol, PowerPointGenericMethods {
    optional var allowPNG: Int { get set }
    optional var encoding: PowerPointMsoEncoding { get set }
}
extension SBObject: PowerPointWebOptions {}

// MARK: PowerPointAdjustment
@objc public protocol PowerPointAdjustment: SBObjectProtocol, PowerPointGenericMethods {
    optional var adjustment_value: Double { get set }
}
extension SBObject: PowerPointAdjustment {}

// MARK: PowerPointCalloutFormat
@objc public protocol PowerPointCalloutFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var accent: Int { get set }
    optional var angle: PowerPointMsoCalloutAngleType { get set }
    optional var autoAttach: Int { get set }
    optional var autoLength: Int { get }
    optional var border: Int { get set }
    optional var calloutFormatLength: Double { get }
    optional var calloutType: PowerPointMsoCalloutType { get set }
    optional var drop: Double { get }
    optional var dropType: PowerPointMsoCalloutDropType { get }
    optional var gap: Double { get set }
    optional func presetDropDropType(dropType: PowerPointMsoCalloutDropType)
}
extension SBObject: PowerPointCalloutFormat {}

// MARK: PowerPointConnectorFormat
@objc public protocol PowerPointConnectorFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var beginConnected: Int { get }
    optional var beginConnectedShape: PowerPointShape { get }
    optional var beginConnectionSite: Int { get }
    optional var connectorType: PowerPointMsoConnectorType { get set }
    optional var endConnected: Int { get }
    optional var endConnectedShape: PowerPointShape { get }
    optional var endConnectionSite: Int { get }
}
extension SBObject: PowerPointConnectorFormat {}

// MARK: PowerPointFillFormat
@objc public protocol PowerPointFillFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional func gradientStops()
    optional var backColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the specified fill background color.
    optional var fillFormatType: PowerPointMsoFillType { get }
    optional var foreColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the specified foreground fill or solid color.
    optional var gradientColorType: PowerPointMsoGradientColorType { get }
    optional var gradientDegree: Double { get }
    optional var gradientStyle: PowerPointMsoGradientStyle { get }
    optional var gradientVariant: Int { get }
    optional var pattern: PowerPointMsoPatternType { get }
    optional var presetGradientType: PowerPointMsoPresetGradientType { get }
    optional var presetTexture: PowerPointMsoPresetTexture { get }
    optional var rotateWithObject: Int { get set } // Returns or sets whether the fill rotates with the specified shape.
    optional var textureAlignment: PowerPointMsoTextureAlignment { get set } // Returns or sets the texture alignment for the specified object.
    optional var textureHorizontalScale: Double { get set } // Returns or sets the texture alignment for the specified object.
    optional var textureName: Int { get }
    optional var textureOffsetX: Double { get set } // Returns or sets the texture alignment for the specified object.
    optional var textureOffsetY: Double { get set } // Returns or sets the texture alignment for the specified object.
    optional var textureTile: Int { get set } // Returns the texture tile style for the specified fill.
    optional var textureVerticalScale: Double { get set } // Returns or sets the texture alignment for the specified object.
    optional var transparency: Double { get set }
    optional var visible: Int { get set }
    optional func deleteGradientStopStopIndex(stopIndex: AnyObject!) // Removes a gradient stop.
    optional func insertGradientStopCustomColor(customColor: AnyObject!, position: Double, transparency: Double, stopIndex: AnyObject!) // Adds a stop to a gradient.
}
extension SBObject: PowerPointFillFormat {}

// MARK: PowerPointGlowFormat
@objc public protocol PowerPointGlowFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var colorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the color for the specified glow format.
    optional var radius: Double { get set } // Returns or sets the length of the radius for the specified glow format.
}
extension SBObject: PowerPointGlowFormat {}

// MARK: PowerPointGradientStop
@objc public protocol PowerPointGradientStop: SBObjectProtocol, PowerPointGenericMethods {
    optional var colorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the color for the specified gradient stop.
    optional var position: Double { get set } // Returns or sets the position for the specified gradient stop expressed as a percent.
    optional var transparency: Double { get set } // Returns or sets a value representing the transparency of the gradient fill expressed as a percent.
}
extension SBObject: PowerPointGradientStop {}

// MARK: PowerPointLineFormat
@objc public protocol PowerPointLineFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var backColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the background color for a patterned line.
    optional var beginArrowHeadLength: PowerPointMsoArrowheadLength { get set }
    optional var beginArrowheadStyle: PowerPointMsoArrowheadStyle { get set }
    optional var beginArrowheadWidth: PowerPointMsoArrowheadWidth { get set }
    optional var dashStyle: PowerPointMsoLineDashStyle { get set }
    optional var endArrowheadLength: PowerPointMsoArrowheadLength { get set }
    optional var endArrowheadStyle: PowerPointMsoArrowheadStyle { get set }
    optional var endArrowheadWidth: PowerPointMsoArrowheadWidth { get set }
    optional var foreColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the foreground color for the line.
    optional var lineFormatPatterned: PowerPointMsoPatternType { get set }
    optional var lineStyle: PowerPointMsoLineStyle { get set }
    optional var lineWeight: Double { get set }
    optional var transparency: Double { get set }
}
extension SBObject: PowerPointLineFormat {}

// MARK: PowerPointLinkFormat
@objc public protocol PowerPointLinkFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var autoUpdate: PowerPointEPPUpdateOption { get set }
    optional var sourceFullName: Int { get set }
}
extension SBObject: PowerPointLinkFormat {}

// MARK: PowerPointOfficeTheme
@objc public protocol PowerPointOfficeTheme: SBObjectProtocol, PowerPointGenericMethods {
    optional var themeColorScheme: PowerPointThemeColorScheme { get } // Returns the color scheme of a Microsoft Office theme.
    optional var themeEffectScheme: PowerPointThemeEffectScheme { get } // Returns the effects scheme of a Microsoft Office theme.
    optional var themeFontScheme: PowerPointThemeFontScheme { get } // Returns the font scheme of a Microsoft Office theme.
}
extension SBObject: PowerPointOfficeTheme {}

// MARK: PowerPointPictureFormat
@objc public protocol PowerPointPictureFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var brightness: Double { get set } // Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
    optional var colorType: PowerPointMsoPictureColorType { get set } // Returns or sets the type of color transformation applied to the specified picture.
    optional var contrast: Double { get set } // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
    optional var cropBottom: Double { get set } // Returns or sets the number of points that are cropped off the bottom of the specified picture.
    optional var cropLeft: Double { get set } // Returns or sets the number of points that are cropped off the left side of the specified picture.
    optional var cropRight: Double { get set } // Returns or sets the number of points that are cropped off the right side of the specified picture.
    optional var cropTop: Double { get set } // Returns or sets the number of points that are cropped off the top of the specified picture.
    optional var transparentBackground: Int { get set } // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.
}
extension SBObject: PowerPointPictureFormat {}

// MARK: PowerPointPlaceholderFormat
@objc public protocol PowerPointPlaceholderFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var containedType: PowerPointMsoShapeType { get }
    optional var placeholderName: Int { get set }
    optional var placeholderType: PowerPointEPPPlaceholderType { get }
}
extension SBObject: PowerPointPlaceholderFormat {}

// MARK: PowerPointReflectionFormat
@objc public protocol PowerPointReflectionFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var reflectionType: PowerPointMsoReflectionType { get set } // Returns or sets the type of the reflection format object.
}
extension SBObject: PowerPointReflectionFormat {}

// MARK: PowerPointShadowFormat
@objc public protocol PowerPointShadowFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var blur: Double { get set }
    optional var foreColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the foreground color for the shadow format.
    optional var obscured: Int { get set }
    optional var offsetX: Double { get set }
    optional var offsetY: Double { get set }
    optional var rotateWithShape: Int { get set } // Returns or sets whether to rotate the shadow when rotating the shape.
    optional var shadowStyle: PowerPointMsoShadowStyle { get set } // Returns or sets the style of shadow formatting to apply to a shape.
    optional var shadowType: PowerPointMsoShadowType { get set }
    optional var size: Double { get set } // Returns or sets the width of the shadow.
    optional var transparency: Double { get set }
    optional var visible: Int { get set }
}
extension SBObject: PowerPointShadowFormat {}

// MARK: PowerPointShapeRange
@objc public protocol PowerPointShapeRange: SBObjectProtocol, PowerPointGenericMethods {
    optional func adjustments()
    optional func shapes()
    optional var animationSettings: PowerPointAnimationSettings { get } // Returns all the special effects that you can apply to the animation of the specified shape.
    optional var autoShapeType: PowerPointMsoAutoShapeType { get set } // Returns or sets the shape type for AutoShapes other than a line, freeform drawing, or connector.
    optional var backgroundStyle: PowerPointMsoBackgroundStyleIndex { get set } // Returns or sets the background style for the specified object.
    optional var blackAndWhiteMode: PowerPointMsoBlackWhiteMode { get set } // Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode.
    optional var calloutFormat: PowerPointCalloutFormat { get } // Returns callout formatting properties for the specified line callout shapes.
    optional var child: Int { get } // Returns whether all shapes in a shape range are child shapes of the same parent.
    optional var connectionSiteCount: Int { get } // Returns the number of connection sites on the specified shape.
    optional var fillFormat: PowerPointFillFormat { get } // Returns the fill format properties for the specified object.
    optional var glowFormat: PowerPointGlowFormat { get } // Returns the glow format for the specified range of shapes.
    optional var hasChild: Int { get }
    optional var hasTable: Int { get } // Returns whether the specified shape is a table.
    optional var hasTextFrame: Int { get } // Returns whether the specified shape has a text frame.
    optional var height: Double { get set } // Returns or sets the height of the specified object.
    optional var horizontalFlip: Int { get } // Returns whether the specified shape is flipped around the horizontal axis.
    optional var isConnector: Int { get } // Determines whether the specified shape is a connector.
    optional var leftPosition: Double { get set } // Returns or sets a value equal to the distance from the left edge of the leftmost shape in the shape range to the left edge of the slide.
    optional var lineFormat: PowerPointLineFormat { get } // Returns line format properties for the specified line or shape border.
    optional var linkFormat: PowerPointLinkFormat { get } // Returns the properties for linked OLE objects.
    optional var lockAspectRatio: Int { get set } // Determines whether the specified shape retains its original proportions when you resize it.
    optional var mediaType: PowerPointEPPMediaType { get } // Returns the OLE media type.
    optional var name: Int { get set } // Identifies the shape on a slide.
    optional var reflectionFormat: PowerPointReflectionFormat { get } // Returns the reflection format for the specified range of shapes.
    optional var rotation: Double { get set } // Returns or sets the number of degrees that the specified shape is rotated around the z-axis.
    optional var shadowFormat: PowerPointShadowFormat { get } // Returns shadow formatting for specified shapes.
    optional var shapeStyle: PowerPointMsoShapeStyleIndex { get set } // Returns or sets the shape style index for the specified object.
    optional var shapeType: PowerPointMsoShapeType { get } // Represents the type of shape or shapes in a range of shapes.
    optional var softEdgeFormat: PowerPointSoftEdgeFormat { get } // Returns the soft edge format for the specified range of shapes.
    optional var textFrame: PowerPointTextFrame { get } // Returns the alignment and anchoring properties for the specified shape or master text style.
    optional var threeDFormat: PowerPointThreeDFormat { get } // Returns 3-D effect formatting properties for the specified shape.
    optional var top: Double { get set } // Returns or sets a value equal to the distance from the top edge of the topmost shape in the shape range to the top edge of the document.
    optional var verticalFlip: Int { get } // Determines whether the specified shape is flipped around the vertical axis.
    optional var visible: Int { get set } // Returns or sets the visibility of the specified object or the formatting applied to the specified object.
    optional var width: Double { get set } // Returns or sets the width of the specified object.
    optional var zOrderPosition: Int { get } // Returns the position of the specified shape in the z-order.
    optional func alignAlignType(alignType: PowerPointMsoAlignCmd, relativeToSlide: AnyObject!) // Aligns the shapes in the specified range of shapes.
    optional func copyShapeRange()
    optional func cutShapeRange()
    optional func distributeDistributeType(distributeType: PowerPointMsoDistributeCmd, relativeToSlide: AnyObject!) // Evenly distributes the shapes in the specified range of shapes. You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire slide or over the space they originally occupy.
    optional func group() -> PowerPointShape // Groups the shapes in the specified shape range.
    optional func regroup() -> PowerPointShape // Regroups the group that the specified shape range belonged to previously.
    optional func ungroup() -> PowerPointShapeRange // Ungroups any grouped shapes in the specified shape or range of shapes.
}
extension SBObject: PowerPointShapeRange {}

// MARK: PowerPointShape
@objc public protocol PowerPointShape: SBObjectProtocol, PowerPointGenericMethods {
    optional func shapes()
    optional func callouts()
    optional func connectors()
    optional func pictures()
    optional func lineShapes()
    optional func placeHolders()
    optional func textBoxes()
    optional func comments()
    optional func shapeTables()
    optional func adjustments()
    optional var animationSettings: PowerPointAnimationSettings { get }
    optional var autoShapeType: PowerPointMsoAutoShapeType { get set }
    optional var backgroundStyle: PowerPointMsoBackgroundStyleIndex { get set } // Returns or sets the background style.
    optional var blackAndWhiteMode: PowerPointMsoBlackWhiteMode { get set }
    optional var child: Int { get } // True if the shape is a child shape.
    optional var connectionSiteCount: Int { get }
    optional var fillFormat: PowerPointFillFormat { get } // Returns a fill format object that contains fill formatting properties for the specified shape.
    optional var glowFormat: PowerPointGlowFormat { get } // Returns the formatting properties for a glow effect.
    optional var hasTable: Int { get }
    optional var hasTextFrame: Int { get }
    optional var height: Double { get set }
    optional var horizontalFlip: Int { get }
    optional var isConnector: Int { get }
    optional var leftPosition: Double { get set }
    optional var lineFormat: PowerPointLineFormat { get }
    optional var linkFormat: PowerPointLinkFormat { get }
    optional var lockAspectRatio: Int { get set }
    optional var mediaType: PowerPointEPPMediaType { get }
    optional var name: Int { get set }
    optional var reflectionFormat: PowerPointReflectionFormat { get } // Returns the formatting properties for a reflection effect.
    optional var rotation: Double { get set }
    optional var shadowFormat: PowerPointShadowFormat { get }
    optional var shapeStyle: PowerPointMsoShapeStyleIndex { get set } // Returns or sets the shape style corresponding to the Quick Styles.
    optional var shapeType: PowerPointMsoShapeType { get }
    optional var softEdgeFormat: PowerPointSoftEdgeFormat { get } // Returns the formatting properties for a soft edge effect.
    optional var textFrame: PowerPointTextFrame { get }
    optional var threeDFormat: PowerPointThreeDFormat { get } // Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape.
    optional var top: Double { get set }
    optional var verticalFlip: Int { get }
    optional var visible: Int { get set }
    optional var width: Double { get set }
    optional var wordArtFormat: PowerPointWordArtFormat { get } // Returns the formatting properties for a word art effect.
    optional var zOrderPosition: Int { get }
    optional func copyShape()
    optional func cutShape()
    optional func saveAsPicturePictureType(pictureType: PowerPointMsoPictureType, fileName: AnyObject!) // Saves the shape in the requested file using the stated graphic format
}
extension SBObject: PowerPointShape {}

// MARK: PowerPointCallout
@objc public protocol PowerPointCallout: PowerPointShape {
    optional var calloutType: PowerPointMsoCalloutType { get }
    optional var calloutFormat: PowerPointCalloutFormat { get }
}
extension SBObject: PowerPointCallout {}

// MARK: PowerPointComment
@objc public protocol PowerPointComment: PowerPointShape {
}
extension SBObject: PowerPointComment {}

// MARK: PowerPointConnector
@objc public protocol PowerPointConnector: PowerPointShape {
    optional var connectorFormat: PowerPointConnectorFormat { get }
    optional var connectorType: PowerPointMsoConnectorType { get }
}
extension SBObject: PowerPointConnector {}

// MARK: PowerPointLineShape
@objc public protocol PowerPointLineShape: PowerPointShape {
    optional var beginLineX: Double { get set }
    optional var beginLineY: Double { get set }
    optional var endLineX: Double { get set }
    optional var endLineY: Double { get set }
}
extension SBObject: PowerPointLineShape {}

// MARK: PowerPointMediaObject
@objc public protocol PowerPointMediaObject: PowerPointShape {
    optional var fileName: Int { get }
}
extension SBObject: PowerPointMediaObject {}

// MARK: PowerPointMedia2Object
@objc public protocol PowerPointMedia2Object: PowerPointShape {
    optional var fileName: Int { get }
    optional var linkToFile: Int { get }
    optional var saveWithDocument: Int { get }
}
extension SBObject: PowerPointMedia2Object {}

// MARK: PowerPointPicture
@objc public protocol PowerPointPicture: PowerPointShape {
    optional var fileName: Int { get }
    optional var linkToFile: Int { get }
    optional var pictureFormat: PowerPointPictureFormat { get }
    optional var saveWithDocument: Int { get }
    optional func scaleHeightFactor(factor: Double, relativeToOriginalSize: AnyObject!, scale: PowerPointMsoScaleFrom)
    optional func scaleWidthFactor(factor: Double, relativeToOriginalSize: AnyObject!, scale: PowerPointMsoScaleFrom)
}
extension SBObject: PowerPointPicture {}

// MARK: PowerPointPlaceHolder
@objc public protocol PowerPointPlaceHolder: PowerPointShape {
    optional var placeHolderFormat: PowerPointPlaceholderFormat { get }
    optional var placeholderType: PowerPointEPPPlaceholderType { get }
}
extension SBObject: PowerPointPlaceHolder {}

// MARK: PowerPointShapeTable
@objc public protocol PowerPointShapeTable: PowerPointShape {
    optional var numberOfColumns: Int { get }
    optional var numberOfRows: Int { get }
    optional var tableObject: PowerPointTable { get }
}
extension SBObject: PowerPointShapeTable {}

// MARK: PowerPointSoftEdgeFormat
@objc public protocol PowerPointSoftEdgeFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var softEdgeType: PowerPointMsoSoftEdgeType { get set } // Returns or sets the type soft edge format object.
}
extension SBObject: PowerPointSoftEdgeFormat {}

// MARK: PowerPointTextBox
@objc public protocol PowerPointTextBox: PowerPointShape {
    optional var textOrientation: PowerPointMsoTextOrientation { get }
}
extension SBObject: PowerPointTextBox {}

// MARK: PowerPointTextColumn
@objc public protocol PowerPointTextColumn: SBObjectProtocol, PowerPointGenericMethods {
    optional var columnNumber: Int { get set } // Returns or sets the index of the text column object.
    optional var spacing: Double { get set } // Returns or sets the spacing between text columns in a text column object.
    optional var textDirection: PowerPointMsoTextDirection { get set } // Returns or sets the direction of text in the text column object.
}
extension SBObject: PowerPointTextColumn {}

// MARK: PowerPointTextFrame
@objc public protocol PowerPointTextFrame: SBObjectProtocol, PowerPointGenericMethods {
    optional var autoSize: PowerPointMsoAutoSize { get set }
    optional var hasText: Int { get } // Returns whether the specified text frame has text.
    optional var horizontalAnchor: PowerPointMsoHorizontalAnchor { get set }
    optional var marginBottom: Double { get set }
    optional var marginLeft: Double { get set }
    optional var marginRight: Double { get set }
    optional var marginTop: Double { get set }
    optional var orientation: PowerPointMsoTextOrientation { get }
    optional var pathFormat: PowerPointMsoPathFormat { get set } // Returns or sets the path type for the specified text frame.
    optional var ruler: PowerPointRuler { get }
    optional var textColumn: PowerPointTextColumn { get } // Returns the text column object that represents the columns within the text frame.
    optional var textOrientation: PowerPointMsoTextOrientation { get set }
    optional var textRange: PowerPointTextRange { get }
    optional var threeDFormat: PowerPointThreeDFormat { get } // Returns the 3-D effect formatting properties for the specified text.
    optional var verticalAnchor: PowerPointMsoVerticalAnchor { get set }
    optional var warpFormat: PowerPointMsoWarpFormat { get set } // Returns or sets the warp type for the specified text frame.
    optional var wordArtStylesFormat: PowerPointMsoPresetTextEffect { get } // Returns or sets a value specifying the text effect for the selected text.
    optional var wordWrap: Int { get set } // Returns or sets text break lines within or past the boundaries of the shape.
    optional var wordartFormat: PowerPointMsoPresetTextEffect { get set } // Returns or sets the WordArt type for the specified text frame.
}
extension SBObject: PowerPointTextFrame {}

// MARK: PowerPointThemeColorScheme
@objc public protocol PowerPointThemeColorScheme: SBObjectProtocol, PowerPointGenericMethods {
    optional func themeColors()
    optional func getCustomColorName(name: AnyObject!) // Returns the custom color for the specified Microsoft Office theme.
    optional func loadThemeColorSchemeFileName(fileName: AnyObject!) // Loads the color scheme of a Microsoft Office theme from a file.
    optional func saveThemeColorSchemeFileName(fileName: AnyObject!) // Saves the color scheme of a Microsoft Office theme to a file.
}
extension SBObject: PowerPointThemeColorScheme {}

// MARK: PowerPointThemeColor
@objc public protocol PowerPointThemeColor: SBObjectProtocol, PowerPointGenericMethods {
    optional var themeColorSchemeIndex: PowerPointMsoThemeColorSchemeIndex { get } // Returns the index value a color scheme of a Microsoft Office theme.
}
extension SBObject: PowerPointThemeColor {}

// MARK: PowerPointThemeEffectScheme
@objc public protocol PowerPointThemeEffectScheme: SBObjectProtocol, PowerPointGenericMethods {
    optional func loadThemeEffectSchemeFileName(fileName: AnyObject!) // Loads the effects scheme of a Microsoft Office theme from a file.
}
extension SBObject: PowerPointThemeEffectScheme {}

// MARK: PowerPointThemeFontScheme
@objc public protocol PowerPointThemeFontScheme: SBObjectProtocol, PowerPointGenericMethods {
    optional func minorThemeFonts()
    optional func majorThemeFonts()
    optional func loadThemeFontSchemeFileName(fileName: AnyObject!) // Loads the font scheme of a Microsoft Office theme from a file.
    optional func saveThemeFontSchemeFileName(fileName: AnyObject!) // Saves the font scheme of a Microsoft Office theme to a file.
}
extension SBObject: PowerPointThemeFontScheme {}

// MARK: PowerPointThemeFont
@objc public protocol PowerPointThemeFont: SBObjectProtocol, PowerPointGenericMethods {
    optional var name: Int { get set } // Returns or sets a value specifying the font to use for a selection.
}
extension SBObject: PowerPointThemeFont {}

// MARK: PowerPointMajorThemeFont
@objc public protocol PowerPointMajorThemeFont: PowerPointThemeFont {
}
extension SBObject: PowerPointMajorThemeFont {}

// MARK: PowerPointMinorThemeFont
@objc public protocol PowerPointMinorThemeFont: PowerPointThemeFont {
}
extension SBObject: PowerPointMinorThemeFont {}

// MARK: PowerPointThreeDFormat
@objc public protocol PowerPointThreeDFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var ZDistance: Double { get set } // Returns or sets a Single that represents the distance from the center of an object or text.
    optional var bevelBottomDepth: Double { get set } // Returns or sets the depth/height of the bottom bevel.
    optional var bevelBottomInset: Double { get set } // Returns or sets the inset size/width for the bottom bevel.
    optional var bevelBottomType: PowerPointMsoBevelType { get set } // Returns or sets the bevel type for the bottom bevel.
    optional var bevelTopDepth: Double { get set } // Returns or sets the depth/height of the top bevel.
    optional var bevelTopInset: Double { get set } // Returns or sets the inset size/width for the top bevel.
    optional var bevelTopType: PowerPointMsoBevelType { get set } // Returns or sets the bevel type for the top bevel.
    optional var contourColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the color for the specified contour.
    optional var contourWidth: Double { get set } // Returns or sets the width of the contour of an object or text.
    optional var depth: Double { get set } // Returns or sets the depth of the shape's extrusion.
    optional var extrusionColorThemeIndex: PowerPointMsoThemeColorIndex { get set } // Returns or sets the color for the specified extrusion.
    optional var extrusionColorType: PowerPointMsoExtrusionColorType { get set } // Returns or sets a value that indicates what will determine the extrusion color.
    optional var fieldOfView: Double { get set } // Returns or sets the amount of perspective for an object or text.
    optional var lightAngle: Double { get set } // Returns or sets the angle of the lighting.
    optional var perspective: Int { get set } // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
    optional var presetCamera: PowerPointMsoPresetCamera { get set } // Returns a constant that represents the camera preset.
    optional var presetExtrusionDirection: PowerPointMsoPresetExtrusionDirection { get set } // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
    optional var presetLightingDirection: PowerPointMsoPresetLightingDirection { get set } // Returns or sets the position of the light source relative to the extrusion.
    optional var presetLightingRig: PowerPointMsoLightRigType { get set } // Returns a constant that represents the lighting preset.
    optional var presetLightingSoftness: PowerPointMsoPresetLightingSoftness { get set } // Returns or sets the intensity of the extrusion lighting.
    optional var presetMaterial: PowerPointMsoPresetMaterial { get set } // Returns or sets the extrusion surface material.
    optional var presetThreeDFormat: PowerPointMsoPresetThreeDFormat { get set } // Returns or sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
    optional var projectText: Int { get set } // Returns or sets whether text on a shape rotates with shape.
    optional var rotationX: Double { get set } // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
    optional var rotationY: Double { get set } // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
    optional var rotationZ: Double { get set } // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
    optional var visible: Int { get set } // Returns or sets if the specified object, or the formatting applied to it, is visible.
}
extension SBObject: PowerPointThreeDFormat {}

// MARK: PowerPointWordArtFormat
@objc public protocol PowerPointWordArtFormat: SBObjectProtocol, PowerPointGenericMethods {
    optional var fontBold: Int { get set }
    optional var fontItalic: Int { get set }
    optional var fontName: Int { get set }
    optional var fontSize: Double { get set }
    optional var kernedPairs: Int { get set }
    optional var normalizedHeight: Int { get set }
    optional var presetShape: PowerPointMsoPresetTextEffectShape { get set }
    optional var rotatedChars: Int { get set }
    optional var textAlignment: PowerPointMsoTextEffectAlignment { get set }
    optional var tracking: Double { get set }
    optional var wordArtStylesFormat: PowerPointMsoPresetTextEffect { get set } // Returns or sets a value specifying the text effect for the selected text.
    optional var wordArtText: Int { get set }
    optional func toggleVerticalText()
}
extension SBObject: PowerPointWordArtFormat {}

// MARK: PowerPointTextRange
@objc public protocol PowerPointTextRange: SBObjectProtocol, PowerPointGenericMethods {
    optional func characters()
    optional func words()
    optional func sentences()
    optional func lines()
    optional func paragraphs()
    optional func textFlows()
    optional var boundsHeight: Double { get }
    optional var boundsWidth: Double { get }
    optional var content: Int { get set }
    optional var font: PowerPointFont { get }
    optional var indentLevel: Int { get set }
    optional var leftBounds: Double { get }
    optional var offset: Int { get }
    optional var paragraphFormat: PowerPointParagraphFormat { get }
    optional var textLength: Int { get }
    optional var topBounds: Double { get }
    optional func addPeriodsTo()
    optional func changeCaseTo(to: PowerPointMsoTextChangeCase)
    optional func copyTextRange()
    optional func cutTextRange()
    optional func getRotatedTextBounds() // Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }
    optional func getTextActionSettingResult(result: PowerPointEPPMouseActivation) -> PowerPointActionSetting
    optional func insertTextTextRangeInsertWhere(insertWhere: PowerPointMsoTextRangeInsertPosition, newText: AnyObject!) // Adds text to existing text.
    optional func pasteTextRange()
    optional func removePeriodsFrom()
}
extension SBObject: PowerPointTextRange {}

// MARK: PowerPointCharacter
@objc public protocol PowerPointCharacter: PowerPointTextRange {
}
extension SBObject: PowerPointCharacter {}

// MARK: PowerPointLine
@objc public protocol PowerPointLine: PowerPointTextRange {
}
extension SBObject: PowerPointLine {}

// MARK: PowerPointParagraph
@objc public protocol PowerPointParagraph: PowerPointTextRange {
}
extension SBObject: PowerPointParagraph {}

// MARK: PowerPointSentence
@objc public protocol PowerPointSentence: PowerPointTextRange {
}
extension SBObject: PowerPointSentence {}

// MARK: PowerPointTextFlow
@objc public protocol PowerPointTextFlow: PowerPointTextRange {
}
extension SBObject: PowerPointTextFlow {}

// MARK: PowerPointWord
@objc public protocol PowerPointWord: PowerPointTextRange {
}
extension SBObject: PowerPointWord {}

// MARK: PowerPointCell
@objc public protocol PowerPointCell: SBObjectProtocol, PowerPointGenericMethods {
    optional var selected: Int { get }
    optional var shape: PowerPointShape { get }
    optional func getBorderEdge(edge: PowerPointPpBorderType) -> PowerPointLineFormat
    optional func mergeMergeWith(mergeWith: PowerPointCell!)
    optional func splitNumberOfRows(numberOfRows: AnyObject!, numberOfColumns: AnyObject!)
}
extension SBObject: PowerPointCell {}

// MARK: PowerPointColumn
@objc public protocol PowerPointColumn: SBObjectProtocol, PowerPointGenericMethods {
    optional func cells()
    optional var width: Double { get set }
}
extension SBObject: PowerPointColumn {}

// MARK: PowerPointRow
@objc public protocol PowerPointRow: SBObjectProtocol, PowerPointGenericMethods {
    optional func cells()
    optional var height: Double { get set }
}
extension SBObject: PowerPointRow {}

// MARK: PowerPointTable
@objc public protocol PowerPointTable: SBObjectProtocol, PowerPointGenericMethods {
    optional func columns()
    optional func rows()
    optional var tableDirection: PowerPointMsoTextDirection { get set }
    optional func getCellFromRow(row: AnyObject!, column: AnyObject!) -> PowerPointCell
}
extension SBObject: PowerPointTable {}

