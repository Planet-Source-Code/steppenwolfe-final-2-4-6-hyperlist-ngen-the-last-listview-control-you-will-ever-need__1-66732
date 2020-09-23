VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/userex/functions/getthememetric.asp
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnwxp/html/xptheming.asp
'http://www.codeproject.com/managedcpp/ManagedUxTheme.asp?df=100&forumid=16513&exp=0&select=1217662

'/* skin any control, (within reason ;o) with xp styles using uxtheme dll
'/* I have translated all the api and structures from the theme headers,
'/* far more is possible with this class, only needs to be developed.

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER    As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS     As Long = &H200
Private Const FORMAT_MESSAGE_FROM_STRING        As Long = &H400
Private Const FORMAT_MESSAGE_FROM_HMODULE       As Long = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM        As Long = &H1000
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY     As Long = &H2000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK     As Long = &HFF
Private Const HWND_DESKTOP                      As Long = &H0
Private Const MAX_INTLIST_COUNT                 As Long = &HA


Private Enum THEME_SIZE
    TS_MIN
    TS_TRUE
    TS_DRAW
End Enum

Private Enum PROPERTYORIGIN
    PO_STATE = 0
    PO_PART = 1
    PO_CLASS = 2
    PO_GLOBAL = 3
    PO_NOTFOUND = 4
End Enum

Private Enum WINDOWTHEMEATTRIBUTETYPE
    WTA_NONCLIENT = 1
End Enum

Private Enum BP_ANIMATIONSTYLE
    BPAS_NONE
    BPAS_LINEAR
    BPAS_CUBIC
    BPAS_SINE
End Enum

'/* non-client enums
Public Enum UxThemeWindowParts
    WP_CAPTION = 1
    WP_SMALLCAPTION = 2
    WP_MINCAPTION = 3
    WP_SMALLMINCAPTION = 4
    WP_MAXCAPTION = 5
    WP_SMALLMAXCAPTION = 6
    WP_FRAMELEFT = 7
    WP_FRAMERIGHT = 8
    WP_FRAMEBOTTOM = 9
    WP_SMALLFRAMELEFT = 10
    WP_SMALLFRAMERIGHT = 11
    WP_SMALLFRAMEBOTTOM = 12
    WP_SYSBUTTON = 13
    WP_MDISYSBUTTON = 14
    WP_MINBUTTON = 15
    WP_MDIMINBUTTON = 16
    WP_MAXBUTTON = 17
    WP_CLOSEBUTTON = 18
    WP_SMALLCLOSEBUTTON = 19
    WP_MDICLOSEBUTTON = 20
    WP_RESTOREBUTTON = 21
    WP_MDIRESTOREBUTTON = 22
    WP_HELPBUTTON = 23
    WP_MDIHELPBUTTON = 24
    WP_HORZSCROLL = 25
    WP_HORZTHUMB = 26
    WP_VERTSCROLL = 27
    WP_VERTTHUMB = 28
    WP_DIALOG = 29
    WP_CAPTIONSIZINGTEMPLATE = 30
    WP_SMALLCAPTIONSIZINGTEMPLATE = 31
    WP_FRAMELEFTSIZINGTEMPLATE = 32
    WP_SMALLFRAMELEFTSIZINGTEMPLATE = 33
    WP_FRAMERIGHTSIZINGTEMPLATE = 34
    WP_SMALLFRAMERIGHTSIZINGTEMPLATE = 35
    WP_FRAMEBOTTOMSIZINGTEMPLATE = 36
    WP_SMALLFRAMEBOTTOMSIZINGTEMPLATE = 37
End Enum

Public Enum UxThemeFrameStates
    FS_ACTIVE = 1
    FS_INACTIVE = 2
End Enum

Public Enum UxThemeCaptionStates
    CS_ACTIVE = 1
    CS_INACTIVE = 2
    CS_DISABLED = 3
End Enum

Public Enum UxThemeMaxCaptionStates
    MXCS_ACTIVE = 1
    MXCS_INACTIVE = 2
    MXCS_DISABLED = 3
End Enum

Public Enum UxThemeMinCaptionStates
    MNCS_ACTIVE = 1
    MNCS_INACTIVE = 2
    MNCS_DISABLED = 3
End Enum

Public Enum UxThemeHorzScrollStates
    HSS_NORMAL = 1
    HSS_HOT = 2
    HSS_PUSHED = 3
    HSS_DISABLED = 4
End Enum

Public Enum UxThemeHorzThumbStates
    HTS_NORMAL = 1
    HTS_HOT = 2
    HTS_PUSHED = 3
    HTS_DISABLED = 4
End Enum

Public Enum UxThemeVertScrollStates
    VSS_NORMAL = 1
    VSS_HOT = 2
    VSS_PUSHED = 3
    VSS_DISABLED = 4
End Enum

Public Enum UxThemeVertThumbStates
    VTS_NORMAL = 1
    VTS_HOT = 2
    VTS_PUSHED = 3
    VTS_DISABLED = 4
End Enum

Public Enum UxThemeSysButtonStates
    SBS_NORMAL = 1
    SBS_HOT = 2
    SBS_PUSHED = 3
    SBS_DISABLED = 4
End Enum

Public Enum UxThemeMinButtonStates
    MINBS_NORMAL = 1
    MINBS_HOT = 2
    MINBS_PUSHED = 3
    MINBS_DISABLED = 4
End Enum

Public Enum UxThemeMaxButtonStates
    MAXBS_NORMAL = 1
    MAXBS_HOT = 2
    MAXBS_PUSHED = 3
    MAXBS_DISABLED = 4
End Enum

Public Enum UxThemeRestoreButtonStates
    RBS_NORMAL = 1
    RBS_HOT = 2
    RBS_PUSHED = 3
    RBS_DISABLED = 4
End Enum

Public Enum UxThemeHelpButtonStates
    HBS_NORMAL = 1
    HBS_HOT = 2
    HBS_PUSHED = 3
    HBS_DISABLED = 4
End Enum

Public Enum UxThemeCloseButtonStates
    CBS_NORMAL = 1
    CBS_HOT = 2
    CBS_PUSHED = 3
    CBS_DISABLED = 4
End Enum

'/* buttons
Public Enum UxThemeButtonParts
    BP_PUSHBUTTON = 1
    BP_RADIOBUTTON = 2
    BP_CHECKBOX = 3
    BP_GROUPBOX = 4
    BP_USERBUTTON = 5
End Enum

Public Enum UxThemePushButtonStates
    PBS_NORMAL = 1
    PBS_HOT = 2
    PBS_PRESSED = 3
    PBS_DISABLED = 4
    PBS_DEFAULTED = 5
End Enum

Public Enum UxThemeRadioButtonStates
    RBS_UNCHECKEDNORMAL = 1
    RBS_UNCHECKEDHOT = 2
    RBS_UNCHECKEDPRESSED = 3
    RBS_UNCHECKEDDISABLED = 4
    RBS_CHECKEDNORMAL = 5
    RBS_CHECKEDHOT = 6
    RBS_CHECKEDPRESSED = 7
    RBS_CHECKEDDISABLED = 8
End Enum

Public Enum UxThemeCheckBoxStates
    CBS_UNCHECKEDNORMAL = 1
    CBS_UNCHECKEDHOT = 2
    CBS_UNCHECKEDPRESSED = 3
    CBS_UNCHECKEDDISABLED = 4
    CBS_CHECKEDNORMAL = 5
    CBS_CHECKEDHOT = 6
    CBS_CHECKEDPRESSED = 7
    CBS_CHECKEDDISABLED = 8
    CBS_MIXEDNORMAL = 9
    CBS_MIXEDHOT = 10
    CBS_MIXEDPRESSED = 11
    CBS_MIXEDDISABLED = 12
End Enum

Public Enum UxThemeGroupBoxStates
    GBS_NORMAL = 1
    GBS_DISABLED = 2
End Enum

'/* rebar
Public Enum UxThemeRebarParts
    RP_GRIPPER = 1
    RP_GRIPPERVERT = 2
    RP_BAND = 3
    RP_CHEVRON = 4
    RP_CHEVRONVERT = 5
End Enum

Public Enum UxThemeChevronStates
    CHEVS_NORMAL = 1
    CHEVS_HOT = 2
    CHEVS_PRESSED = 3
End Enum

'/* rebar
Public Enum UxThemeToolBarParts
    TP_BUTTON = 1
    TP_DROPDOWNBUTTON = 2
    TP_SPLITBUTTON = 3
    TP_SPLITBUTTONDROPDOWN = 4
    TP_SEPARATOR = 5
    TP_SEPARATORVERT = 6
End Enum

Public Enum UxThemeToolBarStates
    TS_NORMAL = 1
    TS_HOT = 2
    TS_PRESSED = 3
    TS_DISABLED = 4
    TS_CHECKED = 5
    TS_HOTCHECKED = 6
End Enum

'/* status
Public Enum UxThemeStatusParts
    SP_PANE = 1
    SP_GRIPPERPANE = 2
    SP_GRIPPER = 3
End Enum

'/* menu
Public Enum UxThemeMenuParts
    MP_MENUITEM = 1
    MP_MENUDROPDOWN = 2
    MP_MENUBARITEM = 3
    MP_MENUBARDROPDOWN = 4
    MP_CHEVRON = 5
    MP_SEPARATOR = 6
End Enum

Public Enum UxThemeMenuStates
    MS_NORMAL = 1
    MS_SELECTED = 2
    MS_DEMOTED = 3
End Enum

'/* listview
Public Enum UxThemeLISTVIEWParts
    LVP_LISTITEM = 1
    LVP_LISTGROUP = 2
    LVP_LISTDETAIL = 3
    LVP_LISTSORTEDDETAIL = 4
    LVP_EMPTYTEXT = 5
End Enum

Public Enum UxThemeLISTITEMStates
    LIS_NORMAL = 1
    LIS_HOT = 2
    LIS_SELECTED = 3
    LIS_DISABLED = 4
    LIS_SELECTEDNOTFOCUS = 5
End Enum

'/* header
Public Enum UxThemeHEADERParts
    HP_HEADERITEM = 1
    HP_HEADERITEMLEFT = 2
    HP_HEADERITEMRIGHT = 3
    HP_HEADERSORTARROW = 4
End Enum

Public Enum UxThemeHEADERITEMStates
    HIS_NORMAL = 1
    HIS_HOT = 2
    HIS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMLEFTStates
    HILS_NORMAL = 1
    HILS_HOT = 2
    HILS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMRIGHTStates
    HIRS_NORMAL = 1
    HIRS_HOT = 2
    HIRS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERSORTARROWStates
    HSAS_SORTEDUP = 1
    HSAS_SORTEDDOWN = 2
End Enum

'/* progressbar
Public Enum UxThemePROGRESSParts
    PP_BAR = 1
    PP_BARVERT = 2
    PP_CHUNK = 3
    PP_CHUNKVERT = 4
End Enum

'/* tab control
Public Enum UsxThemeTABParts
    TABP_TABITEM = 1
    TABP_TABITEMLEFTEDGE = 2
    TABP_TABITEMRIGHTEDGE = 3
    TABP_TABITEMBOTHEDGE = 4
    TABP_TOPTABITEM = 5
    TABP_TOPTABITEMLEFTEDGE = 6
    TABP_TOPTABITEMRIGHTEDGE = 7
    TABP_TOPTABITEMBOTHEDGE = 8
    TABP_PANE = 9
    TABP_BODY = 10
End Enum

Public Enum UxThemeTABITEMStates
    TIS_NORMAL = 1
    TIS_HOT = 2
    TIS_SELECTED = 3
    TIS_DISABLED = 4
    TIS_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMLEFTEDGEStates
    TILES_NORMAL = 1
    TILES_HOT = 2
    TILES_SELECTED = 3
    TILES_DISABLED = 4
    TILES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMRIGHTEDGEStates
    TIRES_NORMAL = 1
    TIRES_HOT = 2
    TIRES_SELECTED = 3
    TIRES_DISABLED = 4
    TIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMBOTHEDGESStates
    TIBES_NORMAL = 1
    TIBES_HOT = 2
    TIBES_SELECTED = 3
    TIBES_DISABLED = 4
    TIBES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMStates
    TTIS_NORMAL = 1
    TTIS_HOT = 2
    TTIS_SELECTED = 3
    TTIS_DISABLED = 4
    TTIS_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMLEFTEDGEStates
    TTILES_NORMAL = 1
    TTILES_HOT = 2
    TTILES_SELECTED = 3
    TTILES_DISABLED = 4
    TTILES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMRIGHTEDGEStates
    TTIRES_NORMAL = 1
    TTIRES_HOT = 2
    TTIRES_SELECTED = 3
    TTIRES_DISABLED = 4
    TTIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMBOTHEDGESStates
    TTIBES_NORMAL = 1
    TTIBES_HOT = 2
    TTIBES_SELECTED = 3
    TTIBES_DISABLED = 4
    TTIBES_FOCUSED = 5
End Enum

'/* trackbar
Public Enum UxThemeTRACKBARParts
    TKP_TRACK = 1
    TKP_TRACKVERT = 2
    TKP_THUMB = 3
    TKP_THUMBBOTTOM = 4
    TKP_THUMBTOP = 5
    TKP_THUMBVERT = 6
    TKP_THUMBLEFT = 7
    TKP_THUMBRIGHT = 8
    TKP_TICS = 9
    TKP_TICSVERT = 10
End Enum

Public Enum UxThemeTRACKBARStates
    TKS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKStates
    TRS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKVERTStates
    TRVS_NORMAL = 1
End Enum

Public Enum UxThemeTHUMBStates
    TUS_NORMAL = 1
    TUS_HOT = 2
    TUS_PRESSED = 3
    TUS_FOCUSED = 4
    TUS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBBOTTOMStates
    TUBS_NORMAL = 1
    TUBS_HOT = 2
    TUBS_PRESSED = 3
    TUBS_FOCUSED = 4
    TUBS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBTOPStates
    TUTS_NORMAL = 1
    TUTS_HOT = 2
    TUTS_PRESSED = 3
    TUTS_FOCUSED = 4
    TUTS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBVERTStates
    TUVS_NORMAL = 1
    TUVS_HOT = 2
    TUVS_PRESSED = 3
    TUVS_FOCUSED = 4
    TUVS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBLEFTStates
    TUVLS_NORMAL = 1
    TUVLS_HOT = 2
    TUVLS_PRESSED = 3
    TUVLS_FOCUSED = 4
    TUVLS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBRIGHTStates
    TUVRS_NORMAL = 1
    TUVRS_HOT = 2
    TUVRS_PRESSED = 3
    TUVRS_FOCUSED = 4
    TUVRS_DISABLED = 5
End Enum

Public Enum UxThemeTICSStates
    TSS_NORMAL = 1
End Enum

Public Enum UxThemeTICSVERTStates
    TSVS_NORMAL = 1
End Enum

'/* tooltips
Public Enum UxThemeTOOLTIPParts
    TTP_STANDARD = 1
    TTP_STANDARDTITLE = 2
    TTP_BALLOON = 3
    TTP_BALLOONTITLE = 4
    TTP_CLOSE = 5
End Enum

Public Enum UxThemeCLOSEStates
    TTCS_NORMAL = 1
    TTCS_HOT = 2
    TTCS_PRESSED = 3
End Enum

Public Enum UxThemeSTANDARDStates
    TTSS_NORMAL = 1
    TTSS_LINK = 2
End Enum

Public Enum UxThemeBALLOONStates
    TTBS_NORMAL = 1
    TTBS_LINK = 2
End Enum

'/* treeview
Public Enum UxThemeTREEVIEWParts
    TVP_TREEITEM = 1
    TVP_GLYPH = 2
    TVP_BRANCH = 3
End Enum

Public Enum UxThemeTREEITEMStates
    TREIS_NORMAL = 1
    TREIS_HOT = 2
    TREIS_SELECTED = 3
    TREIS_DISABLED = 4
    TREIS_SELECTEDNOTFOCUS = 5
End Enum

Public Enum UxThemeGLYPHStates
    GLPS_CLOSED = 1
    GLPS_OPENED = 2
End Enum

'/* spin button
Public Enum UxThemeSPINStates
    SPNP_UP = 1
    SPNP_DOWN = 2
    SPNP_UPHORZ = 3
    SPNP_DOWNHORZ = 4
End Enum

Public Enum UxThemeUPStates
    UPS_NORMAL = 1
    UPS_HOT = 2
    UPS_PRESSED = 3
    UPS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNStates
    DNS_NORMAL = 1
    DNS_HOT = 2
    DNS_PRESSED = 3
    DNS_DISABLED = 4
End Enum

Public Enum UxThemeUPHORZStates
    UPHZS_NORMAL = 1
    UPHZS_HOT = 2
    UPHZS_PRESSED = 3
    UPHZS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNHORZStates
    DNHZS_NORMAL = 1
    DNHZS_HOT = 2
    DNHZS_PRESSED = 3
    DNHZS_DISABLED = 4
End Enum

'/* page
Public Enum UxThemePageParts
    PGRP_UP = 1
    PGRP_DOWN = 2
    PGRP_UPHORZ = 3
    PGRP_DOWNHORZ = 4
End Enum

'/* scrollbar
Public Enum UxThemeSCROLLBARParts
    SBP_ARROWBTN = 1
    SBP_THUMBBTNHORZ = 2
    SBP_THUMBBTNVERT = 3
    SBP_LOWERTRACKHORZ = 4
    SBP_UPPERTRACKHORZ = 5
    SBP_LOWERTRACKVERT = 6
    SBP_UPPERTRACKVERT = 7
    SBP_GRIPPERHORZ = 8
    SBP_GRIPPERVERT = 9
    SBP_SIZEBOX = 10
End Enum

Public Enum UxThemeARROWBTNStates
    ABS_UPNORMAL = 1
    ABS_UPHOT = 2
    ABS_UPPRESSED = 3
    ABS_UPDISABLED = 4
    ABS_DOWNNORMAL = 5
    ABS_DOWNHOT = 6
    ABS_DOWNPRESSED = 7
    ABS_DOWNDISABLED = 8
    ABS_LEFTNORMAL = 9
    ABS_LEFTHOT = 10
    ABS_LEFTPRESSED = 11
    ABS_LEFTDISABLED = 12
    ABS_RIGHTNORMAL = 13
    ABS_RIGHTHOT = 14
    ABS_RIGHTPRESSED = 15
    ABS_RIGHTDISABLED = 16
End Enum

Public Enum UxThemeSCROLLBARStates
    SCRBS_NORMAL = 1
    SCRBS_HOT = 2
    SCRBS_PRESSED = 3
    SCRBS_DISABLED = 4
End Enum

Public Enum UxThemeSIZEBOXStates
    SZB_RIGHTALIGN = 1
    SZB_LEFTALIGN = 2
End Enum

'/* edit control
Public Enum UxThemeEDITParts
    EP_EDITTEXT = 1
    EP_CARET = 2
End Enum

Public Enum UxThemeEDITTEXTStates
    ETS_NORMAL = 1
    ETS_HOT = 2
    ETS_SELECTED = 3
    ETS_DISABLED = 4
    ETS_FOCUSED = 5
    ETS_READONLY = 6
    ETS_ASSIST = 7
End Enum

'/* combobox
Public Enum UxThemeComboBoxParts
    CP_DROPDOWNBUTTON = 1
End Enum

Public Enum UxThemeComboBoxStates
    CBXS_NORMAL = 1
    CBXS_HOT = 2
    CBXS_PRESSED = 3
    CBXS_DISABLED = 4
End Enum

'/* taskbar clock
Public Enum UxThemeCLOCKParts
    CLP_TIME = 1
End Enum

Public Enum UxThemeCLOCKStates
    CLS_NORMAL = 1
End Enum

'/* systray
Public Enum UxThemeTRAYNOTIFYParts
    TNP_BACKGROUND = 1
    TNP_ANIMBACKGROUND = 2
End Enum

'/* taskbar
Public Enum UxThemeTASKBARParts
    TBP_BACKGROUNDBOTTOM = 1
    TBP_BACKGROUNDRIGHT = 2
    TBP_BACKGROUNDTOP = 3
    TBP_BACKGROUNDLEFT = 4
    TBP_SIZINGBARBOTTOM = 5
    TBP_SIZINGBARRIGHT = 6
    TBP_SIZINGBARTOP = 7
    TBP_SIZINGBARLEFT = 8
End Enum

'/* taskband
Public Enum UxThemeTASKBANDParts
    TDP_GROUPCOUNT = 1
    TDP_FLASHBUTTON = 2
    TDP_FLASHBUTTONGROUPMENU = 3
End Enum

'/* startpanel
Public Enum UxThemeSTARTPANELParts
    SPP_USERPANE = 1
    SPP_MOREPROGRAMS = 2
    SPP_MOREPROGRAMSARROW = 3
    SPP_PROGLIST = 4
    SPP_PROGLISTSEPARATOR = 5
    SPP_PLACESLIST = 6
    SPP_PLACESLISTSEPARATOR = 7
    SPP_LOGOFF = 8
    SPP_LOGOFFBUTTONS = 9
    SPP_USERPICTURE = 10
    SPP_PREVIEW = 11
End Enum

Public Enum UxThemeMOREPROGRAMSARROWStates
    SPS_NORMAL = 1
    SPS_HOT = 2
    SPS_PRESSED = 3
End Enum

Public Enum UxThemeLOGOFFBUTTONSStates
    SPLS_NORMAL = 1
    SPLS_HOT = 2
    SPLS_PRESSED = 3
End Enum

'/* explorerbar
Public Enum UxThemeEXPLORERBARParts
    EBP_HEADERBACKGROUND = 1
    EBP_HEADERCLOSE = 2
    EBP_HEADERPIN = 3
    EBP_IEBARMENU = 4
    EBP_NORMALGROUPBACKGROUND = 5
    EBP_NORMALGROUPCOLLAPSE = 6
    EBP_NORMALGROUPEXPAND = 7
    EBP_NORMALGROUPHEAD = 8
    EBP_SPECIALGROUPBACKGROUND = 9
    EBP_SPECIALGROUPCOLLAPSE = 10
    EBP_SPECIALGROUPEXPAND = 11
    EBP_SPECIALGROUPHEAD = 12
End Enum

Public Enum UxThemeHEADERCLOSEStates
    EBHC_NORMAL = 1
    EBHC_HOT = 2
    EBHC_PRESSED = 3
End Enum

Public Enum UxThemeHEADERPINStates
    EBHP_NORMAL = 1
    EBHP_HOT = 2
    EBHP_PRESSED = 3
    EBHP_SELECTEDNORMAL = 4
    EBHP_SELECTEDHOT = 5
    EBHP_SELECTEDPRESSED = 6
End Enum

Public Enum UxThemeIEBARMENUStates
    EBM_NORMAL = 1
    EBM_HOT = 2
    EBM_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPCOLLAPSEStates
    EBNGC_NORMAL = 1
    EBNGC_HOT = 2
    EBNGC_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPEXPANDStates
    EBNGE_NORMAL = 1
    EBNGE_HOT = 2
    EBNGE_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPCOLLAPSEStates
    EBSGC_NORMAL = 1
    EBSGC_HOT = 2
    EBSGC_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPEXPANDStates
    EBSGE_NORMAL = 1
    EBSGE_HOT = 2
    EBSGE_PRESSED = 3
End Enum

'/* taskband
Public Enum UxThemeMENUBANDParts
    MDP_NEWAPPBUTTON = 1
    MDP_SEPERATOR = 2
End Enum

Public Enum UxThemeMENUBANDStates
    MDS_NORMAL = 1
    MDS_HOT = 2
    MDS_PRESSED = 3
    MDS_DISABLED = 4
    MDS_CHECKED = 5
    MDS_HOTCHECKED = 6
End Enum

Public Enum THEMESIZE
    TS_MIN
    TS_TRUE
    TS_DRAW
End Enum

Public Enum UxDrawTextFlags
    UxDT_TOP = &H0
    UxDT_LEFT = &H0
    UxDT_CENTER = &H1
    UxDT_RIGHT = &H2
    UxDT_VCENTER = &H4
    UxDT_BOTTOM = &H8
    UxDT_WORDBREAK = &H10
    UxDT_SINGLELINE = &H20
    UxDT_EXPANDTABS = &H40
    UxDT_TABSTOP = &H80
    UxDT_NOCLIP = &H100
    UxDT_EXTERNALLEADING = &H200
    UxDT_CALCRECT = &H400
    UxDT_NOPREFIX = &H800
    UxDT_INTERNAL = &H1000
    UxDT_EDITCONTROL = &H2000
    UxDT_PATH_ELLIPSIS = &H4000
    UxDT_END_ELLIPSIS = &H8000
    UxDT_MODIFYSTRING = &H10000
    UxDT_RTLREADING = &H20000
    UxDT_WORD_ELLIPSIS = &H40000
    UxDT_NOFULLWIDTHCHARBREAK = &H80000
    UxDT_HIDEPREFIX = &H100000
    UxDT_PREFIXONLY = &H200000
End Enum

Public Enum UxDrawEdgeEdgeTypes
    UxBDR_RAISEDOUTER = &H1
    UxBDR_SUNKENOUTER = &H2
    UxBDR_RAISEDINNER = &H4
    UxBDR_SUNKENINNER = &H8
    UxBDR_OUTER = (UxBDR_RAISEDOUTER Or UxBDR_SUNKENOUTER)
    UxBDR_INNER = (UxBDR_RAISEDINNER Or UxBDR_SUNKENINNER)
    UxBDR_RAISED = (UxBDR_RAISEDOUTER Or UxBDR_RAISEDINNER)
    UxBDR_SUNKEN = (UxBDR_SUNKENOUTER Or UxBDR_SUNKENINNER)
    UxEDGE_RAISED = (UxBDR_RAISEDOUTER Or UxBDR_RAISEDINNER)
    UxEDGE_SUNKEN = (UxBDR_SUNKENOUTER Or UxBDR_SUNKENINNER)
    UxEDGE_ETCHED = (UxBDR_SUNKENOUTER Or UxBDR_RAISEDINNER)
    UxEDGE_BUMP = (UxBDR_RAISEDOUTER Or UxBDR_SUNKENINNER)
End Enum

Public Enum UxDrawEdgeBorderFlags
    UxBF_LEFT = &H1
    UxBF_TOP = &H2
    UxBF_RIGHT = &H4
    UxBF_BOTTOM = &H8
    UxBF_TOPLEFT = (UxBF_TOP Or UxBF_LEFT)
    UxBF_TOPRIGHT = (UxBF_TOP Or UxBF_RIGHT)
    UxBF_BOTTOMLEFT = (UxBF_BOTTOM Or UxBF_LEFT)
    UxBF_BOTTOMRIGHT = (UxBF_BOTTOM Or UxBF_RIGHT)
    UxBF_RECT = (UxBF_LEFT Or UxBF_TOP Or UxBF_RIGHT Or UxBF_BOTTOM)
    UxBF_DIAGONAL = &H10
    UxBF_DIAGONAL_ENDTOPRIGHT = (UxBF_DIAGONAL Or UxBF_TOP Or UxBF_RIGHT)
    UxBF_DIAGONAL_ENDTOPLEFT = (UxBF_DIAGONAL Or UxBF_TOP Or UxBF_LEFT)
    UxBF_DIAGONAL_ENDBOTTOMLEFT = (UxBF_DIAGONAL Or UxBF_BOTTOM Or UxBF_LEFT)
    UxBF_DIAGONAL_ENDBOTTOMRIGHT = (UxBF_DIAGONAL Or UxBF_BOTTOM Or UxBF_RIGHT)
    UxBF_MIDDLE = &H800
    UxBF_SOFT = &H1000
    UxBF_ADJUST = &H2000
    UxBF_FLAT = &H4000
    UxBF_MONO = &H8000
End Enum

'/* grayed-out string
Public Enum UxDrawTextAdditionalFlags
    UxDTT_GRAYED = &H1
End Enum

Public Enum UxThemeName
    ["HomeStead"]
    ["NormalColor"]
    ["Metallic"]
End Enum

Public Enum UxClassName
    ["Window"]
    ["Button"]
    ["Rebar"]
    ["Toolbar"]
    ["Status"]
    ["Menu"]
    ["ListView"]
    ["Header"]
    ["Progress"]
    ["Tab"]
    ["Trackbar"]
    ["Tooltip"]
    ["Treeview"]
    ["Spin"]
    ["Page"]
    ["Scrollbar"]
    ["Edit"]
    ["ComboBox"]
    ["Clock"]
    ["TrayNotify"]
    ["TaskBar"]
    ["TaskBand"]
    ["StartPanel"]
    ["ExplorerBar"]
    ["MenuBand"]
End Enum


Private Type SIZE
    cx                                          As Long
    cy                                          As Long
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    right                                       As Long
    bottom                                      As Long
End Type

Private Type POINT
    X                                           As Long
    Y                                           As Long
End Type

Private Type BLENDFUNCTION
  BlendOp                                       As Byte
  BlendFlags                                    As Byte
  SourceConstantAlpha                           As Byte
  AlphaFormat                                   As Byte
End Type

Private Type BP_ANIMATIONPARAMS
    cbSize                                      As Long
    dwFlags                                     As Long
    style                                       As BP_ANIMATIONSTYLE
    dwDuration                                  As Long
End Type

Private Type BP_PAINTPARAMS
    cbSize                                      As Long
    dwFlags                                     As Long
    prcExclude                                  As RECT
    pBlendFunction                              As BLENDFUNCTION
End Type

Private Type DTBGOPTS
    dwSize                                      As Long
    dwFlags                                     As Long
    rcClip                                      As RECT
End Type

Private Type DTTOPTS
    dwSize                                      As Long
    dwFlags                                     As Long
    crText                                      As Long
    crBorder                                    As Long
    crShadow                                    As Long
    iTextShadowType                             As Long
    ptShadowOffset                              As POINT
    iBorderSize                                 As Long
    iFontPropId                                 As Long
    iColorPropId                                As Long
    iStateId                                    As Long
    fApplyOverlay                               As Long
    iiGlowSize                                  As Long
    pfnDrawTextCallback                         As Long
    lParam                                      As Long
End Type

Private Type INTLIST
    iValueCount                                 As Long
    iValues(9)                                  As Long
End Type

Private Type MARGINS
    cxLeftWidth                                 As Long
    cxRightWidth                                As Long
    cyTopHeight                                 As Long
    cyBottomHeight                              As Long
End Type

Private Type WTA_OPTIONS
    dwFlags                                     As Long
    dwMask                                      As Long
End Type

Private Type TEXTMETRIC
    tmHeight                                    As Long
    tmAscent                                    As Long
    tmDescent                                   As Long
    tmInternalLeading                           As Long
    tmExternalLeading                           As Long
    tmAveCharWidth                              As Long
    tmMaxCharWidth                              As Long
    tmWeight                                    As Long
    tmOverhang                                  As Long
    tmDigitizedAspectX                          As Long
    tmDigitizedAspectY                          As Long
    tmFirstChar                                 As Long
    TtmLastChar                                 As Long
    tmDefaultChar                               As Long
    tmBreakChar                                 As Long
    tmItalic                                    As Byte
    tmUnderlined                                As Byte
    tmStruckOut                                 As Byte
    tmPitchAndFamily                            As Byte
    tmCharSet                                   As Byte
End Type

Private Type LOGFONT
    lfHeight                                    As Long
    lfWidth                                     As Long
    lfEscapement                                As Long
    lfOrientation                               As Long
    lfWeight                                    As Long
    lfItalic                                    As Byte
    lfUnderline                                 As Byte
    lfStrikeOut                                 As Byte
    lfCharSet                                   As Byte
    lfOutPrecision                              As Byte
    lfClipPrecision                             As Byte
    lfQuality                                   As Byte
    lfPitchAndFamily                            As Byte
    lfFaceName(32)                              As Byte
End Type


Private Declare Function GetThemeTextMetrics Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                             ByVal hdc As Long, _
                                                             ByVal iPartId As Long, _
                                                             ByVal iStateId As Long, _
                                                             ByRef ptm As TEXTMETRIC) As Long

Private Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                   ByVal hdc As Long, _
                                                                   ByVal iPartId As Long, _
                                                                   ByVal iStateId As Long, _
                                                                   ByRef pRect As RECT, _
                                                                   ByVal hrgn As Long) As Long

Private Declare Function HitTestThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                   ByVal hdc As Long, _
                                                                   ByVal iPartId As Long, _
                                                                   ByVal iStateId As Long, _
                                                                   ByVal dwOptions As Long, _
                                                                   ByRef pRect As RECT, _
                                                                   ByVal hrgn As Long, _
                                                                   ByRef ptTest As POINT, _
                                                                   ByVal pwHitTestCode As Long) As Long

Private Declare Function IsThemePartDefined Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                               ByVal iPartId As Long, _
                                                               ByVal iStateId As Long) As Long

Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                                  ByVal iPartId As Long, _
                                                                                  ByVal iStateId As Long) As Long

Private Declare Function GetThemeColor Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                          ByVal iPartId As Long, _
                                                          ByVal iStateId As Long, _
                                                          ByVal iPropId As Long, _
                                                          ByRef pColor As Long, _
                                                          ByRef piVal As Long) As Long

Private Declare Function GetThemeMetric Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                           ByVal hdc As Long, _
                                                           ByVal iPartId As Long, _
                                                           ByVal iStateId As Long, _
                                                           ByVal iPropId As Long, _
                                                           ByRef piVal As Long) As Long

Private Declare Function GetThemeString Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                           ByVal iPartId As Long, _
                                                           ByVal iStateId As Long, _
                                                           ByVal iPropId As Long, _
                                                           ByRef pszBuff As Long, _
                                                           ByVal cchMaxBuffChars As Long) As Long

Private Declare Function GetThemeBool Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                         ByVal iPartId As Long, _
                                                         ByVal iStateId As Long, _
                                                         ByVal iPropId As Long, _
                                                         ByRef pfVal As Long) As Long

Private Declare Function GetThemeInt Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                        ByVal iPartId As Long, _
                                                        ByVal iStateId As Long, _
                                                        ByVal iPropId As Long, _
                                                        ByRef piVal As Long) As Long


Private Declare Function GetThemeEnumValue Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                              ByVal iPartId As Long, _
                                                              ByVal iStateId As Long, _
                                                              ByVal iPropId As Long, _
                                                              ByRef piVal As Long) As Long

Private Declare Function GetThemePosition Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                             ByVal iPartId As Long, _
                                                             ByVal iStateId As Long, _
                                                             ByVal iPropId As Long, _
                                                             ByRef pPoint As POINT) As Long

Private Declare Function GetThemeFont Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                         ByVal hdc As Long, _
                                                         ByVal iPartId As Long, _
                                                         ByVal iStateId As Long, _
                                                         ByVal iPropId As Long, _
                                                         ByRef pFont As LOGFONT) As Long

Private Declare Function GetThemeRect Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                         ByVal iPartId As Long, _
                                                         ByVal iStateId As Long, _
                                                         ByVal iPropId As Long, _
                                                         ByRef prc As RECT) As Long

Private Declare Function GetThemeIntList Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                            ByVal iPartId As Long, _
                                                            ByVal iStateId As Long, _
                                                            ByRef pIntList As INTLIST) As Long

Private Declare Function GetThemePropertyOrigin Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                   ByVal iPartId As Long, _
                                                                   ByVal iStateId As Long, _
                                                                   ByVal iPropId As Long, _
                                                                   ByRef pOrigin As PROPERTYORIGIN) As Long

Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                           ByVal pszSubAppName As Long, _
                                                           ByVal pszSubIdList As Long) As Long

Private Declare Function GetThemeFilename Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                             ByVal iPartId As Long, _
                                                             ByVal iStateId As Long, _
                                                             ByVal iPropId As Long, _
                                                             ByVal pszThemeFileName As Long, _
                                                             ByVal cchMaxBuffChars As Long) As Long

Private Declare Function GetThemeSysColor Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                             ByVal iColorId As Long) As Long

Private Declare Function GetThemeSysColorBrush Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                  ByVal iColorId As Long) As Long

Private Declare Function GetThemeSysBool Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                            ByVal iBoolId As Long) As Long

Private Declare Function GetThemeSysSize Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                            ByVal iSizeId As Long) As Long

Private Declare Function GetThemeSysFont Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                            ByVal iFontId As Long, _
                                                            plf As LOGFONT) As Long

Private Declare Function GetThemeSysString Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                              ByVal iStringId As Long, _
                                                              ByVal pszStringBuff As Long, _
                                                              ByVal pszStringBuff As Long) As Long

Private Declare Function GetThemeSysInt Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                           ByVal iIntIdas As Long, _
                                                           ByRef piValue As Long) As Long

Private Declare Function IsThemeActive Lib "uxtheme.dll" () As Long

Private Declare Function GetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long) As Long

Private Declare Function EnableThemeDialogTexture Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                                     ByVal dwFlags As Long) As Long

Private Declare Function IsThemeDialogTextureEnabled Lib "uxtheme.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetThemeAppProperties Lib "uxtheme.dll" () As Long

Private Declare Function SetThemeAppProperties Lib "uxtheme.dll" (ByVal dwFlags As Long) As Long

Private Declare Function GetThemeDocumentationProperty Lib "uxtheme.dll" (ByVal pszThemeName As Long, _
                                                                          ByVal pszPropertyName As Long, _
                                                                          ByVal pszValueBuff As Long, _
                                                                          ByVal cchMaxValChars As Long) As Long

Private Declare Function EnableTheming Lib "uxtheme.dll" (ByVal fEnable As Long) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                                                              lpSource As Any, _
                                                                              ByVal dwMessageId As Long, _
                                                                              ByVal dwLanguageId As Long, _
                                                                              ByVal lpBuffer As String, _
                                                                              ByVal nSize As Long, _
                                                                              Arguments As Long) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                          ByVal pszClassList As Long) As Long

Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                ByVal lHDC As Long, _
                                                                ByVal iPartId As Long, _
                                                                ByVal iStateId As Long, _
                                                                pRect As RECT, _
                                                                pClipRect As RECT) As Long

Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                                      ByVal hdc As Long, _
                                                                      prc As RECT) As Long

Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                                          ByVal hdc As Long, _
                                                                          ByVal iPartId As Long, _
                                                                          ByVal iStateId As Long, _
                                                                          pBoundingRect As RECT, _
                                                                          pContentRect As RECT) As Long

Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                          ByVal hdc As Long, _
                                                          ByVal iPartId As Long, _
                                                          ByVal iStateId As Long, _
                                                          ByVal pszText As Long, _
                                                          ByVal iCharCount As Long, _
                                                          ByVal dwTextFlag As Long, _
                                                          ByVal dwTextFlags2 As Long, _
                                                          pRect As RECT) As Long

Private Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                          ByVal hdc As Long, _
                                                          ByVal iPartId As Long, _
                                                          ByVal iStateId As Long, _
                                                          pRect As RECT, _
                                                          ByVal hIml As Long, _
                                                          ByVal iImageIndex As Long) As Long

Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                             ByVal hdc As Long, _
                                                             ByVal iPartId As Long, _
                                                             ByVal iStateId As Long, _
                                                             prc As RECT, _
                                                             ByVal eSize As THEMESIZE, _
                                                             psz As SIZE) As Long

Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                               ByVal hdc As Long, _
                                                               ByVal iPartId As Long, _
                                                               ByVal iStateId As Long, _
                                                               ByVal pszText As Long, _
                                                               ByVal iCharCount As Long, _
                                                               ByVal dwTextFlags As UxDrawTextFlags, _
                                                               pBoundingRect As RECT, _
                                                               pExtentRect As RECT) As Long

Private Declare Function DrawThemeEdge Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                                          ByVal hdc As Long, _
                                                          ByVal iPartId As Long, _
                                                          ByVal iStateId As Long, _
                                                          pDestRect As RECT, _
                                                          ByVal uEdge As UxDrawEdgeEdgeTypes, _
                                                          ByVal uFlags As UxDrawEdgeBorderFlags, _
                                                          pContentRect As RECT) As Long

Private Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" (ByVal hIml As Long, _
                                                                    ByVal i As Long, _
                                                                    prcImage As RECT) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
                                                                ByVal dwMaxNameChars As Long, _
                                                                ByVal pszColorBuff As Long, _
                                                                ByVal cchMaxColorChars As Long, _
                                                                ByVal pszSizeBuff As Long, _
                                                                ByVal cchMaxSizeChars As Long) As Long

Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long


Private m_lParentHwnd                           As Long
Private m_lIconIndex                            As Long
Private m_lPartId                               As Long
Private m_lStateId                              As Long
Private m_lImlHwnd                              As Long
Private m_lHdc                                  As Long
Private m_lLeft                                 As Long
Private m_lTop                                  As Long
Private m_lWidth                                As Long
Private m_lHeight                               As Long
Private m_sText                                 As String
Private m_sThemeName                            As String
Private m_sClass                                As String
Private m_eTextAlign                            As UxDrawTextFlags
Private m_tRect                                 As RECT


Public Property Get ParentHwnd() As Long
    ParentHwnd = m_lParentHwnd
End Property

Public Property Let ParentHwnd(ByVal PropVal As Long)
    m_lParentHwnd = PropVal
End Property

Public Property Get hdc() As Long
    hdc = m_lHdc
End Property

Public Property Let hdc(ByVal PropVal As Long)
    m_lHdc = PropVal
End Property

Public Property Get Class() As String
    Class = m_sClass
End Property

Public Property Let Class(ByVal PropVal As String)
    m_sClass = PropVal
End Property

Public Property Get State() As Long
    State = m_lStateId
End Property

Public Property Let State(ByVal PropVal As Long)
    m_lStateId = PropVal
End Property

Public Property Get top() As Long
    top = m_lTop
End Property

Public Property Let top(ByVal PropVal As Long)
    m_lTop = PropVal
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Let Height(ByVal PropVal As Long)
    m_lHeight = PropVal
End Property

Public Property Get left() As Long
    left = m_lLeft
End Property

Public Property Let left(ByVal PropVal As Long)
    m_lLeft = PropVal
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Let Width(ByVal PropVal As Long)
    m_lWidth = PropVal
End Property

Public Property Get PartID() As Long
    PartID = m_lPartId
End Property

Public Property Let PartID(ByVal PropVal As Long)
    m_lPartId = PropVal
End Property

Public Property Get Text() As String
    Text = m_sText
End Property

Public Property Let Text(ByVal PropVal As String)
    m_sText = PropVal
End Property

Public Property Get TextAlign() As UxDrawTextFlags
    TextAlign = m_eTextAlign
End Property

Public Property Let TextAlign(ByVal PropVal As UxDrawTextFlags)
    m_eTextAlign = PropVal
End Property

Public Property Get IconIndex() As Long
    IconIndex = m_lIconIndex
End Property

Public Property Let IconIndex(ByVal PropVal As Long)
    m_lIconIndex = PropVal
End Property

Public Property Get ImlHwnd() As Long
    ImlHwnd = m_lImlHwnd
End Property

Public Property Let ImlHwnd(ByVal PropVal As Long)
    m_lImlHwnd = PropVal
End Property

Public Function IsXPThemed() As Boolean
    IsXPThemed = CBool(IsAppThemed())
End Function

Public Property Get ThemeName() As String

    If Len(m_sThemeName) = 0 Then
        ThemeName = GetThemeName
    End If
    ThemeName = m_sThemeName
    
End Property

Public Sub DrawPart()

Dim lhTheme     As Long
Dim tRect       As RECT

    lhTheme = OpenThemeData(m_lParentHwnd, StrPtr(m_sClass))
    With m_tRect
        .left = m_lLeft
        .top = m_lTop
        .right = m_lLeft + m_lWidth
        .bottom = m_lTop + m_lHeight
    End With
    '/* test theme load
    If (lhTheme) Then
        '/* draw part
        DrawThemeParentBackground m_lParentHwnd, m_lHdc, tRect
        '/* clean up
        CloseThemeData lhTheme
    End If
    
End Sub

Public Function UxDrawPart(Optional ByVal bPaintParent As Boolean) As Boolean

Dim lhTheme     As Long
Dim lRet        As Long

On Error GoTo Handler

    lhTheme = OpenThemeData(m_lParentHwnd, StrPtr(m_sClass))
    If lhTheme Then
        '/* create rect
        With m_tRect
            .left = m_lLeft
            .top = m_lTop
            .right = m_lLeft + m_lWidth
            .bottom = m_lTop + m_lHeight
        End With
        '/* paint parent bg
        If bPaintParent Then
            lRet = DrawThemeParentBackground(m_lParentHwnd, m_lHdc, m_tRect)
            If Not lRet = 0 Then GoTo Handler
        End If
        '/* draw part
        lRet = DrawThemeBackground(lhTheme, m_lHdc, m_lPartId, m_lStateId, m_tRect, m_tRect)
        If Not lRet = 0 Then GoTo Handler
    End If
    UxDrawPart = True

Handler:
    CloseThemeData lhTheme

End Function

Public Function UxDrawIcon(ByVal lIndex As Long) As Boolean
'/* seems to require 32 bit icons, use sys image list or vbaccelerator iml

Dim lRet        As Long
Dim lWidth      As Long
Dim lhTheme     As Long
Dim tIconR      As RECT
Dim tImlR       As RECT

On Error GoTo Handler

    lhTheme = OpenThemeData(m_lParentHwnd, StrPtr(m_sClass))
    If lhTheme Then
        If (m_lIconIndex > -1) And m_lImlHwnd Then
            ImageList_GetImageRect m_lImlHwnd, m_lIconIndex, tImlR
            lWidth = tImlR.right - tImlR.left + 4
            lRet = GetThemeBackgroundContentRect(lhTheme, m_lHdc, m_lPartId, m_lStateId, m_tRect, tIconR)
            ImageList_GetImageRect m_lImlHwnd, m_lIconIndex, tImlR
            With tIconR
                .left = .left + 2
                .top = .top + 2
                .right = .left + tImlR.right - tImlR.left
                .bottom = .top + tImlR.bottom - tImlR.top
            End With
            lRet = DrawThemeIcon(lhTheme, m_lHdc, m_lPartId, m_lStateId, tIconR, m_lImlHwnd, m_lIconIndex)
            If Not lRet = 0 Then GoTo Handler
        End If
    End If
    UxDrawIcon = True

Handler:
    CloseThemeData lhTheme
    
End Function

Public Function UxDrawText(ByVal sText As String, _
                           Optional ByVal lOffset As Long, _
                           Optional eAlign As UxDrawTextFlags, _
                           Optional bUseExtent As Boolean) As Boolean

'/* there is also a DrawThemeTextEx, looks interesting..
Dim lRet        As Long
Dim lhTheme     As Long
Dim tText       As RECT
Dim tCont       As RECT
Dim tRect       As RECT

On Error GoTo Handler

    lhTheme = OpenThemeData(m_lParentHwnd, StrPtr(m_sClass))
    If lhTheme Then
        If Len(m_sText) > 0 Then
            lRet = GetThemeBackgroundContentRect(lhTheme, m_lHdc, m_lPartId, m_lStateId, m_tRect, tText)
            If Not lRet = 0 Then GoTo Handler
            If bUseExtent Then
                LSet tRect = m_tRect
                '/* get parent rect
                lRet = GetThemeBackgroundContentRect(lhTheme, m_lHdc, m_lPartId, m_lStateId, tRect, tCont)
                '/* get text rect
                lRet = GetThemeTextExtent(lhTheme, m_lHdc, m_lPartId, m_lStateId, StrPtr(m_sText), -1, eAlign, tRect, tText)
                '/* calc text size
                If lRet = 0 Then
                    With tCont
                        If ((.bottom - .top) < (tText.bottom - tText.top)) Then
                            tRect.bottom = tRect.bottom + ((tText.bottom - tText.top) - (.bottom - .top))
                        End If
                        If ((.right - .left - lOffset) < (tText.right - tText.left + 8)) Then
                            tRect.right = tRect.right + ((tText.right - tText.left + 8) - (.right - .left - lOffset))
                        End If
                    End With
                End If
            End If
            '/* offset
            With tText
                .left = .left + lOffset
                .right = .right + lOffset
            End With
            '/* draw text
            lRet = DrawThemeText(lhTheme, m_lHdc, m_lPartId, m_lStateId, StrPtr(m_sText), -1, eAlign, 0&, tText)
            If Not lRet = 0 Then GoTo Handler
        End If
    End If
    UxDrawText = True

Handler:
    CloseThemeData lhTheme

End Function

Public Function UxDrawEdge(ByVal eEdge As UxDrawEdgeEdgeTypes, _
                           ByVal eFlags As UxDrawEdgeBorderFlags) As Boolean

Dim lRet        As Long
Dim lhTheme     As Long
Dim tTextR      As RECT

On Error GoTo Handler

    lhTheme = OpenThemeData(m_lParentHwnd, StrPtr(m_sClass))
    If lhTheme Then
        lRet = GetThemeBackgroundContentRect(lhTheme, m_lHdc, m_lPartId, m_lStateId, m_tRect, m_tRect)
        If Not lRet = 0 Then GoTo Handler
        lRet = DrawThemeEdge(lhTheme, m_lHdc, m_lPartId, m_lStateId, m_tRect, eEdge, eFlags, tTextR)
        If Not lRet = 0 Then GoTo Handler
    End If
    UxDrawEdge = True

Handler:
    CloseThemeData lhTheme

End Function

Private Function GetThemeName() As UxThemeName

Dim bClrNm()    As Byte
Dim bTmeNm()    As Byte
Dim lhTheme     As Long
Dim lPtrTme     As Long
Dim lPtrClr     As Long

On Error GoTo Handler

    If m_lParentHwnd = 0 Then
        GetThemeName = "None"
        Exit Function
    ElseIf Len(m_sClass) = 0 Then
        GetThemeName = "None"
        Exit Function
    End If
    
    lhTheme = OpenThemeData(m_lParentHwnd, StrPtr(m_sClass))
    If lhTheme = 0 Then GoTo Handler
    
    ReDim bTmeNm(0 To 260 * 2) As Byte
    lPtrTme = VarPtr(bTmeNm(0))
    ReDim bClrNm(0 To 260 * 2) As Byte
    lPtrClr = VarPtr(bClrNm(0))
    GetCurrentThemeName lPtrTme, 260, lPtrClr, 260, 0, 0
    m_sThemeName = bClrNm
    If Not InStr(m_sThemeName, vbNullChar) = 0 Then
        m_sThemeName = left$(m_sThemeName, InStr(m_sThemeName, vbNullChar) - 1)
    End If
    
On Error GoTo 0
Exit Function

Handler:
    CloseThemeData lhTheme
    m_sThemeName = "None"

End Function
