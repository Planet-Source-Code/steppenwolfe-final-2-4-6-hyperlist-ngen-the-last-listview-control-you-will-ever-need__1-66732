VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHyperList"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-> go big, or go home..

                                                                           
'       _/    _/  _/              _/      _/    _/_/_/  _/_/_/_/  _/      _/
'      _/    _/  _/              _/_/    _/  _/        _/        _/_/    _/
'     _/_/_/_/  _/    _/_/_/_/  _/  _/  _/  _/  _/_/  _/_/_/    _/  _/  _/
'    _/    _/  _/              _/    _/_/  _/    _/  _/        _/    _/_/
'   _/    _/  _/_/_/_/        _/      _/    _/_/_/  _/_/_/_/  _/      _/


'****************************************************************************************
'*  HyperList G2!   HL-NGEN - Virtual Listview Hybrid 2.1  (dll version)                *
'*                                                                                      *
'*  Created:        (V1) June 23, 2006                                                  *
'*  Updated:        October 3, 2006                                                     *
'*  Purpose:        Ultra-Fast Virtual Listview Hybrid (dll version)                    *
'*  Functions:      (listed)                                                            *
'*  Revision:       2.1.0                                                               *
'*  Compile:        Native                                                              *
'*  Author:         John Underhill (Steppenwolfe)                                       *
'*                                                                                      *
'****************************************************************************************

' ~*** 47 Exposed Functions ***~

'/~ AlphaSelectorBar            - use alpha bar effect              [in -byte | in -bool (2) | out -bool]
'/~ BackgroundPicture           - background image                  [in -string | enum | out -bool]
'/~ CheckAll                    - mark checkboxes                   [in -none | out -bool]
'/~ ClearList                   - clear items                       [in -none  | out -bool]
'/~ ColumnAdd                   - add column                        [in -long (2) | in -string |in -enum (3) | out -bool]
'/~ ColumnAutosize              - autosize columns                  [in -long | in-bool | out -bool]
'/~ ColumnClear                 - remove all columns                [in -none | out -bool]
'/~ ColumnLastFit               - fit last column                   [in -none | out -bool]
'/~ ColumnRemove                - remove column                     [in -long | out -bool]
'/~ ColumnReorder               - reorder columns                   [in -long | out -bool]
'/~ ColumnSizeToItems           - size column to items              [in -long | out -bool]
'/~ CopyItemToClipboard         - copy item to clipboard            [in -long | out -bool]
'/~ ImlHeaderAddBmp             - add header bmp                    [in -long | out -bool]
'/~ ImlHeaderAddIcon            - add header icon                   [in -long | out -bool]
'/~ ImlLargeAddBmp              - add large bmp                     [in -long | out -bool]
'/~ ImlLargeAddIcon             - add large icon                    [in -long | out -bool]
'/~ ImlSmallAddBmp              - add small bmp                     [in -long | out -bool]
'/~ ImlSmallAddIcon             - add small icon                    [in -long | out -bool]
'/~ ImlStateAddBmp              - add state bmp                     [in -long | out -bool]
'/~ ImlStateAddIcon             - add state icon                    [in -long | out -bool]
'/~ InitImlHeader               - set header iml                    [in -none| out -bool]
'/~ InitImlLarge                - set large iml                     [in -none | out -bool]
'/~ InitImlSmall                - set small iml                     [in -long () | out -bool]
'/~ InitImlState                - set state iml                     [in -long | out -bool]
'/~ InitList                    - init listview                     [in -long (2) | out -bool]
'/~ ItemAdd                     - add list item                     [in -long (3) string (2) | out -bool]
'/~ ItemEnsureVisible           - scroll to item                    [in -long | out -bool]
'/~ ItemRemove                  - remove item                       [in -long | out -bool]
'/~ ItemsSort                   - sort items                        [in -long | in -bool | out -bool]
'/~ ItemTopIndex                - top item                          [in -none | out -long]
'/~ ListRefresh                 - refresh listview                  [in -none | out -none]
'/~ LoadArray                   - load items array                  [in -none | out -bool]
'/~ LoadFromFile                - load list from file               [in -string () | out -bool]
'/~ RemoveDuplicates            - remove list duplicates            [in -none | out -bool]
'/~ RowDecoration               - row colors                        [in -long (3) in -enum | in -bool | out -bool]
'/~ SaveToFile                  - save list items                   [in -string | out -bool]
'/~ SetItemCount                - init list                         [in -long | out -bool]
'/~ SkinCheckBox                - load checkboxskin                 [in -enum | in -bool | out -bool]
'/~ SkinHeaders                 - load header skin                  [in -none | out -bool]
'/~ SkinScrollBars              - load scrollbar skin               [in -none | out -bool]
'/~ SubIconIndex                - set subitem icon                  [in -long (2) | out -bool]
'/~ SubItemsAdd                 - add subitem                       [in -long (2) | in -string | out -bool]
'/~ UnCheckAll                  - unmark all checkboxes             [in -none | out -bool]
'/~ UnSkinAll                   - remove all skins                  [in -none | out -bool]
'/~ UnSkinCheckBox              - remove checkbox skin              [in -none | out -bool]
'/~ UnSkinHeaders               - remove header skin                [in -none | out -bool]
'/~ UnSkinScrollBars            - remove scrollbar skin             [in -none | out -bool]

' ~*** 66 Exposed Properties ***~

'/~ AlphaBarTheme               - use alpha bar theme colors        [bool]
'/~ AlphaBarTransparency        - alpha bar transparency index      [byte]
'/~ AlphaBarActive              - use alpha bar                     [bool]
'/~ AlphaThemeBackClr           - use theme backcolor               [bool]
'/~ Appearance                  - listview appearence               [enum]
'/~ AutoArrange                 - autoarrange list items            [bool]
'/~ BackColor                   - list backcolor                    [long]
'/~ BorderStyle                 - list border style                 [enum]
'/~ Checkboxes                  - use checkboxes                    [bool]
'/~ CheckBoxSkinStyle           - checkbox skin style               [enum]
'/~ Checked                     - checkbox state                    [bool]
'/~ ColumnAlign                 - column alignment                  [long]
'/~ ColumnCount                 - column count                      [long]
'/~ ColumnHeight                - column height                     [long]
'/~ ColumnIcon                  - column icon                       [long]
'/~ ColumnTag                   - column tag                        [string]
'/~ ColumnText                  - column text                       [string]
'/~ ColumnWidth                 - column width                      [long]
'/~ Count                       - item count                        [long]
'/~ CustomDraw                  - use custom draw                   [bool]
'/~ Enabled                     - listview enabled                  [bool]
'/~ Font                        - listview font                     [obj]
'/~ ForeColor                   - listview forecolor                [long]
'/~ FullRowSelect               - fullrow select                    [bool]
'/~ GridLines                   - listview gridlines                [bool]
'/~ HeaderColor                 - header color (non skin)           [long]
'/~ HeaderCustom                - use custom header colors          [bool]
'/~ HeaderDragDrop              - header drag and drop              [bool]
'/~ HeaderFixedWidth            - columns fixed width               [bool]
'/~ HeaderFlat                  - flat headers (non skin)           [bool]
'/~ HeaderForeColor             - header forecolor                  [long]
'/~ HeaderHide                  - hide headers                      [bool]
'/~ HeaderHighLite              - header font highlite              [long]
'/~ HideSelection               - hide item                         [bool]
'/~ IconSpaceX                  - icon shift position X             [long]
'/~ IconSpaceY                  - icon shift position Y             [long]
'/~ InfoTips                    - use info tips                     [bool]
'/~ ItemBorderSelect            - use item border                   [bool]
'/~ ItemChecked                 - item check state                  [bool]
'/~ ItemFocused                 - item focus state                  [bool]
'/~ ItemGhosted                 - item ghosted                      [bool]
'/~ ItemIcon                    - item icon index                   [long]
'/~ ItemIndent                  - item indent                       [long]
'/~ ItemSelected                - item selected state               [bool]
'/~ ItemsSorted                 - item sorted state                 [bool]
'/~ ItemText                    - item text                         [string]
'/~ LabelEdit                   - label edit                        [bool]
'/~ LabelTips                   - label tips                        [bool]
'/~ ListMode                    - list access mode                  [enum]
'/~ MultiSelect                 - item multi select                 [bool]
'/~ OneClickActivate            - edit one click activate           [bool]
'/~ ScrollBarFlat               - flat scrollbar (non skin)         [bool]
'/~ SelectedCount               - selected items count              [long]
'/~ StructPtr                   - array pointer                     [long]
'/~ SubItemIcon                 - subitem icon index                [long]
'/~ SubItemImages               - use subitem images                [bool]
'/~ SubItemText                 - subitem text                      [string]
'/~ ThemeColor                  - theme base color                  [long]
'/~ ThemeLuminence              - theme luminence                   [enum]
'/~ TrackSelected               - track selected item               [bool]
'/~ UnderlineHot                - underline hot item                [bool]
'/~ UseThemeColors              - use skin theme color              [bool]
'/~ ViewMode                    - listview style                    [enum]
'/~ Visible                     - listview visible                  [bool]
'/~ WordWrap                    - item text wordwrap                [bool]
'/~ XPColors                    - use xp color offset               [bool]

'~*** Notes ***~

'-! Disclaimer
'/~ Obviously no warranty or liability, or any responsibility in any way imaginable, is expressed or implied.
'/~ Use this software in your projects in any way you wish under a relaxed GNU, but all responsibilities
'/~ are yours entirely. If this control or class is ported to a commercial project, I expect to be notified,
'/~ and at a minimum, the appropriate credit must be given in the help/about or other suitable area of the
'/~ software, ex. 'HL-NGEN listview class provided by John Underhill of NSPowertools.com'.

'-* Credits/Cudos
'~ A big thanks to Zhu Jin Yong for adding unicode support to the Hyperlist control class.
'~ A shout out to Carles 'da man!' PV, for his awesome api listview:
'~ http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56021&lngWId=1
'~ Much of the original HyperList class was derived from Carles example, and without Carles great demo of
'~ styles and methods this demonstration would have been much harder, (if not impossible), to create.
'~ Steve 'big Steve' McMahon, and his unparalleled listview control:
'~ http://www.vbaccelerator.com/home/VB/Code/Controls/ListView/article.asp
'~ if it is possible to do it with a listview, Steve has demonstrated it with this control, and as always,
'~ a great wealth of information and inspiration lies in the source code of this control.
'~ Rohan 'the Sort Monster' RDE, for his incredible QSort routines:
'~ http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63800&lngWId=1
'~ The Qsort routines are a simply inspirational use of refined logic and methods.
'~ and of course, as per usual, a (weak and confusing) reference per M$:
'~ http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/listview_overview.asp?frame=true
'~ and.. http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/messages/lvm_geteditcontrol.asp

'-> History
'/~ June 23, 2006 - V1 released as a simple demonstration of how to use a virtual listview with internal storage.
'/~ October 3, 2006 - V2 released with full feature set, skinning and 3 access modes.
'/~ I got a number of requests to clean up v1, and add features, so here it is..

'-! Compiling and Distribution
'/~ When compiling a library for release, -always- set the version compatability to binary.
'/~ Project > Properties > Component > Version Compatabilty > 'Binary Compatability'.
'/~ If you plan to distribute compiled dll, Rename it! Use a unique name, with a app/company name prefix,
'/~ ex. sgrhlst.dll. If you do not, you are taking a risk that someone elses version will overwrite
'/~ your own, and your app will stop working!
'/~ The compile switches are set, but you could change them to your liking and test responsiveness
'/~ under different option settings.
'/~ If you use this in a commercial product, I expect to be notified, and a mention in the about/help or
'/~ other suitable spot is a basic professional courtesy, otherwise you are free to use this any way
'/~ you like [excepting malicious].
'/~ You can distribute this project to other vb related websites, but headers must remain intact,
'/~ and proper credit must be given.

'-> 98/ME users
'/~ Should be good on legacy OS, any problems, just let me know, -but be specific [where/when/how].

'-> Using Hyperlist
'-> A bit about virtual listviews.. They are 'virtual' because very little information is actually stored
'/~ in the listview itself. The application is almost entirely responsible for maintaining item state, text, icons
'/~ checkboxes etc. When you add an item to a standard listview, you are actually populating a complex
'/~ type structure with many different branches and sub structures. Data used for sorting, spacing, icons etc, are all
'/~ stored in this compound structure, and often it is data you do not need. This is why a standard listviews access
'/~ methods are both slow and consume a great deal of memory, because wether you will use the information or not, it is
'/~ stored and contstantly queried through the listviews internal methods. The virtual list however, relies on
'/~ callback interfaces to retrieve data about an item. When an item scrolls into view, the parent window gets a
'/~ message requesting the items data. In this way, you can store only the information you need, and supply
'/~ a specific items elements only when it becomes visible.
'/~ Virtual listviews are commonly used in C++ to display very large databases. The database itself holds all the
'/~ information, and item data is fetched from the application through a callback to the hosting object.
'/~ The problem with this method though, is that there is no persistant data, and that a database and connectors
'/~ must be available to the application. With no internal storage facility, and said requirements, a virtual listviews
'/~ implementation was usually very specific, and did not find its way into many applications.
'/~ Hyperlist has overcome this limitation, with an internal storage container, much of the raw speed of a virtual
'/~ listview has been maintained (several thousand times faster then a standard listview), but most of the
'/~ properties and methods of a standard listview are also accessable. Data is stored in several different
'/~ ways, depending on the access method..

'-> Data Access modes:
'/~ There are three different data access modes in HL: Custom Draw, Hyper Mode, and DataBase.
'/~ Each has different access methods and is geared towards specific design criteria:
'-> Hyper mode:
'/~ The fastest method utilizing internal storage.
'/~ This method uses a type structure to hold the data. The arrays in the struct can either be populated directly
'/~ in place of your applications data arrays, or array pointers can be swapped into the struct. The application
'/~ builds a copy of two structures, items and subitems [HLIStc and HLISubItm], these types are also in the
'/~ clsHyperList class. Once the application local structure has been built, the struct is copied into the class
'/~ by copying in the structure pointer via the StructPtr property. example:
'/~ CopyMemory m_lPointer, ByVal VarPtrArray(m_HLIStc), 4&          <- copy local struct to pointer
'/~ cHyperList.StructPtr = m_lPointer                               <- copy pointer into class
'/~ cHyperList.LoadArray                                            <- dimension the in-class data struct
'/~ cHyperList.SetItemCount UBound(m_HLIStc(0).item) + 1            <- set the item count and instantiate
'/~ In this way your application data can be transferred into the listview almost instantaneously.
'/~ Access is also available through the wrapper with the methods [InitList and ItemAdd].

'-> Custom Draw mode:
'/~ The most flexible method, combining both speed and per item properties.
'/~ The custom draw mode uses a class array to store item data. Properties for each item, including font/forecolor
'/~ backcolor/icon/key, and even per cell formatting can be set in each class instance.
'/~ Access method is first dimensioning array to the size of the application data array, then populating
'/~ each class member with desired properies, local build example:
'/~ Redim Preserve m_cListItems(0 To 999)                           <- dimension the local class array
'/~ Set m_cListItems(999) = New clsListItem                         <- instantiate members
'/~ m_cListItems(999).Add 0, "key1", "item text", lIcon, lSubIcon   <- add items
'/~ CopyMemBr m_lPointer, ByVal VarPtrArray(m_cListItems), 4&       <- copy array to local pointer
'/~ cHyperList.StructPtr = m_lPointer                               <- copy pointer to class
'/~ cHyperList.LoadArray                                            <- dimension the in-class array
'/~ cHyperList.SetItemCount UBound(m_cListItems) + 1                <- set the item count and instantiate
'/~ The clsListItem class is a public member, so can be instantiated directly from the application, or
'/~ the internal methods can be used [InitList and ItemAdd].

'-> DataBase mode:
'/~ This mode is in keeping with the traditional virtual listviews access method. A callback event [eHIndirect]
'/~ is used to fetch the data from an indexed database. Events have also been added for labeledit, item and column
'/~ click events.

'-> All three access methods require the SetItemCount call [ubound(itemsarray) + 1] to initialize the listviews items.

'-> Features
'/~ There are a lot of properties (over 110 functions/properties), so I suggest some experimenting to get their correct
'/~ implementation for those that are not demonstrated. Most properties of a listview that can be implemented are available,
'/~ some things though were not possible, (like grouping), because they require the listview to maintain item state data.

'-> Skinning and Theming
'/~ 5 seperate skin themes for checkbox/header/scrollbars. Checkbox also has an added xp style.
'/~ Skins can all be colorized with user selected color, using the [ThemeLuminence ThemeColor and UseThemes]
'/~ properties. Skin styles can be mixed and matched between components by setting their individual properties,
'/~ or applied universally using the [SkinAll] function.

'-@ steppenwolfe_2000@yahoo.com
'-> Cheers,
'-> John



Implements GXISubclass

Private Const NEG1                              As Long = -1
Private Const m0                                As Long = &H0
Private Const m1                                As Long = &H1
Private Const m2                                As Long = &H2
Private Const m4                                As Long = &H4
Private Const m8                                As Long = &H8
Private Const m32                               As Long = &H20

Private Const CCM_FIRST                         As Long = &H2000
Private Const CCM_SETUNICODEFORMAT              As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT              As Long = (CCM_FIRST + 6)

Private Const CDIS_SELECTED                     As Long = &H1
Private Const CDIS_GRAYED                       As Long = &H2
Private Const CDIS_DISABLED                     As Long = &H4
Private Const CDIS_CHECKED                      As Long = &H8
Private Const CDIS_FOCUS                        As Long = &H10
Private Const CDIS_DEFAULT                      As Long = &H20
Private Const CDIS_HOT                          As Long = &H40
Private Const CDIS_MARKED                       As Long = &H80
Private Const CDIS_INDETERMINATE                As Long = &H100

Private Const CLR_DEFAULT                       As Long = -16777216
Private Const CLR_HILIGHT                       As Long = -16777216
Private Const CLR_NONE                          As Long = -1

Private Const H_MAX                             As Long = &HFFFF + 1
Private Const CDDS_PREPAINT                     As Long = &H1
Private Const CDDS_POSTPAINT                    As Long = &H2
Private Const CDDS_PREERASE                     As Long = &H3
Private Const CDDS_POSTERASE                    As Long = &H4
Private Const CDDS_ITEM                         As Long = &H10000
Private Const CDDS_ITEMPREPAINT                 As Long = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_ITEMPOSTPAINT                As Long = CDDS_ITEM Or CDDS_POSTPAINT
Private Const CDDS_SUBITEM                      As Long = &H20000

Private Const CDRF_DODEFAULT                    As Long = &H0
Private Const CDRF_NOTIFYITEMDRAW               As Long = &H20
Private Const CDRF_NOTIFYSUBITEMDRAW            As Long = &H20

Private Const DST_ICON                          As Long = &H3
Private Const DSS_DISABLED                      As Long = &H20

Private Const ES_UPPERCASE                      As Long = &H8&
Private Const ES_LOWERCASE                      As Long = &H10

Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const GWL_STYLE                         As Long = (-16)
Private Const GWL_EXSTYLE                       As Long = (-20)

Private Const HDF_LEFT                          As Long = 0
Private Const HDF_RIGHT                         As Long = 1
Private Const HDF_CENTER                        As Long = 2
Private Const HDF_IMAGE                         As Long = &H800
Private Const HDF_STRING                        As Long = &H4000
Private Const HDF_BITMAP_ON_RIGHT               As Long = &H1000

Private Const HDI_WIDTH                         As Long = &H1
Private Const HDI_TEXT                          As Long = &H2
Private Const HDI_FORMAT                        As Long = &H4
Private Const HDI_IMAGE                         As Long = &H20

Private Const HDM_FIRST                         As Long = &H1200
Private Const HDM_GETITEMCOUNT                  As Long = (HDM_FIRST + 0)
Private Const HDM_SETIMAGELIST                  As Long = (HDM_FIRST + 8)
#If Unicode Then
Private Const HDM_GETITEMW                      As Long = (HDM_FIRST + 11)
Private Const HDM_SETITEMW                      As Long = (HDM_FIRST + 12)
Private Const HDM_INSERTITEMW                   As Long = (HDM_FIRST + 10)
Private Const HDM_GETITEM                       As Long = HDM_GETITEMW
Private Const HDM_SETITEM                       As Long = HDM_SETITEMW
Private Const HDM_INSERTITEM                    As Long = HDM_INSERTITEMW
#Else
Private Const HDM_GETITEMA                      As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEMA                      As Long = (HDM_FIRST + 4)
Private Const HDM_INSERTITEMA                   As Long = (HDM_FIRST + 1)
Private Const HDM_GETITEM                       As Long = HDM_GETITEMA
Private Const HDM_SETITEM                       As Long = HDM_SETITEMA
Private Const HDM_INSERTITEM                    As Long = HDM_INSERTITEMA
#End If

Private Const HDN_FIRST                         As Long = -300
#If Unicode Then
Private Const HDN_ITEMCLICKW                    As Long = (HDN_FIRST - 22)
Private Const HDN_ITEMCHANGINGW                 As Long = (HDN_FIRST - 20)
Private Const HDN_ITEMCHANGEDW                  As Long = (HDN_FIRST - 21)
Private Const HDN_ITEMCLICK                     As Long = HDN_ITEMCLICKW
Private Const HDN_ITEMCHANGING                  As Long = HDN_ITEMCHANGINGW
Private Const HDN_ITEMCHANGED                   As Long = HDN_ITEMCHANGEDW
#Else
Private Const HDN_ITEMCLICKA                    As Long = (HDN_FIRST - 2)
Private Const HDN_ITEMCHANGINGA                 As Long = (HDN_FIRST - 0)
Private Const HDN_ITEMCHANGEDA                  As Long = (HDN_FIRST - 1)
Private Const HDN_ITEMCLICK                     As Long = HDN_ITEMCLICKA
Private Const HDN_ITEMCHANGING                  As Long = HDN_ITEMCHANGINGA
Private Const HDN_ITEMCHANGED                   As Long = HDN_ITEMCHANGEDA
#End If

Private Const HDS_BUTTONS                       As Long = &H2

Private Const ICC_LISTVIEW_CLASSES              As Long = &H1

Private Const ILC_MASK                          As Long = &H1
Private Const ILC_COLOR32                       As Long = &H20

Private Const ILD_NORMAL                        As Long = &H0
Private Const ILD_TRANSPARENT                   As Long = &H1
Private Const ILD_BLEND25                       As Long = &H2
Private Const ILD_SELECTED                      As Long = &H4
Private Const ILD_FOCUS                         As Long = &H4
Private Const ILD_MASK                          As Long = &H10&
Private Const ILD_IMAGE                         As Long = &H20&
Private Const ILD_ROP                           As Long = &H40&
Private Const ILD_OVERLAYMASK                   As Long = 3840&


Private Const LVBKIF_SOURCE_NONE                As Long = &H0
Private Const LVBKIF_SOURCE_HBITMAP             As Long = &H1
Private Const LVBKIF_SOURCE_URL                 As Long = &H2
Private Const LVBKIF_SOURCE_MASK                As Long = &H3
Private Const LVBKIF_STYLE_NORMAL               As Long = &H0
Private Const LVBKIF_STYLE_TILE                 As Long = &H10
Private Const LVBKIF_STYLE_MASK                 As Long = &H10

Private Const LVCF_FMT                          As Long = &H1
Private Const LVCF_WIDTH                        As Long = &H2
Private Const LVCF_TEXT                         As Long = &H4
Private Const LVCF_SUBITEM                      As Long = &H8
Private Const LVCF_IMAGE                        As Long = &H10
Private Const LVCF_ORDER                        As Long = &H20

Private Const LVIS_CHECKED                      As Long = &H2000&
Private Const LVIS_UNCHECKED                    As Long = &H1000&
Private Const LVIS_CHKCLICK                     As Long = &HFFFE

Private Const LVM_FIRST                         As Long = &H1000
Private Const LVM_GETBKCOLOR                    As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR                    As Long = (LVM_FIRST + 1)
Private Const LVM_SETIMAGELIST                  As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT                  As Long = (LVM_FIRST + 4)
Private Const LVM_ENSUREVISIBLE                 As Long = (LVM_FIRST + 19)
Private Const LVM_GETEDITCONTROL                As Long = (LVM_FIRST + 24)
Private Const LVM_DELETECOLUMN                  As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH                As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH                As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER                     As Long = (LVM_FIRST + 31)
Private Const LVM_GETTEXTCOLOR                  As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR                  As Long = (LVM_FIRST + 36)
Private Const LVM_SETTEXTBKCOLOR                As Long = (LVM_FIRST + 38)
Private Const LVM_SETITEMSTATE                  As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE                  As Long = (LVM_FIRST + 44)
Private Const LVM_GETSELECTEDCOUNT              As Long = (LVM_FIRST + 50)
Private Const LVM_SETICONSPACING                As Long = (LVM_FIRST + 53)

#If Unicode Then
Private Const LVM_GETITEMW                      As Long = (LVM_FIRST + 75)
Private Const LVM_SETITEMW                      As Long = (LVM_FIRST + 76)
Private Const LVM_INSERTITEMW                   As Long = (LVM_FIRST + 77)
Private Const LVM_FINDITEMW                     As Long = (LVM_FIRST + 83)
Private Const LVM_GETCOLUMNW                    As Long = (LVM_FIRST + 95)
Private Const LVM_SETCOLUMNW                    As Long = (LVM_FIRST + 96)
Private Const LVM_INSERTCOLUMNW                 As Long = (LVM_FIRST + 97)
Private Const LVM_GETITEMTEXTW                  As Long = (LVM_FIRST + 115)
Private Const LVM_SETITEMTEXTW                  As Long = (LVM_FIRST + 116)
Private Const LVM_EDITLABELW                    As Long = (LVM_FIRST + 118)
Private Const LVM_SETBKIMAGEW                   As Long = (LVM_FIRST + 138)
Private Const LVM_GETBKIMAGEW                   As Long = (LVM_FIRST + 139)
Private Const LVM_SETBKIMAGE                    As Long = LVM_SETBKIMAGEW
Private Const LVM_GETBKIMAGE                    As Long = LVM_GETBKIMAGEW
Private Const LVM_INSERTCOLUMN                  As Long = LVM_INSERTCOLUMNW
Private Const LVM_INSERTITEM                    As Long = LVM_INSERTITEMW
Private Const LVM_SETCOLUMN                     As Long = LVM_SETCOLUMNW
Private Const LVM_SETITEM                       As Long = LVM_SETITEMW
Private Const LVM_SETITEMTEXT                   As Long = LVM_SETITEMTEXTW
Private Const LVM_EDITLABEL                     As Long = LVM_EDITLABELW
Private Const LVM_FINDITEM                      As Long = LVM_FINDITEMW
Private Const LVM_GETCOLUMN                     As Long = LVM_GETCOLUMNW
Private Const LVM_GETITEM                       As Long = LVM_GETITEMW
Private Const LVM_GETITEMTEXT                   As Long = LVM_GETITEMTEXTW
#Else
Private Const LVM_GETITEMA                      As Long = (LVM_FIRST + 5)
Private Const LVM_SETITEMA                      As Long = (LVM_FIRST + 6)
Private Const LVM_INSERTITEMA                   As Long = (LVM_FIRST + 7)
Private Const LVM_FINDITEMA                     As Long = (LVM_FIRST + 13)
Private Const LVM_EDITLABELA                    As Long = (LVM_FIRST + 23)
Private Const LVM_GETCOLUMNA                    As Long = (LVM_FIRST + 25)
Private Const LVM_SETCOLUMNA                    As Long = (LVM_FIRST + 26)
Private Const LVM_INSERTCOLUMNA                 As Long = (LVM_FIRST + 27)
Private Const LVM_GETITEMTEXTA                  As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMTEXTA                  As Long = (LVM_FIRST + 46)
Private Const LVM_SETBKIMAGEA                   As Long = (LVM_FIRST + 68)
Private Const LVM_GETBKIMAGEA                   As Long = (LVM_FIRST + 69)
Private Const LVM_INSERTCOLUMN                  As Long = LVM_INSERTCOLUMNA
Private Const LVM_INSERTITEM                    As Long = LVM_INSERTITEMA
Private Const LVM_SETBKIMAGE                    As Long = LVM_SETBKIMAGEA
Private Const LVM_GETBKIMAGE                    As Long = LVM_GETBKIMAGEA
Private Const LVM_SETCOLUMN                     As Long = LVM_SETCOLUMNA
Private Const LVM_SETITEM                       As Long = LVM_SETITEMA
Private Const LVM_SETITEMTEXT                   As Long = LVM_SETITEMTEXTA
Private Const LVM_EDITLABEL                     As Long = LVM_EDITLABELA
Private Const LVM_FINDITEM                      As Long = LVM_FINDITEMA
Private Const LVM_GETCOLUMN                     As Long = LVM_GETCOLUMNA
Private Const LVM_GETITEM                       As Long = LVM_GETITEMA
Private Const LVM_GETITEMTEXT                   As Long = LVM_GETITEMTEXTA
#End If

Private Const LVN_FIRST                         As Long = -100&
Private Const LVN_LAST                          As Long = -199&

#If Unicode Then
Private Const LVN_GETDISPINFOW                  As Long = (LVN_FIRST - 77)
Private Const LVN_SETDISPINFOW                  As Long = (LVN_FIRST - 78)
Private Const LVN_BEGINLABELEDITW               As Long = (LVN_FIRST - 75)
Private Const LVN_ENDLABELEDITW                 As Long = (LVN_FIRST - 76)
Private Const LVN_GETDISPINFO                   As Long = LVN_GETDISPINFOW
Private Const LVN_SETDISPINFO                   As Long = LVN_SETDISPINFOW
Private Const LVN_BEGINLABELEDIT                As Long = LVN_BEGINLABELEDITW
Private Const LVN_ENDLABELEDIT                  As Long = LVN_ENDLABELEDITW
#Else
Private Const LVN_SETDISPINFOA                  As Long = (LVN_FIRST - 51)
Private Const LVN_GETDISPINFOA                  As Long = (LVN_FIRST - 50)
Private Const LVN_BEGINLABELEDITA               As Long = (LVN_FIRST - 5)
Private Const LVN_ENDLABELEDITA                 As Long = (LVN_FIRST - 6)
Private Const LVN_GETDISPINFO                   As Long = LVN_GETDISPINFOA
Private Const LVN_SETDISPINFO                   As Long = LVN_SETDISPINFOA
Private Const LVN_BEGINLABELEDIT                As Long = LVN_BEGINLABELEDITA
Private Const LVN_ENDLABELEDIT                  As Long = LVN_ENDLABELEDITA
#End If

Private Const LVS_EX_UNDERLINEHOT               As Long = &H800&
Private Const LVS_EX_TRACKSELECT                As Long = &H8&
Private Const LVS_AUTOARRANGE                   As Long = &H100
Private Const LVS_EDITLABELS                    As Long = &H200
Private Const LVS_EX_BORDERSELECT               As Long = &H8000&
Private Const LVS_EX_INFOTIP                    As Long = &H400&
Private Const LVS_ICON                          As Long = &H0
Private Const LVS_REPORT                        As Long = &H1
Private Const LVS_SMALLICON                     As Long = &H2
Private Const LVS_LIST                          As Long = &H3
Private Const LVS_EX_GRIDLINES                  As Long = &H1&
Private Const LVS_EX_CHECKBOXES                 As Long = &H4&
Private Const LVS_EX_HEADERDRAGDROP             As Long = &H10&
Private Const LVS_EX_FULLROWSELECT              As Long = &H20&
Private Const LVS_EX_ONECLICKACTIVATE           As Long = &H40&
Private Const LVS_EX_FLATSB                     As Long = &H100&
Private Const LVS_EX_LABELTIP                   As Long = &H4000&
Private Const LVS_SINGLESEL                     As Long = &H4
Private Const LVS_SHOWSELALWAYS                 As Long = &H8
Private Const LVS_SORTASCENDING                 As Long = &H10
Private Const LVS_SHAREIMAGELISTS               As Long = &H40
Private Const LVS_OWNERDATA                     As Long = &H1000
Private Const LVS_NOCOLUMNHEADER                As Long = &H4000

Private Const LVSCW_AUTOSIZE                    As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER          As Long = -2

Private Const LVSIL_NORMAL                      As Long = 0
Private Const LVSIL_SMALL                       As Long = 1
Private Const LVSIL_STATE                       As Long = 2

Private Const NM_FIRST                          As Long = H_MAX
Private Const NM_CUSTOMDRAW                     As Long = (NM_FIRST - 12)
Private Const NM_CLICK                          As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK                         As Long = (NM_FIRST - 3)
Private Const NM_RETURN                         As Long = (NM_FIRST - 4)
Private Const NM_KILLFOCUS                      As Long = (NM_FIRST - 8)
Private Const NM_RCLICK                         As Long = (NM_FIRST - 5)


Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOOWNERZORDER                 As Long = &H200
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_FRAMECHANGED                  As Long = &H20

Private Const WC_LISTVIEW                       As String = "SysListView32"

Private Const WM_SETFONT                        As Long = &H30
Private Const WM_SETFOCUS                       As Long = &H7
Private Const WM_NOTIFY                         As Long = &H4E
Private Const WM_KEYDOWN                        As Long = &H100
Private Const WM_KEYUP                          As Long = &H101
Private Const WM_CHAR                           As Long = &H102
Private Const WM_MOUSEMOVE                      As Long = &H200
Private Const WM_LBUTTONUP                      As Long = &H202
Private Const WM_LBUTTONDOWN                    As Long = &H201
Private Const WM_RBUTTONDOWN                    As Long = &H204
Private Const WM_RBUTTONUP                      As Long = &H205
Private Const WM_MBUTTONDOWN                    As Long = &H207
Private Const WM_MBUTTONUP                      As Long = &H208
Private Const WM_TIMER                          As Long = &H113&

Private Const WS_EX_WINDOWEDGE                  As Long = &H100&
Private Const WS_EX_CLIENTEDGE                  As Long = &H200&
Private Const WS_EX_STATICEDGE                  As Long = &H20000
Private Const WS_TABSTOP                        As Long = &H10000
Private Const WS_THICKFRAME                     As Long = &H40000
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_VISIBLE                        As Long = &H10000000
Private Const WS_CHILD                          As Long = &H40000000
Private Const SW_HIDE                           As Long = &H0
Private Const SW_NORMAL                         As Long = &H1

Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0&
    SM_CYSCREEN = 1&
    SM_CXVSCROLL = 2&
    SM_CYHSCROLL = 3&
    SM_CYCAPTION = 4&
    SM_CXBORDER = 5&
    SM_CYBORDER = 6&
    SM_CYVTHUMB = 9&
    SM_CXHTHUMB = 10&
    SM_CXICON = 11&
    SM_CYICON = 12&
    SM_CXCURSOR = 13&
    SM_CYCURSOR = 14&
    SM_CYMENU = 15&
    SM_CXFULLSCREEN = 16&
    SM_CYFULLSCREEN = 17&
    SM_CYKANJIWINDOW = 18&
    SM_MOUSEPRESENT = 19&
    SM_CYVSCROLL = 20&
    SM_CXHSCROLL = 21&
    SM_CXMIN = 28&
    SM_CYMIN = 29&
    SM_CXSIZE = 30&
    SM_CYSIZE = 31&
    SM_CXFRAME = 32&
    SM_CYFRAME = 33&
    SM_CXMINTRACK = 34&
    SM_CYMINTRACK = 35&
    SM_CXSMICON = 49&
    SM_CYSMICON = 50&
    SM_CYSMCAPTION = 51&
    SM_CXMINIMIZED = 57&
    SM_CYMINIMIZED = 58&
    SM_CXMAXTRACK = 59&
    SM_CYMAXTRACK = 60&
    SM_CXMAXIMIZED = 61&
    SM_CYMAXIMIZED = 62&
End Enum

Private Enum LVM_SETITEMCOUNT_LPARAM
    LVSICF_NOINVALIDATEALL = &H1
    LVSICF_NOSCROLL = &H2
End Enum

Private Enum TT_NOTIFICATIONS
    TTN_FIRST = -520&
    TTN_LAST = -549&
    TTN_GETDISPINFO = (TTN_FIRST - 0)
End Enum

Private Enum LISTVIEW_MESSAGES
    LVM_DELETEALLITEMS = (LVM_FIRST + 9)
    LVM_GETITEMRECT = (LVM_FIRST + 14)
    LVM_HITTEST = (LVM_FIRST + 18)
    LVM_SCROLL = (LVM_FIRST + 20)
    LVM_GETTOPINDEX = (LVM_FIRST + 39)
    LVM_SETITEMCOUNT = (LVM_FIRST + 47)
    LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
    LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
    LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
End Enum

Private Enum LV_ITEM_MASK
    LVIF_TEXT = &H1
    LVIF_IMAGE = &H2
    LVIF_PARAM = &H4
    LVIF_STATE = &H8
    LVIF_INDENT = &H10
    LVIF_NORECOMPUTE = &H800
    LVIF_DI_SETITEM = &H1000
End Enum

Private Enum LV_ITEM_STATE
    LVIS_FOCUSED = &H1
    LVIS_SELECTED = &H2
    LVIS_CUT = &H4
    LVIS_DROPHILITED = &H8
    LVIS_OVERLAYMASK = &HF00
    LVIS_STATEIMAGEMASK = &HF000
    LVIS_ALL = LVIS_FOCUSED Or LVIS_SELECTED Or LVIS_CUT Or LVIS_DROPHILITED Or LVIS_OVERLAYMASK Or LVIS_STATEIMAGEMASK
End Enum

Private Enum LVN_NOTIFY
    LVN_ITEMCHANGING = (LVN_FIRST - 0)
    LVN_ITEMCHANGED = (LVN_FIRST - 1)
    LVN_INSERTITEM = (LVN_FIRST - 2)
    LVN_DELETEITEM = (LVN_FIRST - 3)
    LVN_DELETEALLITEMS = (LVN_FIRST - 4)
    LVN_COLUMNCLICK = (LVN_FIRST - 8)
    LVN_BEGINDRAG = (LVN_FIRST - 9)
    LVN_BEGINRDRAG = (LVN_FIRST - 11)
    LVN_ODCACHEHINT = (LVN_FIRST - 13)
    LVN_ITEMACTIVATE = (LVN_FIRST - 14)
    LVN_ODSTATECHANGED = (LVN_FIRST - 15)
    LVN_ODFINDITEM = (LVN_FIRST - 52)
    LVN_KEYDOWN = (LVN_FIRST - 55)
    LVN_MARQUEEBEGIN = (LVN_FIRST - 56)
End Enum


Public Enum ECGTextAlignFlags
    DT_TOP = &H0&
    DT_LEFT = &H0&
    DT_CENTER = &H1&
    DT_RIGHT = &H2&
    DT_VCENTER = &H4&
    DT_BOTTOM = &H8&
    DT_WORDBREAK = &H10&
    DT_SINGLELINE = &H20&
    DT_EXPANDTABS = &H40&
    DT_TABSTOP = &H80&
    DT_NOCLIP = &H100&
    DT_EXTERNALLEADING = &H200&
    DT_CALCRECT = &H400&
    DT_NOPREFIX = &H800&
    DT_INTERNAL = &H1000&
    DT_EDITCONTROL = &H2000&
    DT_PATH_ELLIPSIS = &H4000&
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Public Enum ECGAppearance
    [Flat] = 1
    [3D] = 4
End Enum

Public Enum ECGBackGroundImage
    [BgNone] = -1
    [BgNormal] = LVBKIF_STYLE_NORMAL
    [BgTile] = LVBKIF_STYLE_TILE
End Enum

Public Enum ECGBorderStyle
    [None] = 0
    [Fixed Single] = 2
    [Thin] = 3
End Enum

Public Enum ECGCheckBoxSkinStyle
    [CheckBoxClassic] = 0
    [CheckBoxEclipse] = 1
    [CheckBoxLime] = 2
    [CheckBoxMetallic] = 3
    [CheckBoxGloss] = 4
    [CheckBoxXP] = 5
End Enum

Public Enum ECGColumnAutosize
    [ColumnItem] = LVSCW_AUTOSIZE
    [ColumnHeader] = LVSCW_AUTOSIZE_USEHEADER
End Enum

Public Enum ECGColumnAlign
    [ColumnLeft] = HDF_LEFT
    [ColumnRight] = HDF_RIGHT
    [ColumnEnter] = HDF_CENTER
End Enum

Public Enum ECGColumnSortTags
    [SortNone] = -1
    [SortDefault] = 0
    [SortDate] = 1
    [SortNumeric] = 2
    [SortAuto] = 3
End Enum

Public Enum ECGHeaderSkinStyle
    [HeaderClassic] = 0
    [HeaderEclipse] = 1
    [HeaderLime] = 2
    [HeaderMetallic] = 3
    [HeaderGloss] = 4
End Enum

Public Enum ECGListMode
    [eCustomDraw] = 0
    [eDatabase] = 1
    [eHyperList] = 2
End Enum

Public Enum ECGStyle
    [StyleReport] = LVS_REPORT
    [StyleIcon] = LVS_ICON
    [StyleSmallIcon] = LVS_SMALLICON
    [StyleList] = LVS_LIST
End Enum

Public Enum ECGRowDecoration
    [RowLine] = 0
    [RowSplit] = 1
    [RowBiLinear] = 2
    [RowChecker] = 3
End Enum

Public Enum ECGScrollBarSkinStyle
    [ScrollClassic] = 0
    [ScrollEclipse] = 1
    [ScrollLime] = 2
    [ScrollMetallic] = 3
    [ScrollGloss] = 4
End Enum

Public Enum ECGThemeLuminence
    [ThemeSoft] = 0
    [ThemePastel] = 1
    [ThemeHard] = 2
End Enum


Private Type tagINITCOMMONCONTROLSEX
    dwSize                                      As Long
    dwICC                                       As Long
End Type

Private Type HLISubItm
    lIcon()                                     As Long
    Text()                                      As String
End Type

Private Type HLIStc
    item()                                      As String
    lIcon()                                     As Long
    SubItem()                                   As HLISubItm
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    right                                       As Long
    bottom                                      As Long
End Type

Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Type LVHITTESTINFO
    pt                                          As POINTAPI
    flags                                       As Long
    iItem                                       As Long
    iSubItem                                    As Long
End Type

Private Type LVCOLUMN
    Mask                                        As Long
    fmt                                         As Long
    cx                                          As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iSubItem                                    As Long
    iImage                                      As Long
    iOrder                                      As Long
End Type

Private Type HDITEM
    Mask                                        As Long
    cxy                                         As Long
    pszText                                     As Long
    hbm                                         As Long
    cchTextMax                                  As Long
    fmt                                         As Long
    lParam                                      As Long
    iImage                                      As Long
    iOrder                                      As Long
End Type

Private Type NMHDR
    hwndFrom                                    As Long
    idfrom                                      As Long
    code                                        As Long
End Type

Private Type NMCUSTOMDRAWINFO
    hdr                                         As NMHDR
    dwDrawStage                                 As Long
    hdc                                         As Long
    rc                                          As RECT
    dwItemSpec                                  As Long
    iItemState                                  As Long
    lItemLParam                                 As Long
End Type

Private Type NMLVCUSTOMDRAW
    nmcmd                                       As NMCUSTOMDRAWINFO
    clrText                                     As Long
    clrTextBk                                   As Long
    iSubItem                                    As Long
End Type

Private Type LV_ITEM
    Mask                                        As LV_ITEM_MASK
    iItem                                       As Long
    iSubItem                                    As Long
    State                                       As LV_ITEM_STATE
    stateMask                                   As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iImage                                      As Long
    lParam                                      As Long
    iIndent                                     As Long
End Type

Private Type LV_DISPINFO
    hdr                                         As NMHDR
    item                                        As LV_ITEM
End Type

Private Type NM_LISTVIEW
    hdr                                         As NMHDR
    iItem                                       As Long
    iSubItem                                    As Long
    uNewState                                   As Long
    uOldState                                   As Long
    uChanged                                    As Long
    ptAction                                    As POINTAPI
    lParam                                      As Long
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

Private Type LVBKIMAGE
    ulFlags                                     As Long
    hbm                                         As Long
    pszImage                                    As String
    cchImageMax                                 As Long
    xOffsetPercent                              As Long
    yOffsetPercent                              As Long
End Type


#If Unicode Then
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, _
                                                  lpString2 As Any) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (lpString As Any) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
#Else
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
#End If

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal x As Long, _
                                                                              ByVal y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hWndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_Add Lib "Comctl32" (ByVal hImageList As Long, _
                                                       ByVal hBitmap As Long, _
                                                       ByVal hBitmapMask As Long) As Long

Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImageList As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_AddIcon Lib "Comctl32" (ByVal hImageList As Long, _
                                                           ByVal hIcon As Long) As Long

Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_Draw Lib "Comctl32" (ByVal hImageList As Long, _
                                                        ByVal lIndex As Long, _
                                                        ByVal hdc As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long, _
                                                        ByVal fStyle As Long) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, _
                                                               ByVal i As Long, _
                                                               ByVal diIgnore As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function ImageList_DrawEx Lib "Comctl32" (ByVal hIml As Long, _
                                                          ByVal i As Long, _
                                                          ByVal hdcDst As Long, _
                                                          ByVal x As Long, _
                                                          ByVal y As Long, _
                                                          ByVal dX As Long, _
                                                          ByVal dy As Long, _
                                                          ByVal rgbBk As Long, _
                                                          ByVal rgbFg As Long, _
                                                          ByVal fStyle As Long) As Long

Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                    ByVal hBrush As Long, _
                                                                    ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, _
                                                                    ByVal wParam As Long, _
                                                                    ByVal x As Long, _
                                                                    ByVal y As Long, _
                                                                    ByVal cx As Long, _
                                                                    ByVal cy As Long, _
                                                                    ByVal fuFlags As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, _
                                                           ByVal HPALETTE As Long, _
                                                           pccolorref As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal fEnable As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Long, _
                                                      ByVal bErase As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Sub CopyMemBv Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Any, _
                                                                    ByVal pSrc As Any, _
                                                                    ByVal lByteLen As Long)

Private Declare Sub CopyMemBr Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                    pSrc As Any, _
                                                                    ByVal lByteLen As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SYSTEM_METRICS) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function InitCommonControlsEx Lib "Comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Declare Sub InitCommonControls Lib "Comctl32" ()

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


Public Event eHItemClick(lItem As Long)
Public Event eHItemCheck(lItem As Long)
Public Event eHColumnClick(Column As Long)
Public Event eHIndirect(ByVal iItem As Long, ByVal iSubItem As Long, ByVal fMask As Long, sText As String, hImage As Long)
Public Event eHLabelChange(ByVal iItem As Long, sText As String)
Public Event eHErrCond(sRtn As String, lErr As Long)

Private m_lDummy                                As Long
Private m_bCheckBoxes                           As Boolean
Private m_bFirstItem                            As Boolean
Private m_bFullRowSelect                        As Boolean
Private m_bGridLines                            As Boolean
Private m_bDragDrop                             As Boolean
Private m_bHeaderFixed                          As Boolean
Private m_bHeaderFlat                           As Boolean
Private m_bHeaderHide                           As Boolean
Private m_bHideSelection                        As Boolean
Private m_bLabelTips                            As Boolean
Private m_bMultiSelect                          As Boolean
Private m_bOneClick                             As Boolean
Private m_bScrollFlat                           As Boolean
Private m_bCheckInit                            As Boolean
Private m_bCustomHeader                         As Boolean
Private m_bSorted                               As Boolean
Private m_bUseSorted                            As Boolean
Private m_bSubItemImage                         As Boolean
Private m_bCustomDraw                           As Boolean
Private m_bEnabled                              As Boolean
Private m_bWordWrap                             As Boolean
Private m_bXPColors                             As Boolean
Private m_bSkinHeader                           As Boolean
Private m_bUseThemeColors                       As Boolean
Private m_bVisible                              As Boolean
Private m_bInfoTips                             As Boolean
Private m_bItemBorderSelect                     As Boolean
Private m_bOneClickActivate                     As Boolean
Private m_bEditLabels                           As Boolean
Private m_bAutoArrange                          As Boolean
Private m_bTrackSelected                        As Boolean
Private m_bUnderlineHot                         As Boolean
Private m_bSkinnedCheck                         As Boolean
Private m_bAlphaSelectorBar                     As Boolean
Private m_bUseCheckBoxTheme                     As Boolean
Private m_bRowUseXP                             As Boolean
Private m_bCustomCheckBox                       As Boolean
Private m_bRowDecoration                        As Boolean
Private m_bAlphaBarTheme                        As Boolean
Private m_bAlphaThemeBackClr                    As Boolean
Private m_bItemActive                           As Boolean
Private m_bteAlphaTransparency                  As Byte
Private m_lImlStateHndl                         As Long
Private m_lIconSpaceX                           As Long
Private m_lIconSpaceY                           As Long
Private m_lCheckHeight                          As Long
Private m_lCheckWidth                           As Long
Private m_lTmpBg                                As Long
Private m_lhMod                                 As Long
Private m_lItemIndent                           As Long
Private m_lParentHwnd                           As Long
Private m_lLVHwnd                               As Long
Private m_lHdrHwnd                              As Long
Private m_lImlHdHndl                            As Long
Private m_lImlSmallHndl                         As Long
Private m_lImlLargeHndl                         As Long
Private m_lItemsCnt                             As Long
Private m_lFont                                 As Long
Private m_lCheckState()                         As Long
Private m_lStrctPtr                             As Long
Private m_lPtr()                                As Long
Private m_lSortArray()                          As Long
Private m_lBackColor                            As Long
Private m_lRowColor()                           As Long
Private m_lRowColorBase                         As Long
Private m_lRowColorOffset                       As Long
Private m_lRowDepth                             As Long
Private m_lTmpForeClr                           As Long
Private m_sngLuminence                          As Single
Private m_sngChkLuminence                       As Single
Private m_sSortArray()                          As String
Private m_eListMode                             As ECGListMode
Private m_eHeaderSkinStyle                      As ECGHeaderSkinStyle
Private m_eThemeLuminence                       As ECGThemeLuminence
Private m_ECGScrollBarSkinStyle                 As ECGScrollBarSkinStyle
Private m_eAppearance                           As ECGAppearance
Private m_eBorderStyle                          As ECGBorderStyle
Private m_eViewMode                             As ECGStyle
Private m_eSortTag                              As ECGColumnSortTags
Private m_eCheckBoxSkinStyle                    As ECGCheckBoxSkinStyle
Private m_eRowDecoration                        As ECGRowDecoration
Private m_oHdrBkClr                             As OLE_COLOR
Private m_oHdrForeClr                           As OLE_COLOR
Private m_oHdrHighLiteClr                       As OLE_COLOR
Private m_oHdrPressedClr                        As OLE_COLOR
Private m_oThemeColor                           As OLE_COLOR
Private c_ColumnTags                            As Collection
Private c_PtrMem                                As Collection
Private WithEvents m_oFont                      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_pISelectorBar                         As StdPicture
Private IChecked                                As StdPicture
Private m_cSelectorBar                          As clsStoreDc
Private m_cChkCheckDc                           As clsStoreDc
Private m_cSkinHeader                           As clsSkinHeader
Private m_cSkinScrollBars                       As clsSkinScrollbars
Private m_HLIStc()                              As HLIStc
Private m_cListItems()                          As clsListItem
Private m_cRender                               As clsRender
Private m_cHListSubclass                        As GXMSubclass


'/~ 2.1 todo ~
'/~ bugfix                          - done
'/~ custom mode                     - done
'/~ hyper mode                      - done
'/~ database mode                   - done
'/~ edit controls                   - > defer to hgrid
'/~ save to file                    - done
'/~ skinned checkbox                - done
'/~ bg image                        - done
'/~ alpha select bar                - done
'/~ disabled icons                  - done
'/~ label edit                      - done
'/~ check all                       - done
'/~ extended sort                   - done
'/~ skin theming                    - done
'/~ row fonts                       - done
'/~ skin headers                    - done
'/~ skin scrollbars                 - done
'/~ var cleanup                     - done
'/~ operand alignment               - done
'/~ create uc version               -
'/~ example project                 -

'**********************************************************************
'*                              OPERATIONAL
'**********************************************************************

Private Sub Class_Initialize()
'/* init library

'Dim i As Long
    
    '/* var alignment test
    'i = VarPtr(m_bCheckBoxes) Mod 8
    'If i Then
    '    Debug.Print "Misaligned by " & i
    'End If

    m_lhMod = LoadLibrary("shell32.dll")
    InitComctl32
    Set m_cHListSubclass = New GXMSubclass
    m_oHdrBkClr = GetSysColor(m_oHdrBkClr And &H1F&)
    m_oHdrForeClr = GetSysColor(vbWindowText And &H1F&)
    m_lTmpBg = -1
    m_bEnabled = True
    m_lBackColor = &HFFFFFF
    m_lTmpForeClr = &H0
    m_oThemeColor = -1
    m_eCheckBoxSkinStyle = CheckBoxXP
    Set c_ColumnTags = New Collection

End Sub

Public Function CreateList(ByRef Style As ECGStyle, _
                           ByVal lWidth As Long, _
                           ByVal lHeight As Long, _
                           ByVal lCntHnd As Long, _
                           ByVal lAppInst As Long) As Boolean

'*/ initialize the listview

Dim lLVStyle    As Long
Dim lExStyle    As Long

    On Error GoTo Handler
    
    '/* destroy existing
    DestroyList
    '/* initial style flags including LVS_OWNERDATA
    '/* this tells the list that all data will be
    '/* managed externally
    lLVStyle = WS_CHILD Or WS_BORDER Or WS_VISIBLE Or Style Or LVS_SORTASCENDING Or LVS_OWNERDATA Or _
               LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_SINGLESEL Or WS_TABSTOP Or Style
    '/* default container style
    lExStyle = GetWindowLong(lCntHnd, GWL_EXSTYLE) And Not WS_EX_CLIENTEDGE
    '/* create the list
    m_lLVHwnd = CreateWindowEx(lExStyle, WC_LISTVIEW, vbNullString, lLVStyle, _
                0&, 0&, lWidth, lHeight, lCntHnd, 0&, lAppInst, ByVal 0&)
    '/* set the style flag
    ViewMode = Style
    '/* default border style
    SetWindowStyle lCntHnd, GWL_STYLE, 0, WS_BORDER
    '/* subclass the list and parent WM_NOTIFY messages
    '/* control callback data is reflected from parent control
    If Not m_lLVHwnd = 0 Then
#If Unicode Then
        SendMessageLong m_lLVHwnd, CCM_SETUNICODEFORMAT, 1&, 0&
        Debug.Assert SendMessageLong(m_lLVHwnd, CCM_GETUNICODEFORMAT, 0&, 0&) <> 0
#End If
        m_lParentHwnd = lCntHnd
        ListAttatch m_lLVHwnd
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("CreateList", Err.Number)

End Function

Private Function InitComctl32() As Boolean
'/* init comctl32 listview class

Dim icc As tagINITCOMMONCONTROLSEX

On Error GoTo Handler
  
  icc.dwSize = Len(icc)
  icc.dwICC = ICC_LISTVIEW_CLASSES
  InitComctl32 = InitCommonControlsEx(icc)

On Error GoTo 0
Exit Function

Handler:
  InitCommonControls

End Function

Public Sub InitList(ByVal lCount As Long, _
                    Optional ByVal lSubItemCt As Long = -1)

'/* init item arrays

Dim lCt As Long

    Select Case m_eListMode
    '/* init class array
    Case eCustomDraw
        ReDim m_cListItems(0 To lCount) As clsListItem
        For lCt = 0 To lCount
            Set m_cListItems(lCt) = New clsListItem
            If Not lSubItemCt = -1 Then
                m_cListItems(lCt).SubItemCount = lSubItemCt
            m_cListItems(lCt).Init
            End If
        Next lCt
    '/* init struct arrays
    Case eHyperList
        ReDim m_HLIStc(0).item(0 To lCount)
        For lCt = 0 To lCount
            If Not lSubItemCt = -1 Then
                ReDim m_HLIStc(0).SubItem(lCt).Text(1 To lSubItemCt)
            End If
        Next lCt
    End Select

End Sub

Public Property Get ListMode() As ECGListMode
'*/ retrieve pointer to the data structure
    ListMode = m_eListMode
End Property

Public Property Let ListMode(ByVal PropVal As ECGListMode)
'*/ add pointer to the data structure
    m_eListMode = PropVal
End Property

Private Function PointerToString(ByVal lpString As Long) As String
'/* get string from pointer

#If Unicode Then
    Dim nLen As Long
    If lpString Then
        nLen = lstrlen(ByVal lpString)
        If nLen Then
            '/* allocate string with nLen chars
            PointerToString = String$(nLen, 0)
            lstrcpyW ByVal StrPtr(PointerToString), ByVal lpString
        End If
    End If
#Else
    Dim nLen As Long
    Dim b()  As Byte
    If lpString Then
        nLen = lstrlen(ByVal lpString)
        If nLen Then
            '/* allocate buffer with nLen bytes
            ReDim b(0 To nLen - 1) As Byte
            '/* copy nLen bytes for ANSI
            CopyMemory b(0), ByVal lpString, nLen
            PointerToString = StrConv(b(), vbUnicode)
        End If
    End If
#End If

End Function

Private Sub StringToPointer(ByVal sText As String, _
                            ByRef lpString As Long)

'/* get pointer from string

#If Unicode Then
    '/* Direct copy address of an Unicode string into a pointer on NT system
    lstrcpyW ByVal lpString, ByVal StrPtr(sText)
#Else
    Dim b()  As Byte
    b = StrConv(sText, vbFromUnicode)
    CopyMemory ByVal lpString, b(0), UBound(b) + 1
#End If

End Sub

Public Property Get StructPtr() As Long
'*/ retrieve pointer to the data structure

    StructPtr = m_lStrctPtr

End Property

Public Property Let StructPtr(ByVal PropVal As Long)
'*/ add pointer to the data structure

    If Not m_lStrctPtr = 0 Then
        DeAllocatePointer "a", True
    End If
    m_lStrctPtr = PropVal

End Property

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

'/* translate to ole color

    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function


'**********************************************************************
'*                              SKINNING
'**********************************************************************

Public Property Get AlphaBarTheme() As Boolean
'/* [get] use theme color
    AlphaBarTheme = m_bAlphaBarTheme
End Property

Public Property Let AlphaBarTheme(ByVal PropVal As Boolean)
'/* [let] use theme color
    m_bAlphaBarTheme = PropVal
End Property

Public Property Get AlphaBarTransparency() As Byte
'/* [get] alpha bar transparency index
    AlphaBarTransparency = m_bteAlphaTransparency
End Property

Public Property Let AlphaBarTransparency(PropVal As Byte)
'/* [let] alpha bar transparency index

    If PropVal < 70 Then
        m_bteAlphaTransparency = 70
    ElseIf PropVal > 240 Then
        m_bteAlphaTransparency = 200
    Else
        m_bteAlphaTransparency = PropVal
    End If
    
End Property

Public Property Get AlphaThemeBackClr() As Boolean
'/* [get] alpha bar themed backcolor
    AlphaThemeBackClr = m_bAlphaThemeBackClr
End Property

Public Property Let AlphaThemeBackClr(ByVal PropVal As Boolean)
'/* [let] alpha bar themed backcolor
    m_bAlphaThemeBackClr = PropVal
End Property

Public Property Get AlphaBarActive() As Boolean
'/* [get] alpha bar state
    AlphaBarActive = m_bAlphaSelectorBar
End Property

Public Property Let AlphaBarActive(ByVal PropVal As Boolean)
'/* [let] alpha bar state
    m_bAlphaSelectorBar = PropVal
End Property

Public Function AlphaSelectorBar(ByVal btTransparency As Byte, _
                                 ByVal bUseThemeColor As Boolean, _
                                 ByVal bUseAlphaBackcolor As Boolean) As Boolean
'/* create alpha selector

On Error GoTo Handler

    m_bAlphaSelectorBar = False
    AlphaBarTransparency = btTransparency
    m_bAlphaBarTheme = bUseThemeColor
    m_bAlphaThemeBackClr = bUseAlphaBackcolor
    
    '/* reset
    If Not m_cSelectorBar Is Nothing Then
        ResetSelectorBar
    End If
    '/* render class
    If m_cRender Is Nothing Then
        Set m_cRender = New clsRender
    End If
    
    Set ISelectorBar = LoadResPicture("SELECTORBAR", vbResBitmap)
    Set m_cSelectorBar = New clsStoreDc
    With m_cSelectorBar
        .UseAlpha = True
        .CreateFromPicture ISelectorBar
        If m_bAlphaBarTheme Then
            .ColorizeImage m_oThemeColor, m_sngLuminence
        End If
    End With
    m_bAlphaSelectorBar = True
    AlphaSelectorBar = True

On Error GoTo 0
Exit Function

Handler:

End Function

Private Function ResetSelectorBar() As Boolean

    m_bAlphaSelectorBar = False
    If Not m_cSelectorBar Is Nothing Then
        Set m_cSelectorBar = Nothing
    End If
    Set ISelectorBar = Nothing
    
End Function

Private Sub DrawAlphaSelectorBar(ByVal lItem As Long)
'/* paint vert scrollbar

Dim lDrawDc     As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long
Dim lSrcDc      As Long
Dim tTmp        As RECT
Dim tRect       As RECT

On Error GoTo Handler

    lItem = lItem + 1
    GetItemRect lItem, tRect
    lSrcDc = GetDC(m_lLVHwnd)
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lDrawDc = CreateCompatibleDC(lTmpDc)
    
    LSet tTmp = tRect
    With tTmp
        OffsetRect tTmp, -.left, -.top
        lBmp = CreateCompatibleBitmap(lTmpDc, .right, .bottom)
    End With
    lBmpOld = SelectObject(lDrawDc, lBmp)
    
    With tTmp
        If m_bCheckBoxes And ((m_eViewMode = StyleReport) Or (m_eViewMode = StyleList)) Then
            tRect.left = tRect.left + 18
            .right = .right - 18
        End If
        '/* left
        m_cRender.Stretch lDrawDc, 0, .top, 3, .bottom, m_cSelectorBar.hdc, 0, 0, 3, m_cSelectorBar.Height, SRCCOPY
        '/* center
        m_cRender.Stretch lDrawDc, 3, 0, .right - 6, .bottom, m_cSelectorBar.hdc, 3, 0, m_cSelectorBar.Width - 6, m_cSelectorBar.Height, SRCCOPY
        '/* right
        m_cRender.Stretch lDrawDc, .right - 3, 0, 3, .bottom, m_cSelectorBar.hdc, m_cSelectorBar.Width - 3, 0, 3, m_cSelectorBar.Height, SRCCOPY
        '/* copy to dest
        m_cRender.AlphaBlit lSrcDc, tRect.left, tRect.top, .right, .bottom, lDrawDc, 0, 0, .right, .bottom, m_bteAlphaTransparency
    End With
    
    '/* cleanup
    ReleaseDC lHwnd, lTmpDc
    ReleaseDC m_lLVHwnd, lSrcDc
    SelectObject lDrawDc, lBmpOld
    DeleteObject lBmp
    DeleteDC lDrawDc

Handler:
    On Error GoTo 0

End Sub

Private Property Get ISelectorBar() As StdPicture
'/* selector bar image
    Set ISelectorBar = m_pISelectorBar
End Property

Private Property Set ISelectorBar(PropVal As StdPicture)
    Set m_pISelectorBar = PropVal
End Property

Public Function SkinCheckBox(ByVal eCheckBoxStyle As ECGCheckBoxSkinStyle, _
                             ByVal bUseThemeColors As Boolean) As Boolean
'/* skin the listview checkboxes
'/~ SendMessageLong m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_STATE, m_lImlStateHndl <- does not work, why?

On Error GoTo Handler

    m_eCheckBoxSkinStyle = eCheckBoxStyle
    m_bUseCheckBoxTheme = bUseThemeColors
    '/* system image sizes
    CheckBoxMetrics
    
    '/* render class
    If m_cRender Is Nothing Then
        Set m_cRender = New clsRender
    End If
    '/* image dc
    If Not m_cChkCheckDc Is Nothing Then
        ResetSkinnedCheckboxes
    Else
        Set m_cChkCheckDc = New clsStoreDc
    End If
    
    '/* load images
    If LoadCheckBoxImages Then
        '/* create dc
        With m_cChkCheckDc
            .CreateFromPicture IChecked
            If m_bUseCheckBoxTheme Then
                .ColorizeImage m_oThemeColor, m_sngLuminence
            End If
        End With
    End If
    
    '/* success
    m_bSkinnedCheck = True
    SkinCheckBox = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SkinCheckBox", Err.Number)

End Function

Public Property Get CheckBoxSkinStyle() As ECGCheckBoxSkinStyle
'*/ return the checkbox style
    CheckBoxSkinStyle = m_eCheckBoxSkinStyle
End Property

Public Property Let CheckBoxSkinStyle(PropVal As ECGCheckBoxSkinStyle)
'*/ change the checkbox style

    If Not m_eCheckBoxSkinStyle = PropVal Then
        If ResetSkinnedCheckboxes Then
            SkinCheckBox PropVal, m_bUseCheckBoxTheme
            ListRefresh
        End If
    End If
    m_eCheckBoxSkinStyle = PropVal
    
End Property

Private Function ResetSkinnedCheckboxes() As Boolean

On Error GoTo Handler

    m_bSkinnedCheck = False
    If Not m_cChkCheckDc Is Nothing Then
        Set m_cChkCheckDc = Nothing
    End If
    Set m_cChkCheckDc = New clsStoreDc
    Set IChecked = Nothing
    '/* success
    ResetSkinnedCheckboxes = True

On Error GoTo 0
Exit Function

Handler:
    
End Function

Private Function LoadCheckBoxImages() As Boolean
'/* load checkbox skin images

On Error GoTo Handler

    Select Case m_eCheckBoxSkinStyle
    '/* classic
    Case 0
        Set IChecked = LoadResPicture("CLASSIC-CHECKBOX", vbResBitmap)
    '/* eclipse
    Case 1
        Set IChecked = LoadResPicture("ECLIPSE-CHECKBOX", vbResBitmap)
    '/* lime
    Case 2
        Set IChecked = LoadResPicture("LIME-CHECKBOX", vbResBitmap)
    '/* metallic
    Case 3
        Set IChecked = LoadResPicture("METALLIC-CHECKBOX", vbResBitmap)
    '/* Gloss
    Case 4
        Set IChecked = LoadResPicture("GLOSS-CHECKBOX", vbResBitmap)
    '/* xp
    Case 5
        Set IChecked = LoadResPicture("XP-CHECKBOX", vbResBitmap)
    End Select
    
    '/* theme settings
    If m_bUseCheckBoxTheme Then
        CheckBoxThemeSettings
    End If
    
    '/* success
    LoadCheckBoxImages = True

On Error GoTo 0
Exit Function

Handler:

End Function

Private Function CheckBoxThemeSettings() As Boolean
'/* checkbox skin theme luminence

    Select Case m_eThemeLuminence
    Case 0
        m_sngChkLuminence = 0.2
    Case 1
        m_sngChkLuminence = 0.4
    Case 2
        m_sngChkLuminence = 1
    End Select
    
End Function

Private Sub CheckBoxMetrics()
'/* checkbox system metrics

    m_lCheckWidth = GetSystemMetrics(SM_CXSMICON)
    m_lCheckHeight = GetSystemMetrics(SM_CYSMICON)
    If (m_lCheckWidth = 0) Or (m_lCheckHeight = 0) Then
        m_lCheckWidth = 16
        m_lCheckHeight = 16
    End If

End Sub

Public Function UnSkinCheckBox() As Boolean
'/* unskin checkbox

On Error GoTo Handler

    ResetSkinnedCheckboxes
    UnSkinCheckBox = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("UnSkinCheckBox", Err.Number)

End Function

Public Function SkinHeaders(ByVal eSkinStyle As ECGHeaderSkinStyle, _
                            ByVal oFontForecolor As OLE_COLOR, _
                            ByVal oFontHighliteColor As OLE_COLOR, _
                            ByVal oFontPressedColor As OLE_COLOR, _
                            ByVal bUseThemeColors As Boolean) As Boolean

'/* use header skin

    m_eHeaderSkinStyle = eSkinStyle
    m_oHdrForeClr = oFontForecolor
    m_oHdrHighLiteClr = oFontHighliteColor
    m_oHdrPressedClr = oFontPressedColor
    m_bUseThemeColors = bUseThemeColors
    If Not m_cSkinHeader Is Nothing Then
        m_cSkinHeader.ResetHeaderSkin
    Else
        Set m_cSkinHeader = New clsSkinHeader
    End If
    With m_cSkinHeader
        .HeaderForeColor = m_oHdrForeClr
        .HeaderHighLite = m_oHdrHighLiteClr
        .HeaderPressed = m_oHdrPressedClr
        .HeaderIml = m_lImlSmallHndl
        .HeaderLuminence = m_eThemeLuminence
        .HeaderSkinStyle = m_eHeaderSkinStyle
        .HeaderThemeColor = m_oThemeColor
        .SetFont m_oFont
        .UseHeaderTheme = m_bUseThemeColors
        .LoadSkin HeaderHwnd
        m_bSkinHeader = True
    End With
    
End Function

Public Function UnSkinHeaders()
'/* reset skinned header

    Set m_cSkinHeader = Nothing
    ListRefresh

End Function

Public Function SkinScrollBars(ByVal eSkinStyle As ECGScrollBarSkinStyle, _
                               ByVal bUseThemeColors As Boolean)
'/* skin scrollbars

    If Not m_cSkinScrollBars Is Nothing Then
        m_cSkinScrollBars.ResetScrollBarSkin
    Else
        Set m_cSkinScrollBars = New clsSkinScrollbars
    End If
    
    m_ECGScrollBarSkinStyle = eSkinStyle
    bUseThemeColors = m_bUseThemeColors
    With m_cSkinScrollBars
        .ScrollBarSkinStyle = m_ECGScrollBarSkinStyle
        .ScrollLuminence = m_eThemeLuminence
        .ScrollThemeColor = m_oThemeColor
        .SkinScrollBar = True
        .UseScrollBarTheme = m_bUseThemeColors
        .LoadSkin m_lLVHwnd, m_lParentHwnd
    End With
    ListRefresh
    Resize
    
End Function

Public Function UnSkinScrollBars()
'/* unskin scrollbars

    Set m_cSkinScrollBars = Nothing
    ListRefresh
    
End Function

Public Property Get ThemeColor() As OLE_COLOR
'/* get theme color
    ThemeColor = m_oThemeColor
End Property

Public Property Let ThemeColor(ByVal PropVal As OLE_COLOR)
'/* set theme color

    m_oThemeColor = PropVal
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.ScrollThemeColor = PropVal
    End If
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderThemeColor = PropVal
    End If
    
End Property

Public Property Get UseThemeColors() As Boolean
'/* get theme status
    UseThemeColors = m_bUseThemeColors
End Property

Public Property Let UseThemeColors(ByVal PropVal As Boolean)
'/* set theme option

    m_bUseThemeColors = PropVal
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.UseScrollBarTheme = PropVal
    End If
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.UseHeaderTheme = PropVal
    End If
    m_bUseCheckBoxTheme = True
    
End Property

Public Property Get ThemeLuminence() As ECGThemeLuminence
'/* get theme luminence
    ThemeLuminence = m_eThemeLuminence
End Property

Public Property Let ThemeLuminence(ByVal PropVal As ECGThemeLuminence)
'/* set theme luminence

    m_eThemeLuminence = PropVal

    Select Case PropVal
    Case 0
        m_sngLuminence = 0.3
    Case 1
        m_sngLuminence = 0.6
    Case 2
        m_sngLuminence = 1
    End Select
    
End Property

Public Function UnSkinAll() As Boolean
'/* remove all skinning

    UnSkinCheckBox
    UnSkinHeaders
    UnSkinScrollBars
    ListRefresh
    
End Function


'**********************************************************************
'*                              LOAD/SAVE
'**********************************************************************

Public Function SaveToFile(ByVal sPath As String) As Boolean
'/* save items list to file

Dim FF      As Integer
Dim lLen    As Long
Dim sTemp   As String
Dim sHeader As String


On Error GoTo Handler

    If Count < 1 Then Exit Function
    '/* invalid file name
    If Not InStr(1, sPath, Chr(46)) > 0 Then
        Exit Function
    End If
    If FileExists(sPath) Then
        DeleteFile sPath
    Else
        '/* test path
        sTemp = left$(sPath, InStrRev(sPath, Chr(92)))
        If Not FileExists(sTemp) Then
            Exit Function
        End If
    End If
    
    '/* test validity -hl mode only
    If Not m_eListMode = eHyperList Then
        GoTo Handler
    ElseIf Not ArrayCheck(m_HLIStc(0).item) Then
        GoTo Handler
    ElseIf Count = 0 Then
        GoTo Handler
    ElseIf Count >= 100000 Then
        GoTo Handler
    End If
    '-> header
    '/* id
    '/* item count
    '-> data
    FF = FreeFile
    Open sPath For Binary Access Write Lock Read As #FF
    '/* 4 byte id string
    Put #FF, , "HL20"
    '/* item count
    Put #FF, , Count
    '/* write data
    Put #FF, , m_HLIStc(0)
    Close #FF
    
    '/* success
    SaveToFile = True

Exit Function

Handler:
    Close #FF
    On Error GoTo 0

End Function

Public Function LoadFromFile(ByVal sPath As String) As Boolean
'/* load items list from file

Dim FF      As Integer
Dim lRows   As Long
Dim lCols   As Long
Dim lCnt    As Long
Dim sTemp   As String

On Error GoTo Handler
    
    '/* invalid file name
    If Not InStr(1, sPath, Chr(46)) > 0 Then
        Exit Function
    End If
    '/* test path
    sTemp = left$(sPath, InStrRev(sPath, Chr(92)))
    If Not FileExists(sTemp) Then
        Exit Function
    End If
    
    '/* clear list data
    ClearList
    '/* switch to hl mode
    m_eListMode = eHyperList
    
    FF = FreeFile
    Open sPath For Binary Access Read Lock Write As #FF
    '/* app id
    sTemp = Space$(4)
    Get #FF, , sTemp
    If Not sTemp = "HL20" Then
        GoTo Handler
    End If

    '/* row count
    Get #FF, , lRows
    If lRows > 0 Then
        '/* init struct
        ReDim m_HLIStc(0) As HLIStc
        '/* read data
        Get #FF, , m_HLIStc(0)
        '/* set count
        SetItemCount lRows
    End If
    Close #FF
    
    '/* success
    LoadFromFile = True

Exit Function

Handler:
    On Error Resume Next
    Close #FF
    On Error GoTo 0
    
End Function

Public Function CopyItemToClipboard()
'/* copy selected item to clipboard

Dim lItem As Long
Dim sTemp As String

    lItem = ItemFocused
    sTemp = ItemText(lItem)
    Clipboard.Clear
    Clipboard.SetText sTemp

End Function

Private Function FileExists(ByVal sPath As String) As Boolean
'/* test file or path

Dim lRes    As Long
Dim sTemp   As String

    sTemp = String$(254, 0)
    lRes = GetShortPathName(sPath, sTemp, 255)
    FileExists = lRes > 0

End Function


'**********************************************************************
'*                              COLUMNS
'**********************************************************************

Public Function ColumnAdd(ByVal lIndex As Long, _
                          ByVal sText As String, _
                          ByVal lWidth As Long, _
                          Optional ByVal eAlign As ECGColumnAlign = [ColumnLeft], _
                          Optional ByVal lIcon As Long = -1, _
                          Optional ByVal ColumnTag As ECGColumnSortTags = SortDefault) As Boolean

'*/ create column headers

Dim bFirst  As Boolean
Dim uLVC    As LVCOLUMN
Dim uHDI    As HDITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    bFirst = (Me.ColumnCount = 0)

    With uLVC
#If Unicode Then
        .pszText = StrPtr(sText)
        .cchTextMax = Len(sText)
#Else
        .pszText = StrPtr(StrConv(sText & vbNullChar, vbFromUnicode))
        .cchTextMax = Len(sText) + 1
#End If
        .cx = lWidth
        .fmt = eAlign
        .Mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_FMT
    End With

    ColumnAdd = (SendMessage(m_lLVHwnd, LVM_INSERTCOLUMN, lIndex, uLVC) > -1)

    If ColumnAdd Then
        If bFirst Then
            m_lHdrHwnd = HeaderHwnd()
            SendMessageLong m_lHdrHwnd, HDM_SETIMAGELIST, 0&, m_lImlHdHndl
        End If
        With uHDI
#If Unicode Then

            .pszText = StrPtr(sText)
            .cchTextMax = Len(sText)

#Else
            .pszText = StrPtr(StrConv(sText + vbNullChar, vbFromUnicode))
            .cchTextMax = Len(sText) + 1
#End If

            .cxy = lWidth
            .iImage = lIcon
            .fmt = HDF_STRING Or eAlign * -(lIndex <> 0) Or HDF_IMAGE * -(lIcon > -1) Or HDF_BITMAP_ON_RIGHT
            .Mask = HDI_TEXT Or HDI_WIDTH Or HDI_IMAGE Or HDI_FORMAT
        End With
        SendMessage m_lHdrHwnd, HDM_SETITEM, lIndex, uHDI
    End If
    
    '/* sort flag
    c_ColumnTags.Add ColumnTag, CStr(lIndex)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnAdd", Err.Number)

End Function

Public Property Get ColumnAlign(ByVal lColumn As Long) As ECGColumnAlign
'*/ retieve a columns text alignment

Const lMask     As Long = &H3
Dim uLVC        As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    uLVC.Mask = LVCF_FMT
    SendMessage m_lLVHwnd, LVM_GETCOLUMN, lColumn, uLVC
    ColumnAlign = (lMask And uLVC.fmt)

End Property

Public Property Let ColumnAlign(ByVal lColumn As Long, _
                                  ByVal eAlign As ECGColumnAlign)
'*/ change a columns text alignment

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uLVC
        .fmt = eAlign * -(Not lColumn = 0)
        .Mask = LVCF_FMT
    End With
    SendMessage m_lLVHwnd, LVM_SETCOLUMN, lColumn, uLVC

End Property

Public Function ColumnAutosize(ByVal lColumn As Long, _
                               Optional ByVal AutosizeType As ECGColumnAutosize = [ColumnItem]) As Boolean
'*/ autosize columns

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ColumnAutosize = CBool(SendMessageLong(m_lLVHwnd, LVM_SETCOLUMNWIDTH, lColumn, AutosizeType))

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnAutosize", Err.Number)

End Function

Public Function ColumnClear() As Boolean
'*/ remove all columns

Dim lCt As Long

    For lCt = ColumnCount To 0 Step -1
        ColumnRemove lCt
    Next lCt

End Function

Public Property Get ColumnCount() As Long
'*/ retieve column count

    If m_lLVHwnd = 0 Then Exit Property
    ColumnCount = SendMessageLong(HeaderHwnd(), HDM_GETITEMCOUNT, 0&, 0&)

End Property

Public Property Get ColumnHeight() As Long

'*/ retrieve a columns height

Dim tHdr            As RECT

On Error GoTo Handler

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    '/* get coordinates
    GetClientRect m_lHdrHwnd, tHdr
    ColumnHeight = tHdr.bottom

Handler:
    On Error GoTo 0
    
End Property

Public Property Get ColumnIcon(ByVal lColumn As Long) As Long
'*/ retieve header icon index

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Then Exit Property
    uLVC.Mask = LVCF_IMAGE
    SendMessage m_lLVHwnd, LVM_GETCOLUMN, lColumn, uLVC
    ColumnIcon = uLVC.iImage

End Property

Public Property Let ColumnIcon(ByVal lColumn As Long, _
                               ByVal lIcon As Long)
'*/ change header icon

Const lMask     As Long = &H3
Dim lAlign      As Long
Dim uHDI        As HDITEM

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uHDI
        .Mask = HDI_FORMAT
        SendMessage m_lHdrHwnd, HDM_GETITEM, lColumn, uHDI
        lAlign = lMask And .fmt
        .iImage = lIcon
        .fmt = HDF_STRING Or lAlign Or HDF_IMAGE * -(lIcon > -1 And m_lImlHdHndl <> 0) Or HDF_BITMAP_ON_RIGHT
        .Mask = HDI_IMAGE * -(lIcon > -1) Or HDI_FORMAT
    End With
    SendMessage m_lHdrHwnd, HDM_SETITEM, lColumn, uHDI

End Property

Public Function ColumnLastFit() As Boolean

Dim lCol    As Long
Dim lParam  As Long

    If m_lLVHwnd = 0 Then Exit Function
    lCol = (ColumnCount - 1)
    SendMessageLong m_lLVHwnd, LVM_SETCOLUMNWIDTH, lCol, LVSCW_AUTOSIZE_USEHEADER

End Function

Public Function ColumnRemove(ByVal lColumn As Long) As Boolean
'*/ remove a column

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ColumnRemove = CBool(SendMessageLong(m_lLVHwnd, LVM_DELETECOLUMN, lColumn, 0&))
    If Me.ColumnCount = 0 Then m_lHdrHwnd = 0

On Error GoTo 0

Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnRemove", Err.Number)

End Function

Public Sub ColumnReorder(ByVal bRemCheckbox As Boolean)
'*/ reorder columns to accomodate checkbox

Dim aWidth()    As Long
Dim lCt         As Long
Dim lUct        As Long
Dim aText()     As String

On Error Resume Next

    lUct = ColumnCount
    ReDim aText(lUct)
    ReDim aWidth(lUct)

    For lCt = 0 To lUct
        aText(lCt) = Trim$(ColumnText(lCt))
        aWidth(lCt) = ColumnWidth(lCt)
    Next lCt

    For lCt = lUct To 0 Step -1
        ColumnRemove lCt
    Next lCt

    If bRemCheckbox Then
        For lCt = 0 To lUct
            ColumnAdd lCt + 1, aText(lCt), aWidth(lCt)
        Next lCt
    Else
        For lCt = 1 To lUct
            ColumnAdd lCt - 1, aText(lCt - 1), aWidth(lCt - 1)
        Next lCt
    End If

On Error GoTo 0

End Sub

Public Function ColumnSizeToItems(Optional bColumnFit As Boolean) As Boolean

Dim lCol    As Long
Dim lParam  As Long

    If m_lLVHwnd = 0 Then Exit Function
    If bColumnFit Then
        lParam = LVSCW_AUTOSIZE_USEHEADER
    Else
        lParam = LVSCW_AUTOSIZE
    End If

    For lCol = 0 To ColumnCount - 1
        SendMessageLong m_lLVHwnd, LVM_SETCOLUMNWIDTH, lCol, lParam
    Next
    
End Function

Public Property Get ColumnTag(ByVal lColumn As Long) As ECGColumnSortTags
'/* get column sort tag

Dim lRet As Long

On Error GoTo Handler

    lRet = c_ColumnTags.item(CStr(lColumn))
    If (lRet < 0) Or lRet > 3 Then
        GoTo Handler
    End If
    ColumnTag = lRet

On Error GoTo 0
Exit Property

Handler:
    ColumnTag = -1
    
End Property

Public Property Get ColumnText(ByVal lColumn As Long) As String
'*/ get a columns heading

Dim uLVC            As LVCOLUMN
Dim aText(261)      As Byte

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uLVC
        .pszText = VarPtr(aText(0))
        .cchTextMax = UBound(aText) + 1
        .Mask = LVCF_TEXT
    End With
    SendMessage m_lLVHwnd, LVM_GETCOLUMN, lColumn, uLVC
    ColumnText = PointerToString(uLVC.pszText)

End Property

Public Property Let ColumnText(ByVal lColumn As Long, _
                                 ByVal sText As String)
'*/ change a columns heading

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uLVC
#If Unicode Then
        .pszText = StrPtr(sText)
        .cchTextMax = Len(sText)
#Else
        .pszText = StrPtr(StrConv(sText & vbNullChar, vbFromUnicode))
        .cchTextMax = Len(sText) + 1
#End If

        .Mask = LVCF_TEXT
    End With
    SendMessage m_lLVHwnd, LVM_SETCOLUMN, lColumn, uLVC

End Property

Public Property Get ColumnWidth(ByVal lColumn As Long) As Long
'*/ retrieve a columns length

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    ColumnWidth = SendMessageLong(m_lLVHwnd, LVM_GETCOLUMNWIDTH, lColumn, 0&)

End Property

Public Property Let ColumnWidth(ByVal lColumn As Long, _
                                  ByVal lWidth As Long)
'*/ change a columns length

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    SendMessageLong m_lLVHwnd, LVM_SETCOLUMNWIDTH, lColumn, lWidth

End Property

Private Function HeaderHwnd() As Long
'*/ return the column header handle

    If m_lLVHwnd = 0 Then Exit Function
    HeaderHwnd = SendMessageLong(m_lLVHwnd, LVM_GETHEADER, 0&, 0&)

End Function

Public Property Get HeaderColor() As OLE_COLOR
'*/ return the header color
    HeaderColor = m_oHdrBkClr
End Property

Public Property Let HeaderColor(PropVal As OLE_COLOR)
'*/ change the header color

    m_oHdrBkClr = PropVal
    If m_bXPColors Then
        m_oHdrBkClr = XPShift(m_oHdrBkClr)
    End If
    If Not m_cSkinHeader Is Nothing Then
        m_cSkinHeader.HeaderThemeColor = PropVal
    End If
    
End Property

Public Property Get HeaderCustom() As Boolean
'*/ return the custom header status
    HeaderCustom = m_bCustomHeader
End Property

Public Property Let HeaderCustom(PropVal As Boolean)
'*/ change the custom header status

    m_bCustomHeader = PropVal
    If Not m_cSkinHeader Is Nothing Then
        m_cSkinHeader.CustomHeader = PropVal
    End If
    
End Property

Public Property Get HeaderDragDrop() As Boolean
'*/ retrieve drag and drop state
    HeaderDragDrop = m_bDragDrop
End Property

Public Property Let HeaderDragDrop(ByVal PropVal As Boolean)
'*/ retrieve drag and drop state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bDragDrop = PropVal
    If m_bDragDrop Then
        SetExtendedStyle LVS_EX_HEADERDRAGDROP, 0
    Else
        SetExtendedStyle 0, LVS_EX_HEADERDRAGDROP
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("HeaderDragDrop", Err.Number)

End Property

Public Property Get HeaderFixedWidth() As Boolean
'*/ retrieve fixed width state
    HeaderFixedWidth = m_bHeaderFixed
End Property

Public Property Let HeaderFixedWidth(ByVal PropVal As Boolean)
'*/ change fixed width state
    m_bHeaderFixed = PropVal
End Property

Public Property Get HeaderFlat() As Boolean
'*/ change width state
    HeaderFlat = m_bHeaderFlat
End Property

Public Property Let HeaderFlat(ByVal PropVal As Boolean)
'*/ change width state

Dim lStyle      As Long
Dim lHwnd       As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bHeaderFlat = PropVal
    lHwnd = HeaderHwnd()
    If lHwnd = 0 Then Exit Property
    lStyle = GetWindowLong(lHwnd, GWL_STYLE)
    If m_bHeaderFlat Then
        lStyle = lStyle And Not HDS_BUTTONS
    Else
        lStyle = lStyle Or HDS_BUTTONS
    End If
    SetWindowLong lHwnd, GWL_STYLE, lStyle

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("HeaderFlat", Err.Number)

End Property

Public Property Get HeaderForeColor() As OLE_COLOR
'*/ return the header forecolor
    HeaderForeColor = m_oHdrForeClr
End Property

Public Property Let HeaderForeColor(PropVal As OLE_COLOR)
'*/ change the header forecolor

    m_oHdrForeClr = PropVal
    If Not m_cSkinHeader Is Nothing Then
        m_cSkinHeader.HeaderForeColor = PropVal
    End If
    
End Property

Public Property Get HeaderHide() As Boolean
'*/ retrieve header visible state
    HeaderHide = m_bHeaderHide
End Property

Public Property Let HeaderHide(ByVal PropVal As Boolean)
'*/ change header visible state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bHeaderHide = PropVal
    If m_bHeaderHide Then
        SetStyle LVS_NOCOLUMNHEADER, 0
    Else
        SetStyle 0, LVS_NOCOLUMNHEADER
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("HeaderHide", Err.Number)

End Property

Public Property Get HeaderHighLite() As OLE_COLOR
'*/ return header highlite color
    HeaderHighLite = m_oHdrHighLiteClr
End Property

Public Property Let HeaderHighLite(PropVal As OLE_COLOR)
'*/ change header highlite color

    m_oHdrHighLiteClr = PropVal
    If Not m_cSkinHeader Is Nothing Then
        m_cSkinHeader.HeaderHighLite = PropVal
    End If
    
End Property


'**********************************************************************
'*                              IMAGELIST
'**********************************************************************

Public Function InitImlHeader() As Boolean
'*/ initialize header imagelist

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlHeader
    m_lImlHdHndl = ImageList_Create(16, 16, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    InitImlHeader = (Not m_lImlHdHndl = 0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("InitImlHeader", Err.Number)

End Function

Public Function ImlHeaderAddBmp(ByVal lBitmap As Long, _
                                Optional ByVal lMaskColor As Long = CLR_NONE) As Long
'*/ add a bitmap to header iml

On Error GoTo Handler

    If m_lImlHdHndl = 0 Then Exit Function
    If Not lMaskColor = CLR_NONE Then
        ImlHeaderAddBmp = ImageList_AddMasked(m_lImlHdHndl, lBitmap, lMaskColor)
    Else
        ImlHeaderAddBmp = ImageList_Add(m_lImlHdHndl, lBitmap, 0&)
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlHeaderAddBmp", Err.Number)

End Function

Public Function ImlHeaderAddIcon(ByVal lIcon As Long) As Long
'*/ add an icon to header iml

On Error GoTo Handler

    If m_lImlHdHndl = 0 Then Exit Function
    ImlHeaderAddIcon = ImageList_AddIcon(m_lImlHdHndl, lIcon)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlHeaderAddIcon", Err.Number)

End Function

Private Function DestroyImlHeader() As Boolean
'*/ destroy header image list

On Error GoTo Handler

    If m_lImlHdHndl = 0 Then Exit Function
    If ImageList_Destroy(m_lImlHdHndl) Then
        DestroyImlHeader = True
        m_lImlHdHndl = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyImlHeader", Err.Number)

End Function

Public Function InitImlLarge(Optional ByVal lImgWidth As Long = 32, _
                             Optional ByVal lImgHeight As Long = 32) As Boolean
'*/ initialize large icons image list

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlLarge
    m_lImlLargeHndl = ImageList_Create(lImgWidth, lImgHeight, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    SendMessageLong m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_NORMAL, m_lImlLargeHndl
    InitImlLarge = (Not m_lImlLargeHndl = 0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlLargeAddIcon", Err.Number)

End Function

Public Function ImlLargeAddBmp(ByVal lBitmap As Long, _
                               Optional ByVal lMaskColor As Long = CLR_NONE) As Long
'*/ add bmp to large image iml

On Error GoTo Handler

    If m_lImlLargeHndl = 0 Then Exit Function
    If Not lMaskColor = CLR_NONE Then
        ImlLargeAddBmp = ImageList_AddMasked(m_lImlLargeHndl, lBitmap, lMaskColor)
    Else
        ImlLargeAddBmp = ImageList_Add(m_lImlLargeHndl, lBitmap, 0&)
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlLargeAddBmp", Err.Number)

End Function

Public Function ImlLargeAddIcon(ByVal lIcon As Long) As Long
'*/ add icon to large image iml

On Error GoTo Handler

    If m_lImlLargeHndl = 0 Then Exit Function
    ImlLargeAddIcon = ImageList_AddIcon(m_lImlLargeHndl, lIcon)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlLargeAddIcon", Err.Number)

End Function

Private Function DestroyImlLarge() As Boolean
'*/ destroy large icons image list

On Error GoTo Handler

    If m_lImlLargeHndl = 0 Then Exit Function
    If ImageList_Destroy(m_lImlLargeHndl) Then
        DestroyImlLarge = True
        m_lImlLargeHndl = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyImlLarge", Err.Number)

End Function

Public Function InitImlSmall(Optional lWidth As Long = 16, _
                             Optional lHeight As Long = 16) As Boolean

'*/ initialize smallicons image list

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlSmall
    m_lImlSmallHndl = ImageList_Create(lWidth, lHeight, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    SendMessageLong m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_SMALL, m_lImlSmallHndl
    InitImlSmall = (Not m_lImlSmallHndl = 0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("InitImlSmall", Err.Number)

End Function

Public Function ImlSmallAddBmp(ByVal lBitmap As Long, _
                               Optional ByVal lMaskColor As Long = CLR_NONE) As Long
'*/ add bmp to small image iml

On Error GoTo Handler

    If m_lImlSmallHndl = 0 Then Exit Function
    If Not lMaskColor = CLR_NONE Then
        ImlSmallAddBmp = ImageList_AddMasked(m_lImlSmallHndl, lBitmap, lMaskColor)
    Else
        ImlSmallAddBmp = ImageList_Add(m_lImlSmallHndl, lBitmap, 0&)
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlSmallAddBmp", Err.Number)

End Function

Public Function ImlSmallAddIcon(ByVal lIcon As Long) As Long
'*/ add icon to small image iml

On Error GoTo Handler

    If m_lImlSmallHndl = 0 Then Exit Function
    ImlSmallAddIcon = ImageList_AddIcon(m_lImlSmallHndl, lIcon)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlSmallAddIcon", Err.Number)

End Function

Private Function DestroyImlSmall() As Boolean
'*/ destroy small icons image list

On Error GoTo Handler

    If m_lImlSmallHndl = 0 Then Exit Function
    If ImageList_Destroy(m_lImlSmallHndl) Then
        DestroyImlSmall = True
        m_lImlSmallHndl = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyImlSmall", Err.Number)

End Function

Public Function InitImlState() As Boolean
'*/ initialize header imagelist


On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlState
    m_lImlStateHndl = ImageList_Create(16, 16, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    InitImlState = (Not m_lImlStateHndl = 0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("InitImlState", Err.Number)

End Function

Public Function ImlStateAddIcon(ByVal lIcon As Long) As Long
'*/ add an icon to header iml

On Error GoTo Handler

    If m_lImlStateHndl = 0 Then Exit Function
    ImlStateAddIcon = ImageList_AddIcon(m_lImlStateHndl, lIcon)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlStateAddIcon", Err.Number)

End Function

Public Function ImlStateAddBmp(ByVal lBitmap As Long, _
                               Optional ByVal lMaskColor As Long = CLR_NONE) As Long
'*/ add a bitmap to header iml

On Error GoTo Handler

    If m_lImlStateHndl = 0 Then Exit Function
    If Not lMaskColor = CLR_NONE Then
        ImlStateAddBmp = ImageList_AddMasked(m_lImlStateHndl, lBitmap, lMaskColor)
    Else
        ImlStateAddBmp = ImageList_Add(m_lImlStateHndl, lBitmap, 0&)
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlStateAddBmp", Err.Number)

End Function

Private Function DestroyImlState() As Boolean
'*/ destroy header image list

On Error GoTo Handler

    If m_lImlStateHndl = 0 Then Exit Function
    If ImageList_Destroy(m_lImlStateHndl) Then
        DestroyImlState = True
        m_lImlStateHndl = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyImlState", Err.Number)

End Function


'**********************************************************************
'*                               LISTITEMS
'**********************************************************************

Public Sub ClearList()
'/* clear all items

    DestroyItems
    DeAllocatePointer "a", True
    SetItemCount 0
    
End Sub

Public Property Get Count() As Long
'*/ [get] item count

    If m_lLVHwnd Then
        Count = SendMessageLong(m_lLVHwnd, LVM_GETITEMCOUNT, 0&, 0&)
    End If

End Property

Public Function ItemAdd(ByVal lIndex As Long, _
                        ByVal sKey As String, _
                        ByVal sText As String, _
                        ByVal lIcon As Long, _
                        ByVal lSmallIcon As Long) As Boolean

'/* add an item

On Error GoTo Handler

    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        lIndex = lIndex - 1
        If lIndex > (Count - 1) Then
            '/* add to class array
            ReDim Preserve m_cListItems(LBound(m_cListItems) To lIndex)
            Set m_cListItems(lIndex) = New clsListItem
        ElseIf m_cListItems(lIndex) Is Nothing Then
            Set m_cListItems(lIndex) = New clsListItem
        End If
        m_cListItems(lIndex).Add (lIndex), sKey, sText, lIcon, lSmallIcon
    '/* hl mode
    Case eHyperList
        lIndex = lIndex - 1
        '/* redim struct item arrays
        If lIndex > (Count - 1) Then
            ReDim Preserve m_HLIStc(0).item(LBound(m_HLIStc(0).item) To lIndex)
            m_HLIStc(0).item(lIndex) = sText
            m_HLIStc(0).lIcon(lIndex) = lIcon
        End If
    End Select
    '/* set new count
    SetItemCount Count + 1
    '/* success
    ItemAdd = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ItemAdd", Err.Number)
    
End Function

Public Property Get ItemChecked(ByVal lItem As Long) As Boolean
'*/ return item checked state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    ItemChecked = ((SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_STATEIMAGEMASK) And &H2000&) = &H2000&)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemChecked", Err.Number)

End Property

Public Property Let ItemChecked(ByVal lItem As Long, _
                                ByVal bChecked As Boolean)
'*/ set item checked state

Dim uLVI As LV_ITEM

On Error GoTo Handler

    '/* toggle in checkbox array
    If m_eListMode = eCustomDraw Or eHyperList Then
        CheckToggle lItem
    End If
    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .stateMask = LVIS_STATEIMAGEMASK
        .State = &H1000& * (1 - bChecked)
        .Mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemChecked", Err.Number)

End Property

Public Function ItemEnsureVisible(ByVal lItem As Long) As Boolean
'*/ move to item index

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ItemEnsureVisible = CBool(SendMessageLong(m_lLVHwnd, LVM_ENSUREVISIBLE, lItem, 0&))

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ItemEnsureVisible", Err.Number)

End Function

Public Property Get ItemFocused(ByVal lItem As Long) As Boolean
'*/ return item focused state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    ItemFocused = CBool(SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_FOCUSED))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemFocused", Err.Number)

End Property

Public Property Let ItemFocused(ByVal lItem As Long, _
                                ByVal bFocused As Boolean)
'*/ change item focused state

Dim uLVI As LV_ITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .stateMask = LVIS_FOCUSED
        .State = -bFocused * LVIS_FOCUSED
        .Mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemFocused", Err.Number)

End Property

Public Property Get ItemGhosted(ByVal lItem As Long) As Boolean
'*/ return item ghosted state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    ItemGhosted = (SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_CUT))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemGhosted", Err.Number)

End Property

Public Property Let ItemGhosted(ByVal lItem As Long, _
                                ByVal bGhosted As Boolean)
'*/ change item ghosted state

Dim uLVI As LV_ITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .stateMask = LVIS_CUT
        .State = LVIS_CUT * -bGhosted
        .Mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemGhosted", Err.Number)

End Property

Public Property Get ItemIcon(ByVal lItem As Long) As Long
'*/ return icon index

Dim uLVI As LV_ITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .iItem = lItem
        .Mask = LVIF_IMAGE
    End With
    SendMessage m_lLVHwnd, LVM_GETITEM, 0&, uLVI
    ItemIcon = uLVI.iImage

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIcon", Err.Number)

End Property

Public Property Let ItemIcon(ByVal lItem As Long, _
                             ByVal lIcon As Long)
'*/ change icon index

Dim uLVI As LV_ITEM

On Error GoTo Handler

    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).Icon = lIcon
    Case eHyperList '***
        m_HLIStc(0).lIcon(lItem) = lIcon
    End Select
    
    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .iItem = lItem
        .iImage = lIcon
        .Mask = LVIS_OVERLAYMASK
        .State = lItem * 256
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, 0&, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIcon", Err.Number)

End Property

Public Property Get ItemIndent() As Long
'*/ return item indent

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    ItemIndent = m_lItemIndent

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIndent", Err.Number)

End Property

Public Property Let ItemIndent(ByVal PropVal As Long)
'*/ change item indent

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_lItemIndent = PropVal

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIndent", Err.Number)

End Property

Public Function ItemRemove(ByVal lItem As Long) As Boolean
'*/ remove an item from the list

Dim lLb         As Long
Dim lUb         As Long
Dim lC          As Long
Dim lP          As Long
Dim lS          As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        If ArrayExists(m_cListItems) Then
            '/* remove item
            Set m_cListItems(lItem) = Nothing
            '/* reset array
            CDResizeArray m_cListItems, lItem
            '/* init list
            SetItemCount Count - 1
        Else
            SetItemCount 0
        End If
    '/* cd mode
    Case eHyperList
        If ArrayExists(m_HLIStc(0).item) Then
            '/* reste array
            HLResizeArray lItem
            '/* init list
            SetItemCount Count - 1
        Else
            SetItemCount 0
        End If
    End Select
    '/* success
    ItemRemove = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ItemRemove", Err.Number)

End Function

Public Property Get ItemSelected(ByVal lItem As Long) As Boolean
'*/ return selected state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    ItemSelected = CBool(SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_SELECTED))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemSelected", Err.Number)

End Property

Public Property Let ItemSelected(ByVal lItem As Long, _
                                 ByVal bSelected As Boolean)
'*/ select an item

Dim uLVI    As LV_ITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .stateMask = LVIS_SELECTED Or -(bSelected And lItem > -1) * LVIS_FOCUSED
        .State = -bSelected * LVIS_SELECTED Or -(lItem > -1) * LVIS_FOCUSED
        .Mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemSelected", Err.Number)

End Property

Public Property Get ItemsSorted() As Boolean
'*/ return sorted mode status
    ItemsSorted = m_bUseSorted
End Property

Public Property Let ItemsSorted(PropVal As Boolean)
'*/ change sorted mode status
    m_bUseSorted = PropVal
End Property

Public Property Get ItemText(ByVal lItem As Long) As String
'*/ return item text

Dim uLVI   As LV_ITEM
Dim a(261) As Byte
Dim lLen   As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .pszText = VarPtr(a(0))
        .cchTextMax = UBound(a)
    End With

    lLen = SendMessage(m_lLVHwnd, LVM_GETITEMTEXT, lItem, uLVI)
    ItemText = PointerToString(uLVI.pszText)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemText", Err.Number)

End Property

Public Property Let ItemText(ByVal lItem As Long, _
                             ByVal sText As String)
'*/ change item text

Dim uLVI As LV_ITEM

On Error GoTo Handler

    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).Text = sText
    Case eHyperList '***
        m_HLIStc(0).item(lItem) = sText
    End Select
    
    If m_lLVHwnd = 0 Then Exit Property
    With uLVI

#If Unicode Then
        .pszText = StrPtr(sText)
        .cchTextMax = Len(sText)
#Else
        .pszText = StrPtr(StrConv(sText & vbNullChar, vbFromUnicode))
        .cchTextMax = Len(sText) + 1
#End If
    End With

    SendMessage m_lLVHwnd, LVM_SETITEMTEXT, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemText", Err.Number)

End Property


'**********************************************************************
'*                              SUBITEMS
'**********************************************************************


Public Function SubItemsAdd(ByVal lIndex As Long, _
                            ByVal lSubItem As Long, _
                            ByVal sText As String) As Boolean

'/* add a subitem

On Error GoTo Handler

    If lSubItem > (ColumnCount - 1) Then Exit Function
    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        If lIndex > Count Then
            Exit Function
        ElseIf m_cListItems(lIndex - 1) Is Nothing Then
            Exit Function
        End If
        lIndex = lIndex - 1
        m_cListItems(lIndex).SubItem lSubItem, sText
    '/* hl mode
    Case eHyperList
        m_HLIStc(0).SubItem(lIndex).Text(lSubItem) = sText
    End Select
    '/* init list
    SetItemCount Count
    '/* success
    SubItemsAdd = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SubItemsAdd", Err.Number)
    
End Function

Public Function SubIconIndex(ByVal lIndex As Long, _
                             ByVal lSubItem As Long, _
                             ByVal lIcon As Long) As Boolean

'/* name subitem icon index

On Error GoTo Handler

    '/* invalid
    If lSubItem > (ColumnCount - 1) Then Exit Function
    
    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        If lIndex > Count Then
            Exit Function
        ElseIf m_cListItems(lIndex - 1) Is Nothing Then
            Exit Function
        End If
        lIndex = lIndex - 1
        m_cListItems(lIndex).SubIcon lSubItem, lIcon
    '/* hl mode
    Case eHyperList
        m_HLIStc(0).SubItem(lIndex).lIcon(lSubItem) = lIcon
    End Select
    '/* init list
    SetItemCount Count
    '/* success
    SubIconIndex = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SubIconIndex", Err.Number)

End Function

Public Property Get SubItemIcon(ByVal lItem As Long, _
                                ByVal lSubItem As Long) As Long
'*/ retrieve subitem icon

Dim uLVI    As LV_ITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .iItem = lItem
        .iSubItem = lSubItem
        .Mask = LVIF_IMAGE
    End With
    SendMessage m_lLVHwnd, LVM_GETITEM, 0&, uLVI
    SubItemIcon = uLVI.iImage

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("SubItemIcon", Err.Number)

End Property

Public Property Let SubItemIcon(ByVal lItem As Long, _
                                ByVal lSubItem As Long, _
                                ByVal lIcon As Long)
'*/ change subitem icon

Dim uLVI    As LV_ITEM

On Error GoTo Handler

    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).SubItemIcon(lSubItem) = lIcon
        
    Case eHyperList '***
        m_HLIStc(0).SubItem(lItem).lIcon(lSubItem) = lIcon
    End Select
    
    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .iItem = lItem
        .iSubItem = lSubItem
        .iImage = lIcon
        .Mask = LVIF_IMAGE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEM, 0&, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("SubItemIcon", Err.Number)

End Property

Public Property Get SubItemImages() As Boolean
'*/ retrieve subitem icon state
    SubItemImages = m_bSubItemImage
End Property

Public Property Let SubItemImages(ByVal PropVal As Boolean)
'*/ change subitem icon state
    m_bSubItemImage = PropVal
End Property

Public Property Get SubItemText(ByVal lItem As Long, _
                                ByVal lSubItem As Long) As String
'*/ retieve subitem text

Dim uLVI        As LV_ITEM
Dim aText(256)  As Byte
Dim lLen        As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .iSubItem = lSubItem
        .pszText = VarPtr(aText(0))
        .cchTextMax = UBound(aText)
        .Mask = LVIF_TEXT
    End With
    lLen = SendMessage(m_lLVHwnd, LVM_GETITEMTEXT, lItem, uLVI)

    If lLen > 0 Then
#If Unicode Then
        SubItemText = aText
        SubItemText = left$(SubItemText, lLen)
#Else
        SubItemText = left$(StrConv(aText(), vbUnicode), lLen)
#End If
    Else
        SubItemText = ""
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("SubItemText", Err.Number)

End Property

Public Property Let SubItemText(ByVal lItem As Long, _
                                ByVal lSubItem As Long, _
                                ByVal sText As String)
'*/ change subitem text

Dim uLVI    As LV_ITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    
    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).SubItemText(lSubItem) = sText
        
    Case eHyperList '***
        m_HLIStc(0).SubItem(lItem).Text(lSubItem) = sText
    End Select
    
    With uLVI
        .iSubItem = lSubItem
#If Unicode Then
        .pszText = StrPtr(sText)
        .cchTextMax = Len(sText)
#Else
        .pszText = StrPtr(StrConv(sText & vbNullChar, vbFromUnicode))
        .cchTextMax = Len(sText) + 1
#End If
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMTEXT, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("SubItemText", Err.Number)

End Property

Public Property Get WordWrap() As Boolean
'*/ [get] word wrap

    If m_lLVHwnd = 0 Then Exit Property
    WordWrap = m_bWordWrap

End Property

Public Property Let WordWrap(PropVal As Boolean)
'*/ [let] word wrap

    If m_lLVHwnd = 0 Then Exit Property
    m_bWordWrap = PropVal

End Property


'**********************************************************************
'*                          LISTVIEW PROPERTIES
'**********************************************************************

Public Property Get Appearance() As ECGAppearance
'*/ retrive list borderstyle
    Appearance = m_eAppearance
End Property

Public Property Let Appearance(ByVal PropVal As ECGAppearance)
'*/ change list borderstyle

    m_eAppearance = PropVal
    SetBorderStyle m_lParentHwnd

End Property

Public Property Get AutoArrange() As Boolean
'*/ [get] auto arrange
    AutoArrange = m_bAutoArrange
End Property

Public Property Let AutoArrange(ByVal PropVal As Boolean)
'*/ [let] auto arrange

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bAutoArrange = PropVal
    If PropVal Then
        SetStyle 0, LVS_AUTOARRANGE
    Else
        SetStyle LVS_AUTOARRANGE, 0
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("AutoArrange", Err.Number)

End Property

Public Property Get BackColor() As OLE_COLOR
'*/ retrieve list backcolor

    BackColor = m_lBackColor

End Property

Public Property Let BackColor(ByVal PropVal As OLE_COLOR)
'*/ change list backcolor

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_lBackColor = PropVal
    OleTranslateColor PropVal, 0&, m_lBackColor
    If m_bXPColors Then
        m_lBackColor = XPShift(m_lBackColor)
    End If
    SendMessageLong m_lLVHwnd, LVM_SETBKCOLOR, 0&, m_lBackColor
    SendMessageLong m_lLVHwnd, LVM_SETTEXTBKCOLOR, 0&, m_lBackColor
    ListRefresh

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("BackColor", Err.Number)

End Property

Public Property Get BorderStyle() As ECGBorderStyle
'*/ retrive list borderstyle
    BorderStyle = m_eBorderStyle
End Property

Public Property Let BorderStyle(ByVal PropVal As ECGBorderStyle)
'*/ change list borderstyle

    m_eBorderStyle = PropVal
    SetBorderStyle m_lParentHwnd

End Property

Public Function CheckAll() As Boolean
'*/ mark all checkboxes

Dim lCt As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    If Not ArrayCheck(m_lCheckState) Then Exit Function
    For lCt = LBound(m_lCheckState) To UBound(m_lCheckState)
        m_lCheckState(lCt) = 1
    Next lCt
    ListRefresh
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("CheckAll", Err.Number)

End Function

Public Function UnCheckAll() As Boolean
'*/ unmark all checkboxes

Dim lCt As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    If Not ArrayCheck(m_lCheckState) Then Exit Function
    For lCt = LBound(m_lCheckState) To UBound(m_lCheckState)
        m_lCheckState(lCt) = 0
    Next lCt
    ListRefresh
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("UnCheckAll", Err.Number)

End Function

Public Property Get CheckBoxes() As Boolean
'*/ retrieve checkbox state
    CheckBoxes = m_bCheckBoxes
End Property

Public Property Let CheckBoxes(ByVal PropVal As Boolean)
'*/ change checkbox state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bCheckBoxes = PropVal
    If m_bCheckBoxes Then
        SetExtendedStyle LVS_EX_CHECKBOXES, 0
    Else
        SetExtendedStyle 0, LVS_EX_CHECKBOXES
    End If
    
    If m_bCheckBoxes Then
        If Not m_bCheckInit Then
            InitCheckBoxes Count
        End If
    End If
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("CheckBoxes", Err.Number)

End Property

Public Property Get CustomDraw() As Boolean
'/* [get] custom draw mode
    CustomDraw = m_bCustomDraw
End Property

Public Property Let CustomDraw(PropVal As Boolean)
'/* [let] custom draw mode
    m_bCustomDraw = PropVal
End Property

Private Function EditHandle() As Long

Dim lEditHnd As Long

    If m_lLVHwnd = 0 Then Exit Function
    lEditHnd = SendMessageLong(m_lLVHwnd, LVM_GETEDITCONTROL, 0&, 0&)
    If Not lEditHnd = 0 Then
        EditHandle = lEditHnd
    End If

End Function

Private Function EditLimitLength(ByVal lLength As Long) As Boolean

Dim lEditHnd As Long

    lEditHnd = EditHandle
    If Not lEditHnd = 0 Then
        SendMessageLong lEditHnd, EM_LIMITTEXT, lLength, 0&
    End If
    
End Function

Private Function EditLowerCase(ByVal lLength As Long) As Boolean

Dim lStyle As Long
Dim lEditHnd    As Long

    lEditHnd = EditHandle
    If Not lEditHnd = 0 Then
        lStyle = GetWindowLong(lEditHnd, GWL_STYLE)
        SetWindowLong lEditHnd, GWL_STYLE, lStyle Or ES_LOWERCASE
    End If
    
End Function

Private Function EditUpperCase(ByVal lLength As Long) As Boolean

Dim lStyle      As Long
Dim lEditHnd    As Long

    lEditHnd = EditHandle
    If Not lEditHnd = 0 Then
        lStyle = GetWindowLong(lEditHnd, GWL_STYLE)
        SetWindowLong lEditHnd, GWL_STYLE, lStyle Or ES_UPPERCASE
    End If
    
End Function

Public Property Get Enabled() As Boolean
'/* [get] toggle listview enable state
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal PropVal As Boolean)
'/* [let] toggle listview enable state

Static bCustom As Boolean

On Error GoTo Handler

    If m_bCustomDraw Then
        bCustom = True
    End If
    
    m_bEnabled = PropVal
    EnableWindow m_lLVHwnd, Abs(PropVal)
    If Not PropVal Then
        m_lTmpBg = m_lBackColor
        BackColor = GetSysColor(vbButtonFace And &H1F&)
        ForeColor = GetSysColor(vbButtonShadow And &H1F&)
        m_bCustomDraw = False
    Else
        If Not m_lTmpBg = -1 Then
            BackColor = m_lTmpBg
            ForeColor = m_lTmpForeClr
        Else
            BackColor = &HFFFFFF
            ForeColor = &H0
        End If
        m_bCustomDraw = bCustom
    End If
    ListRefresh
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("Enabled", Err.Number)
    
End Property

Public Property Get Font() As StdFont
'*/ retrieve list font

On Error GoTo Handler
'Attribute DrawImage.VB_Description = "Draws an Image from the image list onto a device context."

    Set Font = m_oFont

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("Font", Err.Number)

End Property

Public Property Set Font(ByVal PropVal As StdFont)
'*/ change list font

Dim uLF     As LOGFONT
Dim lChar   As Long
Dim lHDC    As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    Set m_oFont = PropVal
    With uLF
        For lChar = 1 To Len(m_oFont.Name)
            .lfFaceName(lChar - 1) = CByte(Asc(Mid$(m_oFont.Name, lChar, 1)))
        Next lChar
        .lfHeight = m_oFont.Size
        .lfItalic = m_oFont.Italic
        .lfWeight = IIf(m_oFont.Bold, FW_BOLD, FW_NORMAL)
        .lfUnderline = m_oFont.Underline
        .lfStrikeOut = m_oFont.Strikethrough
        .lfCharSet = m_oFont.Charset
    End With
    DestroyFont
    m_lFont = CreateFontIndirect(uLF)
    SendMessageLong m_lLVHwnd, WM_SETFONT, m_lFont, True

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("Font", Err.Number)

End Property

Private Function DestroyFont() As Boolean
'*/ font cleanup

On Error GoTo Handler

    If m_lFont Then
        If DeleteObject(m_lFont) Then
            DestroyFont = True
            m_lFont = 0
        End If
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyFont", Err.Number)

End Function

Public Property Get ForeColor() As OLE_COLOR
'*/ retrieve list forecolor

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    ForeColor = SendMessageLong(m_lLVHwnd, LVM_GETTEXTCOLOR, 0&, 0&)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ForeColor", Err.Number)

End Property

Public Property Let ForeColor(ByVal PropVal As OLE_COLOR)
'*/ change list forecolor

Dim lColor As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    OleTranslateColor PropVal, 0&, lColor
    SendMessageLong m_lLVHwnd, LVM_SETTEXTCOLOR, 0&, lColor
    ListRefresh

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ForeColor", Err.Number)

End Property

Public Property Get FullRowSelect() As Boolean
'*/ retrieve full row select state
    FullRowSelect = m_bFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal PropVal As Boolean)
'*/ change full row select state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bFullRowSelect = PropVal
    If m_bFullRowSelect Then
        SetExtendedStyle LVS_EX_FULLROWSELECT, 0
    Else
        SetExtendedStyle 0&, LVS_EX_FULLROWSELECT
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("FullRowSelect", Err.Number)

End Property

Private Sub GetItemRect(ByVal iIndex As Long, _
                        ByRef tRect As RECT, _
                        Optional ByVal bSubItem As Boolean)

'/* get rect struct of row item

    If m_lLVHwnd = 0 Then Exit Sub
    If Not bSubItem Then
        SendMessage m_lLVHwnd, LVM_GETITEMRECT, iIndex - 1, tRect
    Else
        SendMessage m_lLVHwnd, LVM_GETSUBITEMRECT, iIndex - 1, tRect
    End If

End Sub

Public Property Get GridLines() As Boolean
'*/ change gridlines state
    GridLines = m_bGridLines
End Property

Public Property Let GridLines(ByVal PropVal As Boolean)
'*/ change gridlines state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bGridLines = PropVal
    If m_bGridLines Then
        SetExtendedStyle LVS_EX_GRIDLINES, 0
    Else
        SetExtendedStyle 0, LVS_EX_GRIDLINES
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("GridLines", Err.Number)

End Property

Public Property Get HideSelection() As Boolean
'*/ retrieve selection visible state
    HideSelection = m_bHideSelection
End Property

Public Property Let HideSelection(ByVal PropVal As Boolean)
'*/ change selection visible state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bHideSelection = PropVal
    If m_bHideSelection Then
        SetStyle 0, LVS_SHOWSELALWAYS
    Else
        SetStyle LVS_SHOWSELALWAYS, 0
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("HideSelection", Err.Number)

End Property

Private Function IconDisabled(ByVal lItem As Long, _
                              ByVal lImgIndex As Long) As Boolean

'/* draw disabled icon

Dim lx          As Long
Dim lY          As Long
Dim lWidth      As Long
Dim lHeight     As Long
Dim lhIcon      As Long
Dim lhImgList   As Long
Dim lHDC        As Long
Dim lFlags      As Long
Dim tRect       As RECT

On Error GoTo Handler

    lHDC = GetDC(m_lLVHwnd)
    '/* get list handle
    If ViewMode = StyleIcon Then
        lhImgList = m_lImlLargeHndl
    Else
        lhImgList = m_lImlSmallHndl
    End If
    '/* icon size
    ImageList_GetIconSize lhImgList, lWidth, lHeight
    '/* get coords
    GetItemRect lItem, tRect
    lx = 2
    '/* checkbox offset
    If m_bCheckBoxes Then
        lx = lx + 16
    End If
    lx = lx + (m_lItemIndent * 4) + m_lIconSpaceX
    lY = tRect.top + ColumnHeight
    If lItem = 0 Then
        lY = lY + 1
    End If
    '/* get icon handle
    lhIcon = ImageList_GetIcon(lhImgList, lImgIndex, 0)
    '/* blend flag
    lFlags = lFlags Or ILD_SELECTED Or ILD_BLEND25
    '/* draw image
    ImageList_DrawEx lhImgList, lImgIndex, lHDC, lx, lY, 15, 15, CLR_NONE, TranslateColor(vbButtonFace), lFlags
    ReleaseDC m_lLVHwnd, lHDC
    
    '/* success
    IconDisabled = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("IconDisabled", Err.Number)
    
End Function

Public Property Get IconSpaceX() As Long
'*/ [get] change icon left align
    IconSpaceX = m_lIconSpaceX
End Property

Public Property Let IconSpaceX(ByVal PropVal As Long)
'*/ [let] change icon left align

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_lIconSpaceX = PropVal
    pSetIconSpacing

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("IconSpaceX", Err.Number)

End Property

Public Property Get IconSpaceY() As Long
'*/ [get] change icon left align
    IconSpaceY = m_lIconSpaceY
End Property

Public Property Let IconSpaceY(ByVal PropVal As Long)
'*/ [let] change icon top align

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_lIconSpaceY = PropVal
    pSetIconSpacing

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("IconSpaceY", Err.Number)

End Property

Private Sub pSetIconSpacing()
'/* set icon indent

Dim cx      As Long
Dim cy      As Long
Dim lXY     As Long
Dim lXYD    As Long
Dim lXYC    As Long

   '/* Set cx=-1, cy=-1 to reset to default and return current settings:
   lXYD = &HFFFFFFFF
   lXYC = SendMessageLong(m_lLVHwnd, LVM_SETICONSPACING, 0, lXYD)
   '/* cX is loword:
   cx = (lXYC And &HFFFF&)
   '/* cY is hiword:
   cy = (lXYC \ &H10000)
   If m_lIconSpaceX > 0 Then cx = m_lIconSpaceX
   If m_lIconSpaceY > 0 Then cy = m_lIconSpaceY
   
   lXY = cx And &H7FFF
   lXY = lXY Or ((cy And &H7FFF) * &H10000)
   SendMessageLong m_lLVHwnd, LVM_SETICONSPACING, 0, lXY
   
End Sub

Public Property Get InfoTips() As Boolean
'/* [get] infotips state
    InfoTips = m_bInfoTips
End Property

Public Property Let InfoTips(ByVal PropVal As Boolean)
'*/ [let] change info tips state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bInfoTips = PropVal
    If PropVal Then
        SetExtendedStyle LVS_EX_INFOTIP, 0
    Else
        SetExtendedStyle 0, LVS_EX_INFOTIP
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("InfoTips", Err.Number)

End Property

Public Property Get ItemBorderSelect() As Boolean
'*/ [get] border select state
    ItemBorderSelect = m_bItemBorderSelect
End Property

Public Property Let ItemBorderSelect(ByVal PropVal As Boolean)
'*/ [let] border select state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bItemBorderSelect = PropVal
    If PropVal Then
        SetExtendedStyle LVS_EX_BORDERSELECT, 0
    Else
        SetExtendedStyle 0, LVS_EX_BORDERSELECT
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemBorderSelect", Err.Number)
    
End Property

Public Function ItemsSort(ByVal lColumn As Long, _
                          ByVal bDescending As Boolean) As Boolean
'*/ sort items in the list

On Error GoTo Handler

    If bDescending Then
        SortControl lColumn, 2
    Else
        SortControl lColumn, 1
    End If
    SetItemCount Count

Handler:
    On Error GoTo 0

End Function

Public Function ItemTopIndex() As Long

Dim lIndex As Long

    If m_lLVHwnd = 0 Then Exit Function
    lIndex = SendMessageLong(m_lLVHwnd, LVM_GETTOPINDEX, 0&, 0&)
    ItemTopIndex = lIndex
    
End Function

Public Property Get LabelEdit() As Boolean
'/* [get] edit state
    LabelEdit = m_bEditLabels
End Property

Public Property Let LabelEdit(ByVal PropVal As Boolean)
'*/ [let] change edit state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bEditLabels = PropVal
    If PropVal Then
        SetStyle LVS_EDITLABELS, 0
    Else
        SetStyle 0, LVS_EDITLABELS
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("LabelEdit", Err.Number)

End Property

Public Property Get LabelTips() As Boolean
'*/ retrieve label tips state
    LabelTips = m_bLabelTips
End Property

Public Property Let LabelTips(ByVal PropVal As Boolean)
'*/ change label tips state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bLabelTips = PropVal
    If PropVal Then
        SetExtendedStyle LVS_EX_LABELTIP, 0
    Else
        SetExtendedStyle 0, LVS_EX_LABELTIP
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("LabelTips", Err.Number)

End Property

Public Property Get MultiSelect() As Boolean
'*/ retrieve multiselect state
    MultiSelect = m_bMultiSelect
End Property

Public Property Let MultiSelect(ByVal PropVal As Boolean)
'*/ change multiselect state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bMultiSelect = PropVal
    If PropVal Then
        SetStyle 0, LVS_SINGLESEL
    Else
        SetStyle LVS_SINGLESEL, 0
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("MultiSelect", Err.Number)

End Property

Public Property Get OneClickActivate() As Boolean
'/* [get] one click label edit
    OneClickActivate = m_bOneClickActivate
End Property

Public Property Let OneClickActivate(ByVal PropVal As Boolean)
'*/ [let] one click edit

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bOneClickActivate = PropVal
    If PropVal Then
        SetExtendedStyle LVS_EX_ONECLICKACTIVATE, 0
    Else
        SetExtendedStyle 0, LVS_EX_ONECLICKACTIVATE
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("OneClickActivate", Err.Number)

End Property

Public Function RemoveDuplicates() As Boolean
'*/ change viewmode state

Dim lCt         As Long
Dim lLb         As Long
Dim lUb         As Long
Dim lTr         As Long
Dim cT          As Collection
Dim cList()     As clsListItem

On Error Resume Next
    
    BuildStringSortArray 0
    Set cT = New Collection
    m_bSorted = False
    
    '/* only unique keys will be added
    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        '/* get bounds
        lLb = LBound(m_sSortArray)
        lUb = UBound(m_sSortArray)
        lCt = lLb
        '/* filter with collection key
        Do
            cT.Add 1, m_sSortArray(lCt)
            If Err.Number = 457 Then
                Set m_cListItems(lCt) = Nothing
            End If
            lCt = lCt + 1
        Loop Until lCt > lUb
        '/* reset the array
        CDResetArray m_cListItems
        '/* init list
        SetItemCount Count
    '/* hl mode
    Case eHyperList
        '/* get bounds
        lLb = LBound(m_HLIStc(0).item)
        lUb = UBound(m_HLIStc(0).item)
        lCt = lLb
        '/* filter
        Do
            cT.Add 1, m_sSortArray(lCt)
            '/* remove from array
            If Err.Number = 457 Then
                HLResizeArray lCt
            End If
            lCt = lCt + 1
        Loop Until lCt > lUb
        '/* reset list
        SetItemCount Count
    End Select
    '/* success
    RemoveDuplicates = True

On Error GoTo 0

End Function

Public Property Get ScrollBarFlat() As Boolean
'*/ retrieve scrollbar state
    ScrollBarFlat = m_bScrollFlat
End Property

Public Property Let ScrollBarFlat(ByVal PropVal As Boolean)
'*/ change scrollbar state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bScrollFlat = PropVal
    If m_bScrollFlat Then
        SetExtendedStyle LVS_EX_FLATSB, 0
    Else
        SetExtendedStyle 0, LVS_EX_FLATSB
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ScrollBarFlat", Err.Number)

End Property

Public Property Get SelectedCount() As Long
'*/ retrieve selected count

    If m_lLVHwnd = 0 Then Exit Property
    SelectedCount = SendMessageLong(m_lLVHwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)

End Property

Public Property Get TrackSelected() As Boolean
'*/ [get] track item
    TrackSelected = m_bTrackSelected
End Property

Public Property Let TrackSelected(ByVal PropVal As Boolean)
'*/ [let] track item

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bTrackSelected = PropVal
    If PropVal Then
        SetExtendedStyle LVS_EX_TRACKSELECT, 0
    Else
        SetExtendedStyle 0, LVS_EX_TRACKSELECT
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("TrackSelected", Err.Number)

End Property

Public Property Get UnderlineHot() As Boolean
'*/ [get] underline hot item
    UnderlineHot = m_bUnderlineHot
End Property

Public Property Let UnderlineHot(ByVal PropVal As Boolean)
'*/ [let] underline hot item

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bUnderlineHot = PropVal
    If PropVal Then
        SetExtendedStyle LVS_EX_UNDERLINEHOT, 0
    Else
        SetExtendedStyle 0, LVS_EX_UNDERLINEHOT
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("UnderlineHot", Err.Number)

End Property

Public Property Get ViewMode() As ECGStyle
'*/ retrieve viewmode state
    ViewMode = m_eViewMode
End Property

Public Property Let ViewMode(ByVal PropVal As ECGStyle)
'*/ change viewmode state

Static bWasSkinned As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    If m_bSkinnedCheck Then
        bWasSkinned = True
    End If
    m_eViewMode = PropVal
    If Not PropVal = StyleReport Then
        m_bSorted = False
    End If
    SetStyle m_eViewMode, (LVS_ICON Or LVS_SMALLICON Or LVS_REPORT Or LVS_LIST)
    If Not PropVal = StyleReport Then
        m_bSkinnedCheck = False
    Else
        m_bSkinnedCheck = bWasSkinned
    End If
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ViewMode", Err.Number)

End Property

Public Property Get Visible() As Boolean
'/* [get] use xp colors
    Visible = m_bVisible
End Property

Public Property Let Visible(PropVal As Boolean)
'/* [let] use xp colors

    m_bVisible = PropVal
    If PropVal Then
        ShowWindow m_lParentHwnd, SW_NORMAL
    Else
        ShowWindow m_lParentHwnd, SW_HIDE
    End If
    
End Property

Private Function VisibleItemCount() As Long

Dim lCt As Long

    If m_lLVHwnd = 0 Then Exit Function
    lCt = SendMessageLong(m_lLVHwnd, LVM_GETCOUNTPERPAGE, 0&, 0&)
    If Not lCt = 0 Then
        VisibleItemCount = lCt
    End If

End Function

Public Property Get XPColors() As Boolean
'/* [get] use xp colors
    XPColors = m_bXPColors
End Property

Public Property Let XPColors(PropVal As Boolean)
'/* [let] use xp colors
    m_bXPColors = PropVal
End Property


'**********************************************************************
'*                              SUPPORT
'**********************************************************************

Public Function BackgroundPicture(ByVal sPath As String, _
                                  ByVal eStyle As ECGBackGroundImage, _
                                  Optional ByVal bCenter As Boolean) As Boolean

'/* set a bg image

Dim lReturn As Long
Dim uLBI    As LVBKIMAGE

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    With uLBI
        .pszImage = sPath & vbNullChar
        .cchImageMax = Len(sPath) + 1
        '/* clear image
        If eStyle = BgNone Then
            .ulFlags = LVBKIF_SOURCE_NONE
            .xOffsetPercent = 3
            .yOffsetPercent = 3
        Else
            .ulFlags = LVBKIF_SOURCE_URL Or eStyle
            '/* center image
            If bCenter And eStyle = BgNormal Then
                .xOffsetPercent = 50
                .yOffsetPercent = 50
            End If
        End If
    End With
    
    lReturn = SendMessage(m_lLVHwnd, (LVM_FIRST + 68), 0, uLBI)
    '/* set font transparent
    If lReturn Then
        SendMessageLong m_lLVHwnd, LVM_SETTEXTBKCOLOR, 0, CLR_NONE
        BackgroundPicture = True
    End If
    '/* success
    BackgroundPicture = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("BackgroundPicture", Err.Number)

End Function

Private Function CheckToggle(ByVal lItem As Long) As Boolean
'/* toggle check state

On Error GoTo Handler

    If pItemChecked(lItem) Then
        m_lCheckState(lItem) = 0
        CheckToggle = 0
    Else
        m_lCheckState(lItem) = 1
        CheckToggle = 1
    End If

Handler:
    On Error GoTo 0

End Function

Public Property Get Checked(ByVal lIndex As Long) As Boolean
'/* get checkbox state
    Checked = (m_lCheckState(lIndex) = 1)
End Property

Public Property Let Checked(ByVal lIndex As Long, _
                            ByVal bChecked As Boolean)
'/* let checkbox state

    If bChecked Then
        m_lCheckState(lIndex) = 1
    Else
        m_lCheckState(lIndex) = 0
    End If

End Property

Private Function InitCheckBoxes(ByVal lCount As Long)
'/* load checkbox state array

    ReDim m_lCheckState(lCount)
    m_bCheckInit = True

End Function

Public Sub ListRefresh()
'/* refresh the listview

    InvalidateRect m_lLVHwnd, ByVal 0&, 0&
    UpdateWindow m_lLVHwnd

End Sub

Private Sub Resize()
'/* resize listview

Dim tRect As RECT

On Error Resume Next

    If m_lLVHwnd = 0 Then Exit Sub
    If m_lParentHwnd = 0 Then Exit Sub
    GetClientRect m_lParentHwnd, tRect
    With tRect
        OffsetRect tRect, -.left, -.top
        SetWindowPos m_lLVHwnd, 0, .left, .top, .right, .bottom, SWP_NOZORDER Or SWP_NOOWNERZORDER
        If Not m_cSkinScrollBars Is Nothing Then
            m_cSkinScrollBars.Resize Count
        End If
    End With

On Error GoTo 0

End Sub

Public Function SetBorderStyle(ByVal lHwnd As Long) As Boolean
'*/ change list borderstyle

On Error GoTo Handler

    If m_eAppearance = Flat Then
        '/* Flat border
        SetWindowStyle lHwnd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
        If Not (m_eBorderStyle = None) Then
            SetWindowStyle lHwnd, GWL_STYLE, WS_BORDER, 0
        Else
            SetWindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER
        End If
    Else
        '/* 3d border
        SetWindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER
        If m_eBorderStyle = [Fixed Single] Then
            SetWindowStyle lHwnd, GWL_EXSTYLE, 0, WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
        Else
            If m_eBorderStyle = Thin Then
                SetWindowStyle lHwnd, GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE
            Else
                SetWindowStyle lHwnd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
            End If
        End If
    End If
    '/* success
    SetBorderStyle = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SetBorderStyle", Err.Number)

End Function

Private Sub SetWindowStyle(ByVal lHwnd As Long, _
                           ByVal lType As Long, _
                           ByVal lStyle As Long, _
                           ByVal lStyleNot As Long)
'*/ change list borderstyle

Dim lNewStyle   As Long

On Error GoTo Handler

    lNewStyle = GetWindowLong(lHwnd, lType)
    lNewStyle = (lNewStyle And Not lStyleNot) Or lStyle
    SetWindowLong lHwnd, lType, lNewStyle

    SetWindowPos lHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetWindowStyle", Err.Number)

End Sub

Private Sub SetStyle(ByVal lStyle As Long, _
                     ByVal lStyleNot As Long)
'*/ change list style params

Dim lNewStyle   As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Sub
    lNewStyle = GetWindowLong(m_lLVHwnd, GWL_STYLE)
    lNewStyle = lNewStyle And Not lStyleNot
    lNewStyle = lNewStyle Or lStyle
    SetWindowLong m_lLVHwnd, GWL_STYLE, lNewStyle

    SetWindowPos m_lLVHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetStyle", Err.Number)

End Sub

Private Sub SetExtendedStyle(ByVal lStyle As Long, _
                             ByVal lStyleNot As Long)
'*/ change list extended style params

Dim lNewStyle   As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Sub
    lNewStyle = SendMessageLong(m_lLVHwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    lNewStyle = lNewStyle And Not lStyleNot
    lNewStyle = lNewStyle Or lStyle
    SendMessageLong m_lLVHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, lNewStyle

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetExtendedStyle", Err.Number)

End Sub

Public Function SetItemCount(ByVal lItems As Long) As Boolean
'/* dimension list to item count

    SendMessage m_lLVHwnd, LVM_SETITEMCOUNT, lItems, LVSICF_NOINVALIDATEALL
    m_lItemsCnt = lItems
    
    If m_bCheckBoxes Then
        If m_bCheckInit Then
            '/* test array
            If ArrayCheck(m_lCheckState) Then
                If Not (UBound(m_lCheckState) = (lItems)) Then
                    InitCheckBoxes lItems
                End If
            End If
        Else
            If Not lItems = -1 Then
                InitCheckBoxes lItems
            End If
        End If
    End If
    '/* success
    SetItemCount = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SetItemCount", Err.Number)

End Function

Private Function pItemChecked(lIndex As Long) As Boolean
'/* determoine check state
On Error GoTo Handler

    pItemChecked = m_lCheckState(lIndex) > 0

Handler:
    On Error GoTo 0

End Function

Private Function XPShift(ByVal lColor As Long, _
                         Optional ByVal Base As Long = &HB0) As Long

'/* xp color shift

Dim lRed        As Long
Dim lBlue       As Long
Dim lGreen      As Long
Dim lDelta      As Long

    lBlue = ((lColor \ &H10000) Mod &H100)
    lGreen = ((lColor \ &H100) Mod &H100)
    lRed = (lColor And &HFF)
    lDelta = &HFF - Base

    lBlue = Base + lBlue * lDelta \ &HFF
    lGreen = Base + lGreen * lDelta \ &HFF
    lRed = Base + lRed * lDelta \ &HFF

    If lRed > 255 Then lRed = 255
    If lGreen > 255 Then lGreen = 255
    If lBlue > 255 Then lBlue = 255

    XPShift = lRed + 256& * lGreen + 65536 * lBlue

End Function



'**********************************************************************
'*                              STORAGE
'**********************************************************************

Private Function ArrayCheck(ByRef vArray As Variant) As Boolean
'/* validity test

On Error Resume Next

    '/* an array
    If Not IsArray(vArray) Then
        GoTo Handler
    '/* not dimensioned
    ElseIf IsError(UBound(vArray)) Then
        GoTo Handler
    '/* no members
    ElseIf UBound(vArray) = -1 Then
        GoTo Handler
    End If
    ArrayCheck = True

Handler:
    On Error GoTo 0

End Function

Private Function ArrayExists(ByRef vArray As Variant) As Boolean
'/* valid array

On Error Resume Next

    If IsError(UBound(vArray)) Then
        GoTo Handler
    End If
    '/* success
    ArrayExists = True

Handler:
    On Error GoTo 0

End Function

Private Sub BuildNumericSortArray(ByVal lColumn As Long)
'/* create a copy of sort items

Dim lUb     As Long
Dim lLb     As Long
Dim lCt     As Long

On Error GoTo Handler

    Erase m_sSortArray
    Erase m_lSortArray
    
    Select Case m_eListMode
    Case eCustomDraw
        lLb = LBound(m_cListItems)
        lUb = UBound(m_cListItems)
        ReDim m_lSortArray(lLb To lUb)
        If lColumn = 0 Then
            Do
                If m_eSortTag = SortDate Then
                    m_lSortArray(lCt) = GetTime(m_cListItems(lCt).Text)
                ElseIf IsNumeric(m_cListItems(lCt).Text) Then
                    m_lSortArray(lCt) = m_cListItems(lCt).Text
                Else
                    m_lSortArray(lCt) = 0
                End If
                lCt = lCt + 1
            Loop Until lCt > lUb
        Else
            Do
                If m_eSortTag = SortDate Then
                    m_lSortArray(lCt) = GetTime(m_cListItems(lCt).SubItemText(lColumn))
                ElseIf IsNumeric(m_cListItems(lCt).SubItemText(lColumn)) Then
                    m_lSortArray(lCt) = CLng(m_cListItems(lCt).SubItemText(lColumn))
                Else
                    m_lSortArray(lCt) = 0
                End If
                lCt = lCt + 1
            Loop Until lCt > lUb
        End If

    Case eHyperList
        lLb = LBound(m_HLIStc(0).item)
        lUb = UBound(m_HLIStc(0).item)
        ReDim m_sSortArray(lLb To lUb)
        If lColumn = 0 Then
            Do
                If m_eSortTag = SortDate Then
                    m_lSortArray(lCt) = GetTime(m_HLIStc(0).item(lCt))
                ElseIf IsNumeric(m_HLIStc(0).item(lCt)) Then
                    m_lSortArray(lCt) = CLng(m_HLIStc(0).item(lCt))
                Else
                    m_lSortArray(lCt) = 0
                End If
                lCt = lCt + 1
            Loop Until lCt > lUb
        Else
            Do
                If m_eSortTag = SortDate Then
                    m_lSortArray(lCt) = GetTime(m_HLIStc(0).SubItem(lCt).Text(lColumn))
                ElseIf IsNumeric(m_HLIStc(0).SubItem(lCt).Text(lColumn)) Then
                    m_lSortArray(lCt) = CLng(m_HLIStc(0).SubItem(lCt).Text(lColumn))
                Else
                    m_lSortArray(lCt) = 0
                End If
                lCt = lCt + 1
            Loop Until lCt > lUb
        End If
    End Select

Handler:
    On Error GoTo 0
    
End Sub

Private Sub BuildStringSortArray(ByVal lColumn As Long)
'/* create a copy of sort items

Dim lUb     As Long
Dim lLb     As Long
Dim lCt     As Long

On Error GoTo Handler

    Erase m_sSortArray
    Erase m_lSortArray
    
    Select Case m_eListMode
    Case eCustomDraw
        lLb = LBound(m_cListItems)
        lUb = UBound(m_cListItems)
        ReDim m_sSortArray(lLb To lUb)
        If lColumn = 0 Then
            Do
                m_sSortArray(lCt) = m_cListItems(lCt).Text
                lCt = lCt + 1
            Loop Until lCt > lUb
        Else
            Do
                m_sSortArray(lCt) = m_cListItems(lCt).SubItemText(lColumn)
                lCt = lCt + 1
            Loop Until lCt > lUb
        End If

    Case eHyperList
        lLb = LBound(m_HLIStc(0).item)
        lUb = UBound(m_HLIStc(0).item)
        ReDim m_sSortArray(lLb To lUb)
        If lColumn = 0 Then
            Do
                m_sSortArray(lCt) = m_HLIStc(0).item(lCt)
                lCt = lCt + 1
            Loop Until lCt > lUb
        Else
            Do
                m_sSortArray(lCt) = m_HLIStc(0).SubItem(lCt).Text(lColumn)
                lCt = lCt + 1
            Loop Until lCt > lUb
        End If
    End Select

Handler:
    On Error GoTo 0
    
End Sub

Private Sub CDResetArray(ByRef cArray() As clsListItem)
'/* reset array with new dimensions

Dim lCt     As Long
Dim lLb     As Long
Dim lUb     As Long
Dim lVl     As Long

    If Not IsArray(cArray) Then Exit Sub
    lLb = LBound(cArray)
    lUb = UBound(cArray)

    If (lUb = -1) Or (lUb - lLb = 0) Then
        Erase cArray
        Exit Sub
    End If
    
    lVl = 0
    For lCt = lLb To lUb
        If Not cArray(lCt) Is Nothing Then
            Set cArray(lVl) = cArray(lCt)
            lVl = lVl + 1
        End If
    Next lCt
    ReDim Preserve cArray(lVl)
   
End Sub

Private Sub CDResizeArray(ByRef cArray() As clsListItem, _
                          ByVal lPos As Long)

'/* redimension array

Dim lCt     As Long
Dim lLb     As Long
Dim lUb     As Long
   
    If Not IsArray(cArray) Then Exit Sub
    lLb = LBound(cArray)
    lUb = UBound(cArray)

    If (lUb = -1) Or (lUb - lLb = 0) Then
        Erase cArray
        Exit Sub
    End If

    '/* if invalid Pos
    If (lPos > lUb) Or (lPos = -1) Then
        lPos = lUb
    ElseIf lPos < lLb Then
        lPos = lLb
    ElseIf lPos = lUb Then
        ReDim Preserve cArray(lUb - 1)
        Exit Sub
    End If
    
    Set cArray(lPos) = Nothing
    For lCt = lPos + 1 To lUb
        Set cArray(lCt - 1) = cArray(lCt)
    Next lCt
    ReDim Preserve cArray(lUb - 1)
   
End Sub

Private Function GetTime(vDate As Variant) As Long

Dim lRet As Long

On Error GoTo Handler

    If IsDate(vDate) Then
        lRet = Format(vDate, "General Number")
    Else
        lRet = 0
    End If
    GetTime = lRet

On Error GoTo 0
Exit Function

Handler:
    GetTime = 0
    On Error GoTo 0
    
End Function

Private Sub HLResizeArray(ByVal lPos As Long)
'/* resize array

    With m_HLIStc(0)
        ResizeArray .item, lPos
        ResizeArray .lIcon, lPos
        ResizeStruct lPos
    End With
    
End Sub

Public Function LoadArray() As Boolean
'*/ load data structure

Dim lSpr    As Long

On Error GoTo Handler

    Set c_PtrMem = New Collection

    Select Case m_eListMode
    Case eCustomDraw
        '/* initialize local struct
        ReDim m_cListItems(0)
        '/* copy the structure from the pointer
        CopyMemory ByVal VarPtrArray(m_cListItems), m_lStrctPtr, 4&
        c_PtrMem.Add m_lStrctPtr, "m_cListItems"
        
    Case eHyperList
        '/* initialize local struct
        ReDim m_HLIStc(0)
        '/* copy the structure from the pointer
        CopyMemory ByVal VarPtrArray(m_HLIStc), m_lStrctPtr, 4&
        c_PtrMem.Add m_lStrctPtr, "m_HLIStc"
    End Select
    LoadArray = True

Handler:
    On Error GoTo 0

End Function

Private Sub QSIInitPtr(ByVal lLb As Long, _
                       ByVal lUb As Long, _
                       ByRef aPtr() As Long)

'/* initialize the pointer array

Dim lC As Long

    Erase aPtr
    ReDim aPtr(lLb To lUb)
    lC = lLb

    Do
        aPtr(lC) = lC
        lC = lC + 1
    Loop Until lC > lUb

End Sub

Private Sub QSINumericSort(ByRef lA() As Long, _
                           ByRef lIdxA() As Long, _
                           ByVal bAsc As Boolean)

'/* based on the awesome indexed sort by Rde (Rohan) w/ mods

Dim lo          As Long
Dim hi          As Long
Dim cnt         As Long
Dim lpStr       As Long
Dim idxItem     As Long
Dim lpS         As Long
Dim lbA         As Long
Dim ubA         As Long
Dim lItem       As Long

    lbA = LBound(lA)
    ubA = UBound(lA)
    '/* pre execution check
    If Not UBound(lA) > 0 Then Exit Sub
    '/* Allow for worst case senario + some
    hi = ((ubA - lbA) \ m8) + m32
    '/* Stack to hold pending lower boundries
    ReDim lbs(m1 To hi) As Long
    '/* Stack to hold pending upper boundries
    ReDim ubs(m1 To hi) As Long
    '/* Cache pointer to the string variable
    lpStr = VarPtr(lItem)
    '/* Cache pointer to the string array
    lpS = VarPtr(lA(lbA)) - (lbA * m4)
                                                                                           
    '/* Get pivot index position
    Do: hi = ((ubA - lbA) \ m2) + lbA
        '/* Grab current value into item
        CopyMemBv lpStr, lpS + (lIdxA(hi) * m4), m4
        '/* Grab current index
        idxItem = lIdxA(hi): lIdxA(hi) = lIdxA(ubA)
        '/* Set bounds
        lo = lbA: hi = ubA
        '/* Storm right in
        Do
            If (lItem > lA(lIdxA(lo))) = bAsc Then
                lIdxA(hi) = lIdxA(lo)
                hi = hi - m1
                Do Until hi = lo
                    If (lA(lIdxA(hi)) > lItem) = bAsc Then
                        lIdxA(lo) = lIdxA(hi)
                        Exit Do
                    End If
                    hi = hi - m1
                Loop
                '/* Found swaps or out of loop
                If hi = lo Then Exit Do
            End If
            lo = lo + m1
        Loop While hi > lo
        '/* Re-assign current
        lIdxA(hi) = idxItem
        If (lbA < lo - m1) Then
            If (ubA > lo + m1) Then cnt = cnt + m1: lbs(cnt) = lo + m1: ubs(cnt) = ubA
            ubA = lo - m1
        ElseIf (ubA > lo + m1) Then
            lbA = lo + m1
        Else
            If cnt = m0 Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - m1
        End If
    Loop: CopyMemBr ByVal lpStr, 0&, m4
    
End Sub

Private Sub QSIStringSort(ByRef sA() As String, _
                          ByRef lIdxA() As Long, _
                          ByVal lCp As Long, _
                          ByVal lDr As Long)

'/* based on the awesome indexed sort by Rde (Rohan) w/ mods

Dim lo          As Long
Dim hi          As Long
Dim cnt         As Long
Dim lpStr       As Long
Dim idxItem     As Long
Dim lpS         As Long
Dim lbA         As Long
Dim ubA         As Long
Dim item        As String

    lbA = LBound(sA)
    ubA = UBound(sA)
    '/* pre execution check
    If Not UBound(sA) > 0 Then Exit Sub
    '/* Allow for worst case senario + some
    hi = ((ubA - lbA) \ m8) + m32
    '/* Stack to hold pending lower boundries
    ReDim lbs(m1 To hi) As Long
    '/* Stack to hold pending upper boundries
    ReDim ubs(m1 To hi) As Long
    '/* Cache pointer to the string variable
    lpStr = VarPtr(item)
    '/* Cache pointer to the string array
    lpS = VarPtr(sA(lbA)) - (lbA * m4)
                                                                                           
    '/* Get pivot index position
    Do: hi = ((ubA - lbA) \ m2) + lbA
        '/* Grab current value into item
        CopyMemBv lpStr, lpS + (lIdxA(hi) * m4), m4
        '/* Grab current index
        idxItem = lIdxA(hi): lIdxA(hi) = lIdxA(ubA)
        '/* Set bounds
        lo = lbA: hi = ubA
        '/* Storm right in
        Do
            If Not StrComp(item, sA(lIdxA(lo)), lCp) = lDr Then
                lIdxA(hi) = lIdxA(lo)
                hi = hi - m1
                Do Until hi = lo
                    If Not StrComp(sA(lIdxA(hi)), item, lCp) = lDr Then
                        lIdxA(lo) = lIdxA(hi)
                        Exit Do
                    End If
                    hi = hi - m1
                Loop
                '/* Found swaps or out of loop
                If hi = lo Then Exit Do
            End If
            lo = lo + m1
        Loop While hi > lo
        '/* Re-assign current
        lIdxA(hi) = idxItem
        If (lbA < lo - m1) Then
            If (ubA > lo + m1) Then cnt = cnt + m1: lbs(cnt) = lo + m1: ubs(cnt) = ubA
            ubA = lo - m1
        ElseIf (ubA > lo + m1) Then
            lbA = lo + m1
        Else
            If cnt = m0 Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - m1
        End If
    Loop: CopyMemBr ByVal lpStr, 0&, m4
    
End Sub

Private Sub ResizeArray(ByRef cArray As Variant, _
                        ByVal lPos As Long)

'/* redimension array

Dim lCt     As Long
Dim lLb     As Long
Dim lUb     As Long
   
    If Not IsArray(cArray) Then Exit Sub
    lLb = LBound(cArray)
    lUb = UBound(cArray)

    If (lUb = -1) Or (lUb - lLb = 0) Then
        Erase cArray
        Exit Sub
    End If

    '/* if invalid Pos
    If (lPos > lUb) Or (lPos = -1) Then
        lPos = lUb
    ElseIf lPos < lLb Then
        lPos = lLb
    ElseIf lPos = lUb Then
        ReDim Preserve cArray(lUb - 1)
        Exit Sub
    End If

    cArray(lPos) = cArray(lUb)
    ReDim Preserve cArray(lUb - 1)
   
End Sub

Private Sub ResizeStruct(ByVal lPos As Long)
'/* reset array

Dim iCt As Long

    With m_HLIStc(0)
        LSet .SubItem(lPos) = .SubItem(UBound(.SubItem))
        ReDim Preserve .SubItem(UBound(.SubItem) - 1)
    End With
    
End Sub

Private Sub ReverseSort()
'/* reverse sort array
Dim lTPtr() As Long
Dim lCt     As Long
Dim lRt     As Long
Dim lLb     As Long
Dim lUb     As Long

On Error GoTo Handler

    lLb = LBound(m_lPtr)
    lUb = UBound(m_lPtr)
    ReDim lTPtr(lLb To lUb)
    lCt = lLb
    lRt = lUb
    Do
        lTPtr(lCt) = m_lPtr(lRt)
        lCt = lCt + 1
        lRt = lRt - 1
    Loop Until lCt > lUb
    Erase m_lPtr
    m_lPtr = lTPtr
    
Handler:

End Sub

Private Function SortControl(ByVal lSortType As Long, _
                             ByVal lColumn As Long) As Boolean

'/* sorting hub
'/* Case - lCp
'/* 1 no case, 0 case(binary)
'/* Order - lDir
'/* 1 ascend, -1 descend

On Error GoTo Handler

    m_eSortTag = ColumnTag(lColumn)
    '/* auto determine sort type
    If m_eSortTag = SortAuto Then
        m_eSortTag = SortSample(lColumn)
    End If
    
    '/* build temp array
    Select Case m_eSortTag
    '/* string sort
    Case SortDefault
        BuildStringSortArray lColumn
        '/* array less then min dimensions
        If Not ArrayCheck(m_sSortArray) Then GoTo Handler
        If UBound(m_sSortArray) < 2 Then GoTo Handler
        '/* default sort
        If lSortType = 0 Then lSortType = 1
        '/* load a new pointer index
        QSIInitPtr LBound(m_sSortArray), UBound(m_sSortArray), m_lPtr

        Select Case lSortType
        '/* ascending case sensitive
        Case 1
            QSIStringSort m_sSortArray, m_lPtr, 0&, 1&
        '/* reverse case sensitive
        Case 2
            QSIStringSort m_sSortArray, m_lPtr, 0&, -1&
        '/* forward case insensitive
        Case 3
            QSIStringSort m_sSortArray, m_lPtr, 1&, 1&
        '/* reverse case insensitive
        Case 4
            QSIStringSort m_sSortArray, m_lPtr, 1&, -1&
        End Select
        SetItemCount Count
        
    '/* numeric and date
    Case SortDate, SortNumeric
        BuildNumericSortArray lColumn
        If Not ArrayCheck(m_lSortArray) Then GoTo Handler
        If UBound(m_lSortArray) < 2 Then GoTo Handler
        If lSortType = 0 Then lSortType = 1
        QSIInitPtr LBound(m_lSortArray), UBound(m_lSortArray), m_lPtr

        Select Case lSortType
        '/* descending
        Case 1
            QSINumericSort m_lSortArray, m_lPtr, True
        '/* ascending
        Case 2
            QSINumericSort m_lSortArray, m_lPtr, False
        End Select
        SetItemCount Count

    Case Else
        GoTo Handler
    End Select

    '/* success
    ListRefresh
    m_bSorted = True
    SortControl = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("Sort_Control", Err.Number)

End Function

Private Function SortSample(ByVal lColumn As Long) As ECGColumnSortTags
'/* auto determine sort type

Dim vItem As Variant

    If lColumn = 0 Then
        vItem = ItemText(1)
    Else
        vItem = SubItemText((1), lColumn)
    End If
    If IsDate(vItem) Then
        SortSample = SortDate
    ElseIf IsNumeric(vItem) Then
        SortSample = SortNumeric
    ElseIf Len(vItem) > 0 Then
        SortSample = SortDefault
    Else
        SortSample = SortNone
    End If
    
End Function

Private Sub SortTest()
'/* test sort

Dim l As Long

    For l = LBound(m_lPtr) To UBound(m_lPtr)
        Debug.Print m_lPtr(l)
    Next l
    
End Sub


'**********************************************************************
'*                              SUBCLASSING
'**********************************************************************

Private Sub ListAttatch(ByVal lHwnd As Long)
'/* attatch messages

    If lHwnd = 0 Then Exit Sub
    With m_cHListSubclass
        .Subclass m_lParentHwnd, Me
        .AddMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
        .AddMessage m_lParentHwnd, WM_SETFOCUS, MSG_BEFORE
        .AddMessage m_lParentHwnd, WM_SIZE, MSG_BEFORE
        .Subclass lHwnd, Me
        .AddMessage lHwnd, WM_TIMER, MSG_BEFORE
        .AddMessage lHwnd, WM_NOTIFY, MSG_BEFORE
    End With
    
End Sub

Private Sub ListDetatch(ByVal lHwnd As Long)
'/* detatch messages

    If lHwnd = 0 Then Exit Sub
    With m_cHListSubclass
        .DeleteMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
        .DeleteMessage m_lParentHwnd, WM_SETFOCUS, MSG_BEFORE
        .DeleteMessage m_lParentHwnd, WM_SIZE, MSG_BEFORE
        .UnSubclass m_lParentHwnd
        .DeleteMessage lHwnd, WM_TIMER, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NOTIFY, MSG_BEFORE
        .UnSubclass lHwnd
    End With

End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                ByRef bHandled As Boolean, _
                                ByRef lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                ByRef lParamUser As Long)

Dim bDesc                   As Boolean
Dim lCt                     As Long
Dim hIcon                   As Long
Dim lCode                   As Long
Dim lHandle                 As Long
Dim lItem                   As Long
Dim lLen                    As Long
Dim lRet                    As Long
Dim sTemp                   As String
Dim tNmhdr                  As NMHDR
Dim tNmList                 As NM_LISTVIEW
Dim tDisp                   As LV_DISPINFO
Dim LVI                     As LV_ITEM
Dim tGRedraw                As NMLVCUSTOMDRAW

On Error GoTo Handler

    '/* focus state
    If lHwnd = m_lParentHwnd Then
        If uMsg = WM_SETFOCUS Then
            SetFocus m_lLVHwnd
            lItem = ItemTopIndex
            ItemEnsureVisible lItem
            ItemSelected(lItem) = True
            Exit Sub
        ElseIf uMsg = WM_SIZE Then
            Resize
            Exit Sub
        End If
    End If

    Select Case uMsg
    Case WM_NOTIFY
        CopyMemory tNmhdr, ByVal lParam, LenB(tNmhdr)
        '/* get msg code and owner
        With tNmhdr
            If .code = TTN_GETDISPINFO Then
                lReturn = 1
                bHandled = True
                Exit Sub
            End If
            lCode = .code
            lHandle = .hwndFrom
        End With

        '/* parent
        Select Case lHandle
        Case m_lHdrHwnd
            If lCode = HDN_ITEMCHANGING Then
                If m_bHeaderFixed Then
                    lReturn = 1
                    bHandled = True
                End If
            End If
            
            '/* custom header colors
            If lCode = NM_CUSTOMDRAW Then
                If Not m_bSkinHeader Then
                    If m_bCustomHeader Then
                        CopyMemory tGRedraw, ByVal lParam, Len(tGRedraw)
                        Select Case tGRedraw.nmcmd.dwDrawStage
                        Case CDDS_PREPAINT
                            lReturn = CDRF_NOTIFYITEMDRAW
                            bHandled = True
                            Exit Sub
                        Case CDDS_ITEMPREPAINT
                            SetTextColor tGRedraw.nmcmd.hdc, m_oHdrForeClr
                            SetBkColor tGRedraw.nmcmd.hdc, m_oHdrBkClr
                            bHandled = True
                            Exit Sub
                        Case CDDS_ITEMPOSTPAINT
                            lReturn = CDRF_DODEFAULT
                            bHandled = True
                            Exit Sub
                        End Select
                    End If
                End If
            End If
            
        '/* list
        Case m_lLVHwnd
            '/* custom rows
            If lCode = NM_CUSTOMDRAW Then
                If m_bEnabled Then
                    '/* using alpha bar and row colors
                    If m_bCustomDraw And m_bAlphaSelectorBar Then
                        CopyMemory tGRedraw, ByVal lParam, Len(tGRedraw)
                        If ((tGRedraw.nmcmd.iItemState And CDIS_FOCUS) = CDIS_FOCUS) And _
                            ((tGRedraw.nmcmd.iItemState And CDIS_SELECTED) = CDIS_SELECTED) Then
                            lRet = AlphaCustomRow(lParam)
                            lReturn = lRet
                            bHandled = True
                        Else
                            lRet = CustomRow(lParam)
                            lReturn = lRet
                            bHandled = True
                        End If
                    '/* alpha only
                    ElseIf m_bAlphaSelectorBar Then
                        lRet = AlphaCustomRow(lParam)
                        lReturn = lRet
                        bHandled = True
                    '/* rows only
                    ElseIf m_bCustomDraw Then
                        lRet = CustomRow(lParam)
                        lReturn = lRet
                        bHandled = True
                    End If
                End If
                
            End If

            Select Case lCode
            '/* set edit flag and start editing
            Case NM_DBLCLK
                If m_bEditLabels Then
                    CopyMemory tNmList, ByVal lParam, Len(tNmList)
                    If m_bCheckBoxes Then
                        If tNmList.ptAction.x > 20 Then
                            SendMessageLong m_lLVHwnd, LVM_EDITLABEL, tNmList.iItem, 0&
                            m_bItemActive = True
                        End If
                    End If
                End If
            
            '/* set edit flag
            Case LVN_BEGINLABELEDIT
                m_bItemActive = True
            
            '/* end label edit
            Case LVN_ENDLABELEDIT
                CopyMemory tDisp, ByVal lParam, Len(tDisp)
                With tDisp
                    If m_bSorted Then
                        lItem = m_lPtr(.item.iItem)
                    Else
                        lItem = .item.iItem
                    End If
                    sTemp = PointerToString(.item.pszText)
                    lLen = lstrlen(sTemp)
                    If lLen > 1 Then
                        Select Case m_eListMode
                            Case eCustomDraw
                                m_cListItems(lItem).Text = sTemp
                            Case eHyperList
                                m_HLIStc(0).item(lItem) = sTemp
                            Case eDatabase
                                RaiseEvent eHLabelChange(lItem, sTemp)
                            End Select
                        End If
                    End With
                    m_bItemActive = False
                    
            '/* column click
            Case LVN_COLUMNCLICK
                If m_bUseSorted Then
                    CopyMemory tNmList, ByVal lParam, LenB(tNmList)
                    RaiseEvent eHColumnClick(tNmList.iSubItem)
                    '/* swap sort icon
                    For lCt = 0 To ColumnCount - 1
                        ColumnIcon(lCt) = -1
                    Next lCt
                    If ColumnIcon(tNmList.iSubItem) = 1 Then
                        ColumnIcon(tNmList.iSubItem) = 0
                        bDesc = True
                    Else
                        ColumnIcon(tNmList.iSubItem) = 1
                    End If
                    '/* column and sort direction
                    If bDesc Then
                        SortControl 1, tNmList.iSubItem
                    Else
                        SortControl 2, tNmList.iSubItem
                    End If
                End If

            '/* item changed
            Case LVN_ITEMCHANGED
                CopyMemory tNmList, ByVal lParam, Len(tNmList)
                With tNmList
                    If .uOldState Then
                        If ((.uNewState And LVIS_STATEIMAGEMASK) <> (.uOldState And LVIS_STATEIMAGEMASK)) Then
                            RaiseEvent eHItemCheck(.iItem)
                        End If
                    Else
                        If Not m_bFirstItem Then
                            If ((.uNewState And LVIS_SELECTED)) Then
                                RaiseEvent eHItemClick(.iItem)
                            End If
                        End If
                    End If
                End With

            '/* checkbox click
            Case LVIS_CHKCLICK
                CopyMemory tNmList, ByVal lParam, LenB(tNmList)
                If tNmList.iSubItem = 0 Then
                    If m_bCheckBoxes Then
                        CopyMemory tNmList, ByVal lParam, Len(tNmList)
                        If m_eViewMode = StyleReport Then
                            If tNmList.ptAction.x < 20 Then
                                If m_bSorted Then
                                    CheckToggle m_lPtr(tNmList.iItem)
                                Else
                                    CheckToggle tNmList.iItem
                                End If
                                ListRefresh
                            End If
                        Else
                            CheckToggle tNmList.iItem
                            ListRefresh
                        End If
                    End If
                End If
                '/* alpha bar
                If m_bAlphaSelectorBar Then
                    If Not m_bItemActive Then
                        DrawAlphaSelectorBar tNmList.iItem
                    End If
                End If
            
            '/* throw edit flag
            Case LVN_ITEMACTIVATE
                If m_bEditLabels Then
                    m_bItemActive = True
                End If
                
            '/* list change callback
            Case LVN_GETDISPINFO
                CopyMemory tDisp, ByVal lParam, LenB(tDisp)
                With tDisp.item
                    Select Case m_eListMode
                    Case eCustomDraw
                        '/* display items by qualified pointer index
                        If m_bUseSorted And m_bSorted Then
                            Select Case .iSubItem
                            Case 0
                                sTemp = m_cListItems(m_lPtr(.iItem)).Text
                            Case Else
                                sTemp = m_cListItems(m_lPtr(.iItem)).SubItemText(.iSubItem)
                            End Select
                        '/* normal retrieval
                        Else
                            Select Case .iSubItem
                            Case 0
                                sTemp = m_cListItems(.iItem).Text
                            Case Else
                                sTemp = m_cListItems(.iItem).SubItemText(.iSubItem)
                            End Select
                        End If
                    Case eHyperList
                        If m_bUseSorted And m_bSorted Then
                            Select Case .iSubItem
                            Case 0
                                sTemp = m_HLIStc(0).item(m_lPtr(.iItem))
                            Case Else
                                sTemp = m_HLIStc(0).SubItem(m_lPtr(.iItem)).Text(.iSubItem)
                            End Select
                        '/* normal retrieval
                        Else
                            Select Case .iSubItem
                            Case 0
                                sTemp = m_HLIStc(0).item(.iItem)
                            Case Else
                                sTemp = m_HLIStc(0).SubItem(.iItem).Text(.iSubItem)
                            End Select
                        End If
                    Case eDatabase
                        RaiseEvent eHIndirect(.iItem, .iSubItem, .Mask, sTemp, .iImage)
                    End Select
                    
                    '/* skinned checkbox
                    If m_eViewMode = StyleReport Then
                        If m_bCheckBoxes Then
                            If m_bSkinnedCheck Then
                                OwnerDrawCheckBox .iItem
                            End If
                        End If
                    End If
                End With

                With tDisp.item
                    If .iSubItem = 0 Then
                        If .Mask And LVIF_TEXT Then
                            '/ copy text
                            If Len(sTemp) > .cchTextMax Then
                                sTemp = left$(sTemp, .cchTextMax)
                            End If
                            StringToPointer sTemp, .pszText
                        End If
                        '/* check state
                        If .Mask And LVIF_IMAGE Then
                            If m_bCheckBoxes Then
                                If m_bSorted Then
                                    Select Case pItemChecked(m_lPtr(.iItem))
                                    Case 0
                                        .State = LVIS_UNCHECKED
                                    Case 1
                                        .State = LVIS_CHECKED
                                    End Select
                                Else
                                    Select Case pItemChecked(.iItem)
                                    Case 0
                                        .State = LVIS_UNCHECKED
                                    Case 1
                                        .State = LVIS_CHECKED
                                    End Select
                                End If
                            End If
                            '/* icon
                            Select Case m_eListMode
                            Case eCustomDraw
                                .iImage = m_cListItems(tDisp.item.iItem).Icon
                            Case eHyperList
                                .iImage = m_HLIStc(0).lIcon(tDisp.item.iItem)
                            Case eDatabase
                                RaiseEvent eHIndirect(.iItem, .iSubItem, .Mask, sTemp, .iImage)
                            End Select
                            If Not m_bEnabled And m_eViewMode = StyleReport Then
                                IconDisabled tDisp.item.iItem, .iImage
                                .iImage = -1
                            End If
                            '/* skinned checkbox
                            If m_bSkinnedCheck Then
                                .Mask = LVIF_IMAGE Or LVIF_TEXT
                                .stateMask = LVIS_OVERLAYMASK
                            Else
                                .Mask = LVIF_IMAGE Or LVIF_TEXT Or LVIF_STATE
                                .stateMask = LVIS_OVERLAYMASK Or LVIS_STATEIMAGEMASK
                            End If
                        End If
                        '/* indent
                        If .Mask And LVIF_INDENT Then
                            .iIndent = m_lItemIndent
                        End If
                    Else
                        If .Mask And LVIF_TEXT Then
                            '/ copy text
                            If Len(sTemp) > .cchTextMax Then
                                sTemp = left$(sTemp, .cchTextMax)
                            End If
                            StringToPointer sTemp, .pszText
                        End If
                        '/* subitem image
                        If m_bSubItemImage Then
                            Select Case m_eListMode
                            Case eCustomDraw
                                .iImage = m_cListItems(tDisp.item.iItem).SubItemIcon(tDisp.item.iSubItem)
                            Case eHyperList
                                .iImage = m_HLIStc(0).SubItem(tDisp.item.iItem).lIcon(tDisp.item.iSubItem)
                            Case eDatabase
                                RaiseEvent eHIndirect(.iItem, .iSubItem, .Mask, sTemp, .iImage)
                            End Select
                            .Mask = LVIF_IMAGE Or LVIF_TEXT
                            .stateMask = LVIS_OVERLAYMASK
                        End If
                    End If
                End With
                '/* copy and forward
                CopyMemory ByVal lParam, tDisp, LenB(tDisp)
            End Select
        End Select
    End Select

Handler:
    On Error GoTo 0
    
End Sub


'**********************************************************************
'*                              CUSTOM DRAW
'**********************************************************************

Private Function AlphaCustomRow(ByVal lParam As Long) As Long
'/* a lot of guesswork..

Dim clrTextBk   As OLE_COLOR
Dim clrText     As OLE_COLOR
Dim tNMLVCD     As NMLVCUSTOMDRAW

    CopyMemory tNMLVCD, ByVal lParam, Len(tNMLVCD)
    With tNMLVCD
        Select Case .nmcmd.dwDrawStage
        Case CDDS_PREPAINT
            AlphaCustomRow = CDRF_NOTIFYITEMDRAW

        Case CDDS_ITEMPREPAINT
            CopyMemory ByVal lParam, tNMLVCD, Len(tNMLVCD)
            With tNMLVCD
                If (.nmcmd.iItemState And CDIS_SELECTED) = CDIS_SELECTED Then
                    If (.nmcmd.iItemState And CDIS_FOCUS) = CDIS_FOCUS Then
                        If m_bAlphaThemeBackClr Then
                            .clrTextBk = XPShift(m_oThemeColor, 190)
                        ElseIf m_bRowDecoration Then
                            RowColors .nmcmd.dwItemSpec, .iSubItem, .clrTextBk, .clrText
                        Else
                            .clrTextBk = m_lBackColor
                        End If
                        ItemSelected(.nmcmd.dwItemSpec) = False
                    End If
                    CopyMemory ByVal lParam, tNMLVCD, Len(tNMLVCD)
                    AlphaCustomRow = CDRF_DODEFAULT
                End If
            End With
        End Select
    End With
    
End Function

Private Function CustomRow(ByVal lParam As Long) As Long
' http://www.codeproject.com/listctrl/lvcustomdraw.asp

Dim clrTextBk   As OLE_COLOR
Dim clrText     As OLE_COLOR
Dim tNMLVCD     As NMLVCUSTOMDRAW

    CopyMemory tNMLVCD, ByVal lParam, Len(tNMLVCD)
    With tNMLVCD
        Select Case .nmcmd.dwDrawStage
        Case CDDS_PREPAINT
            CustomRow = CDRF_NOTIFYITEMDRAW
          
        Case CDDS_ITEMPREPAINT
            CustomRow = CDRF_NOTIFYSUBITEMDRAW
            
        Case CDDS_ITEMPREPAINT Or CDDS_SUBITEM
            With m_cListItems(tNMLVCD.nmcmd.dwItemSpec)
                If Not .CellCustom Then
                    clrTextBk = .BackColor
                    clrText = .ForeColor
                Else
                    clrTextBk = .CellBackColor(tNMLVCD.iSubItem)
                    clrText = .CellForeColor(tNMLVCD.iSubItem)
                End If
                If Not .Font Is Nothing Then
                    RowFont tNMLVCD.nmcmd.hdc, .Font
                End If
                If m_bRowDecoration Then
                    RowColors tNMLVCD.nmcmd.dwItemSpec, tNMLVCD.iSubItem, clrTextBk, clrText
                End If
            End With
            
            If Not clrTextBk = -1 Then
                OleTranslateColor clrTextBk, 0, .clrTextBk
            End If

            If Not clrText = -1 Then
                OleTranslateColor clrText, 0, .clrText
            End If
            If m_bWordWrap Then
                With tNMLVCD.nmcmd
                '    DrawWordBreak .dwItemSpec, tNMLVCD.iSubItem, .hdc, .rc
                End With
            End If
            CopyMemory ByVal lParam, tNMLVCD, Len(tNMLVCD)
            CustomRow = CDRF_DODEFAULT
        End Select
    End With

End Function

Private Sub DrawWordBreak(ByVal lItem As Long, _
                          ByVal lSubItem As Long, _
                          ByVal lHDC As Long, _
                          ByRef tRect As RECT)

'/* row with wordwrap - todo* not working yet

Dim hBr     As Long
Dim lHeight As Long
Dim sText   As String

    If m_eViewMode = StyleReport Then
        With tRect
            lHeight = .bottom - .top
            If lHeight > 16 Then
                hBr = CreateSolidBrush(TranslateColor(vbWhite))
                FillRect lHDC, tRect, hBr
                DeleteObject hBr
                sText = SubItemText(lItem, lSubItem)
#If Unicode Then
                DrawText lHDC, StrPtr(sText), -1, tRect, DT_LEFT Or DT_WORDBREAK Or DT_SINGLELINE
#Else
                DrawText lHDC, sText, 0&, tRect, DT_LEFT Or DT_WORDBREAK Or DT_CALCRECT
#End If
            End If
        End With
    End If

End Sub

Private Function IndexToStateImageMask(ByVal lState As Long) As Long
'/* state mask translate
   IndexToStateImageMask = lState * (2 ^ 12)
End Function

Private Function OwnerDrawCheckBox(ByVal lItem As Long) As Boolean
'/* render custom checkboxes

Dim lYOffset    As Long
Dim lXOffset    As Long
Dim lIOffset    As Long
Dim lROffset    As Long
Dim lHDC        As Long
Dim lCkPtr      As Long
Dim tRect       As RECT
Dim tOffst      As RECT
    
    lHDC = GetDC(m_lLVHwnd)
    GetItemRect lItem, tRect, True
    With tRect
        If Not lItem = 0 Then
            lYOffset = ColumnHeight + 1
        Else
            lYOffset = ColumnHeight + 3
            GetItemRect lItem + 1, tOffst, True
            .left = tOffst.left
        End If
        lYOffset = lYOffset + (.top)
        lXOffset = m_lItemIndent + 2
        lROffset = ((.bottom - .top) - m_lCheckHeight) / 2
        '/* check array pointer
        If m_bSorted Then
            lCkPtr = m_lPtr(lItem)
        Else
            lCkPtr = lItem
        End If
        If m_bEnabled = False Then
            lIOffset = (13 * 2)
        ElseIf pItemChecked(lCkPtr) Then
            lIOffset = (13 * 1)
        Else
            lIOffset = 0
        End If
        m_cRender.Blit lHDC, (lXOffset + .left), lYOffset, (13 + .left), 13, m_cChkCheckDc.hdc, lIOffset, 0, SRCCOPY
    End With
    ReleaseDC m_lLVHwnd, lHDC

End Function

Private Function RowColors(ByVal lRow As Long, _
                           ByVal lSubItem As Long, _
                           ByRef lRowClr As Long, _
                           ByRef lRowForeClr As Long)

'/* apply row color patterns

On Error GoTo Handler

    If Not m_bRowDecoration Then Exit Function
    
    If Not ArrayCheck(m_lRowColor) Then
        RowDecoration m_eRowDecoration, m_lRowColorBase, m_lRowColorOffset, m_bRowUseXP, m_lRowDepth
    Else
        If m_eListMode = eCustomDraw Then
            If (lRow < LBound(m_cListItems)) Or (lRow > UBound(m_cListItems)) Then
                RowDecoration m_eRowDecoration, m_lRowColorBase, m_lRowColorOffset, m_bRowUseXP, m_lRowDepth
            End If
        ElseIf m_eListMode = eHyperList Then
            If (lRow < LBound(m_HLIStc(0).item)) Or (lRow > UBound(m_HLIStc(0).item)) Then
                RowDecoration m_eRowDecoration, m_lRowColorBase, m_lRowColorOffset, m_bRowUseXP, m_lRowDepth
            End If
        Else
            Exit Function
        End If
    End If
    
    Select Case m_eRowDecoration
    Case RowLine
        lRowClr = m_lRowColor(lRow)
    
    Case RowChecker
        If (lSubItem Mod 2) Then
            If m_lRowColor(lRow) = m_lRowColorBase Then
                lRowClr = m_lRowColorOffset
            Else
                lRowClr = m_lRowColorBase
            End If
        Else
            lRowClr = m_lRowColor(lRow)
        End If
    End Select

Handler:
    On Error GoTo 0
    
End Function

Public Function RowDecoration(ByVal eRowDecoration As ECGRowDecoration, _
                              ByVal lBaseClr As Long, _
                              ByVal lOffsetClr As Long, _
                              ByVal bXPColors As Boolean, _
                              Optional ByVal lRowDepth As Long) As Boolean

'/* build custom row color arrays

Dim bCs As Boolean
Dim lCt As Long
Dim lCr As Long
Dim lUb As Long

On Error GoTo Handler

    lUb = (Count - 1)
    m_eRowDecoration = eRowDecoration
    m_lRowDepth = lRowDepth
    m_bRowUseXP = bXPColors
    
    If bXPColors Then
        m_lRowColorBase = XPShift(lBaseClr, 120)
        m_lRowColorOffset = XPShift(lOffsetClr, 120)
    Else
        m_lRowColorBase = lBaseClr
        m_lRowColorOffset = lOffsetClr
    End If
    
    If lUb = -1 Then
        m_bRowDecoration = True
        m_bCustomDraw = True
        Exit Function
    End If
    
    ReDim m_lRowColor(lUb)
    Do
        If lCr > lRowDepth Then
            lCr = 0
            bCs = Not bCs
        End If
        If bCs Then
            m_lRowColor(lCt) = m_lRowColorBase
        Else
            m_lRowColor(lCt) = m_lRowColorOffset
        End If
        lCr = lCr + 1
        lCt = lCt + 1
    Loop Until lCt > lUb

    m_bRowDecoration = True
    m_bCustomDraw = True

Handler:
    On Error GoTo 0

End Function

Private Function RowFont(ByVal lHDC As Long, _
                        ByVal oFont As StdFont)
'*/ change list font

Dim uLF     As LOGFONT
Dim lChar   As Long
Dim lFont   As Long

On Error GoTo Handler

    Set m_oFont = oFont
    With uLF
        For lChar = 1 To Len(m_oFont.Name)
            .lfFaceName(lChar - 1) = CByte(Asc(Mid$(m_oFont.Name, lChar, 1)))
        Next lChar
        .lfHeight = m_oFont.Size
        .lfItalic = m_oFont.Italic
        .lfWeight = IIf(m_oFont.Bold, FW_BOLD, FW_NORMAL)
        .lfUnderline = m_oFont.Underline
        .lfStrikeOut = m_oFont.Strikethrough
        .lfCharSet = m_oFont.Charset
    End With
    lFont = CreateFontIndirect(uLF)
    SelectObject lHDC, lFont
    DeleteObject lFont

Handler:
    On Error GoTo 0

End Function


'**********************************************************************
'*                              CLEANUP
'**********************************************************************

Private Function DeAllocatePointer(ByVal sKey As String, _
                                   Optional ByVal bPurge As Boolean) As Boolean

'/* resolve or purge memory pointers

Dim lPtr    As Long
Dim lC      As Long

On Error GoTo Handler

    If Not bPurge Then
        '/* get the pointer
        lPtr = c_PtrMem.item(sKey)
        If lPtr = 0 Then GoTo Handler
        '/* release the memory
        CopyMemory ByVal lPtr, 0&, 4&
    Else
        '/* destroy the struct last
        For lC = c_PtrMem.Count To 1 Step -1
            If Not CLng(c_PtrMem.item(lC)) = 0 Then
                lPtr = CLng(c_PtrMem.item(lC))
                CopyMemory ByVal lPtr, 0&, 4&
            End If
        Next lC
        m_lStrctPtr = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DeAllocatePointer", Err.Number)

End Function

Private Function DestroyList() As Boolean
'/* cleanup

    DestroyImlHeader
    DestroyImlSmall
    DestroyImlLarge
    DestroyImlState
    If m_lLVHwnd Then
        If DestroyWindow(m_lLVHwnd) Then
            DestroyList = True
            m_lLVHwnd = 0
        End If
    End If

End Function

Private Sub DestroyItems()
'/* destroy item classes

Dim lCt     As Long
Dim lUBound As Long

On Error GoTo Handler

    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        '/* destroy classes
        lUBound = UBound(m_cListItems)
        For lCt = 0 To lUBound
            Set m_cListItems(lCt) = Nothing
        Next lCt
    '/* hl mode
    Case eHyperList
        '/* rease structs
        Erase m_HLIStc(0).item
        Erase m_HLIStc(0).lIcon
        Erase m_HLIStc(0).SubItem
        ReDim m_HLIStc(0)
    End Select
    
    If Not c_ColumnTags Is Nothing Then Set c_ColumnTags = Nothing
    If Not c_PtrMem Is Nothing Then Set c_PtrMem = Nothing

Handler:
    On Error GoTo 0
    
End Sub

Private Sub DestroyImages()
'/* destroy images

    If Not IChecked Is Nothing Then Set IChecked = Nothing
    If Not m_cChkCheckDc Is Nothing Then Set m_cChkCheckDc = Nothing
    If Not m_cRender Is Nothing Then Set m_cRender = Nothing

End Sub

Private Sub Class_Terminate()

    ListDetatch m_lLVHwnd
    DestroyItems
    DestroyList
    DestroyImages
    DeAllocatePointer "a", True
    Set m_cHListSubclass = Nothing
    If Not (m_lhMod = 0) Then
        FreeLibrary m_lhMod
    End If
    
End Sub
