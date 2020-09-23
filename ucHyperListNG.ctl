VERSION 5.00
Begin VB.UserControl ucHyperListNG 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucHyperListNG.ctx":0000
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1125
   End
End
Attribute VB_Name = "ucHyperListNG"
Attribute VB_GlobalNameSpace = False
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
'*  HyperList G2!   HL-NGEN - Virtual Listview Hybrid 2.4  (uc version)                 *
'*                                                                                      *
'*  Created:        (V1) June 23, 2006                                                  *
'*  HL NGen:        (V2) October 3, 2006                                                *
'*  Updated:        November 14, 2006                                                   *
'*  Purpose:        Ultra-Fast Virtual Listview Hybrid                                  *
'*  Functions:      (listed)                                                            *
'*  Revision:       2.4.6                                                               *
'*  Compile:        Native                                                              *
'*  Author:         John Underhill (Steppenwolfe)                                       *
'*                                                                                      *
'****************************************************************************************


' ~*** Exposed Functions ***~

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
'/~ Find                        - search for an item                [in -string | in -bool(3) | in -long (1) | out -long]
'/~ ImlHeaderAddBmp             - add header bmp                    [in -long | out -bool]
'/~ ImlHeaderAddIcon            - add header icon                   [in -long | out -bool]
'/~ ImlLargeAddBmp              - add large bmp                     [in -long | out -bool]
'/~ ImlLargeAddIcon             - add large icon                    [in -long | out -bool]
'/~ ImlSmallAddBmp              - add small bmp                     [in -long | out -bool]
'/~ ImlSmallAddIcon             - add small icon                    [in -long | out -bool]
'/~ ImlStateAddBmp              - add state bmp                     [in -long | out -bool]
'/~ ImlStateAddIcon             - add state icon                    [in -long | out -bool]
'/~ InitImlHeader               - set header iml                    [in -none | out -bool]
'/~ InitImlLarge                - set large iml                     [in -none | out -bool]
'/~ InitImlSmall                - set small iml                     [in -long () | out -bool]
'/~ InitImlState                - set state iml                     [in -long | out -bool]
'/~ InitList                    - init listview                     [in -long (2) | out -bool]
'/~ IsUnicode                   - test a string for unicode         [in -string | out -bool]
'/~ ItemAdd                     - add list item                     [in -long (3) string (2) | out -bool]
'/~ ItemEnsureVisible           - scroll to item                    [in -long | out -bool]
'/~ ItemRemove                  - remove item                       [in -long | out -bool]
'/~ ItemsSort                   - sort items                        [in -long | in -bool | out -bool]
'/~ ItemRedraw                  - redraw items                      [in -none | out -bool]
'/~ ItemTopIndex                - top item                          [in -none | out -long]
'/~ ListRefresh                 - refresh listview                  [in -none | out -none]
'/~ LoadArray                   - load items array                  [in -none | out -bool]
'/~ LoadFromFile                - load list from file               [in -string () | out -bool]
'/~ Refresh                     - refresh list                      [in -none | out -bool]
'/~ RemoveDuplicates            - remove list duplicates            [in -none | out -bool]
'/~ Resize                      - flag a resize                     [in -none | out -none]
'/~ RowDecoration               - row colors                        [in -long (3) in -enum | in -bool | out -bool]
'/~ SaveToFile                  - save list items                   [in -string | out -bool]
'/~ SetFocus                    - give the list focus               [in -bool | out -bool]
'/~ SetItemCount                - init list                         [in -long | out -bool]
'/~ SkinCheckBox                - load checkboxskin                 [in -enum | in -bool | out -bool]
'/~ SkinHeaders                 - load header skin                  [in -none | out -bool]
'/~ SkinScrollBars              - load scrollbar skin               [in -none | out -bool]
'/~ SkinXPHeader                - xp style header                   [in -none | out -bool]
'/~ SubIconIndex                - set subitem icon                  [in -long (2) | out -bool]
'/~ SubItemsAdd                 - add subitem                       [in -long (2) | in -string | out -bool]
'/~ UnCheckAll                  - unmark all checkboxes             [in -none | out -bool]
'/~ UnSkinAll                   - remove all skins                  [in -none | out -bool]
'/~ UnSkinCheckBox              - remove checkbox skin              [in -none | out -bool]
'/~ UnSkinHeaders               - remove header skin                [in -none | out -bool]
'/~ UnSkinScrollBars            - remove scrollbar skin             [in -none | out -bool]
'/~ UnSkinXPHeader              - unload xp header style            [in -none | out -bool]


' ~*** Exposed Properties ***~

'/~ AlphaBarTheme               - use alpha bar theme colors        [bool]
'/~ AlphaBarTransparency        - alpha bar transparency index      [byte]
'/~ AlphaBarActive              - use alpha bar                     [bool]
'/~ AlphaThemeBackClr           - use theme backcolor               [bool]
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
'/~ Focus                       - focus item                        [bool]
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
'/~ HeaderPressed               - header font pressed               [long]
'/~ Height                      - listview height                   [long]
'/~ HideSelection               - hide item                         [bool]
'/~ IconSpaceX                  - icon shift position X             [long]
'/~ IconSpaceY                  - icon shift position Y             [long]
'/~ InfoTips                    - use info tips                     [bool]
'/~ InsensitiveSort             - use case insensitive sort         [bool]
'/~ IsWinNT                     - nt4 or above                      [bool]
'/~ IsWinXP                     - in xp operating system            [bool]
'/~ ItemBorderSelect            - use item border                   [bool]
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
'/~ OLEDragMode                 - ole drag mode                     [enum]
'/~ OLEDropMode                 - ole drop mode                     [enum]
'/~ OneClickActivate            - edit one click activate           [bool]
'/~ ScaleHeight                 - list scaleheight                  [long]
'/~ ScaleMode                   - list scale measurement            [enum]
'/~ ScaleWidth                  - list scale width                  [long]
'/~ ScrollBarFlat               - flat scrollbar (non skin)         [bool]
'/~ SelectedCount               - selected items count              [long]
'/~ StructPtr                   - array pointer                     [long]
'/~ SubItemsEdit                - edit subitems                     [bool]
'/~ SubItemIcon                 - subitem icon index                [long]
'/~ SubItemImages               - use subitem images                [bool]
'/~ SubItemText                 - subitem text                      [string]
'/~ TextAlignment               - right or left text align          [enum]
'/~ ThemeColor                  - theme base color                  [long]
'/~ ThemeLuminence              - theme luminence                   [enum]
'/~ TrackSelected               - track selected item               [bool]
'/~ UnderlineHot                - underline hot item                [bool]
'/~ UseCellColor                - enable per cell colors            [bool]
'/~ UseCellFont                 - enable per cell fonts             [bool]
'/~ UseThemeColors              - use skin theme color              [bool]
'/~ UseUnicode                  - enable unicode mode               [bool]
'/~ ViewMode                    - listview style                    [enum]
'/~ Visible                     - toggle control visibility         [bool]
'/~ WordWrap                    - item text wordwrap                [bool]
'/~ XPColors                    - use xp color offset               [bool]

'~*** Notes ***~

'-! Disclaimer
'/~ HyperList Copyright 2006 © John Underhill, All Rights Reserved.
'/~ Obviously no warranty or liability, or any responsibility in any way imaginable, is expressed or implied.
'/~ Use this software in your personal projects in any way you wish under a relaxed GNU, but all responsibilities
'/~ are yours entirely. If this control or class is ported to a commercial project, I expect to be notified,
'/~ and at a minimum, (determined by me), the appropriate credit must be given in the help/about or other
'/~ suitable area of the software, ex. 'HL-NGEN listview class provided by John Underhill of NSPowertools.com'.

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
'/~ When compiling a library/ocx for release, -always- set the version compatability to binary.
'/~ Project > Properties > Component > Version Compatabilty > 'Binary Compatability'.
'/~ If you plan to distribute compiled dll, Rename it! Use a unique name, with a app/company name prefix,
'/~ ex. sgrhlst.ocx. If you do not, you are taking a risk that someone elses version will overwrite
'/~ your own, and your app will stop working!
'/~ The compile switches are set, but you could change them to your liking and test responsiveness
'/~ under different option settings.
'/~ You can distribute this project to other vb related websites, but all headers must remain intact,
'/~ and proper credit must be given.

'-? Recommendations
'/~ *Important: Most of the visual elements should be defined at design time, many of the property changes in demo,
'/~ should not be given as user style options, and continually flipping internal style bits, will cause
'/~ a listview to get 'glitchy'. In est: in application, define the desired styles for the list, then leave it alone.
'/~ Turn alpha bar, and custom row off when using Icon mode, too much flicker.
'/~ Skinned checkbox is not supported on non default large icon spacing, (no way to track location).
'/~ Use a reasonable number of listitems in your application. More then a couple thousand, is probably a bad UI design.
'/~ Don't use background picture with custom draw or skinning, m$ implementation doesn't play well with these features.
'/~ The more features you turn on, the longer it takes to draw, examples are near the limit.
'/~ Processing unicode strings within your application is your job, follow the links, it is
'/~ a bit complicated, but opens up a large market for your applications. The best link I have found overall is:
'/~ http://www.unicodeactivex.com/UnicodeTutorialVb.htm
'/~ I would also recommend you take a good look at sGrid II:
'/~ http://www.vbaccelerator.com/home/VB/Code/Controls/S_Grid_2/S_Grid_2/article.asp
'/~ Steve's grid and listview, though not completely unicode aware, they have much of the groundwork written in,
'/~ you just need to go that extra mile.

'-> 98/ME users
'/~ Unicode separation means this should be working in legacy OS now. Of course, there is no unicode support
'/~ in 98/ME, (forward complaints to-> bill@microsoft.com).

'-> Using Hyperlist
'-> A bit about virtual listviews.. They are 'virtual' because very little information is actually stored
'/~ in the listview itself. The application is almost entirely responsible for maintaining item state, text, icons
'/~ checkboxes etc. When you add an item to a standard listview, you are actually populating a complex
'/~ type structure with many different branches and sub structures. Data used for sorting, spacing, icons etc, are all
'/~ stored in this compound structure, and often it is data you do not need. This is why a standard listviews access
'/~ methods are both slow and consume a great deal of memory, because wether you will use the information or not, it is
'/~ stored and contstantly queried through the listviews internal methods. The virtual list however, relies on
'/~ a callback interface to retrieve data about an item. When an item scrolls into view, the parent window gets a
'/~ message requesting the items data. In this way, you can store only the information you need, and supply
'/~ a specific items elements only when it becomes visible.
'/~ Virtual listviews are commonly used in C++ to display very large databases. The database itself holds all the
'/~ information, and item data is fetched from the application through a callback to the hosting object.
'/~ The problem with this method though, is that there is no persistant data, and that a database and connectors
'/~ must be available to the application. With no internal storage facility, and said requirements, a virtual listviews
'/~ implementation was usually very specific, and did not find its way into many applications.
'/~ Hyperlist has overcome this limitation, with an internal storage container, much of the raw speed of a virtual
'/~ listview has been maintained (several hundred times faster then a standard listview), but most of the
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
'/~ ucHyperListNG control. Once the application local structure has been built, the struct is copied into the uc
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
'/~ each class member with the desired properies, local build example:
'/~ Redim Preserve m_cListItems(0 To 999)                           <- dimension the local class array
'/~ Set m_cListItems(999) = New clsListItem                         <- instantiate members
'/~ m_cListItems(999).Add 0, "key1", "item text", lIcon, lSubIcon   <- add items
'/~ CopyMemory m_lPointer, ByVal VarPtrArray(m_cListItems), 4&      <- copy array to local pointer
'/~ cHyperList.StructPtr = m_lPointer                               <- copy pointer to class
'/~ cHyperList.LoadArray                                            <- dimension the in-class array
'/~ cHyperList.SetItemCount UBound(m_cListItems) + 1                <- set the item count and instantiate
'/~ The clsListItem class is a public member, so can be instantiated directly from the application, or
'/~ the internal methods can be used [InitList and ItemAdd].

'-> DataBase mode:
'/~ This mode is in keeping with the traditional virtual listviews access method. A callback event [eHIndirect]
'/~ is used to fetch the data from an indexed database. Events have also been added for labeledit, item and column
'/~ click events.

'-> Properties/Methods
'/~ I tried to make property names as intuitive as possible, and provided property/function descriptions,
'/~ aside from that, you will just have to do some exploration of capabilities. As with any control, there is
'/~ a bit of a learning curve, but if you know how to use a listview, and consult the description header of
'/~ this control, most tasks should be fairly straightforward.

'/~ A couple of basic examples:

'/~ Checked Items:
'/~     For lCt = 0 To Count - 1
'/~         If Checked(lCt) Then
'/~             'whatever
'/~         End If
'/~     Next lCt

'/~ Item Text:
'/~     For lCt = 0 To Count - 1
'/~         If ItemText(lCt) = "123" Then
'/~             'whatever
'/~         End If
'/~     Next lCt

'-> Features
'/~ There are a lot of properties (over 110 functions/properties), so I suggest some experimenting to get their correct
'/~ implementation for those that are not demonstrated. Most properties of a listview that can be implemented are available,
'/~ some things though were not possible, (like grouping), because they require the listview to maintain item state data.

'-> Skinning and Theming
'/~ 5 seperate skin themes for checkbox/header/scrollbars. Checkbox also has an added xp style.
'/~ Skins can all be colorized with user selected color, using the [ThemeLuminence and ThemeColor] properties.
'/~ Skin styles can be mixed and matched between components by setting their individual properties.

'-@ Bugfixes Ver. 2.1
'/~ Checkbox positions now calculated through GetItemRect directly, offset method, (though faster), caused
'/~ checkbox misalignment. Skinned checkboxes now centered in row by horizontal offset, and adjustable X offset.
'/~ RemoveDuplicates function overhauled. Was setting item count using Count property, giving incorrect dimension.
'/~ ResizeArray routines were returning one member too many after removal, causing empty item on post-sort.
'/~ Sort autotag was exiting after RemoveDuplicates due to logic failure, effecting pre-sort exit, logic revised.
'/~ Columns dissapear after drag, remedied by capturing HDN_ENDDRAG message and responding with Resize call.
'/~ Background image proceedure repaired. Requires formatted url, getshortpath call used to obtain call
'/~ compliant path.
'/~ Alphabar trigger moved past checkbox boundary.
'/~ Winproc overhauled and organized.
'/~ ellipses added to header when resized past text boundary
'/~ in-proceedure unicode checks made by os version check, switched from compiler options
'/~ minimum first column width set, to bypass skinned checkbox overpainting.
'/~ fixed alpha bar length to cell size when full row select is false

'-# Version 2.2 October 11, 2006
'/~ Added- full ole drag and drop functionality.
'/~ Added- subitem edit functionality.
'/~ Added- checkbox horizontal offset property.
'/~ Added- scale measure properties.
'/~ Added- ole mode properties.

'-@ Bugfixes Ver. 2.2
'-> broken during conversion to optional unicode
'/~ item text                   - fixed
'/~ subitem edit                - fixed
'/~ sorting                     - fixed
'/~ column text                 - fixed
'/~ skin checkbox in report     - fixed
'/~ bg image broke (again)      - fixed
'/~ Complete rebuild of unicode subsystem, unicode is on demand, so 98/ME should work now, (here's hoping).
'/~ Unicode only works with installed unicode-friendly fonts, (only a couple: Ariel Wide by default),
'/~ so when font was changed and unicode engaged, font always comes up Ariel. Work around with on demand
'/~ unicode implementation with OS checks and defaults in place.
'/~ Subitem edit does not support unicode, (via api textbox, should work, but..). If using unicode strings in subitems
'/~ turn option off, as I have no plans to change this.
'/~ Forecolor/Backcolor property completed for custom row mode, was not engaged to callback in custom mode.
'/~ Report mode indent fixed by shifting placement within DISPINFO callback structure.
'/~ Subitem edit box subclassing used to respond to keyboard events, focus issues partially resolved.
'/~ Checkboxes removed from small icon mode. Checkbox logic revised.
'/~ When using unicode strings in non-unicode mode, icon mode only partially refreshes on focus event.
'/~ This is error internal to win32 listview, and can not be resolved. Test for unicode (IsUnicode), and use the
'/~ proper unicode mode.
'/~ Fixed sort on first column click, and revised sort icon logic. Added case insensitive sort property.

'-# Version 2.3 October 18, 2006
'/~ Unicode on demand. Unicode support is now fully integrated, with os checks, and method seperation. User
'/~ invokable at design or runtime. Seperation of methods also means legacy OS (98/ME) should work fine with
'/~ the control in non-unicode mode, (let me know..).
'/~ All relevant api have been mirrored to wide versions throughout, controls are created with wide class
'/~ versions on NT compatable systems, ansi only on legacy 98/ME.
'/~ Checkboxes added to Icon and List modes.
'/~ Cell level font and color formatting revised and working.
'/~ Complete rewrite of the skinheader class.
'/~ Edit control implemented as createwindow'd control.
'/~ Added xpstyle class, headers now have optional xp style rendering (demo). Class can be exported and used
'/~ to apply xp styles to almost anything.
'/~ The unicode features still have a few todo's, a couple graphics tweaks, and some spm's, next version should be
'/~ the last.

'-@ Bugfixes Ver. 2.3
'/~ Fixed text ellipses in both modes. Finished unicode compliance across all classes.
'/~ Sort pointers added to icons and other relevant places, rewrite of listitemdraw routines.
'/~ Subitem edit working in unicode now, (thanks Zhu).
'/~ Slow load of report view from icon view solved, skinned checkbox logic revised.
'/~ Fixed alphapar painting past last column by adjusting dc length.

'-# Version 2.4 October 23, 2006
'/~ Added unicode support to subclasser. Now will subclass with wide api on nt based system.
'/~ Went through project and mirrored wide api versions and os checks throughout.
'/~ Added Find function and demo.
'/~ Solved focus backcolor issues with custom draw.

'-@ Bugfixes Ver. 2.4
'/~ Italian version (maybe others?) of xp sp2 crashing when loading skinned headers. Rewrote the
'/~ columntext routine with hditem structs instead of lvitem, seems to work.
'/~ Subitem edit box moving out of place if column resized with edit box showing. Changed so column
'/~ size change closes edit window.
'/~ Painting into background picture after column sort. Changed to refreshing list with erase bit after sort, and
'/~ gif with transparent bg can not be used with this feature, (m$ problem).
'/~ Alphabar truncated when horizontal scrollbar in place, adjusted drawalphabar dc len to compensate.
'/~ Centered column text jumping when icon changing in another column or loses focus, adjusted
'/~ logic in clsSkinHeader::ColumnIcon routine.

'-@ Bugfixes Ver. 2.4.2
'/~ Consolidated initlist routine so one call initialization, and internal only access is now possible.
'/~ Fixed SubItem icons, got broken in 2.4.
'/~ Revised itemadd, itemremove, subitemadd routines, and provided an internal methods demo.
'/~ Alphabar not working in report mode without checkboxes, adjusted logic in wndproc.
'/~ Font color change after re-enabling list issue, adjusted enabled property.
'/~ Flicker when multiselecting items, expanded LVN_KEYDOWN switch member in wndproc with conditional redraw,
'/~ works with ansi mode, wide mode still flickery(?).
'/~ Check items with spacebar was not working, adjusted added listrefresh to routine in wndproc.
'/~ Fixed checkbox offset default, and list frame style issues.
'/~ lstrlena returns wrong length in clsskinheader:drawtext under some circumstances, causing centered
'/~ header text to jump around, changed to len in that routine.
'/~ Cleaned up unused vars and declares in all classes, (except xp styles demo).
'/~ Added property browser descriptions and sorted out viewable properties.
'/~ After removeduplicates call, list will get a little glitchy, only happens in customdraw mode
'/~ so I am guessing it is something to do with the release of class members in the m_clistitems array.
'/~ Tried a couple rewrites, but, no go. If you figure it out, let me know..

'-@ Bugfixes Ver. 2.4.3
'/~ Skinned checkbox routines completely rewritten. Now using state image list rather then blit method.
'/~ This makes for a cleaner draw and less overhead. Also, columns with checkboxes can now be moved without
'/~ issue, and checkbox overpaints are fixed in icon view.
'/~ Added state list drawn disabled icons.
'/~ Subitem edit caused ghost on list with bg image when textbox unloaded. Added a conditional refresh to list
'/~ to compensate. I still would not recommend bg image with custom draw though, m$ listview implementation is
'/~ a little buggy.
'/~ List refresh on items changed to invalidate only effected areas, resulting in a substantial reduction in
'/~ list flicker.
'/~ When scrolling on the horizontal, alphabar was not painted to new area. Solved by refreshing list
'/~ when non client area is left-clicked.
'/~ If droping item column out of default position, disabled icons did not render in position. Modified
'/~ icon disabled routine to calculate by item relative position.
'/~ When column dragged with mouse pointer outside of column area, column would sometimes dissapear after drop.
'/~ Added a refresh sub to clsskinheader, called by uc refresh routine to compensate.
'/~ Added a hot column insertion mark to column drag and drop.
'/~ Two small notes: you could use checkbox routines to change checkbox on most ms controls. Also, you
'/~ can use this listview in C# or Vb .Net, simply compile binary, and add it as a com object to your toolbox.
'/~ This will be the last update for a while, hopefully most of the bugs are gone..

'-@ Bugfixes Ver. 2.4.4
'/~ Compile failure when compiling project group. Adjusted control binary dependency setting, (remember to
'/~ set it to 'Binary Compatability' on final compile! ..as per compile instruction above.)
'/~ Edit box stationary on list vertical scroll. Added 'EditBoxMove' routine to compensate.
'/~ InvalidateRect api does not seem to like rect pointer, (client coords are interpreted wrong). Made EraseRect
'/~ api call to compensate.
'/~ Added WS_CLIPCHILDREN to lv parent style for cleaner painting and editbox focus issues.
'/~ Swapped clsSkinHeader class with new one from grid control, (better, more options).
'/~ Added XP style headers to skins.

'-@ Bugfixes Ver. 2.4.5
'/~ Updated the find routine with sortpointers. Updated the a couple skinheader routines. This will be the
'/~ last update of this control.


'-> enjoy

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

Private Const CF_TEXT                           As Long = &H1
Private Const CF_UNICODETEXT                    As Long = &HD
Private Const CF_OEMTEXT                        As Long = &H7

Private Const DST_ICON                          As Long = &H3
Private Const DSS_DISABLED                      As Long = &H20

Private Const EM_REPLACESEL                     As Long = &HC2

Private Const CW_USEDEFAULT                     As Long = &H80000000

Private Const ES_LEFT                           As Long = &H0
Private Const ES_UPPERCASE                      As Long = &H8&
Private Const ES_LOWERCASE                      As Long = &H10


Private Const EM_LIMITTEXT                      As Long = &H415
Private Const EM_GETSEL                         As Long = &HB0
Private Const EM_SETSEL                         As Long = &HB1

Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const GW_HWNDNEXT                       As Long = &H2
Private Const GW_HWNDPREV                       As Long = &H3

Private Const GWL_STYLE                         As Long = (-16)
Private Const GWL_EXSTYLE                       As Long = (-20)
Private Const GWL_HINSTANCE                     As Long = (-6)

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
Private Const HDM_INSERTITEMA                   As Long = (HDM_FIRST + 1)
Private Const HDM_GETITEMA                      As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEMA                      As Long = (HDM_FIRST + 4)
Private Const HDM_GETITEMRECT                   As Long = (HDM_FIRST + 7)
Private Const HDM_SETIMAGELIST                  As Long = (HDM_FIRST + 8)
Private Const HDM_INSERTITEMW                   As Long = (HDM_FIRST + 10)
Private Const HDM_GETITEMW                      As Long = (HDM_FIRST + 11)
Private Const HDM_SETITEMW                      As Long = (HDM_FIRST + 12)
Private Const HDM_SETHOTDIVIDER                 As Long = (HDM_FIRST + 19)

Private Const H_MAX                             As Long = &HFFFF + 1
Private Const HDN_FIRST                         As Long = H_MAX - 300
Private Const HDN_BEGINDRAG                     As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG                       As Long = (HDN_FIRST - 11)
Private Const HDN_FILTERCHANGE                  As Long = (HDN_FIRST - 12)
Private Const HDN_FILTERBTNCLICK                As Long = (HDN_FIRST - 13)
Private Const HDN_ITEMCHANGINGW                 As Long = (HDN_FIRST - 20)
Private Const HDN_ITEMCHANGEDW                  As Long = (HDN_FIRST - 21)
Private Const HDN_ITEMCLICKW                    As Long = (HDN_FIRST - 22)
Private Const HDN_ITEMDBLCLICKW                 As Long = (HDN_FIRST - 23)
Private Const HDN_DIVIDERDBLCLICKW              As Long = (HDN_FIRST - 25)
Private Const HDN_BEGINTRACKW                   As Long = (HDN_FIRST - 26)
Private Const HDN_ENDTRACKW                     As Long = (HDN_FIRST - 27)
Private Const HDN_TRACKW                        As Long = (HDN_FIRST - 28)
Private Const HDN_ITEMCHANGINGA                 As Long = (HDN_FIRST - 0)
Private Const HDN_ITEMCHANGEDA                  As Long = (HDN_FIRST - 1)
Private Const HDN_ITEMCLICKA                    As Long = (HDN_FIRST - 2)
Private Const HDN_ITEMDBLCLICKA                 As Long = (HDN_FIRST - 3)
Private Const HDN_DIVIDERDBLCLICKA              As Long = (HDN_FIRST - 5)
Private Const HDN_BEGINTRACKA                   As Long = (HDN_FIRST - 6)
Private Const HDN_ENDTRACKA                     As Long = (HDN_FIRST - 7)
Private Const HDN_TRACKA                        As Long = (HDN_FIRST - 8)


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

Private Const LOGPIXELSY                        As Long = 90

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

Private Const LVHT_NOWHERE                      As Long = &H1
Private Const LVHT_ONITEMICON                   As Long = &H2
Private Const LVHT_ONITEMLABEL                  As Long = &H4
Private Const LVHT_ONITEMSTATEICON              As Long = &H8
Private Const LVHT_ONITEM                       As Long = _
    (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

Private Const LVHT_ABOVE                        As Long = &H8
Private Const LVHT_BELOW                        As Long = &H10
Private Const LVHT_TORIGHT                      As Long = &H20
Private Const LVHT_TOLEFT                       As Long = &H40

Private Const LVIR_BOUNDS                       As Long = &H0
Private Const LVIR_ICON                         As Long = &H1
Private Const LVIR_LABEL                        As Long = &H2
Private Const LVIR_SELECTBOUNDS                 As Long = &H3

Private Const LVIS_UNCHECKED                    As Long = &H1000&
Private Const LVIS_CHECKED                      As Long = &H2000&
Private Const LVIS_DISABLED                     As Long = &H3000&
Private Const LVIS_CHKCLICK                     As Long = &HFFFE

Private Const LVM_FIRST                         As Long = &H1000
Private Const LVM_GETBKCOLOR                    As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR                    As Long = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST                  As Long = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST                  As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT                  As Long = (LVM_FIRST + 4)
Private Const LVM_ENSUREVISIBLE                 As Long = (LVM_FIRST + 19)
Private Const LVM_REDRAWITEMS                   As Long = (LVM_FIRST + 21)
Private Const LVM_GETEDITCONTROL                As Long = (LVM_FIRST + 24)
Private Const LVM_DELETECOLUMN                  As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH                As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH                As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER                     As Long = (LVM_FIRST + 31)
Private Const LVM_GETTEXTCOLOR                  As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR                  As Long = (LVM_FIRST + 36)
Private Const LVM_SETTEXTBKCOLOR                As Long = (LVM_FIRST + 38)
Private Const LVM_GETCOUNTPERPAGE               As Long = (LVM_FIRST + 40)
Private Const LVM_SETITEMSTATE                  As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE                  As Long = (LVM_FIRST + 44)
Private Const LVM_GETSELECTEDCOUNT              As Long = (LVM_FIRST + 50)
Private Const LVM_SETICONSPACING                As Long = (LVM_FIRST + 53)
Private Const LVM_SUBITEMHITTEST                As Long = (LVM_FIRST + 57)
Private Const LVM_SETCOLUMNORDERARRAY           As Long = (LVM_FIRST + 58)
Private Const LVM_GETCOLUMNORDERARRAY           As Long = (LVM_FIRST + 59)
Private Const LVM_GETSELECTIONMARK              As Long = (LVM_FIRST + 66)
Private Const LVM_SETSELECTIONMARK              As Long = (LVM_FIRST + 67)
Private Const LVM_SETBKIMAGEA                   As Long = (LVM_FIRST + 68)
Private Const LVM_GETBKIMAGEA                   As Long = (LVM_FIRST + 69)
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
Private Const LVM_GETITEMA                      As Long = (LVM_FIRST + 5)
Private Const LVM_SETITEMA                      As Long = (LVM_FIRST + 6)
Private Const LVM_INSERTITEMA                   As Long = (LVM_FIRST + 7)
Private Const LVM_FINDITEMA                     As Long = (LVM_FIRST + 13)
Private Const LVM_EDITLABELA                    As Long = (LVM_FIRST + 23)
Private Const LVM_GETCOLUMNA                    As Long = (LVM_FIRST + 25)
Private Const LVM_SETCOLUMNA                    As Long = (LVM_FIRST + 26)
Private Const LVM_INSERTCOLUMNA                 As Long = (LVM_FIRST + 27)
Private Const LVM_UPDATE                        As Long = (LVM_FIRST + 42)
Private Const LVM_GETITEMTEXTA                  As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMTEXTA                  As Long = (LVM_FIRST + 46)

Private Const LVN_FIRST                         As Long = -100&
Private Const LVN_LAST                          As Long = -199&
Private Const LVN_BEGINLABELEDITA               As Long = (LVN_FIRST - 5)
Private Const LVN_ENDLABELEDITA                 As Long = (LVN_FIRST - 6)
Private Const LVN_SCROLLCHANGE                  As Long = (LVN_FIRST - 7) '<- undocumented
Private Const LVN_GETDISPINFOA                  As Long = (LVN_FIRST - 50)
Private Const LVN_SETDISPINFOA                  As Long = (LVN_FIRST - 51)
Private Const LVN_BEGINLABELEDITW               As Long = (LVN_FIRST - 75)
Private Const LVN_ENDLABELEDITW                 As Long = (LVN_FIRST - 76)
Private Const LVN_GETDISPINFOW                  As Long = (LVN_FIRST - 77)
Private Const LVN_SETDISPINFOW                  As Long = (LVN_FIRST - 78)

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

Private Const MA_ACTIVATE                       As Long = &H1
Private Const MA_ACTIVATEANDEAT                 As Long = &H2
Private Const MA_NOACTIVATE                     As Long = &H3
Private Const MA_NOACTIVATEANDEAT               As Long = &H4

Private Const NM_FIRST                          As Long = H_MAX
Private Const NM_CLICK                          As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK                         As Long = (NM_FIRST - 3)
Private Const NM_RETURN                         As Long = (NM_FIRST - 4)
Private Const NM_RCLICK                         As Long = (NM_FIRST - 5)
Private Const NM_KILLFOCUS                      As Long = (NM_FIRST - 8)
Private Const NM_CUSTOMDRAW                     As Long = (NM_FIRST - 12)

Private Const PRP_APT                           As Long = 130
Private Const PRP_BRDSTL                        As Long = 1
Private Const PRP_CHKSTL                        As Long = 5
Private Const PRP_ITMIND                        As Long = 0
Private Const PRP_LSTMDE                        As Long = 1
Private Const PRP_TMCLR                         As Long = &H9C541F
Private Const PRP_TMLMC                         As Long = 0
Private Const PRP_TXTALN                        As Long = 0
Private Const PRP_VWEMDE                        As Long = 1

Private Const SB_LINEDOWN                       As Long = 1
Private Const SB_LINELEFT                       As Long = 0
Private Const SB_LINERIGHT                      As Long = 1
Private Const SB_LINEUP                         As Long = 0

Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_NOOWNERZORDER                 As Long = &H200

Private Const VER_PLATFORM_WIN32_NT             As Integer = 2

Private Const VK_TAB                            As Long = &H9
Private Const VK_CONTROL                        As Long = &H11
Private Const VK_ESCAPE                         As Long = &H1B
Private Const VK_ENTER                          As Long = &HD

Private Const WC_LISTVIEW                       As String = "SysListView32"

Private Const WM_SETFOCUS                       As Long = &H7
Private Const WM_KILLFOCUS                      As Long = &H8
Private Const WM_SETFONT                        As Long = &H30
Private Const WM_GETFONT                        As Long = &H31
Private Const WM_SETTEXT                        As Long = &HC
Private Const WM_GETTEXT                        As Long = &HD
Private Const WM_GETTEXTLENGTH                  As Long = &HE
Private Const WM_PAINT                          As Long = &HF
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
Private Const WM_USER                           As Long = &H400
Private Const WM_TIMER                          As Long = &H113&
Private Const WM_VSCROLL                        As Long = &H115
Private Const WM_HSCROLL                        As Long = &H114

Private Const WS_OVERLAPPED                     As Long = &H0
Private Const WS_POPUP                          As Long = &H80000000
Private Const WS_CHILD                          As Long = &H40000000
Private Const WS_MINIMIZE                       As Long = &H20000000
Private Const WS_VISIBLE                        As Long = &H10000000
Private Const WS_DISABLED                       As Long = &H8000000
Private Const WS_CLIPSIBLINGS                   As Long = &H4000000
Private Const WS_CLIPCHILDREN                   As Long = &H2000000
Private Const WS_MAXIMIZE                       As Long = &H1000000
Private Const WS_CAPTION                        As Long = &HC00000
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_DLGFRAME                       As Long = &H400000
Private Const WS_VSCROLL                        As Long = &H200000
Private Const WS_HSCROLL                        As Long = &H100000
Private Const WS_SYSMENU                        As Long = &H80000
Private Const WS_THICKFRAME                     As Long = &H40000
Private Const WS_GROUP                          As Long = &H20000
Private Const WS_TABSTOP                        As Long = &H10000
Private Const WS_MINIMIZEBOX                    As Long = &H20000
Private Const WS_MAXIMIZEBOX                    As Long = &H10000
Private Const WS_EX_DLGMODALFRAME               As Long = &H1
Private Const WS_EX_NOPARENTNOTIFY              As Long = &H4
Private Const WS_EX_TOPMOST                     As Long = &H8
Private Const WS_EX_ACCEPTFILES                 As Long = &H10
Private Const WS_EX_TRANSPARENT                 As Long = &H20
Private Const WS_EX_MDICHILD                    As Long = &H40
Private Const WS_EX_TOOLWINDOW                  As Long = &H80
Private Const WS_EX_WINDOWEDGE                  As Long = &H100
Private Const WS_EX_CLIENTEDGE                  As Long = &H200
Private Const WS_EX_CONTEXTHELP                 As Long = &H400
Private Const WS_EX_RIGHT                       As Long = &H1000
Private Const WS_EX_LEFT                        As Long = &H0
Private Const WS_EX_RTLREADING                  As Long = &H2000
Private Const WS_EX_LTRREADING                  As Long = &H0
Private Const WS_EX_LEFTSCROLLBAR               As Long = &H4000
Private Const WS_EX_RIGHTSCROLLBAR              As Long = &H0
Private Const WS_EX_CONTROLPARENT               As Long = &H10000
Private Const WS_EX_STATICEDGE                  As Long = &H20000
Private Const WS_EX_APPWINDOW                   As Long = &H40000

Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = _
(WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Private Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

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

Public Enum ECTTextAlignFlags
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

Public Enum EDCDropConstants
    vbOLEDropNone
    vbOLEDropManual
End Enum

Public Enum ETXAlign
    [AlignLeft] = 0
    [AlignRight] = 1
End Enum

Public Enum EBGBackGroundImage
    [BgNone] = -1
    [BgNormal] = LVBKIF_STYLE_NORMAL
    [BgTile] = LVBKIF_STYLE_TILE
End Enum

Public Enum EBSBorderStyle
    [None] = 0
    [Thin] = 1
    [Thick] = 2
End Enum

Public Enum ECSCheckBoxSkinStyle
    [CheckBoxClassic] = 0
    [CheckBoxEclipse] = 1
    [CheckBoxLime] = 2
    [CheckBoxMetallic] = 3
    [CheckBoxGloss] = 4
    [CheckBoxXP] = 5
End Enum

Public Enum ECAColumnAutosize
    [ColumnItem] = LVSCW_AUTOSIZE
    [ColumnHeader] = LVSCW_AUTOSIZE_USEHEADER
End Enum

Public Enum ECAColumnAlign
    [ColumnLeft] = HDF_LEFT
    [Columnright] = HDF_RIGHT
    [ColumnCenter] = HDF_CENTER
End Enum

Public Enum ECSColumnSortTags
    [SortNone] = -1
    [SortDefault] = 0
    [SortDate] = 1
    [SortNumeric] = 2
    [SortAuto] = 3
End Enum

Public Enum EHSHeaderSkinStyle
    [HeaderClassic] = 0
    [HeaderEclipse] = 1
    [HeaderLime] = 2
    [HeaderMetallic] = 3
    [HeaderGloss] = 4
    [HeaderXP] = 5
End Enum

Public Enum ELMListMode
    [eCustomDraw] = 0
    [eDatabase] = 1
    [eHyperList] = 2
End Enum

Public Enum ELSStyle
    [StyleReport] = LVS_REPORT
    [StyleIcon] = LVS_ICON
    [StyleSmallIcon] = LVS_SMALLICON
    [StyleList] = LVS_LIST
End Enum

Public Enum ERDRowDecoration
    [RowLine] = 0
    [RowSplit] = 1
    [RowBiLinear] = 2
    [RowChecker] = 3
End Enum

Public Enum ESBScrollBarSkinStyle
    [ScrollClassic] = 0
    [ScrollEclipse] = 1
    [ScrollLime] = 2
    [ScrollMetallic] = 3
    [ScrollGloss] = 4
End Enum

Public Enum ESTThemeLuminence
    [ThemeSoft] = 0
    [ThemePastel] = 1
    [ThemeHard] = 2
End Enum


Private Type HLISubItm
    lIcon()                                     As Long
    Text()                                      As String
End Type

Private Type HLIStc
    Item()                                      As String
    lIcon()                                     As Long
    SubItem()                                   As HLISubItm
End Type

Private Type tagINITCOMMONCONTROLSEX
    dwSize                                      As Long
    dwICC                                       As Long
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    right                                       As Long
    bottom                                      As Long
End Type

Private Type POINTAPI
    X                                           As Long
    Y                                           As Long
End Type

Private Type LVHITTESTINFO
    pt                                          As POINTAPI
    flags                                       As Long
    iItem                                       As Long
    iSubItem                                    As Long
End Type

Private Type LVCOLUMN
   Mask                                         As Long
   fmt                                          As Long
   cx                                           As Long
   pszText                                      As Long
   cchTextMax                                   As Long
   iSubItem                                     As Long
   iImage                                       As Long
   iOrder                                       As Long
End Type

Private Type HDITEM
    Mask                                        As Long
    cxy                                         As Long
    pszText                                     As String
    hbm                                         As Long
    cchTextMax                                  As Long
    fmt                                         As Long
    lParam                                      As Long
    iImage                                      As Long
    iOrder                                      As Long
End Type

Private Type HDITEMW
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

Private Type LOGBRUSH
    lbStyle                                     As Long
    lbColor                                     As Long
    lbHatch                                     As Long
End Type

Private Type NMHDR
    hwndFrom                                    As Long
    idfrom                                      As Long
    code                                        As Long
End Type

Private Type NMHEADER
    hdr                                         As NMHDR
    iItem                                       As Long
    iButton                                     As Long
    lPtrHDItem                                  As Long
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

Private Type LVITEM
    Mask                                        As Long
    iItem                                       As Long
    iSubItem                                    As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iImage                                      As Long
    lParam                                      As Long
    iIndent                                     As Long
End Type

Private Type LVITEMW
    Mask                                        As Long
    iItem                                       As Long
    iSubItem                                    As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iImage                                      As Long
    lParam                                      As Long
    iIndent                                     As Long
End Type

Private Type NMLVDISPINFO
    hdr                                         As NMHDR
    Item                                        As LVITEM
End Type

Private Type NMLVDISPINFOW
    hdr                                         As NMHDR
    Item                                        As LVITEMW
End Type

Private Type NMLISTVIEW
    hdr                                         As NMHDR
    iItem                                       As Long
    iSubItem                                    As Long
    uNewState                                   As LV_ITEM_STATE
    uOldState                                   As LV_ITEM_STATE
    uChanged                                    As LV_ITEM_STATE
    ptAction                                    As POINTAPI
    lParam                                      As Long
End Type

Private Type NMLVKEYDOWN
    hdr                                         As NMHDR
    wVKey                                       As Integer
    flags1                                      As Integer
    flags2                                      As Integer
End Type

Private Type LVBKIMAGE
    ulFlags                                     As Long
    hbm                                         As Long
    pszImage                                    As String
    cchImageMax                                 As Long
    xOffsetPercent                              As Long
    yOffsetPercent                              As Long
End Type

Private Type LVBKIMAGEW
    ulFlags                                     As Long
    hbm                                         As Long
    pszImage                                    As Long
    cchImageMax                                 As Long
    xOffsetPercent                              As Long
    yOffsetPercent                              As Long
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

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type


Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    ByVal lParam As Long) As Long

Private Declare Function PostMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    ByVal lParam As Long) As Long

Private Declare Function GetTextExtentPoint32A Lib "gdi32" (ByVal hdc As Long, _
                                                            ByVal lpsz As String, _
                                                            ByVal cbString As Long, _
                                                            lpSize As POINTAPI) As Long

Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As Long, _
                                                            ByVal lpsz As Long, _
                                                            ByVal cbString As Long, _
                                                            lpSize As POINTAPI) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" (lpDest As Any, _
                                                  lpSource As Any) As Long


Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, _
                                                  lpString2 As Any) As Long

Private Declare Function lstrtoptr Lib "kernel32" Alias "lstrcpyA" (ByVal lpDest As Long, _
                                                                    ByVal lpSource As String) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As String, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As Long, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowTextA Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As String) As Long

Private Declare Function SetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As Long) As Long

Private Declare Function GetWindowTextLengthA Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowTextLengthW Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowTextA Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As String, _
                                                          ByVal cch As Long) As Long

Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As Long, _
                                                          ByVal cch As Long) As Long

Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As String, _
                                                       ByVal lpWindowName As String, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal X As Long, _
                                                       ByVal Y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As Long, _
                                                       ByVal lpWindowName As Long, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal X As Long, ByVal Y As Long, _
                                                       ByVal nWidth As Long, ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Function PathCompactPathA Lib "shlwapi.dll" (ByVal hdc As Long, _
                                                             ByVal pszPath As String, _
                                                             ByVal dX As Long) As Long

Private Declare Function PathCompactPathW Lib "shlwapi.dll" (ByVal hdc As Long, _
                                                             ByVal pszPath As Long, _
                                                             ByVal dX As Long) As Long

Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lLongPath As String, _
                                                           ByVal lShortPath As String, _
                                                           ByVal lBuffer As Long) As Long

Private Declare Function GetShortPathNameW Lib "kernel32" (ByVal lLongPath As Long, _
                                                           ByVal lShortPath As Long, _
                                                           ByVal lBuffer As Long) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Private Declare Function CopyStringA Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, _
                                                                      ByVal OldString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_Add Lib "Comctl32" (ByVal hImagelist As Long, _
                                                       ByVal hBitmap As Long, _
                                                       ByVal hBitmapMask As Long) As Long

Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImagelist As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_AddIcon Lib "Comctl32" (ByVal hImagelist As Long, _
                                                           ByVal hIcon As Long) As Long

Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImagelist As Long) As Long

Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, _
                                                               ByVal i As Long, _
                                                               ByVal diIgnore As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function ImageList_DrawEx Lib "Comctl32" (ByVal hIml As Long, _
                                                          ByVal i As Long, _
                                                          ByVal hdcDst As Long, _
                                                          ByVal X As Long, _
                                                          ByVal Y As Long, _
                                                          ByVal dX As Long, _
                                                          ByVal dy As Long, _
                                                          ByVal rgbBk As Long, _
                                                          ByVal rgbFg As Long, _
                                                          ByVal fStyle As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, _
                                                           ByVal HPALETTE As Long, _
                                                           pccolorref As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal fEnable As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Long, _
                                                      ByVal bErase As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

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
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long

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

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal wCmd As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long) As Long

Private Declare Function GetFocus Lib "user32" () As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long) As Long

Private Declare Function EraseRect Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                        lpRect As RECT, _
                                                                        ByVal bErase As Long) As Long


Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal hrgn As Long) As Long

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, _
                                                      ByVal X1 As Long, _
                                                      ByVal Y1 As Long, _
                                                      ByVal X2 As Long, _
                                                      ByVal Y2 As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long


Public Event eHItemClick(ByVal lItem As Long)
Public Event eHDragging(ByVal lItem As Long)
Public Event eHDragComplete(ByVal lSource As Long, ByVal lTarget As Long)
Public Event eHItemCheck(ByVal lItem As Long)
Public Event eHColumnClick(ByVal Column As Long)
Public Event eHIndirect(ByVal iItem As Long, ByVal iSubItem As Long, ByVal fMask As Long, sText As String, hImage As Long)
Public Event eHLabelChange(ByVal iItem As Long, ByVal iSubItem As Long, ByVal sText As String)
Public Event eHErrCond(ByVal sRtn As String, ByVal lErr As Long)


Private m_bteAlphaTransparency                  As Byte
Private m_sngLuminence                          As Single
Private m_sngChkLuminence                       As Single
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
Private m_bSkinScrollBars                       As Boolean
Private m_bUseThemeColors                       As Boolean
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
Private m_bRowDecoration                        As Boolean
Private m_bAlphaBarTheme                        As Boolean
Private m_bAlphaThemeBackClr                    As Boolean
Private m_bItemActive                           As Boolean
Private m_bAlphaIsLoaded                        As Boolean
Private m_bIsNt                                 As Boolean
Private m_bSubItemsEdit                         As Boolean
Private m_bUseCellFont                          As Boolean
Private m_bUseCellColor                         As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bIsXp                                 As Boolean
Private m_bInsensitiveSort                      As Boolean
Private m_bEditFontSet                          As Boolean
Private m_bStopSearch                           As Boolean
Private m_bBackgroundBg                         As Boolean
Private m_bTimerActive                          As Boolean
Private m_bTravelLeft                           As Boolean
Private m_lSafeTimer                            As Long
Private m_lHorzPos                              As Long
Private m_lHotColumn                            As Long
Private m_lColumnOffset                         As Long
Private m_lHdc()                                As Long
Private m_lBmp()                                As Long
Private m_lBmpOld()                             As Long
Private m_lEditItem                             As Long
Private m_lSmallIconX                           As Long
Private m_lSmallIconY                           As Long
Private m_lLargeIconX                           As Long
Private m_lLargeIconY                           As Long
Private m_lSubItemEdit                          As Long
Private m_lEditHwnd                             As Long
Private m_lCheckBoxSkinOffsetX                  As Long
Private m_lColumnHeight                         As Long
Private m_lIconDefaultX                         As Long
Private m_lSelectedItem                         As Long
Private m_lImlStateHndl                         As Long
Private m_lIconSpaceX                           As Long
Private m_lIconSpaceY                           As Long
Private m_lCheckHeight                          As Long
Private m_lCheckWidth                           As Long
Private m_lTmpBackClr                           As Long
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
Private m_lRowColor()                           As Long
Private m_lRowColorBase                         As Long
Private m_lRowColorOffset                       As Long
Private m_lRowDepth                             As Long
Private m_lTmpForeClr                           As Long
Private m_sSortArray()                          As String
Private m_eAlignment                            As ETXAlign
Private m_eListMode                             As ELMListMode
Private m_eHeaderSkinStyle                      As EHSHeaderSkinStyle
Private m_eThemeLuminence                       As ESTThemeLuminence
Private m_eScrollBarSkinStyle                   As ESBScrollBarSkinStyle
Private m_eBorderStyle                          As EBSBorderStyle
Private m_eViewMode                             As ELSStyle
Private m_eSortTag                              As ECSColumnSortTags
Private m_eCheckBoxSkinStyle                    As ECSCheckBoxSkinStyle
Private m_eRowDecoration                        As ERDRowDecoration
Private m_eOLEDragMode                          As OLEDragConstants
Private m_oBackColor                            As OLE_COLOR
Private m_oForeColor                            As OLE_COLOR
Private m_oHdrBkClr                             As OLE_COLOR
Private m_oHdrForeClr                           As OLE_COLOR
Private m_oHdrHighLiteClr                       As OLE_COLOR
Private m_oHdrPressedClr                        As OLE_COLOR
Private m_oThemeColor                           As OLE_COLOR
Private c_ColumnTags                            As Collection
Private c_PtrMem                                As Collection
Private m_oFont                                 As StdFont
Private m_pISelectorBar                         As StdPicture
Private m_IChecked                              As StdPicture
Private m_IChkDisabled                          As StdPicture
Private m_IUnChecked                            As StdPicture
Private m_IDivider                              As StdPicture
Private m_tRStr                                 As RECT
Private m_cSelectorBar                          As clsStoreDc
Private m_cChkCheckDc                           As clsStoreDc
Private m_cChkUnCheckDc                         As clsStoreDc
Private m_cChkDisableDc                         As clsStoreDc
Private m_cDivider                              As clsStoreDc
Private m_cSkinHeader                           As clsSkinHeader
Private m_cXPHeader                             As clsXPHeader
Private m_HLIStc()                              As HLIStc
Private m_cListItems()                          As clsListItem
Private m_cRender                               As clsRender
Private m_cDrag                                 As clsImageDrag
Private m_cSkinScrollBars                       As clsSkinScrollbars
Attribute m_cSkinScrollBars.VB_VarHelpID = -1
Private m_cHListSubclass                        As GXMSubclass



'/~ 2.1 todo ~
'/~ bugfix                          - done
'/~ custom mode                     - done
'/~ hyper mode                      - done
'/~ database mode                   - done
'/~ edit controls                   - done
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
'/~ create uc version               - done
'/~ example project                 - done


'/~ 2.2 todo ~
'/~ bugfix                          - done
'/~ fix bg image                    - done
'/~ item drag and drop              - done
'/~ subitem edit                    - done
'/~ expand events                   - done
'/~ min column size                 - done
'/~ expand error tracking           - done


'/~ 2.3 todo ~
'/~ bugfix                          - done
    '-> item focus backcolor        - todo
    '-> selection flicker           - done
    '-> report item indent          - done
    '-> 1st click sort              - done
'/~ optional unicode                - done
'/~ editbox highlite                - done
'/~ editbox hot keys                - done
'/~ icon view checkbox hit test     - done
'/~ skin checkbox in all views      - done
'/~ cell fonts                      - done
'/~ xp header theme support         - done
'/~ skin header text alignment      - done


'/~ 2.4 todo ~
'/~ bugfix                          - done
    '-> checkbox flicker            - done
    '-> bg focus color              - done
    '-> unicode subitem edit        - done
    '-> icon to report repaint      - done
    '-> setfocus issues             - done
    '-> alphabar overpaint          - done
    '-> text ellipses               - done
    '-> sort item ordering          - done
    '-> leak check                  - done
'/~ finish unicode                  - done
'/~ right align text                - done
'/~ find function                   - done


'/~ 2.4.2 todo ~
'/~ bugfix
    '-> subitem icons fix           - done
    '-> subitem edit font           - done
    '-> fix scrollbars example      - done
    '-> fix checkbox offset prop    - done
    '-> fix frame style init        - done
    '-> clean up enums              - done
    '-> operand alignment           - done
    '-> fix prop descriptions       - done
    '-> class/var cleanup           - done
    '-> byref/byval cleanup         - done
'/~ manual access mode              - done
'/~ methods demo                    - done


'/~ 2.4.3 todo ~
'/~ bugfix
    '-> subitem edit with bg pic    - done
    '-> skin checkbox header move   - done
    '-> checkboxes w/ state iml     - done
    '-> alpha on horz scroll        - done
    '-> limit refresh area          - done
'/~ insertion mark                  - done


'/~ 2.4.4 todo ~
'/~ bugfix
    '-> subitem edit scroll move    - done
    '-> mousewheel issue            - done
    '-> list/item refresh method    - done
    '-> example compile change      - done
    '-> alpha scroll/hdr reset      - done
'/~ row drop example                - done
'/~ xp header skin style            - done

'**********************************************************************
'*                              OPERATION
'**********************************************************************

Private Sub UserControl_Initialize()
'/* init control

'Dim i As Long
    
    '/* operand alignment test:
    '/* important* you should group vars together by type
    '/* and test operand alignment. aligning variables gave
    '/* an increase in speed to project of about 17%.
    '/* On cd mode item load was .017 down to 0.13 +/- avg.
    '/* On bool, int, long, should return 0 or 4,
    '/* on doubles return 0. If they are out of alignment
    '/* add a variable to pad to starting bit boundary.
    '/* Test on first var of each type.
    
    ' i = VarPtr(m_lEditItem) Mod 8
    ' If i Then
    '     Debug.Print "Misaligned by " & i
    ' End If

    m_lhMod = LoadLibrary("shell32.dll")
    InitComctl32
    VersionCheck
    Set m_cHListSubclass = New GXMSubclass
    m_oHdrBkClr = GetSysColor(m_oHdrBkClr And &H1F&)
    m_oHdrForeClr = GetSysColor(vbWindowText And &H1F&)
    m_oBackColor = &HFFFFFF
    m_oThemeColor = -1
    m_lTmpForeClr = &H0
    m_lTmpBackClr = -1
    m_bEnabled = True
    m_eCheckBoxSkinStyle = CheckBoxXP
    BorderStyle = Thin
    
    Set c_ColumnTags = New Collection
    Set m_oFont = New StdFont
    Set m_cDrag = New clsImageDrag

End Sub

Private Function VersionCheck() As Boolean

Dim tVer    As OSVERSIONINFO

    tVer.dwVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    m_bIsNt = ((tVer.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    If (tVer.dwMajorVersion >= 5) Then
        m_bIsXp = True
    End If
    If Not m_bIsNt Then
        m_bUseUnicode = False
    End If
    VersionCheck = m_bIsNt

End Function

Public Property Get IsWinNT() As Boolean
Attribute IsWinNT.VB_MemberFlags = "400"
    IsWinNT = m_bIsNt
End Property

Public Property Get IsWinXP() As Boolean
Attribute IsWinXP.VB_MemberFlags = "400"
    IsWinXP = m_bIsXp
End Property

Public Function CreateList() As Boolean

'*/ initialize the listview

Dim lLVStyle    As Long
Dim tRect       As RECT

On Error GoTo Handler

    '/* destroy existing
    DestroyList
    m_lParentHwnd = UserControl.hwnd
    GetClientRect m_lParentHwnd, tRect
    
    '/* initial style flags including LVS_OWNERDATA
    '/* this tells the list that all data will be
    '/* managed externally
    lLVStyle = WS_CHILD Or WS_BORDER Or WS_VISIBLE Or LVS_SORTASCENDING Or LVS_OWNERDATA Or _
               LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_SINGLESEL Or WS_TABSTOP Or LVS_REPORT Or WS_CLIPCHILDREN
    '/* create listview
    If m_bIsNt Then
        With tRect
            m_lLVHwnd = CreateWindowExW(0&, StrPtr(WC_LISTVIEW), StrPtr(""), lLVStyle, _
                0&, 0&, (.right - .left), (.bottom - .top), m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
        End With
    Else
        With tRect
            m_lLVHwnd = CreateWindowExA(0&, WC_LISTVIEW, vbNullString, lLVStyle, _
                0&, 0&, .right - .left, .bottom - .top, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
        End With
    End If

    lLVStyle = WS_CHILD Or WS_BORDER Or WS_CLIPSIBLINGS
    '/* create the edit box
    If m_bIsNt Then
        m_lEditHwnd = CreateWindowExW(0&, StrPtr("edit"), StrPtr(""), _
            lLVStyle, 0&, 0&, 0&, 0&, m_lLVHwnd, 0&, App.hInstance, ByVal 0&)
    Else
        m_lEditHwnd = CreateWindowExA(0&, "edit", "", _
            lLVStyle, 0&, 0&, 0&, 0&, m_lLVHwnd, 0&, App.hInstance, ByVal 0&)
    End If

    SetUnicode True
    '/* default border style
    SetBorderStyle m_lLVHwnd, None
    '/* subclass the list and parent WM_NOTIFY messages
    '/* control callback data is reflected from parent control
    If Not m_lLVHwnd = 0 Then
        ListAttatch
    End If
    
    '/* default icon spacing
    m_lIconDefaultX = IconSpaceX

    If Not m_lEditHwnd = 0 Then
        EditSetFont
        m_bSubItemsEdit = False
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("CreateList", Err.Number)

End Function

Private Function SetUnicode(ByVal bEnable As Boolean) As Boolean

Dim lRet As Long

    If m_lLVHwnd = 0 Then Exit Function
    If m_bIsNt Then
        If bEnable Then
            If Not UnicodeState Then
                lRet = SendMessageLongW(m_lLVHwnd, CCM_SETUNICODEFORMAT, 1&, 0&)
            End If
        Else
            If UnicodeState Then
                lRet = SendMessageLongW(m_lLVHwnd, CCM_SETUNICODEFORMAT, 0&, 0&)
            End If
        End If
    End If
    SetUnicode = (lRet = 0)
    
End Function

Private Function UnicodeState() As Boolean

    If m_lLVHwnd = 0 Then Exit Function
    UnicodeState = SendMessageLongW(m_lLVHwnd, CCM_GETUNICODEFORMAT, 0&, 0&) <> 0
    
End Function

Private Function InitComctl32() As Boolean
'/* init comctl32 listview class

Dim icc As tagINITCOMMONCONTROLSEX

On Error GoTo Handler
  
    icc.dwSize = Len(icc)
    icc.dwICC = ICC_LISTVIEW_CLASSES
    InitComctl32 = InitCommonControlsEx(icc)
    If Not InitComctl32 Then GoTo Handler
    
On Error GoTo 0
Exit Function

Handler:
    InitCommonControls
    
End Function

Public Sub InitList(ByVal lCount As Long, _
                    Optional ByVal lSubItemCt As Long = -1)

'/* internally init item arrays

Dim lCt As Long

On Error GoTo Handler

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
        ReDim m_HLIStc(0)
        ReDim m_HLIStc(0).Item(0 To lCount)
        ReDim m_HLIStc(0).lIcon(0 To lCount)
        ReDim m_HLIStc(0).SubItem(0 To lCount)
        For lCt = 0 To lCount
            If Not lSubItemCt = -1 Then
                ReDim m_HLIStc(0).SubItem(lCt).Text(1 To lSubItemCt)
            End If
        Next lCt
    End Select

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("InitList", Err.Number)

End Sub

Public Property Get ListMode() As ELMListMode
Attribute ListMode.VB_Description = "[enum] list data input mode"
'*/ retrieve pointer to the data structure
    ListMode = m_eListMode
End Property

Public Property Let ListMode(ByVal PropVal As ELMListMode)
'*/ add pointer to the data structure
    m_eListMode = PropVal
    PropertyChanged "ListMode"
End Property

Private Function PointerToString(ByVal lpString As Long) As String
'/* get string from pointer

Dim lLen As Long

On Error GoTo Handler

    If m_bIsNt Then
        If lpString Then
            If m_bUseUnicode Then
                lLen = lstrlenW(lpString)
            Else
                lLen = lstrlenA(ByVal lpString)
            End If
            If lLen Then
                '_/* allocate string with lLen chars
                PointerToString = String$(lLen, Chr$(0))
                If m_bUseUnicode Then
                    lstrcpyW ByVal StrPtr(PointerToString), ByVal lpString
                Else
                    lstrcpyA ByVal StrPtr(PointerToString), ByVal lpString
                End If
            End If
        End If
    Else
        Dim b()  As Byte
        If lpString Then
            lLen = lstrlenA(ByVal lpString)
            If lLen Then
                '_/* allocate buffer with lLen bytes
                ReDim b(0 To lLen - 1) As Byte
                '_/* copy lLen bytes for ANSI
                CopyMemory b(0), ByVal lpString, lLen
                PointerToString = StrConv(b(), vbUnicode)
            End If
        End If
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("PointerToString", Err.Number)

End Function

Private Sub StringToPointer(ByVal sText As String, _
                            ByRef lpString As Long)

'/* get pointer from string

On Error GoTo Handler

    If m_bUseUnicode Then
        '/* Direct copy address of an Unicode string into a pointer on NT system
        lstrcpyW ByVal lpString, ByVal StrPtr(sText)
    Else
        Dim b()  As Byte
        b = StrConv(sText, vbFromUnicode)
        CopyMemory ByVal lpString, b(0), UBound(b) + 1
    End If

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("StringToPointer", Err.Number)

End Sub

Public Function IsUnicode(ByVal sText As String) As Boolean
'/* good link: http://www.unicodeactivex.com/UnicodeTutorialVb.htm

Dim iLen    As Long
Dim bLen    As Long
Dim bMap()  As Byte

On Error GoTo Handler

    If LenB(sText) Then
        bMap = sText
        bLen = UBound(bMap)
        For iLen = 1 To bLen Step 2
            If (bMap(iLen) > 0) Then
                IsUnicode = True
                Exit For
            End If
        Next
    End If

Handler:

End Function

Public Property Get TextAlignment() As ETXAlign
Attribute TextAlignment.VB_Description = "[enum] listview text alignment"
Attribute TextAlignment.VB_MemberFlags = "1004"
    TextAlignment = m_eAlignment
End Property

Public Property Let TextAlignment(ByVal PropVal As ETXAlign)

Dim lStyle      As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    If m_bUseUnicode Then
        lStyle = GetWindowLongW(m_lLVHwnd, GWL_EXSTYLE)
    Else
        lStyle = GetWindowLongA(m_lLVHwnd, GWL_EXSTYLE)
    End If
    
    If PropVal = AlignLeft Then
        lStyle = lStyle And Not WS_EX_RTLREADING
        lStyle = lStyle Or WS_EX_LTRREADING
    ElseIf PropVal = AlignRight Then
        lStyle = lStyle And Not WS_EX_LTRREADING
        lStyle = lStyle Or WS_EX_RTLREADING
    End If
    
    If m_bUseUnicode Then
        SetWindowLongW m_lLVHwnd, GWL_EXSTYLE, lStyle
    Else
        SetWindowLongA m_lLVHwnd, GWL_EXSTYLE, lStyle
    End If

    m_eAlignment = PropVal
    PropertyChanged "TextAlignment"


On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("TextAlignment", Err.Number)

End Property

Public Property Get StructPtr() As Long
Attribute StructPtr.VB_MemberFlags = "400"
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
                                Optional ByVal hPal As Long = 0) As Long

'/* translate to ole color

    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Public Property Get UseUnicode() As Boolean
Attribute UseUnicode.VB_Description = "[bool] enable unicode processing"
    UseUnicode = m_bUseUnicode
End Property

Public Property Let UseUnicode(PropVal As Boolean)

    If m_bIsNt Then
        If PropVal Then
            SetUnicode True
            EditSetFont
            m_bUseUnicode = True
        Else
            SetUnicode False
            EditSetFont
            m_bUseUnicode = False
        End If
        ListRefresh True
    End If
    PropertyChanged "UseUnicode"
    
End Property

Public Property Get Focus() As Boolean
Attribute Focus.VB_MemberFlags = "400"

    Focus = (GetFocus() = m_lLVHwnd)

End Property

Public Property Let Focus(ByVal bTrue As Boolean)

Dim lMsg As Long

    If bTrue Then
        lMsg = WM_SETFOCUS
    Else
        lMsg = WM_KILLFOCUS
    End If
    If m_bUseUnicode Then
        PostMessageW m_lLVHwnd, lMsg, 0&, 0&
    Else
        PostMessageA m_lLVHwnd, lMsg, 0&, 0&
    End If
    
End Property

Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "[enum] list item ole drag mode"
    OLEDragMode = m_eOLEDragMode
End Property

Public Property Let OLEDragMode(ByVal PropVal As OLEDragConstants)
    m_eOLEDragMode = PropVal
    PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDropMode() As EDCDropConstants
Attribute OLEDropMode.VB_Description = "[enum] list item ole drag mode"
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal PropVal As EDCDropConstants)
    UserControl.OLEDropMode = PropVal
    PropertyChanged "OLEDropMode"
End Property

Public Sub Refresh()

    If m_bUseUnicode Then
        SendMessageLongW m_lLVHwnd, LVM_UPDATE, 0&, 0&
    Else
        SendMessageLongA m_lLVHwnd, LVM_UPDATE, 0&, 0&
    End If
    
End Sub

Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "[enum] listview scale mode"
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal eMode As ScaleModeConstants)
    UserControl.ScaleMode = eMode
    PropertyChanged "ScaleMode"
End Property

Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_MemberFlags = "400"
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_MemberFlags = "400"
    ScaleHeight = UserControl.ScaleHeight
End Property

'**********************************************************************
'*                              SKINNING
'**********************************************************************

Public Property Get AlphaBarTheme() As Boolean
Attribute AlphaBarTheme.VB_Description = "[bool] use theme colors on alphabar"
'/* [get] use theme color
    AlphaBarTheme = m_bAlphaBarTheme
End Property

Public Property Let AlphaBarTheme(ByVal PropVal As Boolean)
'/* [let] use theme color
    m_bAlphaBarTheme = PropVal
    ListRefresh
    PropertyChanged "AlphaBarTheme"
End Property

Public Property Get AlphaBarTransparency() As Byte
Attribute AlphaBarTransparency.VB_Description = "[byte] alpha bar transparency index"
'/* [get] alpha bar transparency index
    AlphaBarTransparency = m_bteAlphaTransparency
End Property

Public Property Let AlphaBarTransparency(ByVal PropVal As Byte)
'/* [let] alpha bar transparency index

    If PropVal < 70 Then
        m_bteAlphaTransparency = 70
    ElseIf PropVal > 240 Then
        m_bteAlphaTransparency = 200
    Else
        m_bteAlphaTransparency = PropVal
    End If
    PropertyChanged "AlphaBarTransparency"
    
End Property

Public Property Get AlphaThemeBackClr() As Boolean
Attribute AlphaThemeBackClr.VB_Description = "[bool] Use alphabar themed backcolor"
'/* [get] alpha bar themed backcolor
    AlphaThemeBackClr = m_bAlphaThemeBackClr
End Property

Public Property Let AlphaThemeBackClr(ByVal PropVal As Boolean)
'/* [let] alpha bar themed backcolor
    m_bAlphaThemeBackClr = PropVal
    PropertyChanged "AlphaThemeBackClr"
End Property

Public Property Get AlphaBarActive() As Boolean
Attribute AlphaBarActive.VB_Description = "[bool] use alpha bar"
Attribute AlphaBarActive.VB_MemberFlags = "4"
'/* [get] alpha bar state
    If m_bAlphaIsLoaded Then
        AlphaBarActive = m_bAlphaSelectorBar
    End If
End Property

Public Property Let AlphaBarActive(ByVal PropVal As Boolean)
'/* [let] alpha bar state
    
    If PropVal And Not m_bAlphaIsLoaded Then
        If AlphaBarTransparency = 0 Then
            AlphaBarTransparency = 100
        End If
        AlphaSelectorBar AlphaBarTransparency, m_bAlphaBarTheme, m_bAlphaThemeBackClr
    End If
    m_bAlphaSelectorBar = PropVal
    ListRefresh
    PropertyChanged "AlphaBarActive"
    
End Property

Public Function AlphaSelectorBar(ByVal btTransparency As Byte, _
                                 ByVal bUseThemeColor As Boolean, _
                                 ByVal bUseAlphaBackcolor As Boolean) As Boolean
'/* create alpha selector

On Error GoTo Handler

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
    m_bAlphaIsLoaded = True
    AlphaSelectorBar = True

On Error GoTo 0
Exit Function

Handler:

End Function

Private Function ResetSelectorBar() As Boolean

    If Not m_cSelectorBar Is Nothing Then
        Set m_cSelectorBar = Nothing
    End If
    Set ISelectorBar = Nothing
    
End Function

Private Sub DrawAlphaSelectorBar(ByVal lItem As Long)
'/* paint alpha bar

Dim lDrawDc     As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long
Dim lSrcDc      As Long
Dim tTmp        As RECT
Dim tRect       As RECT

On Error GoTo Handler

    GetItemRect lItem, tRect
    If m_eViewMode = StyleReport Then
        If Not m_bFullRowSelect Then
            tRect.right = ColumnWidth(0)
        ElseIf m_lColumnOffset > 0 Then
            tRect.left = m_lColumnOffset
            tRect.right = ColumnWidth(0) + m_lColumnOffset
        End If
    End If

    lSrcDc = GetDC(m_lLVHwnd)
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lDrawDc = CreateCompatibleDC(lTmpDc)
    
    LSet tTmp = tRect
    '/* store the rect
    LSet m_tRStr = tRect
    With tTmp
        OffsetRect tTmp, -.left, -.top
        lBmp = CreateCompatibleBitmap(lTmpDc, .right, .bottom)
    End With
    lBmpOld = SelectObject(lDrawDc, lBmp)

    '/* repaint
    With tTmp
        If m_bCheckBoxes And ((m_eViewMode = StyleReport) Or (m_eViewMode = StyleList)) Then
            tRect.left = (tRect.left + 18)
            .right = (.right - 18)
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

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("DrawAlphaSelectorBar", Err.Number)

End Sub

Private Property Get ISelectorBar() As StdPicture
'/* selector bar image
    Set ISelectorBar = m_pISelectorBar
End Property

Private Property Set ISelectorBar(ByVal PropVal As StdPicture)
    Set m_pISelectorBar = PropVal
End Property

Public Function SkinCheckBox(ByVal eCheckBoxStyle As ECSCheckBoxSkinStyle, _
                             ByVal bUseThemeColors As Boolean) As Boolean

'/* skin the listview checkboxes

Dim lMask   As Long

On Error GoTo Handler

    '/* initialize checkboxes
    If Not m_bCheckBoxes Then
        CheckBoxes = True
    End If
    '/* initialize imagelist
    If Not InitImlState Then
        GoTo Handler
    End If
    m_eCheckBoxSkinStyle = eCheckBoxStyle
    m_bUseCheckBoxTheme = bUseThemeColors
    '/* system image sizes
    CheckBoxMetrics

    '/* load images
    If LoadCheckBoxImages Then
        '/* image dc's
        Set m_cChkCheckDc = New clsStoreDc
        Set m_cChkUnCheckDc = New clsStoreDc
        Set m_cChkDisableDc = New clsStoreDc
        '/* create dc's
        If m_bUseCheckBoxTheme Then
            With m_cChkUnCheckDc
                '/* create image dc
                .CreateFromPicture m_IUnChecked
                '/* colorize
                .ColorizeImage m_oThemeColor, m_sngLuminence
                '/* new mask color
                lMask = GetMask(.hdc)
                '/* extract bitmap handle
                ImlStateAddBmp .ExtractBitmap, lMask
            End With
            With m_cChkCheckDc
                .CreateFromPicture m_IChecked
                .ColorizeImage m_oThemeColor, m_sngLuminence
                ImlStateAddBmp .ExtractBitmap, lMask
            End With
            With m_cChkDisableDc
                .CreateFromPicture m_IChkDisabled
                .ColorizeImage m_oThemeColor, m_sngLuminence
                ImlStateAddBmp .ExtractBitmap, lMask
            End With
            Set m_cChkCheckDc = Nothing
            Set m_cChkUnCheckDc = Nothing
            Set m_cChkDisableDc = Nothing
        Else
            ImlStateAddBmp m_IUnChecked.Handle, &HFF00FF
            ImlStateAddBmp m_IChecked.Handle, &HFF00FF
            ImlStateAddBmp m_IChkDisabled.Handle, &HFF00FF
        End If
    Else
        GoTo Handler
    End If

    '/* success
    m_bSkinnedCheck = True
    SkinCheckBox = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SkinCheckBox", Err.Number)

End Function

Private Function GetMask(ByVal lHdc As Long) As Long

    GetMask = GetPixel(lHdc, 0, 0)

End Function

Public Property Get CheckBoxSkinStyle() As ECSCheckBoxSkinStyle
Attribute CheckBoxSkinStyle.VB_Description = "[enum] skinned checkbox style"
'*/ return the checkbox style
    CheckBoxSkinStyle = m_eCheckBoxSkinStyle
End Property

Public Property Let CheckBoxSkinStyle(ByVal PropVal As ECSCheckBoxSkinStyle)
'*/ change the checkbox style

    If Not m_lLVHwnd = 0 Then
        If Not m_eCheckBoxSkinStyle = PropVal Then
            If ResetSkinnedCheckboxes Then
                SkinCheckBox PropVal, m_bUseCheckBoxTheme
                ListRefresh
            End If
        End If
    End If
    m_eCheckBoxSkinStyle = PropVal
    PropertyChanged "CheckBoxSkinStyle"
    
End Property

Private Function ResetSkinnedCheckboxes() As Boolean

On Error GoTo Handler

    m_bSkinnedCheck = False
    DestroyImlState
    If Not m_IChecked Is Nothing Then Set m_IChecked = Nothing
    If Not m_IUnChecked Is Nothing Then Set m_IUnChecked = Nothing
    If Not m_IChkDisabled Is Nothing Then Set m_IChkDisabled = Nothing
    InitImlState
    
    '/* success
    ResetSkinnedCheckboxes = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ResetSkinnedCheckboxes", Err.Number)

End Function

Private Function LoadCheckBoxImages() As Boolean
'/* load checkbox skin images

On Error GoTo Handler

    ResetSkinnedCheckboxes
    
    Select Case m_eCheckBoxSkinStyle
    '/* classic
    Case 0
        Set m_IChecked = LoadResPicture("CLASSIC-CHECKED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("CLASSIC-UNCHECKED", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("CLASSIC-CHKDISABLED", vbResBitmap)
    '/* eclipse
    Case 1
        Set m_IChecked = LoadResPicture("ECLIPSE-CHECKED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("ECLIPSE-UNCHECKED", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("ECLIPSE-CHKDISABLED", vbResBitmap)
    '/* lime
    Case 2
        Set m_IChecked = LoadResPicture("LIME-CHECKED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("LIME-UNCHECKED", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("LIME-CHKDISABLED", vbResBitmap)
    '/* metallic
    Case 3
        Set m_IChecked = LoadResPicture("METALLIC-CHECKED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("METALLIC-UNCHECKED", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("METALLIC-CHKDISABLED", vbResBitmap)
    '/* Gloss
    Case 4
        Set m_IChecked = LoadResPicture("GLOSS-CHECKED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("GLOSS-UNCHECKED", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("GLOSS-CHKDISABLED", vbResBitmap)
    '/* xp
    Case 5
        Set m_IChecked = LoadResPicture("XP-CHECKED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("XP-UNCHECKED", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("XP-CHKDISABLED", vbResBitmap)
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
    RaiseEvent eHErrCond("LoadCheckBoxImages", Err.Number)

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

Public Function SkinXPHeader() As Boolean
'/* xp header demo

On Error GoTo Handler

    If m_bIsXp Then
        If Not m_cXPHeader Is Nothing Then
            Set m_cXPHeader = Nothing
        End If
        Set m_cXPHeader = New clsXPHeader
        With m_cXPHeader
            .UseUnicode = m_bUseUnicode
            .LoadXpSkin m_lLVHwnd, UserControl.Parent.hwnd
        End With
    End If
    SkinXPHeader = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SkinXPHeader", Err.Number)

End Function

Public Function UnSkinXPHeader() As Boolean
'/* unload xp header style

On Error GoTo Handler

    If m_bIsXp Then
        If Not m_cXPHeader Is Nothing Then
            With m_cXPHeader
                .UnLoadXpSkin
            End With
            Set m_cXPHeader = Nothing
        End If
    End If
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SkinXPHeader", Err.Number)

End Function

Public Function SkinHeaders(ByVal eSkinStyle As EHSHeaderSkinStyle, _
                            ByVal oFontForecolor As OLE_COLOR, _
                            ByVal oFontHighliteColor As OLE_COLOR, _
                            ByVal oFontPressedColor As OLE_COLOR, _
                            ByVal bUseThemeColors As Boolean) As Boolean

'/* use header skin
On Error GoTo Handler

    m_eHeaderSkinStyle = eSkinStyle
    m_oHdrForeClr = oFontForecolor
    m_oHdrHighLiteClr = oFontHighliteColor
    m_oHdrPressedClr = oFontPressedColor
    m_bUseThemeColors = bUseThemeColors
    
    '/* divider
    If Not m_cDivider Is Nothing Then
        Set m_cDivider = Nothing
        Set m_IDivider = Nothing
    End If
    Set m_cDivider = New clsStoreDc
    Set m_IDivider = LoadResPicture("DIVIDER", vbResBitmap)
    m_cDivider.CreateFromPicture m_IDivider
    
    '/* skin params
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
        .UseUnicode = m_bUseUnicode
        .SetFont m_oFont
        .UseHeaderTheme = m_bUseThemeColors
        .LoadSkin m_lLVHwnd
        m_bSkinHeader = True
    End With

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SkinHeaders", Err.Number)

End Function

Public Function UnSkinHeaders()
'/* reset skinned header

    Set m_cSkinHeader = Nothing
    ListRefresh True

End Function

Public Function SkinScrollBars(ByVal eSkinStyle As ESBScrollBarSkinStyle, _
                               ByVal bUseThemeColors As Boolean)
'/* skin scrollbars

On Error GoTo Handler

    If Not m_cSkinScrollBars Is Nothing Then
        m_cSkinScrollBars.ResetScrollBarSkin
    Else
        Set m_cSkinScrollBars = New clsSkinScrollbars
    End If
    
    m_eScrollBarSkinStyle = eSkinStyle
    bUseThemeColors = m_bUseThemeColors
    With m_cSkinScrollBars
        .ScrollBarSkinStyle = m_eScrollBarSkinStyle
        .ScrollLuminence = m_eThemeLuminence
        .ScrollThemeColor = m_oThemeColor
        .SkinScrollBar = True
        .UseScrollBarTheme = m_bUseThemeColors
        .LoadSkin m_lLVHwnd, m_lParentHwnd
    End With
    m_bSkinScrollBars = True
    ListRefresh
    Resize
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SkinScrollBars", Err.Number)

End Function

Public Function UnSkinScrollBars()
'/* unskin scrollbars

On Error GoTo Handler

    Set m_cSkinScrollBars = Nothing
    m_bSkinScrollBars = False
    ListRefresh

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("UnSkinScrollBars", Err.Number)

End Function

Public Property Get ThemeColor() As OLE_COLOR
Attribute ThemeColor.VB_Description = "[ole color] skin theme color base"
'/* get theme color
    ThemeColor = m_oThemeColor
End Property

Public Property Let ThemeColor(ByVal PropVal As OLE_COLOR)
'/* set theme color

    m_oThemeColor = PropVal
    PropertyChanged "ThemeColor"
    
End Property

Public Property Get UseThemeColors() As Boolean
Attribute UseThemeColors.VB_Description = "[bool] apply theme colors to skin elements"
'/* get theme status
    UseThemeColors = m_bUseThemeColors
End Property

Public Property Let UseThemeColors(ByVal PropVal As Boolean)
'/* set theme option

    m_bUseThemeColors = PropVal
    m_bUseCheckBoxTheme = PropVal
    PropertyChanged "UseThemeColors"
    
End Property

Public Property Get ThemeLuminence() As ESTThemeLuminence
Attribute ThemeLuminence.VB_Description = "[enum]  skin theme luminence offset"
'/* get theme luminence
    ThemeLuminence = m_eThemeLuminence
End Property

Public Property Let ThemeLuminence(ByVal PropVal As ESTThemeLuminence)
'/* set theme luminence

    m_eThemeLuminence = PropVal

    Select Case PropVal
    Case 0
        m_sngLuminence = 0.2
    Case 1
        m_sngLuminence = 0.4
    Case 2
        m_sngLuminence = 0.7
    End Select
    PropertyChanged "ThemeLuminence"
    
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
'/* save items list to file -hypermode only

Dim FF      As Integer
Dim sTemp   As String

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
    ElseIf Not ArrayCheck(m_HLIStc(0).Item) Then
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
 
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("SaveToFile", Err.Number)
    Close #FF
    On Error GoTo 0

End Function

Public Function LoadFromFile(ByVal sPath As String) As Boolean
'/* load items list from file

Dim FF      As Integer
Dim lRows   As Long
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

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("LoadFromFile", Err.Number)
    On Error Resume Next
    Close #FF
    On Error GoTo 0
    
End Function

Public Function CopyItemToClipboard()
'/* copy selected item to clipboard

Dim sTemp As String

On Error GoTo Handler

    sTemp = ItemText(m_lSelectedItem)
    Clipboard.Clear
    Clipboard.SetText sTemp

Handler:
    RaiseEvent eHErrCond("CopyItemToClipboard", Err.Number)
    On Error GoTo 0

End Function

Private Function FileExists(ByVal sPath As String) As Boolean
'/* test file or path

Dim lRes    As Long
Dim sTemp   As String

    sTemp = String$(254, Chr$(0))
    If m_bUseUnicode Then
        lRes = GetShortPathNameW(StrPtr(sPath), StrPtr(sTemp), 255)
    Else
        lRes = GetShortPathNameA(sPath, sTemp, 255)
    End If
    FileExists = lRes > 0

End Function


'**********************************************************************
'*                              COLUMNS
'**********************************************************************

Public Function ColumnAdd(ByVal lIndex As Long, _
                          ByVal sText As String, _
                          ByVal lWidth As Long, _
                          Optional ByVal eAlign As ECAColumnAlign = [ColumnLeft], _
                          Optional ByVal lIcon As Long = -1, _
                          Optional ByVal ColumnTag As ECSColumnSortTags = SortDefault) As Boolean

'*/ create column headers

Dim bFirst  As Boolean
Dim uLVC    As LVCOLUMN
Dim uHDI    As HDITEM
Dim uHDW    As HDITEMW

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    bFirst = (Me.ColumnCount = 0)
    If c_ColumnTags Is Nothing Then
        Set c_ColumnTags = New Collection
    End If
    '/* unicode nt
    With uLVC
        .pszText = StrPtr(sText)
        .cchTextMax = Len(sText)
        .cx = lWidth
        .fmt = eAlign
        .Mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_FMT
    End With
        
    If m_bUseUnicode Then
        ColumnAdd = (SendMessageW(m_lLVHwnd, LVM_INSERTCOLUMNW, lIndex, uLVC) > -1)
    '/* no unicode
    Else
        ColumnAdd = (SendMessageA(m_lLVHwnd, LVM_INSERTCOLUMNA, lIndex, uLVC) > -1)
    End If
    
    If ColumnAdd Then
        If bFirst Then
            m_lHdrHwnd = HeaderHwnd()
            If Not m_lImlHdHndl = 0 Then
                If m_bUseUnicode Then
                    SendMessageLongW m_lHdrHwnd, HDM_SETIMAGELIST, 0&, m_lImlHdHndl
                Else
                    SendMessageLongA m_lHdrHwnd, HDM_SETIMAGELIST, 0&, m_lImlHdHndl
                End If
            End If
        End If
        '/* unicode nt
        If m_bUseUnicode Then
            With uHDW
                .pszText = StrPtr(sText)
                .cchTextMax = Len(sText)
                .cxy = lWidth
                .iImage = lIcon
                .fmt = HDF_STRING Or eAlign * -(lIndex <> 0) Or HDF_IMAGE * -(lIcon > -1) Or HDF_BITMAP_ON_RIGHT
                .Mask = HDI_TEXT Or HDI_WIDTH Or HDI_IMAGE Or HDI_FORMAT
            End With
            SendMessageW m_lHdrHwnd, HDM_SETITEMW, lIndex, uHDW
        Else
            With uHDI
                .pszText = sText
                .cchTextMax = Len(sText)
                .cxy = lWidth
                .iImage = lIcon
                .fmt = HDF_STRING Or eAlign * -(lIndex <> 0) Or HDF_IMAGE * -(lIcon > -1) Or HDF_BITMAP_ON_RIGHT
                .Mask = HDI_TEXT Or HDI_WIDTH Or HDI_IMAGE Or HDI_FORMAT
            End With
            SendMessageA m_lHdrHwnd, HDM_SETITEMA, lIndex, uHDI
        End If
    End If
    
    If m_lColumnHeight = 0 Then
        m_lColumnHeight = ColumnHeight
    End If
    '/* sort flag
    c_ColumnTags.Add ColumnTag, CStr(lIndex)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnAdd", Err.Number)

End Function

Public Property Get ColumnAlign(ByVal lColumn As Long) As ECAColumnAlign
'*/ retieve a columns text alignment

Const CALGN     As Long = &H3
Dim uLVC        As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    uLVC.Mask = LVCF_FMT
    If m_bUseUnicode Then
        SendMessageW m_lLVHwnd, LVM_GETCOLUMNW, lColumn, uLVC
    Else
        SendMessageA m_lLVHwnd, LVM_GETCOLUMNA, lColumn, uLVC
    End If
    ColumnAlign = (CALGN And uLVC.fmt)
    
End Property

Public Property Let ColumnAlign(ByVal lColumn As Long, _
                                ByVal eAlign As ECAColumnAlign)
'*/ change a columns text alignment

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uLVC
        .fmt = eAlign * -(Not lColumn = 0)
        .Mask = LVCF_FMT
    End With
    If m_bUseUnicode Then
        SendMessageW m_lLVHwnd, LVM_SETCOLUMNW, lColumn, uLVC
    Else
        SendMessageA m_lLVHwnd, LVM_SETCOLUMNA, lColumn, uLVC
    End If
    ListRefresh True

End Property

Public Function ColumnAutosize(ByVal lColumn As Long, _
                               Optional ByVal AutosizeType As ECAColumnAutosize = [ColumnItem]) As Boolean
'*/ autosize columns

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ColumnAutosize = CBool(SendMessageLongA(m_lLVHwnd, LVM_SETCOLUMNWIDTH, lColumn, AutosizeType))

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
    ColumnCount = SendMessageLongA(HeaderHwnd(), HDM_GETITEMCOUNT, 0&, 0&)

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
    If m_bUseUnicode Then
        SendMessageW m_lLVHwnd, LVM_GETCOLUMNW, lColumn, uLVC
    Else
        SendMessageA m_lLVHwnd, LVM_GETCOLUMNA, lColumn, uLVC
    End If
    ColumnIcon = uLVC.iImage
    
End Property

Public Property Let ColumnIcon(ByVal lColumn As Long, _
                               ByVal lIcon As Long)
'*/ change header icon

Const lMask     As Long = &H3
Dim lAlign      As Long
Dim uHDI        As HDITEM

    If (m_lLVHwnd = 0) Or (m_lHdrHwnd = 0) Then Exit Property
    With uHDI
        .Mask = HDI_FORMAT
        If m_bUseUnicode Then
            SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, uHDI
        Else
            SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, uHDI
        End If
        lAlign = lMask And .fmt
        .iImage = lIcon
        .fmt = HDF_STRING Or lAlign Or HDF_IMAGE * -(lIcon > -1 And m_lImlHdHndl <> 0) Or HDF_BITMAP_ON_RIGHT
        .Mask = HDI_IMAGE * -(lIcon > -1) Or HDI_FORMAT
    End With
    If m_bUseUnicode Then
        SendMessageW m_lHdrHwnd, HDM_SETITEMW, lColumn, uHDI
    Else
        SendMessageA m_lHdrHwnd, HDM_SETITEMA, lColumn, uHDI
    End If

End Property

Private Sub ColumnIconReset()

Dim lCt As Long

On Error GoTo Handler

    For lCt = 0 To ColumnCount - 1
        ColumnIcon(lCt) = -1
    Next lCt

Handler:
    On Error GoTo 0

End Sub

Public Function ColumnLastFit() As Boolean

Dim lCol    As Long

    If m_lLVHwnd = 0 Then Exit Function
    lCol = (ColumnCount - 1)
    If m_bUseUnicode Then
        SendMessageLongW m_lLVHwnd, LVM_SETCOLUMNWIDTH, lCol, LVSCW_AUTOSIZE_USEHEADER
    Else
        SendMessageLongA m_lLVHwnd, LVM_SETCOLUMNWIDTH, lCol, LVSCW_AUTOSIZE_USEHEADER
    End If

End Function

Public Function ColumnRemove(ByVal lColumn As Long) As Boolean
'*/ remove a column

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    If m_bUseUnicode Then
        ColumnRemove = CBool(SendMessageLongW(m_lLVHwnd, LVM_DELETECOLUMN, lColumn, 0&))
    Else
        ColumnRemove = CBool(SendMessageLongA(m_lLVHwnd, LVM_DELETECOLUMN, lColumn, 0&))
    End If
    If Me.ColumnCount = 0 Then m_lHdrHwnd = 0

On Error GoTo 0

Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnRemove", Err.Number)

End Function

Private Property Get ColumnIndex(ByVal lColumn As Long) As Long

Dim tLVI     As LVCOLUMN

    tLVI.Mask = LVCF_ORDER
    If Not (SendMessageA(m_lLVHwnd, LVM_GETCOLUMNA, lColumn, tLVI) = 0) Then
        ColumnIndex = tLVI.iOrder
    End If

End Property

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
    
    ColumnIconReset
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

Public Function ColumnSizeToItems(Optional ByVal bColumnFit As Boolean) As Boolean
'/* size columns to longest items

Dim lCol    As Long
Dim lParam  As Long

    If m_lLVHwnd = 0 Then Exit Function
    If bColumnFit Then
        lParam = LVSCW_AUTOSIZE_USEHEADER
    Else
        lParam = LVSCW_AUTOSIZE
    End If

    For lCol = 0 To ColumnCount - 1
        If m_bUseUnicode Then
            SendMessageLongW m_lLVHwnd, LVM_SETCOLUMNWIDTH, lCol, lParam
        Else
            SendMessageLongA m_lLVHwnd, LVM_SETCOLUMNWIDTH, lCol, lParam
        End If
    Next
    
End Function

Public Property Get ColumnTag(ByVal lColumn As Long) As ECSColumnSortTags
'/* get column sort tag

Dim lRet As Long

On Error GoTo Handler

    lRet = c_ColumnTags.Item(CStr(lColumn))
    If (lRet < 0) Or lRet > 3 Then
        GoTo Handler
    End If
    ColumnTag = lRet

On Error GoTo 0
Exit Property

Handler:
    ColumnTag = -1
    
End Property

Private Property Get ColumnText(ByVal lColumn As Long) As String
'*/ get a columns heading

Dim lLen        As Long
Dim aText(261)  As Byte
Dim uLVC        As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property

    If m_bUseUnicode Then
        With uLVC
            .pszText = VarPtr(aText(0))
            .cchTextMax = UBound(aText) + 1
            .Mask = LVCF_TEXT
        End With
        SendMessageW m_lLVHwnd, LVM_GETCOLUMNW, lColumn, uLVC
        ColumnText = PointerToString(uLVC.pszText)
    Else
        With uLVC
            .pszText = VarPtr(aText(0))
            .cchTextMax = UBound(aText)
            .Mask = LVCF_TEXT
        End With
        SendMessageA m_lLVHwnd, LVM_GETCOLUMNA, lColumn, uLVC
        ColumnText = StrConv(aText(), vbUnicode)
        lLen = InStr(ColumnText, vbNullChar)
        If lLen Then
            ColumnText = left$(ColumnText, lLen - 1)
        End If
    End If
    
End Property

Public Property Let ColumnText(ByVal lColumn As Long, _
                               ByVal sText As String)
'*/ change a columns heading

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    If m_bUseUnicode Then
    With uLVC
        .pszText = StrPtr(sText)
        .cchTextMax = Len(sText)
        .Mask = LVCF_TEXT
    End With
        SendMessageW m_lLVHwnd, LVM_SETCOLUMNW, lColumn, uLVC
    Else
    With uLVC
        .pszText = sText
        .cchTextMax = Len(sText)
        .Mask = LVCF_TEXT
    End With
        SendMessageA m_lLVHwnd, LVM_SETCOLUMNA, lColumn, uLVC
    End If

End Property

Public Property Get ColumnWidth(ByVal lColumn As Long) As Long
'*/ retrieve a columns length

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    ColumnWidth = SendMessageLongA(m_lLVHwnd, LVM_GETCOLUMNWIDTH, lColumn, 0&)

End Property

Public Property Let ColumnWidth(ByVal lColumn As Long, _
                                ByVal lWidth As Long)
'*/ change a columns length

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    SendMessageLongA m_lLVHwnd, LVM_SETCOLUMNWIDTH, lColumn, lWidth

End Property

Private Function HeaderHwnd() As Long
'*/ return the column header handle

    If m_lLVHwnd = 0 Then Exit Function
    m_lHdrHwnd = SendMessageLongA(m_lLVHwnd, LVM_GETHEADER, 0&, 0&)
    HeaderHwnd = m_lHdrHwnd

End Function

Public Property Get HeaderColor() As OLE_COLOR
Attribute HeaderColor.VB_Description = "[ole color] non skinned header color"
'*/ return the header color
    HeaderColor = m_oHdrBkClr
End Property

Public Property Let HeaderColor(ByVal PropVal As OLE_COLOR)
'*/ change the header color

    If m_bXPColors Then
        m_oHdrBkClr = XPShift(PropVal)
    End If
    m_oHdrBkClr = PropVal
    ListRefresh True
    PropertyChanged "HeaderColor"
    
End Property

Public Property Get HeaderCustom() As Boolean
Attribute HeaderCustom.VB_Description = "[bool] enable non skinned custom header colors"
'*/ return the custom header status
    HeaderCustom = m_bCustomHeader
End Property

Public Property Let HeaderCustom(ByVal PropVal As Boolean)
'*/ change the custom header status

    m_bCustomHeader = PropVal
    PropertyChanged "HeaderCustom"
    
End Property

Public Property Get HeaderDragDrop() As Boolean
Attribute HeaderDragDrop.VB_Description = "[bool] enable header drag and drop functionality"
'*/ retrieve drag and drop state
    HeaderDragDrop = m_bDragDrop
End Property

Public Property Let HeaderDragDrop(ByVal PropVal As Boolean)
'*/ retrieve drag and drop state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_HEADERDRAGDROP, 0
        Else
            SetExtendedStyle 0, LVS_EX_HEADERDRAGDROP
        End If
    End If
    m_bDragDrop = PropVal
    ListRefresh
    PropertyChanged "HeaderDragDrop"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("HeaderDragDrop", Err.Number)

End Property

Public Property Get HeaderFixedWidth() As Boolean
Attribute HeaderFixedWidth.VB_Description = "[bool] disable header size change"
'*/ retrieve fixed width state
    HeaderFixedWidth = m_bHeaderFixed
End Property

Public Property Let HeaderFixedWidth(ByVal PropVal As Boolean)
'*/ change fixed width state
    m_bHeaderFixed = PropVal
    ListRefresh
    PropertyChanged "HeaderFixedWidth"
End Property

Public Property Get HeaderFlat() As Boolean
Attribute HeaderFlat.VB_Description = "[bool] non skinned use flat header style"
'*/ change width state
    HeaderFlat = m_bHeaderFlat
End Property

Public Property Let HeaderFlat(ByVal PropVal As Boolean)
'*/ change header style

Dim lStyle      As Long
Dim lHwnd       As Long

On Error GoTo Handler

    ColumnIconReset
    If Not m_lLVHwnd = 0 Then
        lHwnd = HeaderHwnd()
        If lHwnd = 0 Then Exit Property
        If m_bUseUnicode Then
            lStyle = GetWindowLongW(lHwnd, GWL_STYLE)
        Else
            lStyle = GetWindowLongA(lHwnd, GWL_STYLE)
        End If
        If PropVal Then
            lStyle = lStyle And Not HDS_BUTTONS
        Else
            lStyle = lStyle Or HDS_BUTTONS
        End If
        If m_bUseUnicode Then
            SetWindowLongW lHwnd, GWL_STYLE, lStyle
        Else
            SetWindowLongA lHwnd, GWL_STYLE, lStyle
        End If
    End If
    m_bHeaderFlat = PropVal
    ListRefresh True
    PropertyChanged "HeaderFlat"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("HeaderFlat", Err.Number)

End Property

Public Property Get HeaderForeColor() As OLE_COLOR
Attribute HeaderForeColor.VB_Description = "[ole color] header font color in custom or skinned modes"
'*/ return the header forecolor
    HeaderForeColor = m_oHdrForeClr
End Property

Public Property Let HeaderForeColor(ByVal PropVal As OLE_COLOR)
'*/ change the header forecolor

    m_oHdrForeClr = PropVal
    If m_bSkinHeader Then
        m_cSkinHeader.HeaderForeColor = PropVal
    End If
    ListRefresh
    PropertyChanged "HeaderForeColor"
    
End Property

Public Property Get HeaderHide() As Boolean
Attribute HeaderHide.VB_MemberFlags = "400"
'*/ retrieve header visible state
    HeaderHide = m_bHeaderHide
End Property

Public Property Let HeaderHide(ByVal PropVal As Boolean)
'*/ change header visible state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetStyle LVS_NOCOLUMNHEADER, 0
        Else
            SetStyle 0, LVS_NOCOLUMNHEADER
        End If
    End If
    m_bHeaderHide = PropVal
    ListRefresh True
    PropertyChanged "HeaderHide"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("HeaderHide", Err.Number)

End Property

Public Property Get HeaderHighLite() As OLE_COLOR
Attribute HeaderHighLite.VB_Description = "[ole color] header font over color in skinned mode"
'*/ return header highlite color
    HeaderHighLite = m_oHdrHighLiteClr
End Property

Public Property Let HeaderHighLite(ByVal PropVal As OLE_COLOR)
'*/ change header highlite color

    If m_bSkinHeader Then
        m_cSkinHeader.HeaderHighLite = PropVal
    End If
    m_oHdrHighLiteClr = PropVal
    ListRefresh
    PropertyChanged "HeaderHighLite"
    
End Property

Public Property Get HeaderPressed() As OLE_COLOR
Attribute HeaderPressed.VB_Description = "[ole color] header font pressed color in skinned mode"
'*/ return header highlite color
    HeaderPressed = m_oHdrPressedClr
End Property

Public Property Let HeaderPressed(ByVal PropVal As OLE_COLOR)
'*/ change header highlite color

    If Not m_cSkinHeader Is Nothing Then
        m_cSkinHeader.HeaderPressed = PropVal
    End If
    m_oHdrPressedClr = PropVal
    ListRefresh
    PropertyChanged "HeaderPressed"
    
End Property

Private Function HotDivider() As Long
'/* header drag mark

Dim lDivPos     As Long
Dim lHdc        As Long
Dim lCol        As Long
Dim lColIdx     As Long
Dim lCount      As Long
Dim lXpos       As Long
Dim tRect       As RECT

    lDivPos = DividerPosition
    lCount = ColumnCount
    
    If lDivPos = -1 Then
        Exit Function
    End If
    '/* 20 second timeout
    If m_lSafeTimer > 2000 Then
        DragStopTimer
        Exit Function
    Else
        m_lSafeTimer = m_lSafeTimer + 1
    End If
    '/* relative column
    If lDivPos = lCount Then
        lCol = (lCount - 1)
    Else
        lCol = lDivPos
    End If
    '/* get rect
    lColIdx = m_cSkinHeader.ColumnIndex(lCol)
    SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, lColIdx, tRect
    '/* position mark
    If lDivPos = 0 Then
        lXpos = 0
    ElseIf lDivPos = lCount Then
        lXpos = tRect.right
    Else
        lXpos = tRect.left
    End If
    '/* draw mark
    lHdc = GetDC(m_lLVHwnd)
    With tRect
        m_cRender.Stretch lHdc, lXpos, 0, 3, 17, m_cDivider.hdc, 0, 0, 3, 16, SRCCOPY
    End With
    '/* cleanup
    ReleaseDC m_lLVHwnd, lHdc
   
End Function

Private Function DividerPosition() As Long

Dim lPos    As Long
Dim tPnt    As POINTAPI
Dim tRect   As RECT

    GetCursorPos tPnt
    ScreenToClient m_lHdrHwnd, tPnt
    GetClientRect m_lHdrHwnd, tRect
    
    With tPnt
        If (.Y > -8) And (.Y < tRect.bottom + 8) Then
            lPos = (.X And &HFFFF&)
            lPos = lPos Or (.Y And &H7FFF) * &H10000
            If (.Y And &H8000) = &H8000 Then
                lPos = lPos Or &H80000000
            End If
            DividerPosition = SendMessageLongA(m_lHdrHwnd, HDM_SETHOTDIVIDER, 1&, lPos)
        Else
            DividerPosition = -1
        End If
    End With

End Function

Private Function GetHorzPos() As Long

Dim tPnt    As POINTAPI

    GetCursorPos tPnt
    ScreenToClient m_lLVHwnd, tPnt
    GetHorzPos = tPnt.X
    
End Function


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

Public Function InitImlLarge(Optional ByVal lWidth As Long = 32, _
                             Optional ByVal lHeight As Long = 32) As Boolean
'*/ initialize large icons image list

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlLarge
    m_lImlLargeHndl = ImageList_Create(lWidth, lHeight, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    SendMessageLongA m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_NORMAL, m_lImlLargeHndl
    InitImlLarge = (Not m_lImlLargeHndl = 0)
    m_lLargeIconX = lWidth
    m_lLargeIconY = lHeight
    
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

Public Function InitImlSmall(Optional ByVal lWidth As Long = 16, _
                             Optional ByVal lHeight As Long = 16) As Boolean

'*/ initialize smallicons image list

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlSmall
    m_lImlSmallHndl = ImageList_Create(lWidth, lHeight, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    SendMessageLongA m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_SMALL, m_lImlSmallHndl
    InitImlSmall = (Not m_lImlSmallHndl = 0)
    m_lSmallIconX = lWidth
    m_lSmallIconY = lHeight
    
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

Dim lRet As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlState
    m_lImlStateHndl = ImageList_Create(16&, 16&, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    lRet = SendMessageLongA(m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_STATE, m_lImlStateHndl)
    InitImlState = (Not lRet = 0)
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("InitImlState", Err.Number)

End Function

Public Sub LoadStateImageList()


End Sub

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

    If m_lLVHwnd = 0 Then Exit Property
    Count = SendMessageLongA(m_lLVHwnd, LVM_GETITEMCOUNT, 0&, 0&)

End Property

Public Function Find(ByVal sText As String, _
                     ByVal lColumn As Long, _
                     ByVal bMatchCase As Boolean, _
                     ByVal bExact As Boolean, _
                     ByVal bDescending As Boolean, _
                     ByVal bFindNext As Boolean, _
                     ByVal bDisplay As Boolean) As Long

Dim lCt     As Long
Dim lUt     As Long
Dim lRs     As Long
Dim lCm     As Long
Static lSp  As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    If Count = 0 Then Exit Function
    lUt = Count - 1
    '/* if you put this many items in, you are a bad programmer
    If lUt > 200000 Then Exit Function
    lRs = -1
    '/* store last found item
    If Not bFindNext Then
        lSp = 0
    Else
        If bDescending Then
            lUt = lSp - 1
            lSp = 0
        Else
            lSp = lSp + 1
        End If
    End If
    
    If bMatchCase Then
        lCm = vbBinaryCompare
    Else
        lCm = vbTextCompare
    End If
    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        If bDescending Then
            If lColumn = 0 Then
                If m_bUseSorted And m_bSorted Then
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(m_lPtr(lCt)).Text Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(m_lPtr(lCt)).Text, sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(lCt).Text Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(lCt).Text, sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            Else
                If m_bUseSorted And m_bSorted Then
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(m_lPtr(lCt)).SubItemText(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(m_lPtr(lCt)).SubItemText(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(lCt).SubItemText(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(lCt).SubItemText(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            End If
        Else
            If lColumn = 0 Then
                If m_bUseSorted And m_bSorted Then
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(m_lPtr(lCt)).Text Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(m_lPtr(lCt)).Text, sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(lCt).Text Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(lCt).Text, sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            Else
                If m_bUseSorted And m_bSorted Then
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(m_lPtr(lCt)).SubItemText(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(m_lPtr(lCt)).SubItemText(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_cListItems(lCt).SubItemText(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_cListItems(lCt).SubItemText(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            End If
        End If
    
    '/* hl mode
    Case eHyperList
        If bDescending Then
            If lColumn = 0 Then
                If m_bUseSorted And m_bSorted Then
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).Item(m_lPtr(lCt)) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).Item(m_lPtr(lCt)), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).Item(lCt) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).Item(lCt), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            Else
                If m_bUseSorted And m_bSorted Then
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).SubItem(m_lPtr(lCt)).Text(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).SubItem(m_lPtr(lCt)).Text(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lUt To lSp Step -1
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).SubItem(lCt).Text(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).SubItem(lCt).Text(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            End If
        Else
            If lColumn = 0 Then
                If m_bUseSorted And m_bSorted Then
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).Item(m_lPtr(lCt)) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).Item(m_lPtr(lCt)), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).Item(lCt) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).Item(lCt), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            Else
                If m_bUseSorted And m_bSorted Then
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).SubItem(m_lPtr(lCt)).Text(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).SubItem(m_lPtr(lCt)).Text(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                Else
                    For lCt = lSp To lUt
                        '/* early exit
                        If m_bStopSearch Then GoTo Handler
                        If bExact Then
                            If sText = m_HLIStc(0).SubItem(lCt).Text(lColumn) Then
                                lRs = lCt
                                Exit For
                            End If
                        Else
                            If InStr(1, m_HLIStc(0).SubItem(lCt).Text(lColumn), sText, lCm) > 0 Then
                                lRs = lCt
                                Exit For
                            End If
                        End If
                    Next lCt
                End If
            End If
        End If
    
    '/* db mode
    Case eDatabase
        If bDescending Then
            If lColumn = 0 Then
                For lCt = lUt To lSp Step -1
                    '/* early exit
                    If m_bStopSearch Then GoTo Handler
                    If bExact Then
                        If sText = ItemText(lCt) Then
                            lRs = lCt
                            Exit For
                        End If
                    Else
                        If InStr(1, ItemText(lCt), sText, lCm) > 0 Then
                            lRs = lCt
                            Exit For
                        End If
                    End If
                Next lCt
            Else
                For lCt = lUt To lSp Step -1
                    '/* early exit
                    If m_bStopSearch Then GoTo Handler
                    If bExact Then
                        If sText = SubItemText(lCt, lColumn) Then
                            lRs = lCt
                            Exit For
                        End If
                    Else
                        If InStr(1, SubItemText(lCt, lColumn), sText, lCm) > 0 Then
                            lRs = lCt
                            Exit For
                        End If
                    End If
                Next lCt
            End If
        Else
            If lColumn = 0 Then
                For lCt = lSp To lUt
                    '/* early exit
                    If m_bStopSearch Then GoTo Handler
                    If bExact Then
                        If sText = ItemText(lCt) Then
                            lRs = lCt
                            Exit For
                        End If
                    Else
                        If InStr(1, ItemText(lCt), sText, lCm) > 0 Then
                            lRs = lCt
                            Exit For
                        End If
                    End If
                Next lCt
            Else
                For lCt = lSp To lUt
                    '/* early exit
                    If m_bStopSearch Then GoTo Handler
                    If bExact Then
                        If sText = SubItemText(lCt, lColumn) Then
                            lRs = lCt
                            Exit For
                        End If
                    Else
                        If InStr(1, SubItemText(lCt, lColumn), sText, lCm) > 0 Then
                            lRs = lCt
                            Exit For
                        End If
                    End If
                Next lCt
            End If
        End If
    End Select
    
    If lRs > -1 Then
        lSp = lRs
        Find = lRs
        If bDisplay Then
            ItemEnsureVisible lRs
            ItemFocused(lRs) = True
            ItemSelected(lRs) = True
        End If
    End If
    
On Error GoTo 0
Exit Function

Handler:

End Function

Public Function ItemAdd(ByVal lIndex As Long, _
                        ByVal sKey As String, _
                        ByVal sText As String, _
                        ByVal lIcon As Long, _
                        ByVal lSmallIcon As Long) As Boolean

'/* add an item

On Error Resume Next

    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        If lIndex > (Count - 1) Then
            '/* add to class array
            ReDim Preserve m_cListItems(0 To lIndex)
            Set m_cListItems(lIndex) = New clsListItem
        ElseIf m_cListItems(lIndex) Is Nothing Then
            ReDim Preserve m_cListItems(0 To UBound(m_cListItems) + 1)
            Set m_cListItems(lIndex) = New clsListItem
        End If
        m_cListItems(lIndex).Add (lIndex), sKey, sText, lIcon, lSmallIcon
    '/* hl mode
    Case eHyperList
        '/* redim struct item arrays
        If lIndex > (Count - 1) Then
            ReDim Preserve m_HLIStc(0).Item(0 To lIndex)
            ReDim Preserve m_HLIStc(0).lIcon(0 To lIndex)
            ReDim Preserve m_HLIStc(0).SubItem(0 To lIndex)
            ReDim Preserve m_HLIStc(0).SubItem(lIndex).Text(1 To (ColumnCount - 1))
            m_HLIStc(0).Item(lIndex) = sText
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

Public Function ItemEnsureVisible(ByVal lItem As Long) As Boolean
'*/ move to item index

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ItemEnsureVisible = CBool(SendMessageLongA(m_lLVHwnd, LVM_ENSUREVISIBLE, lItem, 0&))

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ItemEnsureVisible", Err.Number)

End Function

Public Property Get ItemFocused(ByVal lItem As Long) As Boolean
Attribute ItemFocused.VB_MemberFlags = "400"
'*/ return item focused state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    ItemFocused = CBool(SendMessageLongA(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_FOCUSED))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemFocused", Err.Number)

End Property

Public Property Let ItemFocused(ByVal lItem As Long, _
                                ByVal bFocused As Boolean)
'*/ change item focused state

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .stateMask = LVIS_FOCUSED
        .State = -bFocused * LVIS_FOCUSED
        .Mask = LVIF_STATE
    End With
    SendMessageA m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemFocused", Err.Number)

End Property

Public Property Get ItemGhosted(ByVal lItem As Long) As Boolean
Attribute ItemGhosted.VB_MemberFlags = "400"
'*/ return item ghosted state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    ItemGhosted = (SendMessageLongA(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_CUT))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemGhosted", Err.Number)

End Property

Public Property Let ItemGhosted(ByVal lItem As Long, _
                                ByVal bGhosted As Boolean)
'*/ change item ghosted state

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .stateMask = LVIS_CUT
        .State = LVIS_CUT * -bGhosted
        .Mask = LVIF_STATE
    End With
    SendMessageA m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemGhosted", Err.Number)

End Property

Public Property Get ItemIcon(ByVal lItem As Long) As Long
Attribute ItemIcon.VB_MemberFlags = "400"
'*/ return icon index

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .iItem = lItem
        .Mask = LVIF_IMAGE
    End With
    If m_bIsNt Then
        SendMessageW m_lLVHwnd, LVM_GETITEMW, 0&, uLVI
    Else
        SendMessageA m_lLVHwnd, LVM_GETITEMA, 0&, uLVI
    End If
    ItemIcon = uLVI.iImage

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIcon", Err.Number)

End Property

Public Property Let ItemIcon(ByVal lItem As Long, _
                             ByVal lIcon As Long)
'*/ change icon index

Dim uLVI As LVITEM

On Error GoTo Handler

    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).Icon = lIcon
    Case eHyperList
        m_HLIStc(0).lIcon(lItem) = lIcon
    End Select
    
    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .iItem = lItem
        .iImage = lIcon
        .Mask = LVIS_OVERLAYMASK
        .State = lItem * 256
    End With
    SendMessageA m_lLVHwnd, LVM_SETITEMSTATE, 0&, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIcon", Err.Number)

End Property

Public Property Get ItemIndent() As Long
Attribute ItemIndent.VB_Description = "[long] report mode indent (skinned checkbox not supported)"
'*/ return item indent

On Error GoTo Handler

    ItemIndent = m_lItemIndent

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIndent", Err.Number)

End Property

Public Property Let ItemIndent(ByVal PropVal As Long)
'*/ change item indent

On Error GoTo Handler

    m_lItemIndent = PropVal
    PropertyChanged "ItemIndent"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemIndent", Err.Number)

End Property

Public Sub ItemRedraw(ByVal iIndex As Long)
    SendMessageLongA m_lLVHwnd, LVM_REDRAWITEMS, iIndex, iIndex
End Sub

Public Function ItemRemove(ByVal lItem As Long) As Boolean
'*/ remove an item from the list

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
        If ArrayExists(m_HLIStc(0).Item) Then
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
Attribute ItemSelected.VB_MemberFlags = "400"
'*/ return selected state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    ItemSelected = CBool(SendMessageLongA(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_SELECTED))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemSelected", Err.Number)

End Property

Public Property Let ItemSelected(ByVal lItem As Long, _
                                 ByVal bSelected As Boolean)
'*/ select an item

Dim uLVI    As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .stateMask = LVIS_SELECTED Or -(bSelected And lItem > -1) * LVIS_FOCUSED
        .State = -bSelected * LVIS_SELECTED Or -(lItem > -1) * LVIS_FOCUSED
        .Mask = LVIF_STATE
    End With
    SendMessageA m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemSelected", Err.Number)

End Property

Public Property Get ItemsSorted() As Boolean
Attribute ItemsSorted.VB_MemberFlags = "400"
'*/ return sorted mode status
    ItemsSorted = m_bUseSorted
End Property

Public Property Let ItemsSorted(ByVal PropVal As Boolean)
'*/ change sorted mode status
    m_bUseSorted = PropVal
End Property

Public Property Get ItemText(ByVal lItem As Long) As String
Attribute ItemText.VB_MemberFlags = "400"
'*/ return item text

Dim uLVI        As LVITEM
Dim uLVW        As LVITEMW
Dim aText(261)  As Byte
Dim lLen        As Long

On Error GoTo Handler


    If m_lLVHwnd = 0 Then Exit Property
    If m_bUseUnicode Then
        With uLVW
            .pszText = VarPtr(aText(0))
            .cchTextMax = UBound(aText) + 1
            .Mask = LVCF_TEXT
        End With
        SendMessageW m_lLVHwnd, LVM_GETITEMTEXTW, lItem, uLVW
        ItemText = PointerToString(uLVW.pszText)
    Else
        With uLVI
            .pszText = VarPtr(aText(0))
            .cchTextMax = UBound(aText)
            .Mask = LVCF_TEXT
        End With
        SendMessageA m_lLVHwnd, LVM_GETITEMTEXTA, lItem, uLVI
        ItemText = StrConv(aText(), vbUnicode)
        lLen = InStr(1, ItemText, vbNullChar)
        If lLen Then
            ItemText = left$(ItemText, lLen - 1)
        End If
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemText", Err.Number)

End Property

Public Property Let ItemText(ByVal lItem As Long, _
                             ByVal sText As String)
'*/ change item text

Dim uLVI As LVITEM
Dim uLVW As LVITEMW

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).Text = sText
    Case eHyperList
        m_HLIStc(0).Item(lItem) = sText
    End Select

    If m_bUseUnicode Then
        With uLVW
            .pszText = StrPtr(sText)
            .cchTextMax = Len(sText)
        End With
        SendMessageW m_lLVHwnd, LVM_SETITEMTEXTW, lItem, uLVW
    Else
        With uLVI
            .pszText = StrPtr(sText)
            .cchTextMax = Len(sText)
        End With
        SendMessageA m_lLVHwnd, LVM_SETITEMTEXTA, lItem, uLVI
    End If

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
Attribute SubItemIcon.VB_Description = "[bool] enable report view subitem icons"
'*/ retrieve subitem icon

Dim uLVI    As LVITEM
Dim uLVW    As LVITEMW

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    If m_bUseUnicode Then
        With uLVW
            .iItem = lItem
            .iSubItem = lSubItem
            .Mask = LVIF_IMAGE
        End With
        SendMessageW m_lLVHwnd, LVM_GETITEMW, 0&, uLVW
        SubItemIcon = uLVW.iImage
    Else
        With uLVI
            .iItem = lItem
            .iSubItem = lSubItem
            .Mask = LVIF_IMAGE
        End With
        SendMessageA m_lLVHwnd, LVM_GETITEMA, 0&, uLVI
        SubItemIcon = uLVI.iImage
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("SubItemIcon", Err.Number)

End Property

Public Property Let SubItemIcon(ByVal lItem As Long, _
                                ByVal lSubItem As Long, _
                                ByVal lIcon As Long)
'*/ change subitem icon

Dim uLVI    As LVITEM
Dim uLVW    As LVITEMW

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).SubItemIcon(lSubItem) = lIcon
    Case eHyperList
        m_HLIStc(0).SubItem(lItem).lIcon(lSubItem) = lIcon
    End Select

    If m_bUseUnicode Then
        With uLVW
            .iItem = lItem
            .iSubItem = lSubItem
            .iImage = lIcon
            .Mask = LVIF_IMAGE
        End With
        SendMessageW m_lLVHwnd, LVM_SETITEMW, 0&, uLVW
    Else
        With uLVI
            .iItem = lItem
            .iSubItem = lSubItem
            .iImage = lIcon
            .Mask = LVIF_IMAGE
        End With
        SendMessageA m_lLVHwnd, LVM_SETITEMA, 0&, uLVI
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("SubItemIcon", Err.Number)

End Property

Public Property Get SubItemImages() As Boolean
Attribute SubItemImages.VB_MemberFlags = "400"
'*/ retrieve subitem icon state
    SubItemImages = m_bSubItemImage
End Property

Public Property Let SubItemImages(ByVal PropVal As Boolean)
'*/ change subitem icon state

    m_bSubItemImage = PropVal
    ListRefresh
    PropertyChanged "SubItemImages"
    
End Property

Public Property Get SubItemText(ByVal lItem As Long, _
                                ByVal lSubItem As Long) As String
Attribute SubItemText.VB_MemberFlags = "400"
'*/ retieve subitem text

Dim uLVI        As LVITEM
Dim uLVW        As LVITEMW
Dim aText(256)  As Byte
Dim lLen        As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    If m_bUseUnicode Then
        With uLVW
            .iSubItem = lSubItem
            .pszText = VarPtr(aText(0))
            .cchTextMax = UBound(aText)
            .Mask = LVIF_TEXT
        End With
        lLen = SendMessageW(m_lLVHwnd, LVM_GETITEMTEXTW, lItem, uLVW)
    Else
        With uLVI
            .iSubItem = lSubItem
            .pszText = VarPtr(aText(0))
            .cchTextMax = UBound(aText)
            .Mask = LVIF_TEXT
        End With
        lLen = SendMessageA(m_lLVHwnd, LVM_GETITEMTEXTA, lItem, uLVI)
    End If

    If lLen > 0 Then
        If m_bUseUnicode Then
            SubItemText = PointerToString(uLVW.pszText)
        Else
            SubItemText = left$(StrConv(aText(), vbUnicode), lLen)
        End If
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

Dim uLVI    As LVITEM
Dim uLVW    As LVITEMW

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    
    Select Case m_eListMode
    Case eCustomDraw
        m_cListItems(lItem).SubItemText(lSubItem) = sText
    Case eHyperList
        m_HLIStc(0).SubItem(lItem).Text(lSubItem) = sText
    End Select
    
    If m_bUseUnicode Then
        With uLVW
            .iSubItem = lSubItem
            .pszText = StrPtr(sText)
            .cchTextMax = Len(sText)
        End With
        SendMessageW m_lLVHwnd, LVM_SETITEMTEXTW, lItem, uLVW
    Else
        With uLVI
            .iSubItem = lSubItem
            .pszText = StrPtr(sText)
            .cchTextMax = Len(sText)
            SendMessageA m_lLVHwnd, LVM_SETITEMTEXTA, lItem, uLVI
        End With
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("SubItemText", Err.Number)

End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_MemberFlags = "400"
'*/ [get] word wrap

    If m_lLVHwnd = 0 Then Exit Property
    WordWrap = m_bWordWrap

End Property

Public Property Let WordWrap(ByVal PropVal As Boolean)
'*/ [let] word wrap

    If m_lLVHwnd = 0 Then Exit Property
    m_bWordWrap = PropVal

End Property


'**********************************************************************
'*                          LISTVIEW PROPERTIES
'**********************************************************************

Public Property Get AutoArrange() As Boolean
Attribute AutoArrange.VB_Description = "[bool] autoarrange listitems"
'*/ [get] auto arrange
    AutoArrange = m_bAutoArrange
End Property

Public Property Let AutoArrange(ByVal PropVal As Boolean)
'*/ [let] auto arrange

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetStyle 0, LVS_AUTOARRANGE
        Else
            SetStyle LVS_AUTOARRANGE, 0
        End If
    End If
    m_bAutoArrange = PropVal
    PropertyChanged "AutoArrange"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("AutoArrange", Err.Number)

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "[ole color] listview backcolor"
'*/ retrieve list backcolor

    BackColor = m_oBackColor

End Property

Public Property Let BackColor(ByVal PropVal As OLE_COLOR)
'*/ change list backcolor

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        OleTranslateColor PropVal, 0&, m_oBackColor
        If m_bXPColors Then
            m_oBackColor = XPShift(PropVal)
        End If
        SendMessageLongA m_lLVHwnd, LVM_SETBKCOLOR, 0&, PropVal
        SendMessageLongA m_lLVHwnd, LVM_SETTEXTBKCOLOR, 0&, PropVal
        ListRefresh
    End If
    m_oBackColor = PropVal
    ListRefresh
    PropertyChanged "BackColor"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("BackColor", Err.Number)

End Property

Public Property Get BorderStyle() As EBSBorderStyle
Attribute BorderStyle.VB_Description = "[enum] listview border style"
'*/ retrive list borderstyle
    BorderStyle = m_eBorderStyle
End Property

Public Property Let BorderStyle(ByVal PropVal As EBSBorderStyle)
'*/ change list borderstyle

    SetBorderStyle UserControl.hwnd, PropVal
    m_eBorderStyle = PropVal
    UserControl_Resize
    ListRefresh True
    PropertyChanged "BorderStyle"

End Property

Private Sub SetBorderStyle(ByVal lHwnd As Long, _
                           ByVal eStyle As EBSBorderStyle)

    Select Case eStyle
    Case [None]
        WindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        WindowStyle lHwnd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
    Case [Thin]
        WindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        WindowStyle lHwnd, GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
    Case [Thick]
        WindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        WindowStyle lHwnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
    End Select

End Sub

Private Sub WindowStyle(ByVal lHwnd As Long, _
                        ByVal lType As Long, _
                        ByVal lStyle As Long, _
                        ByVal lStyleNot As Long)

Dim lNewStyle As Long
    
    If m_bUseUnicode Then
        lNewStyle = GetWindowLongW(lHwnd, lType)
    Else
        lNewStyle = GetWindowLongA(lHwnd, lType)
    End If
    lNewStyle = (lNewStyle And Not lStyleNot) Or lStyle
    If m_bUseUnicode Then
        SetWindowLongW lHwnd, lType, lNewStyle
    Else
        SetWindowLongA lHwnd, lType, lNewStyle
    End If
    SetWindowPos lHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or _
        SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Sub

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
Attribute CheckBoxes.VB_Description = "[bool] use listview checkboxes"
'*/ retrieve checkbox state
    CheckBoxes = m_bCheckBoxes
End Property

Public Property Let CheckBoxes(ByVal PropVal As Boolean)
'*/ change checkbox state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_CHECKBOXES, 0
        Else
            SetExtendedStyle 0, LVS_EX_CHECKBOXES
        End If
        If m_bCheckBoxes Then
            If Not m_bCheckInit Then
                InitCheckBoxes Count
            End If
        End If
    End If
    m_bCheckBoxes = PropVal
    ListRefresh True
    PropertyChanged "CheckBoxes"
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("CheckBoxes", Err.Number)

End Property

Public Property Get CustomDraw() As Boolean
Attribute CustomDraw.VB_Description = "[bool] use custom draw effects"
'/* [get] custom draw mode
    CustomDraw = m_bCustomDraw
End Property

Public Property Let CustomDraw(ByVal PropVal As Boolean)
'/* [let] custom draw mode
    m_bCustomDraw = PropVal
    ListRefresh True
    PropertyChanged "CustomDraw"
End Property

Private Function EditHandle() As Long

Dim lEditHnd As Long

    If m_lLVHwnd = 0 Then Exit Function
    lEditHnd = SendMessageLongA(m_lLVHwnd, LVM_GETEDITCONTROL, 0&, 0&)
    If Not lEditHnd = 0 Then
        EditHandle = lEditHnd
    End If

End Function

Public Function EditLimitLength(ByVal lLength As Long) As Boolean

Dim lEditHnd As Long

    lEditHnd = EditHandle
    If Not lEditHnd = 0 Then
        SendMessageLongA lEditHnd, EM_LIMITTEXT, lLength, 0&
    End If
    
End Function

Public Function EditLowerCase() As Boolean

Dim lStyle As Long
Dim lEditHnd    As Long

    lEditHnd = EditHandle
    If Not lEditHnd = 0 Then
        If m_bUseUnicode Then
            lStyle = GetWindowLongW(lEditHnd, GWL_STYLE)
        Else
            lStyle = GetWindowLongA(lEditHnd, GWL_STYLE)
        End If
        If m_bUseUnicode Then
            SetWindowLongW lEditHnd, GWL_STYLE, lStyle Or ES_LOWERCASE
        Else
            SetWindowLongA lEditHnd, GWL_STYLE, lStyle Or ES_LOWERCASE
        End If
    End If
    
End Function

Public Function EditUpperCase() As Boolean

Dim lStyle      As Long
Dim lEditHnd    As Long

    lEditHnd = EditHandle
    If Not lEditHnd = 0 Then
        If m_bUseUnicode Then
            lStyle = GetWindowLongW(lEditHnd, GWL_STYLE)
        Else
            lStyle = GetWindowLongA(lEditHnd, GWL_STYLE)
        End If
        If m_bUseUnicode Then
            SetWindowLongW lEditHnd, GWL_STYLE, lStyle Or ES_UPPERCASE
        Else
            SetWindowLongA lEditHnd, GWL_STYLE, lStyle Or ES_UPPERCASE
        End If
    End If
    
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "[bool] listview enabled state"
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
    
    If Not m_lLVHwnd = 0 Then
        EnableWindow m_lLVHwnd, Abs(PropVal)
        If Not PropVal Then
            m_lTmpBackClr = m_oBackColor
            m_lTmpForeClr = m_oForeColor
            BackColor = GetSysColor(vbButtonFace And &H1F&)
            ForeColor = GetSysColor(vbButtonShadow And &H1F&)
            m_bCustomDraw = False
        Else
            If Not m_lTmpBackClr = -1 Then
                BackColor = m_lTmpBackClr
                ForeColor = m_lTmpForeClr
            Else
                BackColor = &HFFFFFF
                ForeColor = &H0
            End If
            m_bCustomDraw = bCustom
        End If
    End If
    m_bEnabled = PropVal
    ListRefresh
    PropertyChanged "Enabled"
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("Enabled", Err.Number)
    
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "[ole font] listview font"
'*/ retrieve list font

On Error GoTo Handler

    Set Font = m_oFont

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("Font", Err.Number)

End Property

Public Property Set Font(ByVal PropVal As StdFont)
'*/ change list font

Dim lChar   As Long
Dim uLF     As LOGFONT

On Error GoTo Handler

    DestroyFont
    Set m_oFont = PropVal
    
    With uLF
         For lChar = 1 To Len(PropVal.Name)
             .lfFaceName(lChar - 1) = CByte(Asc(Mid$(PropVal.Name, lChar, 1)))
         Next lChar
         .lfHeight = -MulDiv(PropVal.SIZE, GetDeviceCaps(UserControl.hdc, LOGPIXELSY), 72)
         .lfItalic = PropVal.Italic
         .lfWeight = IIf(PropVal.Bold, FW_BOLD, FW_NORMAL)
         .lfUnderline = PropVal.Underline
         .lfStrikeOut = PropVal.Strikethrough
         If m_bUseUnicode Then
            .lfCharSet = 134
         Else
            .lfCharSet = 3
         End If
    End With
    
    If m_bUseUnicode Then
        m_lFont = CreateFontIndirectW(uLF)
    Else
        m_lFont = CreateFontIndirectA(uLF)
    End If
    If Not m_lLVHwnd = 0 Then
        SendMessageLongA m_lLVHwnd, WM_SETFONT, m_lFont, True
    End If
    ListRefresh True
    PropertyChanged "Font"
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("Font", Err.Number)

End Property

Private Function DestroyFont() As Boolean
'*/ font cleanup

On Error GoTo Handler

    If Not m_lFont = 0 Then
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

Public Property Get UseCellFont() As Boolean
Attribute UseCellFont.VB_Description = "[bool] custom draw mode use per cell fonts"
'/* [get] use cell font status
    UseCellFont = m_bUseCellFont
End Property

Public Property Let UseCellFont(ByVal PropVal As Boolean)
'/* [let] use per cell font

    m_bUseCellFont = PropVal
    ListRefresh
    PropertyChanged "UseCellFont"
    
End Property

Public Property Get UseCellColor() As Boolean
Attribute UseCellColor.VB_Description = "[bool] custom draw mode use per cell color"
'/* [get] use cell color
    UseCellColor = m_bUseCellColor
End Property

Public Property Let UseCellColor(ByVal PropVal As Boolean)
'/* [let] use per cell color

    m_bUseCellColor = PropVal
    ListRefresh
    PropertyChanged "UseCellColor"
    
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "[ole color] listview forecolor"
'*/ retrieve list forecolor

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        ForeColor = SendMessageLongA(m_lLVHwnd, LVM_GETTEXTCOLOR, 0&, 0&)
    End If
    ForeColor = m_oForeColor

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ForeColor", Err.Number)

End Property

Public Property Let ForeColor(ByVal PropVal As OLE_COLOR)
'*/ change list forecolor

On Error GoTo Handler

    m_oForeColor = PropVal
    If Not m_lLVHwnd = 0 Then
        OleTranslateColor PropVal, 0&, m_oForeColor
        SendMessageLongA m_lLVHwnd, LVM_SETTEXTCOLOR, 0&, m_oForeColor
    End If
    '
    ListRefresh
    PropertyChanged "ForeColor"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ForeColor", Err.Number)

End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "[bool] report mode full row select"
'*/ retrieve full row select state
    FullRowSelect = m_bFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal PropVal As Boolean)
'*/ change full row select state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_FULLROWSELECT, 0
        Else
            SetExtendedStyle 0&, LVS_EX_FULLROWSELECT
        End If
    End If
    m_bFullRowSelect = PropVal
    ListRefresh True
    PropertyChanged "FullRowSelect"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("FullRowSelect", Err.Number)

End Property

Private Sub GetItemRect(ByVal lIndex As Long, _
                        ByRef tRect As RECT, _
                        Optional ByVal lSubItem As Long)

'/* get rect struct of row item

    If m_lLVHwnd = 0 Then Exit Sub
    If lSubItem = 0 Then
        tRect.left = LVIR_BOUNDS
        SendMessageA m_lLVHwnd, LVM_GETITEMRECT, lIndex, tRect
    Else
        tRect.left = LVIR_LABEL
        tRect.top = lSubItem
        SendMessageA m_lLVHwnd, LVM_GETSUBITEMRECT, lIndex, tRect
    End If

End Sub

Private Function ColumnItemOffset(ByVal lIndex As Long) As Long

Dim tRect As RECT

    If m_lLVHwnd = 0 Then Exit Function
    tRect.left = LVIR_LABEL Or LVIR_BOUNDS
    SendMessageA m_lLVHwnd, LVM_GETITEMRECT, lIndex, tRect
    ColumnItemOffset = tRect.left - (m_lSmallIconX + 18)

End Function

Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "[bool] report mode gridlines"
'*/ change gridlines state
    GridLines = m_bGridLines
End Property

Public Property Let GridLines(ByVal PropVal As Boolean)
'*/ change gridlines state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_GRIDLINES, 0
        Else
            SetExtendedStyle 0, LVS_EX_GRIDLINES
        End If
    End If
    m_bGridLines = PropVal
    ListRefresh
    PropertyChanged "GridLines"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("GridLines", Err.Number)

End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_MemberFlags = "400"
'*/ retrieve selection visible state
    HideSelection = m_bHideSelection
End Property

Public Property Let HideSelection(ByVal PropVal As Boolean)
'*/ change selection visible state

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bHideSelection = PropVal
    If PropVal Then
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

Dim lX          As Long
Dim lY          As Long
Dim lWidth      As Long
Dim lHeight     As Long
Dim lhIcon      As Long
Dim lhImgList   As Long
Dim lHdc        As Long
Dim lFlags      As Long
Dim tRect       As RECT

On Error GoTo Handler

    lHdc = GetDC(m_lLVHwnd)
    '/* get list handle
    If ViewMode = StyleIcon Then
        lhImgList = m_lImlLargeHndl
    Else
        lhImgList = m_lImlSmallHndl
    End If
    '/* icon size
    ImageList_GetIconSize lhImgList, lWidth, lHeight
    '/* get coords
    tRect.left = LVIR_LABEL Or LVIR_BOUNDS
    SendMessageA m_lLVHwnd, LVM_GETITEMRECT, lItem, tRect
    lX = tRect.left - 32
    '/* checkbox offset
    If m_bCheckBoxes Then
        lX = lX + 16
    End If
    
    '/* get icon handle
    lhIcon = ImageList_GetIcon(lhImgList, lImgIndex, 0)
    '/* blend flag
    lFlags = lFlags Or ILD_SELECTED Or ILD_BLEND25
    '/* draw image
    With tRect
        lY = ((.bottom - .top) - lHeight) / 2
        ImageList_DrawEx lhImgList, lImgIndex, lHdc, lX, (.top + lY), 15, 15, CLR_NONE, TranslateColor(vbButtonFace), lFlags
    End With
    ReleaseDC m_lLVHwnd, lHdc
    
    '/* success
    IconDisabled = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("IconDisabled", Err.Number)
    
End Function

Public Property Get IconSpaceX() As Long
Attribute IconSpaceX.VB_Description = "[long] icon view item horizontal offset (skinned checkboxes not supported)"
'*/ [get] change icon left align
    IconSpaceX = m_lIconSpaceX
End Property

Public Property Let IconSpaceX(ByVal PropVal As Long)
'*/ [let] change icon left align

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        m_lIconSpaceX = PropVal
        'SetIconSpacing
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("IconSpaceX", Err.Number)

End Property

Public Property Get IconSpaceY() As Long
Attribute IconSpaceY.VB_Description = "[long] icon view item vertical offset (skinned checkboxes not supported)"
'*/ [get] change icon left align
    IconSpaceY = m_lIconSpaceY
End Property

Public Property Let IconSpaceY(ByVal PropVal As Long)
'*/ [let] change icon top align

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        m_lIconSpaceY = PropVal
        'SetIconSpacing
    End If
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("IconSpaceY", Err.Number)

End Property

Public Property Get InfoTips() As Boolean
Attribute InfoTips.VB_Description = "[bool] enable mouse over item tips"
'/* [get] infotips state
    InfoTips = m_bInfoTips
End Property

Public Property Let InfoTips(ByVal PropVal As Boolean)
'*/ [let] change info tips state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_INFOTIP, 0
        Else
            SetExtendedStyle 0, LVS_EX_INFOTIP
        End If
    End If
    m_bInfoTips = PropVal
    PropertyChanged "InfoTips"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("InfoTips", Err.Number)

End Property

Public Property Get ItemBorderSelect() As Boolean
Attribute ItemBorderSelect.VB_Description = "[bool] draw rect around icon in icon view (custom draw mode not supported)"
'*/ [get] border select state
    ItemBorderSelect = m_bItemBorderSelect
End Property

Public Property Let ItemBorderSelect(ByVal PropVal As Boolean)
'*/ [let] border select state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If m_eViewMode = StyleIcon Then
            If PropVal Then
                SetExtendedStyle LVS_EX_BORDERSELECT, 0
            Else
                SetExtendedStyle 0, LVS_EX_BORDERSELECT
            End If
        End If
    End If
    m_bItemBorderSelect = PropVal
    ListRefresh
    PropertyChanged "ItemBorderSelect"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ItemBorderSelect", Err.Number)
    
End Property

Private Function ItemHitTest() As Long
'/* subitem hit test

Dim tLVH    As LVHITTESTINFO
Dim tPoint  As POINTAPI
Dim lIndex  As Long

    GetCursorPos tPoint
    ScreenToClient m_lLVHwnd, tPoint
    lIndex = -1
    LSet tLVH.pt = tPoint
    SendMessageA m_lLVHwnd, LVM_HITTEST, 0&, tLVH
    If (tLVH.iItem <= 0) Then
        If (tLVH.flags And LVHT_NOWHERE) = LVHT_NOWHERE Then
            lIndex = -1
        Else
            lIndex = tLVH.iItem
        End If
    Else
        lIndex = tLVH.iItem
    End If
    ItemHitTest = lIndex

End Function

Private Function LargeIconHitTest() As Long

Dim lItems      As Long
Dim lRows       As Long
Dim lHeight     As Long
Dim lWidth      As Long
Dim lDepth      As Long
Dim lCX         As Long
Dim lCY         As Long
Dim lCt         As Long
Dim tPoint      As POINTAPI
Dim tRect       As RECT

On Error Resume Next

    '/* cursor pos
    GetCursorPos tPoint
    ScreenToClient m_lLVHwnd, tPoint
    '/* items per row
    For lCt = 0 To 10
        GetItemRect lCt, tRect
        If tRect.top > 0 Then
            lDepth = lCt
            lHeight = tRect.top
            Exit For
        End If
    Next lCt
    lItems = (Count)
    lDepth = ((Width / Screen.TwipsPerPixelX) / 82)
    '/* row count
    lRows = (lItems / lDepth)
    '/* item width
    lWidth = ((Width / Screen.TwipsPerPixelX) / lDepth)
    lHeight = 82
    lCX = Abs(tPoint.X / lWidth)
    lCY = Abs(tPoint.Y / lHeight)
    If Not lCY = 0 Then
        lCY = (lCY * lDepth)
    End If
    lCt = (lCY + lCX)
    
    GetItemRect lCt, tRect
    With tRect
        .right = .left + 20
        .top = .top + 20
        .bottom = .top + 20
    End With

    If PtInRect(tRect, tPoint.X, tPoint.Y) = 0 Then
        lCt = -1
    End If
    
    LargeIconHitTest = lCt

End Function

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

Public Property Get InsensitiveSort() As Boolean
Attribute InsensitiveSort.VB_MemberFlags = "400"
'/* [get] case insensitive sort
    InsensitiveSort = m_bInsensitiveSort
End Property

Public Property Let InsensitiveSort(ByVal PropVal As Boolean)
'/* [let] case insensitive sort
    m_bInsensitiveSort = PropVal
End Property

Public Function ItemTopIndex() As Long
'/* get first list item index

Dim lIndex As Long

    If Not m_lLVHwnd = 0 Then
        lIndex = SendMessageLongA(m_lLVHwnd, LVM_GETTOPINDEX, 0&, 0&)
    End If
    ItemTopIndex = lIndex
    
End Function

Public Property Get LabelEdit() As Boolean
Attribute LabelEdit.VB_Description = "[bool] enable item label editing"
'/* [get] edit state
    LabelEdit = m_bEditLabels
End Property

Public Property Let LabelEdit(ByVal PropVal As Boolean)
'*/ [let] change edit state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetStyle LVS_EDITLABELS, 0
        Else
            SetStyle 0, LVS_EDITLABELS
        End If
    End If
    m_bEditLabels = PropVal
    PropertyChanged "LabelEdit"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("LabelEdit", Err.Number)

End Property

Public Property Get SubItemsEdit() As Boolean
Attribute SubItemsEdit.VB_Description = "[bool] enable subitem editing"
'/* [get] subitem edit state
    SubItemsEdit = m_bSubItemsEdit
End Property

Public Property Let SubItemsEdit(ByVal PropVal As Boolean)
'*/ [let] change subitem edit state
    m_bSubItemsEdit = PropVal
    PropertyChanged "SubItemsEdit"
End Property

Public Property Get LabelTips() As Boolean
Attribute LabelTips.VB_Description = "[bool] enable tips for item labels"
'*/ retrieve label tips state
    LabelTips = m_bLabelTips
End Property

Public Property Let LabelTips(ByVal PropVal As Boolean)
'*/ change label tips state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_LABELTIP, 0
        Else
            SetExtendedStyle 0, LVS_EX_LABELTIP
        End If
    End If
    m_bLabelTips = PropVal
    PropertyChanged "LabelTips"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("LabelTips", Err.Number)

End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "[bool] enable selection of multiple items"
'*/ retrieve multiselect state
    MultiSelect = m_bMultiSelect
End Property

Public Property Let MultiSelect(ByVal PropVal As Boolean)
'*/ change multiselect state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetStyle 0, LVS_SINGLESEL
        Else
            SetStyle LVS_SINGLESEL, 0
        End If
    End If
    m_bMultiSelect = PropVal
    ListRefresh
    PropertyChanged "MultiSelect"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("MultiSelect", Err.Number)

End Property

Public Property Get OneClickActivate() As Boolean
Attribute OneClickActivate.VB_MemberFlags = "400"
'/* [get] one click label edit
    OneClickActivate = m_bOneClickActivate
End Property

Public Property Let OneClickActivate(ByVal PropVal As Boolean)
'*/ [let] one click edit

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_ONECLICKACTIVATE, 0
        Else
            SetExtendedStyle 0, LVS_EX_ONECLICKACTIVATE
        End If
    End If
    m_bOneClickActivate = PropVal
    PropertyChanged "OneClickActivate"

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
Dim cT          As Collection

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
            Err.Clear
            lCt = lCt + 1
        Loop Until lCt > lUb
        '/* reset the array
        CDResetArray m_cListItems
        '/* init list
        SetItemCount (UBound(m_cListItems) + 1)
        
    '/* hl mode
    Case eHyperList
        '/* get bounds
        lLb = LBound(m_HLIStc(0).Item)
        lUb = UBound(m_HLIStc(0).Item)
        lCt = lLb
        '/* filter
        Do
            cT.Add 1, m_sSortArray(lCt)
            '/* remove from array
            If Err.Number = 457 Then
                HLResizeArray lCt
            End If
            Err.Clear
            lCt = lCt + 1
        Loop Until lCt > lUb
        '/* reset list
        SetItemCount (UBound(m_HLIStc(0).Item) + 1)
    End Select
    '/* success
    m_bSorted = False
    RemoveDuplicates = True

On Error GoTo 0

End Function

Public Property Get ScrollBarFlat() As Boolean
Attribute ScrollBarFlat.VB_Description = "[bool] enable flat scrollbar style (non skinned mode)"
'*/ retrieve scrollbar state
    ScrollBarFlat = m_bScrollFlat
End Property

Public Property Let ScrollBarFlat(ByVal PropVal As Boolean)
'*/ change scrollbar state

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_FLATSB, 0
        Else
            SetExtendedStyle 0, LVS_EX_FLATSB
        End If
    End If
    m_bScrollFlat = PropVal
    ListRefresh True
    PropertyChanged "ScrollBarFlat"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ScrollBarFlat", Err.Number)

End Property

Public Property Get SelectedCount() As Long
Attribute SelectedCount.VB_MemberFlags = "400"
'*/ retrieve selected count

    If m_lLVHwnd = 0 Then Exit Property
    SelectedCount = SendMessageLongA(m_lLVHwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)

End Property

Private Function SubItemHitTest() As Long
'/* subitem hit test

Dim tLVH    As LVHITTESTINFO
Dim tPoint  As POINTAPI

    GetCursorPos tPoint
    ScreenToClient m_lLVHwnd, tPoint
    With tLVH
        .pt.X = tPoint.X
        .pt.Y = tPoint.Y
        .flags = LVHT_ONITEM
    End With

   SendMessageA m_lLVHwnd, LVM_SUBITEMHITTEST, 0&, tLVH
   SubItemHitTest = tLVH.iSubItem

End Function

Public Property Get TrackSelected() As Boolean
Attribute TrackSelected.VB_Description = "[bool] track items on mouse over"
'*/ [get] track item
    TrackSelected = m_bTrackSelected
End Property

Public Property Let TrackSelected(ByVal PropVal As Boolean)
'*/ [let] track item

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_TRACKSELECT, 0
        Else
            SetExtendedStyle 0, LVS_EX_TRACKSELECT
        End If
    End If
    m_bTrackSelected = PropVal
    PropertyChanged "TrackSelected"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("TrackSelected", Err.Number)

End Property

Public Property Get UnderlineHot() As Boolean
Attribute UnderlineHot.VB_MemberFlags = "400"
'*/ [get] underline hot item
    UnderlineHot = m_bUnderlineHot
End Property

Public Property Let UnderlineHot(ByVal PropVal As Boolean)
'*/ [let] underline hot item

On Error GoTo Handler

    If Not m_lLVHwnd = 0 Then
        If PropVal Then
            SetExtendedStyle LVS_EX_UNDERLINEHOT, 0
        Else
            SetExtendedStyle 0, LVS_EX_UNDERLINEHOT
        End If
    End If
    m_bUnderlineHot = PropVal
    PropertyChanged "UnderlineHot"

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("UnderlineHot", Err.Number)

End Property

Public Property Get ViewMode() As ELSStyle
Attribute ViewMode.VB_Description = "[enum] listview viewmode"
'*/ retrieve viewmode state
    ViewMode = m_eViewMode
End Property

Public Property Let ViewMode(ByVal PropVal As ELSStyle)
'*/ change viewmode state

On Error GoTo Handler

    If PropVal = StyleReport Then
        m_bSorted = False
    End If
    If m_bItemBorderSelect Then
        If Not PropVal = StyleIcon Then
            ItemBorderSelect = False
        End If
    End If

    m_eViewMode = PropVal
    If Not m_lLVHwnd = 0 Then
        SetStyle PropVal, (LVS_ICON Or LVS_SMALLICON Or LVS_REPORT Or LVS_LIST)
    End If
    
    PropertyChanged "ViewMode"
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("ViewMode", Err.Number)

End Property

Public Function VisibleItemCount() As Long
'/* items in view
Dim lCt As Long

    If m_lLVHwnd = 0 Then Exit Function
    lCt = SendMessageLongA(m_lLVHwnd, LVM_GETCOUNTPERPAGE, 0&, 0&)
    If Not lCt = 0 Then
        VisibleItemCount = lCt
    End If

End Function

Public Property Get XPColors() As Boolean
Attribute XPColors.VB_Description = "[bool] use xp color offsets on skin themes"
'/* [get] use xp colors
    XPColors = m_bXPColors
End Property

Public Property Let XPColors(ByVal PropVal As Boolean)
'/* [let] use xp colors
    m_bXPColors = PropVal
    PropertyChanged "XPColors"
End Property


'**********************************************************************
'*                              SUPPORT
'**********************************************************************

Public Function BackgroundPicture(ByVal sPath As String, _
                                  ByVal eStyle As EBGBackGroundImage, _
                                  Optional ByVal bCenter As Boolean) As Boolean

'/* set a bg image

Dim lReturn As Long
Dim uLBI    As LVBKIMAGE
Dim uLBW    As LVBKIMAGEW

On Error GoTo Handler

    sPath = ShortPath(sPath)
    If LenB(sPath) = 0 Then GoTo Handler
    If m_lLVHwnd = 0 Then GoTo Handler
    
    If m_bUseUnicode Then
        With uLBW
            sPath = sPath & vbNullChar
            .pszImage = StrPtr(sPath)
            .cchImageMax = lstrlenW(StrPtr(sPath)) + 1
            '/* clear image
            If eStyle = BgNone Then
                .ulFlags = LVBKIF_SOURCE_NONE
                .xOffsetPercent = 3
                .yOffsetPercent = 3
                m_bBackgroundBg = False
            Else
                .ulFlags = LVBKIF_SOURCE_URL Or eStyle
                '/* center image
                If bCenter And eStyle = BgNormal Then
                    .xOffsetPercent = 50
                    .yOffsetPercent = 50
                End If
                m_bBackgroundBg = True
            End If
        End With
        lReturn = SendMessageW(m_lLVHwnd, LVM_SETBKIMAGEW, 0&, uLBW)
        '/* set font transparent
        If lReturn Then
            SendMessageLongW m_lLVHwnd, LVM_SETTEXTBKCOLOR, 0&, CLR_NONE
        End If
    Else
        With uLBI
            .pszImage = sPath & vbNullChar
            .cchImageMax = Len(sPath) + 1
            '/* clear image
            If eStyle = BgNone Then
                .ulFlags = LVBKIF_SOURCE_NONE
                .xOffsetPercent = 3
                .yOffsetPercent = 3
                m_bBackgroundBg = False
            Else
                .ulFlags = LVBKIF_SOURCE_URL Or eStyle
                '/* center image
                If bCenter And eStyle = BgNormal Then
                    .xOffsetPercent = 50
                    .yOffsetPercent = 50
                End If
                m_bBackgroundBg = True
            End If
        End With
        lReturn = SendMessageA(m_lLVHwnd, LVM_SETBKIMAGEA, 0&, uLBI)
        '/* set font transparent
        If lReturn Then
            SendMessageLongA m_lLVHwnd, LVM_SETTEXTBKCOLOR, 0&, CLR_NONE
        End If
    End If
    
    '/* success
    BackgroundPicture = True
    
On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("BackgroundPicture", Err.Number)

End Function

Private Function ShortPath(ByVal sPath As String) As String

Dim lRes    As Long
Dim sTemp   As String

    sTemp = String$(165, Chr$(0))
    If m_bUseUnicode Then
        lRes = GetShortPathNameW(StrPtr(sPath), StrPtr(sTemp), 164)
    Else
        lRes = GetShortPathNameA(sPath, sTemp, 164)
    End If
    ShortPath = left$(sTemp, lRes)

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

    If m_bUseSorted And m_bSorted Then
        Checked = (m_lCheckState(m_lPtr(lIndex)) = 1)
    Else
        Checked = (m_lCheckState(lIndex) = 1)
    End If
    
End Property

Public Property Let Checked(ByVal lIndex As Long, _
                            ByVal bChecked As Boolean)
'/* let checkbox state
Dim lChk    As Long

    If bChecked Then
        lChk = 1
    Else
        lChk = 0
    End If
    If m_bUseSorted And m_bSorted Then
        m_lCheckState(m_lPtr(lIndex)) = lChk
    Else
        m_lCheckState(lIndex) = lChk
    End If

End Property

Private Function InitCheckBoxes(ByVal lCount As Long)
'/* load checkbox state array

    ReDim m_lCheckState(lCount)
    m_bCheckInit = True

End Function

Public Sub ListRefresh(Optional ByVal bErase As Boolean)
'/* refresh the listview

Dim tRect As RECT

    GetClientRect m_lLVHwnd, tRect
    If bErase Then
        EraseRect m_lLVHwnd, tRect, 1&
    Else
        EraseRect m_lLVHwnd, tRect, 0&
    End If

End Sub

Private Sub RefreshRegion(ByVal lItem As Long, _
                          Optional ByVal bCheck As Boolean)

Dim tRect As RECT

    '/* calculate rect
    If m_lLVHwnd = 0 Then Exit Sub
    tRect.left = LVIR_LABEL Or LVIR_BOUNDS
    SendMessageA m_lLVHwnd, LVM_GETITEMRECT, lItem, tRect
    If bCheck Then
        With tRect
            .left = .left - 34
            .right = .left + 18
        End With
    Else
        With tRect
            .left = 0
            .right = (UserControl.Width - 18)
        End With
    End If
    '/* invalidate region
    EraseRect m_lLVHwnd, tRect, 0&
    
End Sub

Private Function RePaint() As Boolean

    If m_bUseUnicode Then
        SendMessageLongW m_lLVHwnd, WM_PAINT, 0&, 0&
    Else
        SendMessageLongA m_lLVHwnd, WM_PAINT, 0&, 0&
    End If
    
End Function

Public Sub Resize()
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
            m_cSkinScrollBars.Resize
        End If
        If Not m_cSkinHeader Is Nothing Then
            m_cSkinHeader.Refresh -1
        End If
    End With

On Error GoTo 0

End Sub

Private Sub SetStyle(ByVal lStyle As Long, _
                     ByVal lStyleNot As Long)
'*/ change list style params

Dim lNewStyle   As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Sub
    If m_bUseUnicode Then
        lNewStyle = GetWindowLongW(m_lLVHwnd, GWL_STYLE)
    Else
        lNewStyle = GetWindowLongA(m_lLVHwnd, GWL_STYLE)
    End If
    
    lNewStyle = lNewStyle And Not lStyleNot
    lNewStyle = lNewStyle Or lStyle
    If m_bUseUnicode Then
        SetWindowLongW m_lLVHwnd, GWL_STYLE, lNewStyle
    Else
        SetWindowLongA m_lLVHwnd, GWL_STYLE, lNewStyle
    End If
    SetWindowPos m_lLVHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or _
        SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

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
    lNewStyle = SendMessageLongA(m_lLVHwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    lNewStyle = lNewStyle And Not lStyleNot
    lNewStyle = lNewStyle Or lStyle
    SendMessageLongA m_lLVHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, lNewStyle

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetExtendedStyle", Err.Number)

End Sub

Public Function SetItemCount(ByVal lItems As Long) As Boolean
'/* dimension list to item count

    If m_lLVHwnd = 0 Then Exit Function
    SendMessageA m_lLVHwnd, LVM_SETITEMCOUNT, lItems, LVSICF_NOINVALIDATEALL
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

Private Function pItemChecked(ByVal lIndex As Long) As Boolean
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
    lCt = 0
    
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
        lLb = LBound(m_HLIStc(0).Item)
        lUb = UBound(m_HLIStc(0).Item)
        ReDim m_sSortArray(lLb To lUb)
        If lColumn = 0 Then
            Do
                If m_eSortTag = SortDate Then
                    m_lSortArray(lCt) = GetTime(m_HLIStc(0).Item(lCt))
                ElseIf IsNumeric(m_HLIStc(0).Item(lCt)) Then
                    m_lSortArray(lCt) = CLng(m_HLIStc(0).Item(lCt))
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
        lLb = LBound(m_HLIStc(0).Item)
        lUb = UBound(m_HLIStc(0).Item)
        ReDim m_sSortArray(lLb To lUb)
        If lColumn = 0 Then
            Do
                m_sSortArray(lCt) = m_HLIStc(0).Item(lCt)
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

    ReDim Preserve cArray(lVl - 1)

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

Private Function GetTime(ByVal vDate As Variant) As Long

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
        ResizeArray .Item, lPos
        If ArrayCheck(.lIcon) Then
            ResizeArray .lIcon, lPos
        End If
        ResizeStruct lPos
    End With
    
End Sub

Public Function LoadArray() As Boolean
'*/ load data structure

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
                           ByVal bDsc As Boolean)

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
            If (lItem > lA(lIdxA(lo))) = bDsc Then
                lIdxA(hi) = lIdxA(lo)
                hi = hi - m1
                Do Until hi = lo
                    If (lA(lIdxA(hi)) > lItem) = bDsc Then
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
Dim Item        As String

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
    lpStr = VarPtr(Item)
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
        Do While hi > lo
            If Not StrComp(Item, sA(lIdxA(lo)), lCp) = lDr Then
                lIdxA(hi) = lIdxA(lo)
                hi = hi - m1
                Do Until hi = lo
                    If Not StrComp(sA(lIdxA(hi)), Item, lCp) = lDr Then
                        lIdxA(lo) = lIdxA(hi)
                        Exit Do
                    End If
                    hi = hi - m1
                Loop
                '/* Found swaps or out of loop
                If hi = lo Then Exit Do
            End If
            lo = lo + m1
        Loop
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

    If ColumnCount = 1 Then Exit Sub
    If lPos = 0 Then
        ReDim m_HLIStc(0).SubItem(0)
        Exit Sub
    End If
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
        If m_bInsensitiveSort Then
            lSortType = lSortType + 2
        End If
        Select Case lSortType
        '/* ascending case sensitive
        Case 1
            QSIStringSort m_sSortArray, m_lPtr, 0&, 1&
        '/* descending case sensitive
        Case 2
            QSIStringSort m_sSortArray, m_lPtr, 0&, -1&
        '/* ascending case insensitive
        Case 3
            QSIStringSort m_sSortArray, m_lPtr, 1&, 1&
        '/* descending case insensitive
        Case 4
            QSIStringSort m_sSortArray, m_lPtr, 1&, -1&
        End Select
        
    '/* numeric and date
    Case SortDate, SortNumeric
        BuildNumericSortArray lColumn
        If Not ArrayCheck(m_lSortArray) Then GoTo Handler
        If UBound(m_lSortArray) < 2 Then GoTo Handler
        If lSortType = 0 Then lSortType = 1
        QSIInitPtr LBound(m_lSortArray), UBound(m_lSortArray), m_lPtr

        Select Case lSortType
        '/* ascending
        Case 1
            QSINumericSort m_lSortArray, m_lPtr, False
        '/* ascending
        Case 2
            QSINumericSort m_lSortArray, m_lPtr, True
        End Select
        
    Case Else
        GoTo Handler
    End Select

    '/* success
    m_bSorted = True
    ListRefresh True
    SortControl = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("Sort_Control", Err.Number)

End Function

Private Function SortSample(ByVal lColumn As Long) As ECSColumnSortTags
'/* auto determine sort type

Dim vItem As Variant

    If lColumn = 0 Then
        vItem = ItemText(ItemTopIndex)
    Else
        vItem = SubItemText((ItemTopIndex), lColumn)
    End If
    If vItem = "" Then
        SortSample = SortDefault
    ElseIf IsDate(vItem) Then
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

Private Function MoveArrayItem(ByVal lIntPos As Long, _
                               ByVal lDstPos As Long) As Boolean

'/* shift an item in array

Dim lCt As Long
Dim lLb As Long
Dim lUb As Long

On Error GoTo Handler

    '/* cd method
    If m_eListMode = eCustomDraw Then
        Dim cLTemp As clsListItem
        lLb = LBound(m_cListItems)
        lUb = UBound(m_cListItems)
        '/* store item
        Set cLTemp = m_cListItems(lIntPos)
        '/* shift arrays
        If lDstPos > lIntPos Then
            For lCt = lIntPos To lDstPos
                If (lCt + 1) > lUb Then Exit For
                Set m_cListItems(lCt) = m_cListItems(lCt + 1)
            Next lCt
        ElseIf lDstPos < lIntPos Then
            For lCt = lIntPos To lDstPos Step -1
                If (lCt - 1) < lLb Then Exit For
                Set m_cListItems(lCt) = m_cListItems(lCt - 1)
            Next lCt
        End If
        '/* copy to dest
        Set m_cListItems(lDstPos) = cLTemp
    '/* hl method
    ElseIf m_eListMode = eHyperList Then
        lLb = LBound(m_HLIStc(0).Item)
        lUb = UBound(m_HLIStc(0).Item)
        Dim tHItem As HLIStc
        '/* store item data
        With tHItem
            ReDim .Item(0)
            ReDim .lIcon(0)
            ReDim .SubItem(0)
            .Item(0) = m_HLIStc(0).Item(lIntPos)
            .lIcon(0) = m_HLIStc(0).lIcon(lIntPos)
            LSet .SubItem(0) = m_HLIStc(0).SubItem(lIntPos)
        End With
        '/* shift arrays
        If lDstPos > lIntPos Then
            For lCt = lIntPos To lDstPos
                If (lCt + 1) > lUb Then Exit For
                m_HLIStc(0).Item(lCt) = m_HLIStc(0).Item(lCt + 1)
                m_HLIStc(0).lIcon(lCt) = m_HLIStc(0).lIcon(lCt + 1)
                LSet m_HLIStc(0).SubItem(lCt) = m_HLIStc(0).SubItem(lCt + 1)
            Next lCt
        ElseIf lDstPos < lIntPos Then
            For lCt = lIntPos To lDstPos Step -1
                If (lCt - 1) < lLb Then Exit For
                m_HLIStc(0).Item(lCt) = m_HLIStc(0).Item(lCt - 1)
                m_HLIStc(0).lIcon(lCt) = m_HLIStc(0).lIcon(lCt - 1)
                LSet m_HLIStc(0).SubItem(lCt) = m_HLIStc(0).SubItem(lCt - 1)
            Next lCt
        End If
        '/* copy to dest
        With tHItem
            m_HLIStc(0).Item(lDstPos) = .Item(0)
            m_HLIStc(0).lIcon(lDstPos) = .lIcon(0)
            LSet m_HLIStc(0).SubItem(lDstPos) = .SubItem(0)
        End With
    End If
    m_bSorted = False
    
Handler:

End Function

Private Sub AddToStringArray(ByRef sArray() As String, _
                             ByVal sItem As String, _
                             Optional ByVal lPos As Long = -1)

Dim lUb     As Long
Dim lTemp   As Long

    If Not ArrayCheck(sArray) Then Exit Sub
    lUb = UBound(sArray)

    If (lPos > lUb) Or (lPos = -1) Then
        ReDim Preserve sArray(lUb + 1)
        sArray(lUb + 1) = sItem
        Exit Sub
    End If
    
    If lPos < 0 Then lPos = 0
    lUb = lUb + 1
    ReDim Preserve sArray(lUb)
    CopyMemory ByVal VarPtr(sArray(lPos + 1)), ByVal VarPtr(sArray(lPos)), (lUb - lPos) * 4
    lTemp = 0
    CopyMemory ByVal VarPtr(sArray(lPos)), lTemp, 4&
    sArray(lPos) = sItem
   
End Sub

Private Sub AddToLongArray(ByRef lArray() As Long, _
                           ByVal lItem As Long, _
                           Optional ByVal lPos As Long = -1)

Dim lUb     As Long
   
    If Not ArrayCheck(lArray) Then Exit Sub
    lUb = UBound(lArray)
   
    If (lPos > lUb) Or (lPos = -1) Then
        ReDim Preserve lArray(lUb + 1)
        lArray(lUb + 1) = lItem
        Exit Sub
    End If
   
    If lPos < 0 Then lPos = 0
    lUb = lUb + 1
    ReDim Preserve lArray(lUb)
    CopyMemory lArray(lPos + 1), lArray(lPos), (lUb - LBound(lArray) - lPos) * Len(lArray(lPos))
    lArray(lPos) = lItem
   
End Sub


'**********************************************************************
'*                              SUBCLASSING
'**********************************************************************

Private Sub ListAttatch()
'/* attatch messages

    If m_lLVHwnd = 0 Then Exit Sub
    With m_cHListSubclass
        If Not m_lParentHwnd = 0 Then
            .Subclass m_lParentHwnd, Me
            .AddMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
            .AddMessage m_lParentHwnd, WM_SETFOCUS, MSG_BEFORE
            .AddMessage m_lParentHwnd, WM_SIZE, MSG_BEFORE
        End If
        .Subclass m_lLVHwnd, Me
        .AddMessage m_lLVHwnd, WM_KEYDOWN, MSG_BEFORE
        .AddMessage m_lLVHwnd, WM_NOTIFY, MSG_BEFORE
        .AddMessage m_lLVHwnd, WM_MOUSEACTIVATE, MSG_BEFORE
        .AddMessage m_lLVHwnd, WM_NCLBUTTONDOWN, MSG_BEFORE
        If Not m_lEditHwnd = 0 Then
            .Subclass m_lEditHwnd, Me
            .AddMessage m_lEditHwnd, WM_KEYDOWN, MSG_BEFORE
        End If
    End With
    
End Sub

Private Sub ListDetatch()
'/* detatch messages

    If m_lLVHwnd = 0 Then Exit Sub
    With m_cHListSubclass
        If Not m_lParentHwnd = 0 Then
            .DeleteMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
            .DeleteMessage m_lParentHwnd, WM_SETFOCUS, MSG_BEFORE
            .DeleteMessage m_lParentHwnd, WM_SIZE, MSG_BEFORE
            .UnSubclass m_lParentHwnd
        End If
        If Not m_lLVHwnd = 0 Then
            .DeleteMessage m_lLVHwnd, WM_KEYDOWN, MSG_BEFORE
            .DeleteMessage m_lLVHwnd, WM_NOTIFY, MSG_BEFORE
            .DeleteMessage m_lLVHwnd, WM_MOUSEACTIVATE, MSG_BEFORE
            .DeleteMessage m_lLVHwnd, WM_NCLBUTTONDOWN, MSG_BEFORE
            .UnSubclass m_lLVHwnd
        End If
        If Not m_lEditHwnd = 0 Then
            .DeleteMessage m_lEditHwnd, WM_KEYDOWN, MSG_BEFORE
            .UnSubclass m_lEditHwnd
        End If
    End With

End Sub

Private Function DragStartTimer() As Boolean
'/* start header drag timer

    If Not m_bTimerActive Then
        m_cHListSubclass.AddMessage m_lParentHwnd, WM_TIMER, MSG_BEFORE
        SetTimer m_lParentHwnd, 1&, 10&, 0&
        m_bTimerActive = True
    End If
    
End Function

Private Function DragStopTimer() As Boolean
'/* stop header drag timer

    If m_bTimerActive Then
        KillTimer m_lParentHwnd, 1&
        m_cHListSubclass.DeleteMessage m_lParentHwnd, WM_TIMER, MSG_BEFORE
        m_bTimerActive = False
        m_lSafeTimer = 0
    End If
    
End Function

Private Sub DrawColumnDivider()

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
Dim lCode                   As Long
Dim lHandle                 As Long
Dim lItem                   As Long
Dim lLen                    As Long
Dim lZHnd                   As Long
Dim sSubText                As String
Dim sTemp                   As String
Dim tNmhdr                  As NMHDR
Dim tHDN                    As NMHEADER
Dim tNmList                 As NMLISTVIEW
Dim tDisp                   As NMLVDISPINFO
Dim tNKey                   As NMLVKEYDOWN
Dim tPoint                  As POINTAPI
Dim tRect                   As RECT

On Error GoTo Handler

    Select Case uMsg
    '/* control focus
    Case WM_SETFOCUS
    Debug.Print lHwnd
        '/* tab without ipao
        If lHwnd = m_lParentHwnd Then
            SetFocus m_lLVHwnd
            lItem = ItemTopIndex
            ItemEnsureVisible lItem
            ItemSelected(lItem) = True
            EditUpdateLabel
            lReturn = 1
            bHandled = True
        End If
        
    '/* size change
    Case WM_SIZE
        If lHwnd = m_lParentHwnd Then
            Resize
            Exit Sub
        End If
        
    '/* handle tab and esc keys
    Case WM_KEYDOWN
        If lHwnd = m_lLVHwnd Then
            '/* tab
            If (wParam = VK_TAB) Then
                '/* next in z-order
                lZHnd = GetWindow(m_lParentHwnd, GW_HWNDNEXT)
                '/* focus
                If Not lZHnd = 0 Then
                    SetFocus lZHnd
                    RePaint
                Else
                    '/* return to parent
                    lZHnd = GetParent(m_lParentHwnd)
                    If Not lZHnd = 0 Then
                        SetFocus lZHnd
                        RePaint
                    End If
                End If
                bHandled = True
                Exit Sub
            ElseIf (wParam = VK_CONTROL) Then
                lReturn = m_cHListSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
                bHandled = True
                Exit Sub
            '/* esc key
            ElseIf (wParam = VK_ESCAPE) Then
                '/* return to parent
                lZHnd = GetParent(m_lParentHwnd)
                If Not lZHnd = 0 Then
                    SetFocus lZHnd
                End If
                bHandled = True
                Exit Sub
            End If
        '/* edit control
        ElseIf lHwnd = m_lEditHwnd Then
            If (wParam = VK_ESCAPE) Then
                EditUpdateLabel
            ElseIf (wParam = VK_ENTER) Then
                EditUpdateLabel
            End If
        End If
        
    Case WM_MOUSEACTIVATE
        If m_bItemActive Then
            If (GetFocus() = m_lEditHwnd) Then
                lReturn = m_cHListSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
                Focus = True
                lReturn = MA_NOACTIVATE
            End If
        Else
            If Not (GetFocus() = m_lLVHwnd) Then
                Focus = True
                lReturn = MA_NOACTIVATE
            Else
                lReturn = m_cHListSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
            End If
        End If
    
    Case WM_NCLBUTTONDOWN
        '/* refresh when scrolled
        If m_bItemActive Then
            
        Else
            If m_bAlphaSelectorBar Then
                RefreshRegion m_lSelectedItem
            End If
        End If
        
        
    Case WM_TIMER
        HotDivider
        
    '/*** begin notifition messages ***
    Case WM_NOTIFY
        CopyMemory tNmhdr, ByVal lParam, Len(tNmhdr)
        '/* get msg code and owner TTN_FIRST
        With tNmhdr
            lCode = .code
            lHandle = .hwndFrom
        End With

        '/* by origin handle
        Select Case lHandle
        '-> header <-
        Case m_lHdrHwnd
            Select Case lCode
            Case HDN_ITEMCHANGINGW, HDN_ITEMCHANGINGA
                If m_bItemActive = True Then
                    EditUpdateLabel
                End If
                If m_bHeaderFixed Then
                    lReturn = 1
                    bHandled = True
                Else
                    '/* trap first column to finite x position
                    '/* to prevent overpaint by checkboxes
                    GetCursorPos tPoint
                    ScreenToClient m_lLVHwnd, tPoint
                    If tPoint.X < 35 Then
                        bHandled = True
                        lReturn = 1
                    End If
                End If
            
            '/* header size change
            Case HDN_BEGINTRACKW, HDN_BEGINTRACKA
                If m_bAlphaSelectorBar Then
                    RefreshRegion m_lSelectedItem
                End If
            
            '/* header start drag
            Case HDN_BEGINDRAG
                CopyMemory tHDN, ByVal lParam, Len(tHDN)
                If m_bDragDrop Then
                    If m_bSkinHeader Then
                        m_cSkinHeader.DragState = True
                        m_lHotColumn = tHDN.iItem
                        DragStartTimer
                    End If
                End If
                
            '/* refresh after drag
            Case HDN_ENDDRAG
                If m_bDragDrop Then
                    If m_bSkinHeader Then
                        m_cSkinHeader.DragState = False
                        DragStopTimer
                        Resize
                    End If
                End If

            '/* [non skin] custom header colors
            Case NM_CUSTOMDRAW
                If m_bEnabled Then
                    lReturn = CustomDrawHeader(lParam)
                End If
                bHandled = True
            End Select
            
        '-> listview <-
        Case m_lLVHwnd
            Select Case lCode
            '/* custom rows
            Case NM_CUSTOMDRAW
                lReturn = CustomDrawRow(lParam)
                bHandled = True
                
            '/* click events
            Case NM_CLICK, NM_RCLICK
                CopyMemory tNmList, ByVal lParam, Len(tNmList)
                With tNmList
                    '/* report
                    If m_eViewMode = StyleReport Then
                        '/* close subitem edit
                        If m_bItemActive Then
                            EditUpdateLabel
                        End If
                        '/* checkboxes
                        If m_bCheckBoxes Then
                            m_lColumnOffset = ColumnItemOffset(.iItem)
                            '/* refresh last row
                            If m_lColumnOffset > 0 Then
                                RefreshRegion m_lSelectedItem, False
                            End If
                            '/* toggle check array
                            If .ptAction.X < (20 + m_lCheckBoxSkinOffsetX + m_lColumnOffset) And _
                                .ptAction.X > m_lColumnOffset Then
                                If m_bSorted Then
                                    CheckToggle m_lPtr(.iItem)
                                Else
                                    CheckToggle .iItem
                                End If
                            End If
                        End If
                        '/* alpha bar
                        If m_bAlphaSelectorBar Then
                            If m_bCheckBoxes Then
                                If .ptAction.X > m_lColumnOffset + 20 Then
                                    '/* toggle focus
                                    If Not .iItem = m_lSelectedItem Then
                                        ItemFocused(.iItem) = True
                                        ItemFocused(m_lSelectedItem) = False
                                    End If
                                    '/* draw bar
                                    DrawAlphaSelectorBar .iItem
                                End If
                            Else
                               DrawAlphaSelectorBar .iItem
                            End If
                        End If
                    '/* icon mode
                    ElseIf m_eViewMode = StyleIcon Then
                        If m_bSorted Then
                            CheckToggle m_lPtr(.iItem)
                        Else
                            CheckToggle .iItem
                        End If
                    '/* list mode
                    ElseIf m_eViewMode = StyleList Then
                        If m_bSorted Then
                            CheckToggle m_lPtr(.iItem)
                        Else
                            CheckToggle .iItem
                        End If
                    End If
                    '/* selected item
                    If m_bUseSorted And m_bSorted Then
                        m_lSelectedItem = m_lPtr(.iItem)
                    Else
                        m_lSelectedItem = .iItem
                    End If
                    '/* refresh checkbox
                    RefreshRegion .iItem, True
                End With
            
            '/* set edit flag and start editing
            Case NM_DBLCLK
                CopyMemory tNmList, ByVal lParam, Len(tNmList)
                m_lSelectedItem = tNmList.iItem
                If Not m_bItemActive Then
                    If m_bEditLabels Then
                        '/* hit test
                        m_lSubItemEdit = SubItemHitTest
                        '/* item clicked
                        If m_lSubItemEdit = 0 Then
                            If m_bCheckBoxes Then
                                If tNmList.ptAction.X > 20 Then
                                    If m_bUseUnicode Then
                                        SendMessageLongW m_lLVHwnd, LVM_EDITLABELW, m_lSelectedItem, 0&
                                    Else
                                        SendMessageLongA m_lLVHwnd, LVM_EDITLABELA, m_lSelectedItem, 0&
                                    End If
                                End If
                            Else
                                If m_bUseUnicode Then
                                    SendMessageLongW m_lLVHwnd, LVM_EDITLABELW, m_lSelectedItem, 0&
                                Else
                                    SendMessageLongA m_lLVHwnd, LVM_EDITLABELA, m_lSelectedItem, 0&
                                End If
                            End If
                            m_bItemActive = True
                        Else
                            If m_bSubItemsEdit Then
                                m_lEditItem = m_lSelectedItem
                                '/* get subitem rect
                                GetItemRect m_lSelectedItem, tRect, m_lSubItemEdit
                                '/* launch api edit window
                                If Not m_lEditHwnd = 0 Then
                                    '/* populate with subitem text
                                    sSubText = SubItemText(m_lEditItem, m_lSubItemEdit)
                                    '/* position editbox
                                    EditSetPosition tRect
                                    EditShow True
                                    '/* get text
                                    EditSetText sSubText
                                    '/* select
                                    EditSelectText
                                    m_bItemActive = True
                                End If
                            End If
                        End If
                    End If
                Else
                    m_bItemActive = False
                End If
                
            '/* drag and drop trigger
            Case LVN_BEGINDRAG, LVN_BEGINRDRAG
                If Not m_bItemActive Then
                    CopyMemory tNmList, ByVal lParam, Len(tNmList)
                    If m_bUseSorted And m_bSorted Then
                        m_lSelectedItem = m_lPtr(tNmList.iItem)
                    Else
                        m_lSelectedItem = tNmList.iItem
                    End If
                    lReturn = 1
                    bHandled = True
                    UserControl.OLEDrag
                End If
                
            '/* set item edit flag
            Case LVN_BEGINLABELEDITW, LVN_BEGINLABELEDITA
                m_bItemActive = True
                bHandled = True
            
            '/* toggle checkbox with spacebar
            Case LVN_KEYDOWN
                CopyMemory tNKey, ByVal lParam, Len(tNKey)
                If tNKey.wVKey = 32 Then
                    If m_bCheckBoxes Then
                        CheckToggle m_lSelectedItem
                    End If
                    ListRefresh
                ElseIf (tNKey.wVKey = 17) Then
                    '/* multiselect
                    If m_bMultiSelect Then
                        '/* conditional left click
                        If GetKeyState(vbLeftButton) > 1 Then
                            ListRefresh
                        End If
                    End If
                    lReturn = 1
                    bHandled = True
                End If
                    
            '/* column click
            Case LVN_COLUMNCLICK
                If m_bUseSorted Then
                    CopyMemory tNmList, ByVal lParam, Len(tNmList)
                    RaiseEvent eHColumnClick(tNmList.iSubItem)
                    '/* swap sort icon
                    ColumnIconReset
                    If ColumnIcon(tNmList.iSubItem) = -1 Then
                        ColumnIcon(tNmList.iSubItem) = 1
                    ElseIf ColumnIcon(tNmList.iSubItem) = 1 Then
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

            '/* end label edit
            Case LVN_ENDLABELEDITW, LVN_ENDLABELEDITA
                CopyMemory tDisp, ByVal lParam, Len(tDisp)
                With tDisp
                    If m_bSorted Then
                        lItem = m_lPtr(.Item.iItem)
                    Else
                        lItem = .Item.iItem
                    End If
                    sTemp = PointerToString(.Item.pszText)
                    If m_bUseUnicode Then
                        lLen = lstrlenW(StrPtr(sTemp))
                    Else
                        lLen = lstrlenA(sTemp)
                    End If
                    If lLen > 1 Then
                        Select Case m_eListMode
                            Case eCustomDraw
                                m_cListItems(lItem).Text = sTemp
                            Case eHyperList
                                m_HLIStc(0).Item(lItem) = sTemp
                            Case eDatabase
                                RaiseEvent eHLabelChange(lItem, 0, sTemp)
                            End Select
                        End If
                    End With
                    m_bItemActive = False

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
                                If m_eViewMode = StyleReport Then
                                    If m_bCustomDraw Or m_bAlphaSelectorBar Then
                                        ItemSelected(.iItem) = True
                                    End If
                                End If
                                RaiseEvent eHItemClick(.iItem)
                            End If
                        End If
                    End If
                End With
            
            'Case LVN_ODCACHEHINT

            'Case LVN_ITEMACTIVATE

            '/* list change callback
            Case LVN_GETDISPINFOW
                CreateListItemsW lParam
                If EditBoxVisible Then
                    EditBoxMove
                Else
                    EditUpdateLabel
                End If
                
            Case LVN_GETDISPINFOA
                CreateListItems lParam
                If EditBoxVisible Then
                    EditBoxMove
                Else
                    EditUpdateLabel
                End If
            End Select
        End Select
    End Select

Handler:
    On Error GoTo 0

End Sub

Private Function CreateListItems(ByRef lParam As Long) As Long

Dim lItem   As Long
Dim sTemp   As String
Dim tDisp   As NMLVDISPINFOW

    CopyMemory tDisp, ByVal lParam, Len(tDisp)
    With tDisp.Item
        '/* sorting pointer
        If m_bUseSorted And m_bSorted Then
            lItem = m_lPtr(.iItem)
        Else
            lItem = .iItem
        End If
        
        '/* list item text
        If ((.Mask And LVIF_TEXT) = LVIF_TEXT) Then
            Select Case m_eListMode
            Case eCustomDraw
                Select Case .iSubItem
                Case 0
                    sTemp = m_cListItems(lItem).Text
                Case Else
                    sTemp = m_cListItems(lItem).SubItemText(.iSubItem)
                End Select
            Case eHyperList
                Select Case .iSubItem
                Case 0
                    sTemp = m_HLIStc(0).Item(lItem)
                Case Else
                    sTemp = m_HLIStc(0).SubItem(lItem).Text(.iSubItem)
                End Select
            Case eDatabase
                RaiseEvent eHIndirect(lItem, .iSubItem, .Mask, sTemp, .iImage)
            End Select
            
            '/ copy text
            If Len(sTemp) > .cchTextMax Then
                sTemp = left$(sTemp, .cchTextMax)
            End If
            lstrtoptr .pszText, sTemp
        End If
        
        '/* indent
        If ((.Mask And LVIF_INDENT) = LVIF_INDENT) Then
            .iIndent = m_lItemIndent
        End If
        
        '/* subitem image
        If .iSubItem > 0 Then
            If m_bSubItemImage Then
                Select Case m_eListMode
                Case eCustomDraw
                    .iImage = m_cListItems(lItem).SubItemIcon(.iSubItem)
                Case eHyperList
                    .iImage = m_HLIStc(0).SubItem(lItem).lIcon(.iSubItem)
                Case eDatabase
                    RaiseEvent eHIndirect(lItem, .iSubItem, .Mask, sTemp, .iImage)
                End Select
                .Mask = LVIF_IMAGE Or LVIF_TEXT
                .stateMask = LVIS_OVERLAYMASK
            End If
        End If
        
        '/* icon
        If ((.Mask And LVIF_IMAGE) = LVIF_IMAGE) Then
            If .iSubItem = 0 Then
                Select Case m_eListMode
                Case eCustomDraw
                    .iImage = m_cListItems(lItem).Icon
                Case eHyperList
                    .iImage = m_HLIStc(0).lIcon(lItem)
                Case eDatabase
                    RaiseEvent eHIndirect(lItem, .iSubItem, .Mask, sTemp, .iImage)
                End Select
            End If
            
            '/* check state
            If m_bCheckBoxes Then
                If m_bEnabled Then
                    Select Case pItemChecked(lItem)
                    Case 0
                        .State = LVIS_UNCHECKED
                    Case 1
                        .State = LVIS_CHECKED
                    End Select
                Else
                .State = LVIS_DISABLED
                End If
            End If
                
            '/* disabled icon
            If Not m_bEnabled And (m_eViewMode = StyleReport) Then
                IconDisabled .iItem, .iImage
                .iImage = -1
            End If
            
            '/* checkboxes
            If m_bCheckBoxes Then
                If Not m_eViewMode = StyleSmallIcon Then
                    .Mask = LVIF_IMAGE Or LVIF_TEXT Or LVIF_STATE
                    .stateMask = LVIS_OVERLAYMASK Or LVIS_STATEIMAGEMASK
                End If
            End If
        End If
    End With
                
    '/* copy and forward
    CopyMemory ByVal lParam, tDisp, Len(tDisp)

End Function

Private Function CreateListItemsW(ByRef lParam As Long) As Long

Dim lItem   As Long
Dim sTemp   As String
Dim tDisp   As NMLVDISPINFOW

    CopyMemory tDisp, ByVal lParam, Len(tDisp)
    With tDisp.Item
        '/* sorting pointer
        If m_bUseSorted And m_bSorted Then
            lItem = m_lPtr(.iItem)
        Else
            lItem = .iItem
        End If
        
        '/* list item text
        If ((.Mask And LVIF_TEXT) = LVIF_TEXT) Then
            Select Case m_eListMode
            Case eCustomDraw
                Select Case .iSubItem
                Case 0
                    sTemp = m_cListItems(lItem).Text
                Case Else
                    sTemp = m_cListItems(lItem).SubItemText(.iSubItem)
                End Select
            Case eHyperList
                Select Case .iSubItem
                Case 0
                    sTemp = m_HLIStc(0).Item(lItem)
                Case Else
                    sTemp = m_HLIStc(0).SubItem(lItem).Text(.iSubItem)
                End Select
            Case eDatabase
                RaiseEvent eHIndirect(lItem, .iSubItem, .Mask, sTemp, .iImage)
            End Select
            
            '/ copy text
            If Len(sTemp) > .cchTextMax Then
                sTemp = left$(sTemp, .cchTextMax)
            End If
            StringToPointer sTemp, .pszText
        End If
        
        '/* indent
        If ((.Mask And LVIF_INDENT) = LVIF_INDENT) Then
            .iIndent = m_lItemIndent
        End If
        
        '/* subitem image
        If .iSubItem > 0 Then
            If m_bSubItemImage Then
                Select Case m_eListMode
                Case eCustomDraw
                    .iImage = m_cListItems(lItem).SubItemIcon(.iSubItem)
                Case eHyperList
                    .iImage = m_HLIStc(0).SubItem(lItem).lIcon(.iSubItem)
                Case eDatabase
                    RaiseEvent eHIndirect(lItem, .iSubItem, .Mask, sTemp, .iImage)
                End Select
                .Mask = LVIF_IMAGE Or LVIF_TEXT
                .stateMask = LVIS_OVERLAYMASK
            End If
        End If
        
        '/* icon
        If ((.Mask And LVIF_IMAGE) = LVIF_IMAGE) Then
            If .iSubItem = 0 Then
                Select Case m_eListMode
                Case eCustomDraw
                    .iImage = m_cListItems(lItem).Icon
                Case eHyperList
                    .iImage = m_HLIStc(0).lIcon(lItem)
                Case eDatabase
                    RaiseEvent eHIndirect(lItem, .iSubItem, .Mask, sTemp, .iImage)
                End Select
            End If
            
            '/* check state
            If m_bCheckBoxes Then
                If m_bEnabled Then
                    Select Case pItemChecked(lItem)
                    Case 0
                        .State = LVIS_UNCHECKED
                    Case 1
                        .State = LVIS_CHECKED
                    End Select
                Else
                    .State = LVIS_DISABLED
                End If
            End If
                
            '/* disabled icon
            If Not m_bEnabled And (m_eViewMode = StyleReport) Then
                IconDisabled .iItem, .iImage
                .iImage = -1
            End If
            
            '/* checkboxes
            If m_bCheckBoxes Then
                If Not m_eViewMode = StyleSmallIcon Then
                    .Mask = LVIF_IMAGE Or LVIF_TEXT Or LVIF_STATE
                    .stateMask = LVIS_OVERLAYMASK Or LVIS_STATEIMAGEMASK
                End If
            End If
        End If
    End With
                
    '/* copy and forward
    CopyMemory ByVal lParam, tDisp, Len(tDisp)

End Function


'**********************************************************************
'*                              CUSTOM DRAW
'**********************************************************************

Private Function CustomDrawRow(ByVal lParam As Long) As Long

Dim lReturn     As Long
Dim tNmLvCd     As NMLVCUSTOMDRAW
    
    CopyMemory tNmLvCd, ByVal lParam, Len(tNmLvCd)
    '/* using alpha bar and row colors
    If m_bCustomDraw And m_bAlphaSelectorBar Then
        With tNmLvCd.nmcmd
            If ((.iItemState And CDIS_SELECTED) = CDIS_SELECTED) Then
                If ((.iItemState And CDIS_FOCUS) = CDIS_FOCUS) Then
                    lReturn = AlphaCustomRow(lParam)
                Else
                    lReturn = CustomRow(lParam)
                End If
            Else
                lReturn = CustomRow(lParam)
            End If
        End With
    '/* alpha only
    ElseIf m_bAlphaSelectorBar Then
        lReturn = AlphaCustomRow(lParam)
    '/* rows only
    ElseIf m_bCustomDraw Then
        lReturn = CustomRow(lParam)
    Else
        lReturn = CDRF_DODEFAULT
    End If
    
    CustomDrawRow = lReturn

End Function

Private Function AlphaCustomRow(ByRef lParam As Long) As Long
'/* a lot of guesswork..

Dim tNmLvCd     As NMLVCUSTOMDRAW

    CopyMemory tNmLvCd, ByVal lParam, Len(tNmLvCd)
    With tNmLvCd
        Select Case .nmcmd.dwDrawStage
        Case CDDS_PREPAINT
            AlphaCustomRow = CDRF_NOTIFYITEMDRAW

        Case CDDS_ITEMPREPAINT
            CopyMemory ByVal lParam, tNmLvCd, Len(tNmLvCd)
            With tNmLvCd
                If (.nmcmd.iItemState And CDIS_SELECTED) = CDIS_SELECTED Then
                    If (.nmcmd.iItemState And CDIS_FOCUS) = CDIS_FOCUS Then
                        If m_bAlphaThemeBackClr Then
                            .clrTextBk = XPShift(m_oThemeColor, 190)
                        ElseIf m_bRowDecoration Then
                            RowColors .nmcmd.dwItemSpec, .iSubItem, .clrTextBk
                        Else
                            .clrTextBk = m_oBackColor
                        End If
                        ItemSelected(.nmcmd.dwItemSpec) = False
                    End If
                    CopyMemory ByVal lParam, tNmLvCd, Len(tNmLvCd)
                End If
            End With
        Case Else
            lParam = CDRF_DODEFAULT
        End Select
    End With
    
End Function

Private Function CustomRow(ByRef lParam As Long) As Long
' http://www.codeproject.com/listctrl/lvcustomdraw.asp

Dim clrTextBk   As OLE_COLOR
Dim clrText     As OLE_COLOR
Dim tNmLvCd     As NMLVCUSTOMDRAW

    CopyMemory tNmLvCd, ByVal lParam, Len(tNmLvCd)
    With tNmLvCd
        Select Case .nmcmd.dwDrawStage
        Case CDDS_PREPAINT
            CustomRow = CDRF_NOTIFYITEMDRAW
            
        Case CDDS_ITEMPREPAINT
              CustomRow = CDRF_NOTIFYSUBITEMDRAW
        
        Case CDDS_ITEMPREPAINT Or CDDS_SUBITEM
        
            With m_cListItems(tNmLvCd.nmcmd.dwItemSpec)
                '/* per cell font
                If m_bUseCellFont Then
                    If Not (.Font Is Nothing) Then
                        RowFont tNmLvCd.nmcmd.hdc, .Font
                    End If
                End If
                '/* custom cell colors
                If .CellCustom And m_bUseCellColor Then
                    clrTextBk = .CellBackColor(tNmLvCd.iSubItem)
                    clrText = .CellForeColor(tNmLvCd.iSubItem)
                '/* row decoration
                ElseIf m_bRowDecoration Then
                    RowColors tNmLvCd.nmcmd.dwItemSpec, tNmLvCd.iSubItem, clrTextBk
                    clrText = m_oForeColor
                Else
                    clrTextBk = .BackColor
                    clrText = .ForeColor
                End If
            End With
            '/* test and apply color formatting
            If Not clrTextBk = -1 Then
                OleTranslateColor clrTextBk, 0, .clrTextBk
            End If
            If Not clrText = -1 Then
                OleTranslateColor clrText, 0, .clrText
            End If
            
            CopyMemory ByVal lParam, tNmLvCd, Len(tNmLvCd)
        Case Else
            lParam = CDRF_DODEFAULT
        End Select
    End With

End Function

Private Function CustomDrawHeader(ByVal lParam As Long) As Long

Dim lReturn     As Long
Dim tNmLvCd     As NMLVCUSTOMDRAW

    CopyMemory tNmLvCd, ByVal lParam, Len(tNmLvCd)
    If Not m_bSkinHeader Then
        If m_bCustomHeader Then
            Select Case tNmLvCd.nmcmd.dwDrawStage
            Case CDDS_PREPAINT
                lReturn = CDRF_NOTIFYITEMDRAW
            Case CDDS_ITEMPREPAINT
                SetTextColor tNmLvCd.nmcmd.hdc, m_oHdrForeClr
                SetBkColor tNmLvCd.nmcmd.hdc, m_oHdrBkClr
            Case CDDS_ITEMPOSTPAINT
                lReturn = CDRF_DODEFAULT
            End Select
        End If
    End If
            
    CustomDrawHeader = lReturn
            
End Function

Private Sub DrawItemText(ByVal lItem As Long, _
                         ByVal sText As String)

'/* custom draw cell text

Dim lTmpDc  As Long
Dim hFntOld As Long
Dim sTemp   As String
Dim tRect   As RECT
Dim tPnt    As POINTAPI

    If Len(sText) = 0 Then Exit Sub
    GetItemRect lItem, tRect
    lTmpDc = GetDC(m_lLVHwnd)
    hFntOld = SelectObject(lTmpDc, m_lFont)
    SetBkMode lTmpDc, 1
    SetTextColor lTmpDc, m_oForeColor

    With tRect
        .left = 36
        .right = (ColumnWidth(0) - 2)
    End With

    If m_bUseUnicode Then
        GetTextExtentPoint32W lTmpDc, StrPtr(StrPtr(sText)), lstrlenW(StrPtr(sText)), tPnt
    Else
        GetTextExtentPoint32A lTmpDc, sText, lstrlenA(sText), tPnt
    End If

    With tRect
        '/* test min size
        If (.right - .left) < (tPnt.X + 40) Then
            sTemp = String(255, Chr$(0))
            sTemp = sText & vbNullChar
            '/* compact text
            If m_bUseUnicode Then
                PathCompactPathW lTmpDc, StrPtr(sTemp), (tPnt.X - 10)
            Else
                PathCompactPathA lTmpDc, sTemp, (tPnt.X - 10)
            End If
        Else
            sTemp = sText
        End If
    End With
    
    If m_bUseUnicode Then
        DrawTextW lTmpDc, StrPtr(sTemp), -1, tRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORDBREAK
    Else
        DrawTextA lTmpDc, sTemp, 0&, tRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    End If

    If Not hFntOld = 0 Then
        'SelectObject m_lHdc, hFntOld
        hFntOld = 0
    End If
    ReleaseDC m_lLVHwnd, lTmpDc

End Sub

Private Sub DrawWordBreak(ByVal lItem As Long, _
                          ByVal lSubItem As Long, _
                          ByVal lHdc As Long, _
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
                FillRect lHdc, tRect, hBr
                DeleteObject hBr
                sText = SubItemText(lItem, lSubItem)
                If m_bUseUnicode Then
                    DrawTextW lHdc, StrPtr(sText), -1, tRect, DT_LEFT Or DT_WORDBREAK Or DT_SINGLELINE
                Else
                    DrawTextA lHdc, sText, 0&, tRect, DT_LEFT Or DT_WORDBREAK Or DT_CALCRECT
                End If
            End If
        End With
    End If

End Sub

Private Function IndexToStateImageMask(ByVal lState As Long) As Long
'/* state mask translate
   IndexToStateImageMask = lState * (2 ^ 12)
End Function

Private Function RowColors(ByVal lRow As Long, _
                           ByVal lSubItem As Long, _
                           ByRef lRowClr As Long)

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
            If (lRow < LBound(m_HLIStc(0).Item)) Or (lRow > UBound(m_HLIStc(0).Item)) Then
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

Public Function RowDecoration(ByVal eRowDecoration As ERDRowDecoration, _
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

Private Function RowFont(ByVal lHdc As Long, _
                         ByVal oFont As StdFont)
'*/ change row font

Dim lChar   As Long
Dim lFont   As Long
Dim uLF     As LOGFONT

On Error GoTo Handler

    With uLF
         For lChar = 1 To Len(oFont.Name)
             .lfFaceName(lChar - 1) = CByte(Asc(Mid$(oFont.Name, lChar, 1)))
         Next lChar
         .lfHeight = -MulDiv(oFont.SIZE, GetDeviceCaps(UserControl.hdc, LOGPIXELSY), 72)
         .lfItalic = oFont.Italic
         .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
         .lfUnderline = oFont.Underline
         .lfStrikeOut = oFont.Strikethrough
         .lfCharSet = oFont.Charset
    End With
    If m_bUseUnicode Then
        lFont = CreateFontIndirectW(uLF)
    Else
        lFont = CreateFontIndirectA(uLF)
    End If
    SelectObject lHdc, lFont
    DeleteObject lFont

Handler:
    On Error GoTo 0

End Function


'**********************************************************************
'*                              EDIT CONTROL
'**********************************************************************
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/editcontrols/abouteditcontrols.asp

Public Function RowsPerPage() As Long
'/* get first row item index

    If Not m_lLVHwnd = 0 Then
        RowsPerPage = SendMessageLongA(m_lLVHwnd, LVM_GETCOUNTPERPAGE, 0&, 0&)
    End If
    
End Function

Public Function RowTopIndex() As Long
'/* get first row item index

    If Not m_lLVHwnd = 0 Then
        RowTopIndex = SendMessageLongA(m_lLVHwnd, LVM_GETTOPINDEX, 0&, 0&)
    End If

End Function

Private Function EditBoxVisible() As Boolean

Dim lRlw As Long
Dim lRct As Long

    lRlw = RowTopIndex
    lRct = lRlw + RowsPerPage
    If Not m_lSelectedItem > lRct Then
        If Not m_lSelectedItem < lRlw Then
            EditBoxVisible = True
        End If
    End If
    
End Function

Private Function EditBoxMove() As Boolean

Dim tRect As RECT

    GetItemRect m_lSelectedItem, tRect, m_lSubItemEdit
    EditSetPosition tRect

End Function

Private Function EditShow(ByVal bVisible As Boolean) As Boolean
'/* show api edit box

Dim lStyle      As Long

On Error GoTo Handler

    If m_lEditHwnd = 0 Then Exit Function
    If m_bUseUnicode Then
        lStyle = GetWindowLongW(m_lEditHwnd, GWL_STYLE)
    Else
        lStyle = GetWindowLongA(m_lEditHwnd, GWL_STYLE)
    End If
    '/* set the style bit
    If bVisible Then
        lStyle = lStyle Or WS_VISIBLE
    Else
        lStyle = lStyle And Not WS_VISIBLE
    End If
    If m_bUseUnicode Then
        SetWindowLongW m_lEditHwnd, GWL_STYLE, lStyle
    Else
        SetWindowLongA m_lEditHwnd, GWL_STYLE, lStyle
    End If
    If bVisible Then
        SetFocus m_lEditHwnd
    Else
        If m_bBackgroundBg Then
            ListRefresh True
        Else
            ListRefresh
        End If
    End If
    EditShow = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("EditShow", Err.Number)
    
End Function

Private Function EditGetText() As String
'/* get edit box text

Dim lLen    As Long
Dim sText   As String

    If m_bUseUnicode Then
        lLen = GetWindowTextLengthW(m_lEditHwnd)
    Else
        lLen = GetWindowTextLengthA(m_lEditHwnd)
    End If
    
    If (lLen > 0) Then
        lLen = lLen + 1
        sText = String$(lLen, Chr$(0))
        If m_bUseUnicode Then
            GetWindowTextW m_lEditHwnd, StrPtr(sText), lLen
        Else
            GetWindowTextA m_lEditHwnd, sText, lLen
        End If
        lLen = InStr(1, sText, vbNullChar)
        If lLen > 0 Then
            sText = left$(sText, (lLen - 1))
        End If
        EditGetText = sText
    Else
        EditGetText = ""
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("EditGetText", Err.Number)

End Function

Private Function EditSetText(ByVal sText As String) As Boolean
'/* set edit box text
'/* thanks to zhu for solving unicode editbox text problem

Dim lPtr As Long

On Error GoTo Handler

    If LenB(sText) = 0 Then
        SetWindowTextA m_lEditHwnd, ""
        Exit Function
    End If
    If m_bUseUnicode Then
        lPtr = StrPtr(sText)
        SetWindowTextW m_lEditHwnd, StrPtr(sText)
    Else
        SetWindowTextA m_lEditHwnd, sText
    End If
    EditSetText = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("EditSetText", Err.Number)

End Function

Private Function EditSetFont() As Boolean
'/* set editbox font

Dim lhFnt As Long

On Error GoTo Handler

    If m_lEditHwnd = 0 Then Exit Function
    If m_bUseUnicode Then
        '/* get listview font handle
        lhFnt = SendMessageLongW(m_lLVHwnd, WM_GETFONT, 0&, 0&)
        If lhFnt = 0 Then Exit Function
        '/* set editbox font
        SendMessageLongW m_lEditHwnd, WM_SETFONT, lhFnt, True
    Else
        lhFnt = SendMessageLongA(m_lLVHwnd, WM_GETFONT, 0&, 0&)
        If lhFnt = 0 Then Exit Function
        SendMessageLongA m_lEditHwnd, WM_SETFONT, lhFnt, True
    End If
    m_bEditFontSet = True
    EditSetFont = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("EditSetFont", Err.Number)

End Function

Private Function EditSetPosition(ByRef tRect As RECT) As Boolean
'/* position edit box

On Error GoTo Handler

    If m_lEditHwnd = 0 Then Exit Function
    '/* position edit window
    InflateRect tRect, 1, 1
    With tRect
        SetWindowPos m_lEditHwnd, 0&, .left, .top + 1, _
            (.right - .left), (.bottom - .top) - 2, SWP_NOZORDER Or SWP_NOOWNERZORDER
    End With

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("EditSetPosition", Err.Number)

End Function

Private Function EditSelectText() As Boolean
'/* select editbox text

Dim lPos As Long

On Error GoTo Handler

    If m_lEditHwnd = 0 Then Exit Function
    If m_bUseUnicode Then
        lPos = GetWindowTextLengthW(m_lEditHwnd)
        SendMessageA m_lEditHwnd, EM_SETSEL, 0&, ByVal lPos
    Else
        lPos = GetWindowTextLengthA(m_lEditHwnd)
        SendMessageW m_lEditHwnd, EM_SETSEL, 0&, ByVal lPos
    End If
    EditSelectText = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("EditSelectText", Err.Number)

End Function

Private Sub EditUpdateLabel()
'/* update listview cell

Dim sSubText As String

    If Not IsWindowVisible(m_lEditHwnd) = 0 Then
        If m_bItemActive Then
            '/* get text
            sSubText = EditGetText
            '/* sorting pointer
            If m_bUseSorted And m_bSorted Then
                m_lEditItem = m_lPtr(m_lEditItem)
            End If
            '/* write to array
            Select Case m_eListMode
            Case eCustomDraw
                m_cListItems(m_lEditItem).SubItemText(m_lSubItemEdit) = sSubText
            Case eHyperList
                m_HLIStc(0).SubItem(m_lEditItem).Text(m_lSubItemEdit) = sSubText
            Case eDatabase
                RaiseEvent eHLabelChange(m_lEditItem, m_lSubItemEdit, sSubText)
            End Select
            '/* hide and clear editbox
            EditShow False
            m_lSubItemEdit = 0
            m_lEditItem = 0
            EditSetText ""
            SetFocus m_lLVHwnd
            m_bItemActive = False
        End If
    End If
        
End Sub


'**********************************************************************
'*                              DRAG AND DROP
'**********************************************************************

Private Sub UserControl_OLECompleteDrag(Effect As Long)
'/* drag completed
    m_cDrag.CompleteDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, _
                                    Effect As Long, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

'/* complete drag operation

Dim lIndex  As Long
Dim tLVHT   As LVHITTESTINFO
Dim tPoint  As POINTAPI

    '/* get target item index
    If (Effect And vbDropEffectMove) = vbDropEffectMove Then
        If ValidateDragData(Data) Then
            GetCursorPos tPoint
            ScreenToClient m_lLVHwnd, tPoint
            lIndex = -1
            LSet tLVHT.pt = tPoint
            SendMessageA m_lLVHwnd, LVM_HITTEST, 0&, tLVHT
            If (tLVHT.iItem <= 0) Then
                If (tLVHT.flags And LVHT_NOWHERE) = LVHT_NOWHERE Then
                    lIndex = FindNearestItem(tPoint)
                Else
                    lIndex = tLVHT.iItem
                End If
            Else
               lIndex = tLVHT.iItem
            End If
            m_cDrag.CompleteDrag
        End If
    End If

    '/* swap items
    If lIndex > -1 Then
        If m_eListMode = eDatabase Then
            RaiseEvent eHDragComplete(m_lSelectedItem, lIndex)
        Else
            MoveArrayItem m_lSelectedItem, lIndex
        End If
    End If
    '/* refresh
    If Not m_cSkinScrollBars Is Nothing Then
        m_cSkinScrollBars.Refresh
    End If
    ItemGhosted(lIndex) = False
    ItemFocused(lIndex) = True
    ListRefresh

End Sub

Private Function OverItem() As Long

Dim lIndex  As Long
Dim tLVHT   As LVHITTESTINFO
Dim tPoint  As POINTAPI

    '/* get target item index
    GetCursorPos tPoint
    ScreenToClient m_lLVHwnd, tPoint
    lIndex = -1
    LSet tLVHT.pt = tPoint
    SendMessageA m_lLVHwnd, LVM_HITTEST, 0&, tLVHT
    If (tLVHT.iItem <= 0) Then
        If (tLVHT.flags And LVHT_NOWHERE) = LVHT_NOWHERE Then
            lIndex = -1
        Else
            lIndex = tLVHT.iItem
        End If
    Else
        lIndex = tLVHT.iItem
    End If
    OverItem = lIndex
    
End Function

Private Function FindNearestItem(ByRef tPoint As POINTAPI) As Long
'/* return closest item index

Dim i           As Long
Dim lX          As Long
Dim lY          As Long
Dim lCt         As Long
Dim lDistSq     As Long
Dim lMinDistSq  As Long
Dim lMinItem    As Long
Dim tRect       As RECT

On Error GoTo Handler

    lMinItem = -1
    lMinDistSq = &H7FFFFFFF
    lCt = Count
    For i = 1 To Count
        If Not m_eViewMode = StyleIcon Then
            tRect.left = LVIR_BOUNDS
        End If
        SendMessageA m_lLVHwnd, LVM_GETITEMRECT, i - 1, tRect
        With tRect
            lX = tPoint.X - (.left + (.right - .left) \ 2)
            lY = tPoint.Y - (.top + (.bottom - .top) \ 2)
        End With
        lDistSq = lX * lX + lY * lY
        If (lDistSq < lMinDistSq) Then
            lMinDistSq = lDistSq
            lMinItem = i
            Exit For
        End If
    Next i
    FindNearestItem = lMinItem

Handler:

End Function

Private Sub UserControl_OLEDragOver(Data As DataObject, _
                                    Effect As Long, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single, _
                                    State As Integer)

'/* move scrollbars during drag

Dim lIndex      As Long
Dim tPoint      As POINTAPI
Dim tRect       As RECT
Static lLstIdx  As Long

    '/* scroll list when required
    If ValidateDragData(Data) Then
        GetCursorPos tPoint
        GetWindowRect m_lLVHwnd, tRect
        '/* vertical scroll
        If LVHasVertical Then
            If Abs(tPoint.Y - tRect.top) < 24 Then
                m_cDrag.HideDragImage True
                LVScrollVertical False
                m_cDrag.HideDragImage False
            ElseIf Abs(tPoint.Y - tRect.bottom) < 24 Then
                m_cDrag.HideDragImage True
                LVScrollVertical True
                m_cDrag.HideDragImage False
            End If
        '/* horizontal scroll
        ElseIf LVHasHorizontal Then
            If Abs(tPoint.X - tRect.left) < 24 Then
                m_cDrag.HideDragImage True
                LVScrollHorizontal False
                m_cDrag.HideDragImage False
            ElseIf Abs(tPoint.X - tRect.right) < 24 Then
                m_cDrag.HideDragImage True
                LVScrollHorizontal True
                m_cDrag.HideDragImage False
            End If
        End If
    End If
    '/* highlite drag over items
    lIndex = OverItem
    If Not OverItem = -1 Then
        If Not lIndex = lLstIdx Then
            ItemGhosted(lIndex) = True
            ItemGhosted(lLstIdx) = False
            lLstIdx = lIndex
        End If
    End If

End Sub

Private Function BuildDragData(ByVal lItem As Long) As String

Dim lCt     As Long
Dim lCol    As Long
Dim sTemp   As String

    lCol = (ColumnCount - 1)
    sTemp = CStr(ItemIcon(lItem)) & vbCrLf
    sTemp = sTemp & ItemText(lItem) & vbCrLf
    For lCt = 1 To lCol
        sTemp = sTemp & SubItemText(lItem, lCt) & vbCrLf
    Next lCt
    sTemp = left$(sTemp, (Len(sTemp) - 1))
    BuildDragData = sTemp
    
End Function

Private Function ValidateDragData(ByRef Data As DataObject) As Boolean
'/* test for valid item

Dim bData() As Byte
Dim iPos    As Long
Dim sData   As String

On Error Resume Next

    bData = Data.GetData(vbCFText)
    sData = bData
    On Error GoTo 0
    iPos = InStr(sData, vbCrLf)
    If iPos > 0 Then
        ValidateDragData = (Len(sData) > 1)
    End If

End Function

Private Function LVScrollVertical(ByVal bDown As Boolean)
'/* scroll vertical

    If bDown Then
        SendMessageLongA m_lLVHwnd, WM_VSCROLL, SB_LINEDOWN, 0
    Else
        SendMessageLongA m_lLVHwnd, WM_VSCROLL, SB_LINEUP, 0
    End If
    
End Function

Private Function LVScrollHorizontal(ByVal bRight As Boolean)
'/* scroll horizontal

    If bRight Then
        SendMessageLongA m_lLVHwnd, WM_HSCROLL, SB_LINERIGHT, 0
    Else
        SendMessageLongA m_lLVHwnd, WM_HSCROLL, SB_LINELEFT, 0
    End If
    
End Function

Private Function LVHasHorizontal() As Boolean
'/* vertical scrollbar test

Dim lStyle  As Long

    If m_bUseUnicode Then
        lStyle = GetWindowLongW(m_lLVHwnd, GWL_STYLE)
    Else
        lStyle = GetWindowLongA(m_lLVHwnd, GWL_STYLE)
    End If
    LVHasHorizontal = (lStyle And WS_HSCROLL) <> 0

End Function

Private Function LVHasVertical() As Boolean
'/* horizontal scrollbar test

Dim lStyle  As Long
    
    If m_bUseUnicode Then
        lStyle = GetWindowLongW(m_lLVHwnd, GWL_STYLE)
    Else
        lStyle = GetWindowLongA(m_lLVHwnd, GWL_STYLE)
    End If
    LVHasVertical = (lStyle And WS_VSCROLL) <> 0

End Function

Private Sub UserControl_OLEGiveFeedback(Effect As Long, _
                                        DefaultCursors As Boolean)

'/* refresh drag image
    m_cDrag.DragDrop

End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, _
                                     AllowedEffects As Long)

'/* initialize drag operation

Dim bData() As Byte
Dim lIcon   As Long
Dim lIcHnd  As Long
Dim sData   As String

On Error GoTo Handler

    AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
    '/* collect data
    sData = BuildDragData(m_lSelectedItem)
    If m_eListMode = eDatabase Then
        RaiseEvent eHDragging(m_lSelectedItem)
    End If
    '/* convert
    bData = sData
    If m_eViewMode = StyleIcon Then
        lIcHnd = m_lImlLargeHndl
    Else
        lIcHnd = m_lImlSmallHndl
    End If
    '/* add item text
    With Data
        .Clear
        .SetData bData, vbCFText
    End With
    '/* start drag
    With m_cDrag
        .Parent = UserControl.hwnd
        .hImagelist = lIcHnd
        .StartDrag lIcon, -8, -8
    End With
    ItemSelected(m_lSelectedItem) = False

Handler:

End Sub


'**********************************************************************
'*                              CLEANUP
'**********************************************************************

Private Function DeAllocatePointer(ByVal sKey As String, _
                                   Optional ByVal bPurge As Boolean) As Boolean

'/* resolve or purge memory pointers

Dim lPtr    As Long
Dim lC      As Long

On Error GoTo Handler

    If c_PtrMem Is Nothing Then Exit Function
    If Not bPurge Then
        '/* get the pointer
        lPtr = c_PtrMem.Item(sKey)
        If lPtr = 0 Then GoTo Handler
        '/* release the memory
        CopyMemory ByVal lPtr, 0&, 4&
    Else
        '/* destroy the struct last
        For lC = c_PtrMem.Count To 1 Step -1
            If Not CLng(c_PtrMem.Item(lC)) = 0 Then
                lPtr = CLng(c_PtrMem.Item(lC))
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
    If Not m_lEditHwnd = 0 Then
        DestroyWindow m_lEditHwnd
        m_lEditHwnd = 0
    End If
    If Not m_lLVHwnd = 0 Then
        If DestroyWindow(m_lLVHwnd) Then
            DestroyList = True
            m_lLVHwnd = 0
        End If
    End If

End Function

Private Sub DestroyItems()
'/* destroy item classes

Dim lCt     As Long
Dim lUb     As Long

On Error GoTo Handler

    Select Case m_eListMode
    '/* cd mode
    Case eCustomDraw
        '/* destroy classes
        lUb = UBound(m_cListItems)
        For lCt = 0 To lUb
            Set m_cListItems(lCt) = Nothing
        Next lCt
    '/* hl mode
    Case eHyperList
        '/* rease structs
        Erase m_HLIStc(0).Item
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

    If Not m_IChecked Is Nothing Then Set m_IChecked = Nothing
    If Not m_cChkCheckDc Is Nothing Then Set m_cChkCheckDc = Nothing
    If Not m_cRender Is Nothing Then Set m_cRender = Nothing
    If Not m_cDrag Is Nothing Then Set m_cDrag = Nothing

End Sub

Private Sub UserControl_Hide()

    If m_bSkinScrollBars Then
        If Not m_cSkinScrollBars Is Nothing Then
            m_cSkinScrollBars.Visible = False
        End If
    End If
    
End Sub

Private Sub UserControl_Show()

    If m_bSkinScrollBars Then
        If Not m_cSkinScrollBars Is Nothing Then
            If m_cSkinScrollBars.Visible = False Then
                m_cSkinScrollBars.Visible = True
                UserControl_Resize
            End If
        End If
    End If
    
End Sub

Private Sub pInitialize()

Dim bRun As Boolean

    bRun = UserControl.Ambient.UserMode
    If bRun Then
        CreateList
    End If
    
End Sub

Private Sub UserControl_Paint()

   If m_lLVHwnd = 0 Then
      If Not (lblName.Caption = UserControl.Extender.Name) Then
         lblName.Caption = UserControl.Extender.Name
      End If
   End If
   
End Sub

Private Sub UserControl_Resize()

Dim tRect As RECT
  
    If Not m_lLVHwnd = 0 Then
        GetClientRect m_lParentHwnd, tRect
        With tRect
            SetWindowPos m_lLVHwnd, 0, .left, .top, .right - .left, .bottom - .top, SWP_NOZORDER Or SWP_NOOWNERZORDER
        End With
    End If
    Resize
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

    pInitialize
    With PropBag
        Dim sFont As New StdFont
        Set Font = .ReadProperty("Font", sFont)
        Set UserControl.Font = Font
        AlphaBarTheme = .ReadProperty("AlphaBarTheme", False)
        AlphaBarTransparency = .ReadProperty("AlphaBarTransparency", PRP_APT)
        AlphaBarActive = .ReadProperty("AlphaBarActive", False)
        AlphaThemeBackClr = .ReadProperty("AlphaThemeBackClr", False)
        AutoArrange = .ReadProperty("AutoArrange", False)
        BackColor = .ReadProperty("BackColor", &HFFFFFF)
        BorderStyle = .ReadProperty("BorderStyle", PRP_BRDSTL)
        CheckBoxes = .ReadProperty("Checkboxes", False)
        CheckBoxSkinStyle = .ReadProperty("CheckBoxSkinStyle", PRP_CHKSTL)
        CustomDraw = .ReadProperty("CustomDraw", False)
        Enabled = .ReadProperty("Enabled", True)
        SubItemsEdit = .ReadProperty("SubItemsEdit", False)
        ForeColor = .ReadProperty("ForeColor", vbWindowText)
        FullRowSelect = .ReadProperty("FullRowSelect", False)
        GridLines = .ReadProperty("GridLines", False)
        HeaderColor = .ReadProperty("HeaderColor", vbButtonFace)
        HeaderCustom = .ReadProperty("HeaderCustom", False)
        HeaderDragDrop = .ReadProperty("HeaderDragDrop", True)
        HeaderFixedWidth = .ReadProperty("HeaderFixedWidth", False)
        HeaderFlat = .ReadProperty("HeaderFlat", False)
        HeaderForeColor = .ReadProperty("HeaderForeColor", vbWindowText)
        HeaderHide = .ReadProperty("HeaderHide", False)
        HeaderHighLite = .ReadProperty("HeaderHighLite", vbButtonShadow)
        HeaderPressed = .ReadProperty("HeaderPressed", vbWindowText)
        InfoTips = .ReadProperty("InfoTips", False)
        ItemBorderSelect = .ReadProperty("ItemBorderSelect", False)
        ItemIndent = .ReadProperty("ItemIndent", PRP_ITMIND)
        LabelEdit = .ReadProperty("LabelEdit", False)
        LabelTips = .ReadProperty("LabelTips", False)
        ListMode = .ReadProperty("ListMode", PRP_LSTMDE)
        MultiSelect = .ReadProperty("MultiSelect", False)
        OLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
        OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
        OneClickActivate = .ReadProperty("OneClickActivate", False)
        ScrollBarFlat = .ReadProperty("ScrollBarFlat", False)
        SubItemImages = .ReadProperty("SubItemImages", False)
        TextAlignment = .ReadProperty("TextAlignment", PRP_TXTALN)
        ThemeColor = .ReadProperty("ThemeColor", PRP_TMCLR)
        ThemeLuminence = .ReadProperty("ThemeLuminence", PRP_TMLMC)
        TrackSelected = .ReadProperty("TrackSelected", False)
        UnderlineHot = .ReadProperty("UnderlineHot", False)
        UseCellColor = .ReadProperty("UseCellColor", False)
        UseCellFont = .ReadProperty("UseCellFont", False)
        UseUnicode = .ReadProperty("UseUnicode", False)
        UseThemeColors = .ReadProperty("UseThemeColors", True)
        ViewMode = .ReadProperty("ViewMode", PRP_VWEMDE)
        XPColors = .ReadProperty("XPColors", True)
    End With

On Error GoTo 0

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

On Error Resume Next

    With PropBag
        Dim sFont As New StdFont
        .WriteProperty "Font", Font, sFont
        .WriteProperty "AlphaBarTheme", AlphaBarTheme, False
        .WriteProperty "AlphaBarTransparency", AlphaBarTransparency, PRP_APT
        .WriteProperty "AlphaBarActive", AlphaBarActive, False
        .WriteProperty "AlphaThemeBackClr", AlphaThemeBackClr, False
        .WriteProperty "AutoArrange", AutoArrange, False
        .WriteProperty "BackColor", BackColor, &HFFFFFF
        .WriteProperty "BorderStyle", BorderStyle, PRP_BRDSTL
        .WriteProperty "Checkboxes", CheckBoxes, False
        .WriteProperty "CheckBoxSkinStyle", CheckBoxSkinStyle, PRP_CHKSTL
        .WriteProperty "CustomDraw", CustomDraw, False
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "SubItemsEdit", SubItemsEdit, False
        .WriteProperty "ForeColor", ForeColor, vbWindowText
        .WriteProperty "FullRowSelect", FullRowSelect, False
        .WriteProperty "GridLines", GridLines, False
        .WriteProperty "HeaderColor", HeaderColor, vbButtonFace
        .WriteProperty "HeaderCustom", HeaderCustom, False
        .WriteProperty "HeaderDragDrop", HeaderDragDrop, True
        .WriteProperty "HeaderFixedWidth", HeaderFixedWidth, False
        .WriteProperty "HeaderFlat", HeaderFlat, False
        .WriteProperty "HeaderForeColor", HeaderForeColor, vbWindowText
        .WriteProperty "HeaderHide", HeaderHide, False
        .WriteProperty "HeaderHighLite", HeaderHighLite, vbButtonShadow
        .WriteProperty "HeaderPressed", HeaderPressed, vbWindowText
        .WriteProperty "InfoTips", InfoTips, False
        .WriteProperty "ItemBorderSelect", ItemBorderSelect, False
        .WriteProperty "ItemIndent", ItemIndent, PRP_ITMIND
        .WriteProperty "LabelEdit", LabelEdit, False
        .WriteProperty "LabelTips", LabelTips, False
        .WriteProperty "ListMode", ListMode, PRP_LSTMDE
        .WriteProperty "MultiSelect", MultiSelect, False
        .WriteProperty "OLEDragMode", OLEDragMode, vbOLEDragManual
        .WriteProperty "OLEDropMode", OLEDropMode, vbOLEDropNone
        .WriteProperty "OneClickActivate", OneClickActivate, False
        .WriteProperty "ScrollBarFlat", ScrollBarFlat, False
        .WriteProperty "SubItemImages", SubItemImages, False
        .WriteProperty "TextAlignment", TextAlignment, PRP_TXTALN
        .WriteProperty "ThemeColor", ThemeColor, PRP_TMCLR
        .WriteProperty "ThemeLuminence", ThemeLuminence, PRP_TMLMC
        .WriteProperty "TrackSelected", TrackSelected, False
        .WriteProperty "UnderlineHot", UnderlineHot, False
        .WriteProperty "UseCellColor", UseCellColor, False
        .WriteProperty "UseCellFont", UseCellFont, False
        .WriteProperty "UseUnicode", UseUnicode, False
        .WriteProperty "UseThemeColors", UseThemeColors, True
        .WriteProperty "ViewMode", ViewMode, PRP_VWEMDE
        .WriteProperty "XPColors", XPColors, False
    End With

On Error GoTo 0

End Sub

Private Sub UserControl_Terminate()

    ListDetatch
    DragStopTimer
    DestroyItems
    DestroyList
    DestroyImages
    DeAllocatePointer "a", True
    Set m_cHListSubclass = Nothing
    If Not (m_lhMod = 0) Then
        FreeLibrary m_lhMod
    End If
    m_lParentHwnd = 0
    m_lLVHwnd = 0
    m_lEditHwnd = 0

End Sub
