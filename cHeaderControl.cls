VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHeaderControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements GXISubclass


Private Const GMEM_FIXED                                As Long = &H0
Private Const GMEM_ZEROINIT                             As Long = &H40
Private Const GWL_STYLE                                 As Long = (-16)
Private Const WS_BORDER                                 As Long = &H800000
Private Const WS_CAPTION                                As Long = &HC00000
Private Const WS_CHILD                                  As Long = &H40000000
Private Const WS_MAXIMIZEBOX                            As Long = &H10000
Private Const WS_MINIMIZE                               As Long = &H20000000
Private Const WS_MINIMIZEBOX                            As Long = &H20000
Private Const WS_OVERLAPPED                             As Long = &H0&
Private Const WS_POPUP                                  As Long = &H80000000
Private Const WS_SYSMENU                                As Long = &H80000
Private Const WS_THICKFRAME                             As Long = &H40000
Private Const WS_VISIBLE                                As Long = &H10000000
Private Const WS_OVERLAPPEDWINDOW                       As Double = _
    (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const SW_HIDE                                   As Long = 0
Private Const SW_SHOW                                   As Long = 5
Private Const FW_NORMAL                                 As Long = 400
Private Const FW_BOLD                                   As Long = 700
Private Const LOGPIXELSY                                As Long = 90
Private Const DT_LEFT                                   As Long = &H0
Private Const DT_CENTER                                 As Long = &H1
Private Const DT_RIGHT                                  As Long = &H2
Private Const DT_VCENTER                                As Long = &H4
Private Const DT_SINGLELINE                             As Long = &H20
Private Const DT_CALCRECT                               As Long = &H400
Private Const DT_END_ELLIPSIS                           As Long = &H8000&
Private Const DT_RTLREADING                             As Long = &H20000
Private Const PS_SOLID                                  As Long = 0
Private Const Transparent                               As Long = 1
Private Const OPAQUE                                    As Long = 2
Private Const ILC_MASK                                  As Long = &H1
Private Const ILC_COLOR32                               As Long = &H20
Private Const ILD_TRANSPARENT                           As Long = &H1
Private Const CCM_FIRST                                 As Long = &H2000
Private Const CCM_SETUNICODEFORMAT                      As Double = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT                      As Double = (CCM_FIRST + 6)
Private Const WC_HEADERA                                As String = "SysHeader32"
Private Const WC_HEADER                                 As String = WC_HEADERA
Private Const HDS_HORZ                                  As Long = &H0
Private Const HDS_BUTTONS                               As Long = &H2
Private Const HDS_HOTTRACK                              As Long = &H4
Private Const HDS_DRAGDROP                              As Long = &H40
Private Const HDS_FULLDRAG                              As Long = &H80
Private Const HDS_FILTERBAR                             As Long = &H100
Private Const HDI_WIDTH                                 As Long = &H1
Private Const HDI_TEXT                                  As Long = &H2
Private Const HDI_FORMAT                                As Long = &H4
Private Const HDI_LPARAM                                As Long = &H8
Private Const HDI_IMAGE                                 As Long = &H20
Private Const HDI_ORDER                                 As Long = &H80
Private Const HDI_FILTER                                As Long = &H100
Private Const HDF_LEFT                                  As Long = 0
Private Const HDF_RIGHT                                 As Long = 1
Private Const HDF_CENTER                                As Long = 2
Private Const HDF_RTLREADING                            As Long = 4
Private Const HDF_SORTUP                                As Long = &H400
Private Const HDF_SORTDOWN                              As Long = &H200
Private Const HDF_BITMAP_ON_RIGHT                       As Long = &H1000
Private Const HDF_IMAGE                                 As Long = &H800
Private Const HDF_OWNERDRAW                             As Long = &H8000
Private Const HDF_STRING                                As Long = &H4000
Private Const HDFT_ISSTRING                             As Long = &H0
Private Const HDM_FIRST                                 As Long = &H1200
Private Const HDM_GETITEMCOUNT                          As Double = (HDM_FIRST + 0)
Private Const HDM_INSERTITEMA                           As Double = (HDM_FIRST + 1)
Private Const HDM_INSERTITEMW                           As Double = (HDM_FIRST + 10)
Private Const HDM_INSERTITEM                            As Long = HDM_INSERTITEMA
Private Const HDM_DELETEITEM                            As Double = (HDM_FIRST + 2)
Private Const HDM_GETITEMA                              As Double = (HDM_FIRST + 3)
Private Const HDM_GETITEMW                              As Double = (HDM_FIRST + 11)
Private Const HDM_GETITEM                               As Long = HDM_GETITEMA
Private Const HDM_SETITEMA                              As Double = (HDM_FIRST + 4)
Private Const HDM_SETITEMW                              As Double = (HDM_FIRST + 12)
Private Const HDM_SETITEM                               As Long = HDM_SETITEMA
Private Const HDM_LAYOUT                                As Double = (HDM_FIRST + 5)
Private Const HDM_HITTEST                               As Double = (HDM_FIRST + 6)
Private Const HDM_GETITEMRECT                           As Double = (HDM_FIRST + 7)
Private Const HDM_SETIMAGELIST                          As Double = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST                          As Double = (HDM_FIRST + 9)
Private Const HDM_CREATEDRAGIMAGE                       As Double = (HDM_FIRST + 16)
Private Const HDM_SETHOTDIVIDER                         As Double = (HDM_FIRST + 19)
Private Const HDM_SETFILTERCHANGETIMEOUT                As Double = (HDM_FIRST + 22)
Private Const HDM_EDITFILTER                            As Double = (HDM_FIRST + 23)
Private Const HDM_CLEARFILTER                           As Double = (HDM_FIRST + 24)
Private Const H_MAX                                     As Long = &HFFFF + 1
Private Const HDN_FIRST                                 As Long = H_MAX - 300
Private Const HDN_ITEMCLICKA                            As Double = (HDN_FIRST - 2)
Private Const HDN_ITEMCLICKW                            As Double = (HDN_FIRST - 22)
Private Const HDN_ITEMDBLCLICKA                         As Double = (HDN_FIRST - 3)
Private Const HDN_ITEMDBLCLICKW                         As Double = (HDN_FIRST - 23)
Private Const HDN_DIVIDERDBLCLICKA                      As Double = (HDN_FIRST - 5)
Private Const HDN_DIVIDERDBLCLICKW                      As Double = (HDN_FIRST - 25)
Private Const HDN_BEGINTRACKA                           As Double = (HDN_FIRST - 6)
Private Const HDN_BEGINTRACKW                           As Double = (HDN_FIRST - 26)
Private Const HDN_ENDTRACKA                             As Double = (HDN_FIRST - 7)
Private Const HDN_ENDTRACKW                             As Double = (HDN_FIRST - 27)
Private Const HDN_TRACKA                                As Double = (HDN_FIRST - 8)
Private Const HDN_TRACKW                                As Double = (HDN_FIRST - 28)
Private Const HDN_BEGINDRAG                             As Double = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG                               As Double = (HDN_FIRST - 11)
Private Const HDN_FILTERCHANGE                          As Double = (HDN_FIRST - 12)
Private Const HDN_FILTERBTNCLICK                        As Double = (HDN_FIRST - 13)
Private Const NM_FIRST                                  As Long = H_MAX
Private Const NM_RCLICK                                 As Double = (NM_FIRST - 5)
Private Const NM_RELEASEDCAPTURE                        As Double = (NM_FIRST - 16)
Private Const WM_USER                                   As Long = &H400
Private Const UM_STARTDRAG                              As Long = WM_USER + &H113
Private Const UM_ENDDRAG                                As Long = WM_USER + &H114
Private Const LVCF_TEXT                                 As Long = &H4
Private Const LVM_FIRST                                 As Long = &H1000
Private Const LVM_GETCOLUMN                             As Long = (LVM_FIRST + 25)
Private Const LVM_GETHEADER                             As Long = (LVM_FIRST + 31)
Private Const LVCF_IMAGE                                As Long = &H10


Public Enum EHdrTextAlign
    HdrTextALignLeft = HDF_LEFT
    HdrTextALignCentre = HDF_CENTER
    HdrTextALignRight = HDF_RIGHT
    HdrTextAlignRTLReading = HDF_RTLREADING
End Enum

Public Enum EHdrSortBitmapTypes
    HdrSortBitmapNone = 0
    HdrSortBitmapUp = HDF_SORTUP
    HdrSortBitmapDown = HDF_SORTDOWN
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum


Private Type RECT
    left                                                    As Long
    top                                                     As Long
    right                                                   As Long
    bottom                                                  As Long
End Type

Private Type POINTAPI
    x                                                       As Long
    y                                                       As Long
End Type

Private Type LOGFONT
    lfHeight                                        As Long
    lfWidth                                         As Long
    lfEscapement                                    As Long
    lfOrientation                                   As Long
    lfWeight                                        As Long
    lfItalic                                        As Byte
    lfUnderline                                     As Byte
    lfStrikeOut                                     As Byte
    lfCharSet                                       As Byte
    lfOutPrecision                                  As Byte
    lfClipPrecision                                 As Byte
    lfQuality                                       As Byte
    lfPitchAndFamily                                As Byte
    lfFaceName(32)                                  As Byte
End Type

Private Type tGroupCol
    lColumn                                                 As Long
    lOrigColSize                                            As Long
    lOrigColOrder                                           As Long
    tR                                                      As RECT
    bMouseOver                                              As Boolean
    bPressed                                                As Boolean
    sXMouseDown                                             As Single
    sYMouseDown                                             As Single
End Type

Private Type HD_HITTESTINFO
    pt                                                      As POINTAPI
    flags                                                   As Long
    iItem                                                   As Long
End Type

Private Type HDITEMA
    Mask                                                    As Long
    cxy                                                     As Long
    pszText                                                 As String
    hbm                                                     As Long
    cchTextMax                                              As Long
    fmt                                                     As Long
    lParam                                                  As Long
    iImage                                                  As Long
    iOrder                                                  As Long
    type                                                    As Long
    pvFilter                                                As Long
End Type

Private Type HDITEMW
    Mask                                                    As Long
    cxy                                                     As Long
    pszText                                                 As Long
    hbm                                                     As Long
    cchTextMax                                              As Long
    fmt                                                     As Long
    lParam                                                  As Long
    iImage                                                  As Long
    iOrder                                                  As Long
type                                                    As Long
    pvFilter                                                As Long
End Type

Private Type NMHDR
    hwndFrom                                                As Long
    idfrom                                                  As Long
    code                                                    As Long
End Type

Private Type HDNOITFYA
    hdr                                                     As NMHDR
    iItem                                                   As Long
    iButton                                                 As Long
    pitem                                                   As HDITEMA
End Type

Private Type HDNOITFYW
    hdr                                                     As NMHDR
    iItem                                                   As Long
    iButton                                                 As Long
    pitem                                                   As HDITEMW
End Type

Private Type NMHEADER
    hdr                                                     As NMHDR
    iItem                                                   As Long
    iButton                                                 As Long
    lPtrHDItem                                              As Long
End Type

Private Type HDTEXTFILTER
    pszText                                                 As Long
    cchTextMax                                              As Long
End Type

Private Type WINDOWPOS
    hwnd                                                    As Long
    hWndInsertAfter                                         As Long
    x                                                       As Long
    y                                                       As Long
    cx                                                      As Long
    cy                                                      As Long
    flags                                                   As Long
End Type

Private Type HDLAYOUT
    lprc                                                    As Long
    lpwpos                                                  As Long
End Type

Private Type PAINTSTRUCT
    hdc                                                     As Long
    fErase                                                  As Boolean
    rcPaint                                                 As RECT
    fRestore                                                As Boolean
    fIncUpdate                                              As Boolean
    rgbReserved(32)                                         As Byte
End Type

Private Type LVCOLUMN
    Mask                                                    As Long
    fmt                                                     As Long
    cx                                                      As Long
    pszText                                                 As Long
    cchTextMax                                              As Long
    iSubItem                                                As Long
    iImage                                                  As Long
    iOrder                                                  As Long
End Type

Private Type BITMAP
    bmType                                                  As Long
    bmWidth                                                 As Long
    bmHeight                                                As Long
    bmWidthBytes                                            As Long
    bmPlanes                                                As Integer
    bmBitsPixel                                             As Integer
    bmBits                                                  As Long
End Type

Private Type HDHITTESTINFO
    pt                                                      As POINTAPI
    flags                                                   As Long
    iItem                                                   As Long
End Type


Private Type TCRD
    lO                                                      As Integer
    Hi                                                      As Integer
End Type

Private Type TCLC
    lBase                                                   As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                                  As Long
    dwFlags                                                 As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                               As Long
    dwHoverTime                                             As Long
End Type



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

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

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As Long, _
                                                       ByVal lpWindowName As Long, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal fEnable As Long) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As RECT, _
                                                      ByVal bErase As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long

Private Declare Function CreateDc Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 lpDeviceName As Any, _
                                                                 lpOutput As Any, _
                                                                 lpInitData As Any) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

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

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal cx As Long, _
                                                          ByVal cy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hIml As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_GetImageCount Lib "Comctl32" (ByVal hImagelist As Long) As Long

Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImagelist As Long) As Long

Private Declare Function ImageList_Draw Lib "Comctl32" (ByVal hIml As Long, _
                                                        ByVal i As Long, _
                                                        ByVal hdcDst As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long, _
                                                        ByVal fStyle As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private Declare Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathA" (ByVal hdc As Long, _
                                                                                     ByVal pszPath As String, _
                                                                                     ByVal dX As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Public Event ColumnWidthChanged(ByVal lColumn As Long, lWidth As Long)
Public Event ColumnWidthChanging(ByVal lColumn As Long, lWidth As Long, bCancel As Boolean)
Public Event StartColumnWidthChange(ByVal lColumn As Long, lWidth As Long, bCancel As Boolean)
Public Event DividerDblClick(ByVal lColumn As Long)
Public Event ColumnClick(ByVal lColumn As Long)
Public Event ColumnDblClick(ByVal lColumn As Long)
Public Event ColumnBeginDrag(ByVal lColumn As Long)
Public Event ColumnManualDragRequest(ByVal lColumn As Long, ByRef bForceManualDragDrop As Boolean)
Public Event ColumnEndDrag(ByVal lColumn As Long, ByVal lOrder As Long)
Public Event ColumnFilterClick(ByVal lColumn As Long)
Public Event ColumnFilterChange(ByVal lColumn As Long, ByVal sFilter As String)
Public Event RightClick(ByVal x As Single, ByVal y As Single)
Public Event Resize()
Public Event RePaint()
Public Event OleDrag()
Public Event ColumnGroupChange(lColumn As Long)
Public Event ColumnUnGroup(lColumn As Long)

Private m_bInitSize                                     As Boolean
Private m_bVisible                                      As Boolean
Private m_bUserMode                                     As Boolean
Private m_bIsNt                                         As Boolean
Private m_bAllowGrouping                                As Boolean
Private m_bHideGroupingBox                              As Boolean
Private m_bSubClass                                     As Boolean
Private m_bHotTrack                                     As Boolean
Private m_bDragReorderColumns                           As Boolean
Private m_bButtons                                      As Boolean
Private m_bFullDrag                                     As Boolean
Private m_bFilterBar                                    As Boolean
Private m_bEnabled                                      As Boolean
Private m_bPreventDrag                                  As Boolean
Private m_bHidingDragImage                              As Boolean
Private m_bCustomHeader                                 As Boolean
Private m_bPainting                                     As Boolean
Private m_bTrackUser32                                  As Boolean
Private m_bSkinLoaded                                   As Boolean
Private m_bHeaderClrAuto                                As Boolean
Private m_bGutterButtonSkinned                          As Boolean
Private m_bUseHeaderTheme                               As Boolean
Private m_lHeaderLuminence                              As Long
Private m_lThemeColor                                   As Long
Private m_lHeaderSkinStyle                              As Long
Private m_lGutterButtonColor                            As Long
Private m_lHeaderBgColor                                As Long
Private m_lTextForeColor                                As Long
Private m_lTextHighLite                                 As Long
Private m_lParentHwnd                                   As Long
Private m_lHdrHandle                                    As Long
Private m_lHeaderWidth                                  As Long
Private m_lHeaderHeight                                 As Long
Private m_lImlHandle                                    As Long
Private m_lHdc()                                        As Long
Private m_lBmp()                                        As Long
Private m_lBmpOld()                                     As Long
Private m_lCurrState                                    As Long
Private m_hFnt                                          As Long
Private m_hIml                                          As Long
Private m_hImlDragImageList                             As Long
Private m_lFilterChangeTimeOut                          As Long
Private m_iGroupedHeaderCount                           As Long
Private m_lCol                                          As Long
Private m_lCXY                                          As Long
Private m_lColOrder                                     As Long
Private m_lR                                            As Long
Private m_iDragCol                                      As Long
Private m_lLastDivider                                  As Long
Private m_lDragCandidateBefore                          As Long
Private m_lDragCandidateAfter                           As Long
Private m_sngLuminence                                  As Single
Private m_sGroupBoxHintText                             As String
Private m_sTag()                                        As String
Private m_cDrag                                         As cImageListDrag
Private m_tULF                                          As LOGFONT
Private m_tR                                            As RECT
Private m_tFont                                         As LOGFONT
Private m_font                                          As StdFont
Private m_pHeader                                       As StdPicture
Private m_tGroupedCol()                                 As tGroupCol
Private WithEvents m_tmrDragDrop                        As clsTimer
Attribute m_tmrDragDrop.VB_VarHelpID = -1
Private m_cRender                                       As clsRender
Private m_cHeaderDc                                     As clsStoreDc
Private m_GXHeader                                      As GXMSubclass


Private Sub Class_Initialize()

    'm_bCustomHeader = True
    '/* Ensure that the common control DLL is loaded
    InitCommonControls
    m_bTrackUser32 = FunctionExported("TrackMouseEvent", "User32")
    '/* default colors
    m_lTextForeColor = &H111111
    m_lTextHighLite = &H676767
    
    '/* Are we running NT
    Dim lVer As Long
    lVer = GetVersion()
    m_bIsNt = ((lVer And &H80000000) = 0)
    
    '/* Enabled default
    m_bEnabled = True
    '/* default to 2s
    m_lFilterChangeTimeOut = 2000
    m_lCol = -1
   ' m_sGroupBoxHintText = DEFAULT_GROUPBOX_HINT_TEXT
    HotTrack = True
    DragReOrderColumns = True
    HasButtons = True
    FullDrag = False

    Set m_cDrag = New cImageListDrag
    Set m_tmrDragDrop = New clsTimer
   
End Sub

Friend Property Get AllowGrouping() As Boolean
    AllowGrouping = m_bAllowGrouping
End Property

Friend Property Let AllowGrouping(ByVal Value As Boolean)

    If Not (m_bAllowGrouping = Value) Then
        m_bAllowGrouping = Value
        If Not (m_bAllowGrouping) Then
            Do While (m_iGroupedHeaderCount > 0)
                ColumnIsGrouped(m_tGroupedCol(1).lColumn) = False
            Loop
        End If
        RaiseEvent Resize
    End If
   
End Property

Friend Property Get ColumnAtIndex(ByVal lIndex As Long) As Long

Dim i As Long

    For i = 0 To ColumnCount - 1
        If (ColumnIndex(i) = lIndex) Then
            ColumnAtIndex = i
            Exit For
        End If
    Next i
   
End Property

Friend Property Get ColumnCount() As Long

    m_lR = SendMessageLongA(m_lHdrHandle, HDM_GETITEMCOUNT, 0&, 0&)
    ColumnCount = m_lR
    
End Property

Friend Property Get ColumnExtraData(ByVal lColumn As Long) As Long

Dim tHI As HDITEMA

    tHI.Mask = HDI_LPARAM
    If (pbGetHeaderItemInfo(lColumn, tHI)) Then
        ColumnExtraData = tHI.lParam
    Else
        debugmsg "Get column extra data error"
    End If

End Property

Friend Property Let ColumnExtraData(ByVal lColumn As Long, _
                                    ByVal lExtraData As Long)

Dim tHI As HDITEMA

    tHI.Mask = HDI_LPARAM
    tHI.lParam = lExtraData
    If Not (pbSetHeaderItemInfo(lColumn, tHI)) Then
        debugmsg "Set column extra data error"
    End If
    
End Property

Friend Property Get ColumnFilter(ByVal lColumn As Long) As String

Dim sFilter As String
Dim tHI     As HDITEMA

    tHI.Mask = HDI_FILTER
    If (pbGetHeaderItemInfo(lColumn, tHI, sFilter)) Then
        ColumnFilter = sFilter
    End If
   
End Property

Friend Property Let ColumnFilter(ByVal lColumn As Long, _
                                 ByVal sFilter As String)

    If Len(sFilter) = 0 Then
        '/* Clear the filter:
        If m_bIsNt Then
            SendMessageLongW m_lHdrHandle, HDM_CLEARFILTER, lColumn, 0
        Else
            SendMessageLongA m_lHdrHandle, HDM_CLEARFILTER, lColumn, 0
        End If
    Else
        '/* Add the text to the filter:
        Dim tHI As HDITEMA
        tHI.Mask = HDI_FILTER
        tHI.type = HDFT_ISSTRING
        If pbSetHeaderItemInfo(lColumn, tHI, sFilter) Then
        Else
            debugmsg "Failed to set filter."
        End If
    End If
    
End Property

Friend Property Get ColumnGroupCount() As Long
    ColumnGroupCount = m_iGroupedHeaderCount
End Property

Friend Property Get ColumnHeader(ByVal lColumn As Long) As String

Dim iPos        As Long
Dim sColHeader  As String
Dim tHI         As HDITEMA

    tHI.cchTextMax = 255
    sColHeader = String$(tHI.cchTextMax, Chr$(0))
    tHI.Mask = HDI_TEXT
    tHI.pszText = sColHeader
    If (pbGetHeaderItemInfo(lColumn, tHI)) Then
        sColHeader = tHI.pszText
        iPos = InStr(sColHeader, Chr$(0))
        If Not (iPos = 0) Then
            ColumnHeader = left$(sColHeader, iPos - 1)
        Else
            ColumnHeader = sColHeader
        End If
    Else
        debugmsg "Get column header text error"
    End If
   
End Property

Friend Property Let ColumnHeader(ByVal lColumn As Long, _
                                 ByVal sText As String)

Dim tHI As HDITEMA

    With tHI
        .cchTextMax = Len(sText)
        If (sText = "") Then
            .pszText = vbNullChar
        Else
            .pszText = sText
        End If
        .Mask = HDI_TEXT
    End With
    If pbSetHeaderItemInfo(lColumn, tHI) Then
        If (ColumnIsGrouped(lColumn)) Then
            RaiseEvent RePaint
        End If
    Else
        debugmsg "Set column header text error"
    End If
   
End Property

Friend Property Get ColumnHeaderFromPoint(ByVal x As Long, _
                                          ByVal y As Long) As Long

Dim lCol    As Long
Dim i       As Long
   
    lCol = -1
    For i = 1 To m_iGroupedHeaderCount
        If Not (PtInRect(m_tGroupedCol(i).tR, x, y) = 0) Then
            ColumnHeaderFromPoint = m_tGroupedCol(i).lColumn
            Exit Property
        End If
    Next i

    If (y >= m_tR.top) And (y <= m_tR.bottom) Then
        Dim xLast As Long
        Dim lWidth As Long
        xLast = m_tR.left
        For i = 0 To ColumnCount - 1
            lCol = ColumnAtIndex(i)
            If Not (ColumnIsGrouped(lCol)) Then
                lWidth = ColumnWidth(lCol)
                If (x >= xLast) And (x < xLast + lWidth) Then
                    ColumnHeaderFromPoint = lCol
                    Exit Property
                End If
                xLast = xLast + lWidth
            End If
        Next i
    End If
    ColumnHeaderFromPoint = lCol
   
End Property

Friend Property Get ColumnImageOnRight(ByVal lColumn As Long) As Boolean

Dim tHI As HDITEMA

   tHI.Mask = HDI_FORMAT
   If (pbGetHeaderItemInfo(lColumn, tHI)) Then
      ColumnImageOnRight = ((tHI.fmt And HDF_BITMAP_ON_RIGHT) = HDF_BITMAP_ON_RIGHT)
   End If
   
End Property

Friend Property Let ColumnImageOnRight(ByVal lColumn As Long, _
                                       ByVal bState As Boolean)

Dim tHI As HDITEMA

    With tHI
        .Mask = HDI_FORMAT
        If (pbGetHeaderItemInfo(lColumn, tHI)) Then
            If bState Then
                .fmt = .fmt Or HDF_BITMAP_ON_RIGHT
            Else
                .fmt = .fmt And Not HDF_BITMAP_ON_RIGHT
            End If
            If (pbSetHeaderItemInfo(lColumn, tHI)) Then
                If (ColumnIsGrouped(lColumn)) Then
                    RaiseEvent RePaint
                End If
            Else
                debugmsg "Failed to set image on right property"
            End If
        End If
    End With
    
End Property

Friend Property Get ColumnGroupOrder(ByVal lColumn As Long) As Long

Dim i   As Long
Dim lr  As Long

    lr = 0
    For i = 1 To m_iGroupedHeaderCount
        If (m_tGroupedCol(i).lColumn = lColumn) Then
            lr = i
            Exit For
        End If
    Next i
    ColumnGroupOrder = lr - 1
   
End Property

Friend Property Let ColumnGroupOrder(ByVal lColumn As Long, _
                                     ByVal lOrder As Long)

Dim i               As Long
Dim lExistingIndex  As Long
Dim tGExisting      As tGroupCol
Dim tHI             As HDITEMA

    '/* Check order index:
    If (lOrder > m_iGroupedHeaderCount) Or (lOrder < -1) Then
        Exit Property
    End If
   
    '/* Find if the column already exists in the groups
    For i = 1 To m_iGroupedHeaderCount
        If (m_tGroupedCol(i).lColumn = lColumn) Then
            lExistingIndex = i
            LSet tGExisting = m_tGroupedCol(i)
            Exit For
        End If
    Next i
            
    If (lOrder = m_iGroupedHeaderCount) Then
        '/* This item wants to be added to the end of the groupings.
        '/* We need to check that it isn't already in the groupings -
        '/* if it is, then you can't add at this order, instead you
        '/* would change the order of the existing item:
        If lExistingIndex > 0 Then
            Exit Property
        End If
        '/* Ok so it wasn't already there, let's add it
        m_iGroupedHeaderCount = m_iGroupedHeaderCount + 1
        ReDim Preserve m_tGroupedCol(1 To m_iGroupedHeaderCount) As tGroupCol
        With m_tGroupedCol(m_iGroupedHeaderCount)
            .lColumn = -1
            .lOrigColOrder = ColumnIndex(lColumn)
            .lOrigColSize = ColumnWidth(lColumn)
            .lColumn = lColumn
        End With
        tHI.Mask = HDI_WIDTH
        tHI.cxy = 0
        pbSetHeaderItemInfo lColumn, tHI
        ColumnIndex(lColumn) = 0
        '/* Resize (we've added something)
        pResize
        '/* Repaint
        RaiseEvent RePaint
    ElseIf (lOrder < 0) Then
        '/* We're removing this column from the grouping
        '/* if it is there
        If (lExistingIndex > 0) Then
            'Debug.Assert "TODO" = ""
        End If
    Else
        '/* Check if this column is already in the grouping:
        If (lExistingIndex > 0) Then
            '/* We are moving this item from the existing location to the new one:
            '/* Check if the column is already at this position:
            If (lExistingIndex = lOrder + 1) Then
            '/* nothing to do
            Else
                '/* Check which direction we're going in:
                If (lExistingIndex - 1 < lOrder) Then
                    '/* Increasing the order.  Move anything from
                    '/* lExistingIndex to lOrder + 1 up one:
                    For i = lExistingIndex To lOrder
                        LSet m_tGroupedCol(i) = m_tGroupedCol(i + 1)
                    Next i
                    '/* Now put this item at lOrder + 1:
                    LSet m_tGroupedCol(lOrder + 1) = tGExisting
                Else
                    '/* Decreasing the order.  Move anything from
                    For i = lExistingIndex - 1 To lOrder + 1 Step -1
                        LSet m_tGroupedCol(i + 1) = m_tGroupedCol(i)
                    Next i
                    '/* Now put this item at lOrder + 1:
                    LSet m_tGroupedCol(lOrder + 1) = tGExisting
                End If
                RaiseEvent RePaint
            End If
        Else
            If (lExistingIndex > 0) Then
                '/* We are inserting this item at the specified index:
                m_iGroupedHeaderCount = m_iGroupedHeaderCount + 1
                ReDim Preserve m_tGroupedCol(1 To m_iGroupedHeaderCount) As tGroupCol
                For i = m_iGroupedHeaderCount To lExistingIndex Step -1
                    LSet m_tGroupedCol(i) = m_tGroupedCol(i - 1)
                Next i
                With m_tGroupedCol(m_iGroupedHeaderCount)
                    .lColumn = -1
                    .lOrigColOrder = ColumnIndex(lColumn)
                    .lOrigColSize = ColumnWidth(lColumn)
                    .lColumn = lColumn
                End With
                tHI.Mask = HDI_WIDTH
                tHI.cxy = 0
                pbSetHeaderItemInfo lColumn, tHI
                ColumnIndex(lColumn) = 0
                '/* Resize (we've added something)
                pResize
                RaiseEvent RePaint
            Else
                '/* inserting this item:
                m_iGroupedHeaderCount = m_iGroupedHeaderCount + 1
                ReDim Preserve m_tGroupedCol(1 To m_iGroupedHeaderCount) As tGroupCol
                For i = m_iGroupedHeaderCount To lOrder + 2 Step -1
                LSet m_tGroupedCol(i) = m_tGroupedCol(i - 1)
                Next i
                With m_tGroupedCol(lOrder + 1)
                    .lColumn = -1
                    .lOrigColOrder = ColumnIndex(lColumn)
                    .lOrigColSize = ColumnWidth(lColumn)
                    .lColumn = lColumn
                End With
                tHI.Mask = HDI_WIDTH
                tHI.cxy = 0
                pbSetHeaderItemInfo lColumn, tHI
                ColumnIndex(lColumn) = 0
                '/* Resize (we've added something)
                pResize
                RaiseEvent RePaint
            End If
        End If
    End If
    
End Property

Friend Property Get ColumnImage(ByVal lColumn As Long) As Long

Dim tHI As HDITEMA

    tHI.Mask = HDI_FORMAT
    If (pbGetHeaderItemInfo(lColumn, tHI)) Then
        If (tHI.fmt And HDF_IMAGE) = HDF_IMAGE Then
            tHI.Mask = HDI_IMAGE
            If (pbGetHeaderItemInfo(lColumn, tHI)) Then
                ColumnImage = tHI.iImage
            Else
                debugmsg "Get column image error"
            End If
        Else
            ColumnImage = -1
        End If
    End If
   
End Property

Friend Property Let ColumnImage(ByVal lColumn As Long, _
                                ByVal lImage As Long)

Dim tHI As HDITEMA

    With tHI
        If Not (ColumnImage(lColumn) = lImage) Then
            .Mask = HDI_FORMAT
            If pbGetHeaderItemInfo(lColumn, tHI) Then
                If (pbValidImage(lImage) < 0) Then
                    .fmt = .fmt Or HDF_IMAGE
                    .Mask = .Mask Or HDI_IMAGE
                    .iImage = lImage
                Else
                    .fmt = .fmt And Not HDF_IMAGE
                End If
                If (pbSetHeaderItemInfo(lColumn, tHI)) Then
                    If (ColumnIsGrouped(lColumn)) Then
                        RaiseEvent RePaint
                    End If
                Else
                    debugmsg "Set column image error"
                End If
            End If
        End If
    End With
   
End Property

Friend Property Get ColumnIndex(ByVal lColumn As Long) As Long

Dim tHI As HDITEMA

    If Not (lColumn = m_lCol) Then
        tHI.Mask = HDI_ORDER
        If (pbGetHeaderItemInfo(lColumn, tHI)) Then
            ColumnIndex = tHI.iOrder
        End If
    Else
        ColumnIndex = m_lColOrder
    End If
   
End Property

Friend Property Get ColumnIsGrouped(ByVal lColumn As Long) As Boolean

Dim i As Long

    For i = 1 To m_iGroupedHeaderCount
        If (m_tGroupedCol(i).lColumn = lColumn) Then
            ColumnIsGrouped = True
            Exit For
        End If
    Next i
   
End Property

Friend Property Let ColumnIsGrouped(ByVal lColumn As Long, _
                                    ByVal bState As Boolean)
    
    ColumnIsGroupedSub lColumn, bState, -1

End Property

Friend Property Let ColumnIndex(ByVal lColumn As Long, _
                                ByVal lOrder As Long)

Dim tHI As HDITEMA

    If Not (ColumnIndex(lColumn) = lOrder) Then
        tHI.Mask = HDI_ORDER
        tHI.iOrder = lOrder
        If Not (pbSetHeaderItemInfo(lColumn, tHI)) Then
            debugmsg "Set column order error"
        End If
    End If
   
End Property

Friend Property Get ColumnOwnerDraw(ByVal lColumn As Long) As Boolean

Dim tHI As HDITEMA

    tHI.Mask = HDI_FORMAT
    If pbGetHeaderItemInfo(lColumn, tHI) Then
        ColumnOwnerDraw = ((tHI.fmt And HDF_OWNERDRAW) = HDF_OWNERDRAW)
    End If
   
End Property

Friend Property Let ColumnOwnerDraw(ByVal lColumn As Long, _
                                    ByVal bState As Boolean)

Dim tHI As HDITEMA

    tHI.Mask = HDI_FORMAT
    If pbGetHeaderItemInfo(lColumn, tHI) Then
        If bState Then
            tHI.fmt = tHI.fmt Or HDF_OWNERDRAW
        Else
            tHI.fmt = tHI.fmt And Not HDF_OWNERDRAW
        End If
        If Not (pbSetHeaderItemInfo(lColumn, tHI)) Then
            debugmsg "Set column owner draw error"
        End If
    End If
   
End Property

Friend Property Get ColumnSortBitmap(ByVal lColumn As Long) As EHdrSortBitmapTypes

Dim tHI As HDITEMA

    tHI.Mask = HDI_FORMAT
    If (pbGetHeaderItemInfo(lColumn, tHI)) Then
        If (tHI.fmt And HdrSortBitmapUp) = HdrSortBitmapUp Then
            ColumnSortBitmap = HdrSortBitmapUp
        ElseIf (tHI.fmt And HdrSortBitmapDown) = HdrSortBitmapDown Then
            ColumnSortBitmap = HdrSortBitmapDown
        Else
            ColumnSortBitmap = HdrSortBitmapNone
        End If
    End If

End Property

Friend Property Let ColumnSortBitmap(ByVal lColumn As Long, _
                                     ByVal eSortBitmapType As EHdrSortBitmapTypes)

Dim tHI As HDITEMA

    tHI.Mask = HDI_FORMAT
    If (pbGetHeaderItemInfo(lColumn, tHI)) Then
        If (eSortBitmapType = HdrSortBitmapDown) Then
            tHI.fmt = tHI.fmt Or HdrSortBitmapDown And Not HdrSortBitmapUp
        ElseIf (eSortBitmapType = HdrSortBitmapUp) Then
            tHI.fmt = tHI.fmt Or HdrSortBitmapUp And Not HdrSortBitmapDown
        Else
            tHI.fmt = tHI.fmt And Not (HdrSortBitmapUp Or HdrSortBitmapDown)
        End If
        If (pbSetHeaderItemInfo(lColumn, tHI)) Then
            If (ColumnIsGrouped(lColumn)) Then
                RaiseEvent RePaint
            End If
        Else
            debugmsg "Failed to set ColumnSortBitmap"
        End If
    End If
   
End Property

Friend Property Get ColumnTag(ByVal lColumn As Long) As String
    ColumnTag = m_sTag(lColumn)
End Property

Friend Property Let ColumnTag(ByVal lColumn As Long, ByVal sTag As String)

    If (lColumn < ColumnCount) Then
        m_sTag(lColumn) = sTag
    Else
        debugmsg "Error setting column tag."
    End If
    
End Property

Friend Property Get ColumnTextAlign(ByVal lColumn As Long) As EHdrTextAlign

Dim tHI As HDITEMA

    tHI.Mask = HDI_FORMAT
    If (pbGetHeaderItemInfo(lColumn, tHI)) Then
        ColumnTextAlign = tHI.fmt And &H7&
    End If
   
End Property

Friend Property Let ColumnTextAlign(ByVal lColumn As Long, _
                                    ByVal eAlign As EHdrTextAlign)

Dim tHI As HDITEMA

    With tHI
        .Mask = HDI_FORMAT
        If (pbGetHeaderItemInfo(lColumn, tHI)) Then
            .fmt = tHI.fmt And Not &H7&
            .fmt = tHI.fmt Or eAlign
            If (pbSetHeaderItemInfo(lColumn, tHI)) Then
                If (ColumnIsGrouped(lColumn)) Then
                    RaiseEvent RePaint
                End If
            Else
            End If
        End If
    End With
    
End Property

Friend Property Get ColumnWidth(ByVal lColumn As Long) As Long

Dim i   As Long
Dim tHI As HDITEMA

    If lColumn = m_lCol Then
        ColumnWidth = m_lCXY
    Else
        If (ColumnIsGrouped(lColumn)) Then
            For i = 1 To m_iGroupedHeaderCount
                If (m_tGroupedCol(i).lColumn = lColumn) Then
                    ColumnWidth = m_tGroupedCol(i).lOrigColSize
                    Exit For
                End If
            Next i
        Else
            tHI.Mask = HDI_WIDTH
            If (pbGetHeaderItemInfo(lColumn, tHI)) Then
                ColumnWidth = tHI.cxy
            Else
                debugmsg "Get column width error"
            End If
        End If
    End If

End Property

Friend Property Let ColumnWidth(ByVal lColumn As Long, _
                                ByVal lWidthPixels As Long)

Dim i   As Long
Dim tHI As HDITEMA

    If Not (ColumnWidth(lColumn) = lWidthPixels) Then
      If (ColumnIsGrouped(lColumn)) Then
         For i = 1 To m_iGroupedHeaderCount
            If (m_tGroupedCol(i).lColumn = lColumn) Then
               m_tGroupedCol(i).lOrigColSize = lWidthPixels
               Exit For
            End If
         Next i
      Else
         tHI.Mask = HDI_WIDTH
         tHI.cxy = lWidthPixels
         If (pbSetHeaderItemInfo(lColumn, tHI)) Then
            RaiseEvent ColumnWidthChanged(lColumn, lWidthPixels)
         Else
            debugmsg "Set column width error"
         End If
      End If
   End If
   
End Property

Friend Property Get ColumnX(ByVal lColumn As Long) As Long

Dim i As Long
Dim x As Long

    For i = 0 To ColumnCount - 1
        If (ColumnIndex(lColumn) = i) Then
            Exit For
        Else
            If Not ColumnIsGrouped(ColumnAtIndex(i)) Then
                x = x + ColumnWidth(ColumnAtIndex(i))
            End If
        End If
    Next i
    ColumnX = x
   
End Property

Friend Property Get CustomHeader() As Boolean
    CustomHeader = m_bCustomHeader
End Property

Friend Property Let CustomHeader(ByVal PropVal As Boolean)

    m_bCustomHeader = PropVal

End Property

Friend Property Get DragImageList(ByVal lColumn As Long)

    ClearDragImageList
    If (ColumnIsGrouped(lColumn)) Then
        Dim i As Long
        Dim lIndex As Long
        Dim lWidth As Long
        Dim lHeight As Long
        '/* Need to create a new image list for dragging:
        For i = 1 To m_iGroupedHeaderCount
            If (m_tGroupedCol(i).lColumn = lColumn) Then
                lIndex = i
                Exit For
            End If
        Next i
        lWidth = m_tGroupedCol(lIndex).tR.right - m_tGroupedCol(lIndex).tR.left
        lHeight = Height
        m_hImlDragImageList = ImageList_Create(lWidth, lHeight, ILC_MASK Or ILC_COLOR32, 1, 1)
        '/* Now need to create an image of the item to be dragged to add to the drag image list:
        Dim lhDCDisp As Long
        Dim lHDC As Long
        Dim lhBmp As Long
        Dim lhBmpOld As Long
        Dim tR As RECT
        Dim hFontOld As Long
        tR.right = lWidth
        tR.bottom = lHeight
        lhDCDisp = CreateDc("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        lHDC = CreateCompatibleDC(lhDCDisp)
        lhBmp = CreateCompatibleBitmap(lhDCDisp, lWidth, lHeight)
        DeleteDC lhDCDisp
        lhBmpOld = SelectObject(lHDC, lhBmp)
        '/* Draw the item onto the DC:
        hFontOld = SelectObject(lHDC, m_hFnt)
        DrawGroupedHeaderItem lHDC, lIndex, tR
        SelectObject lHDC, hFontOld
        '/* Select the bitmap out:
        SelectObject lHDC, lhBmp
        '/* Clear up DC:
        DeleteDC lHDC
        '/* Add the bitmap to the ImageList:
        ImageList_AddMasked m_hImlDragImageList, lhBmp, &H10101
        '/* Clear up the bitmap:
        DeleteObject lhBmp
    Else
        If m_bIsNt Then
            m_hImlDragImageList = SendMessageLongW(m_lHdrHandle, HDM_CREATEDRAGIMAGE, lColumn, 0)
        Else
            m_hImlDragImageList = SendMessageLongA(m_lHdrHandle, HDM_CREATEDRAGIMAGE, lColumn, 0)
        End If
    End If
    DragImageList = m_hImlDragImageList

End Property

Friend Property Get DragReOrderColumns() As Boolean
    DragReOrderColumns = m_bDragReorderColumns
End Property

Friend Property Let DragReOrderColumns(ByVal bState As Boolean)

    If Not (m_bDragReorderColumns = bState) Then
        m_bDragReorderColumns = bState
        If Not (m_lHdrHandle = 0) Then
            pSetStyle HDS_DRAGDROP, bState
        End If
    End If
   
End Property

Friend Property Get Enabled() As Boolean
   Enabled = m_bEnabled
End Property

Friend Property Let Enabled(ByVal bEnabled As Boolean)

Static bResetOnEnable As Boolean

    If Not (m_bEnabled = bEnabled) Then
        m_bEnabled = bEnabled
        If Not m_bEnabled Then
            If m_bButtons Then
                HasButtons = False
                m_bButtons = True
                bResetOnEnable = True
            End If
        Else
            If bResetOnEnable Then
                If m_bButtons Then
                    m_bButtons = False
                    HasButtons = True
                End If
            End If
        End If
        EnableWindow m_lHdrHandle, Abs(m_bEnabled)
    End If

End Property

Friend Property Get FilterBar() As Boolean
    FilterBar = m_bFilterBar
End Property

Friend Property Let FilterBar(ByVal bFilterBar As Boolean)

    If Not (m_bFilterBar = bFilterBar) Then
        m_bFilterBar = bFilterBar
        If Not (m_lHdrHandle = 0) Then
            pSetStyle HDS_FILTERBAR, bFilterBar
            pResize
        End If
    End If
   
End Property

Friend Property Get FilterChangeTimeOut() As Long
    FilterChangeTimeOut = m_lFilterChangeTimeOut
End Property

Friend Property Let FilterChangeTimeOut(ByVal lTimeOutMs As Long)

    m_lFilterChangeTimeOut = lTimeOutMs
    If m_bIsNt Then
        SendMessageLongW m_lHdrHandle, HDM_SETFILTERCHANGETIMEOUT, 0&, lTimeOutMs
    Else
        SendMessageLongA m_lHdrHandle, HDM_SETFILTERCHANGETIMEOUT, 0&, lTimeOutMs
    End If
   
End Property

Friend Property Get FullDrag() As Boolean
    FullDrag = m_bFullDrag
End Property

Friend Property Let FullDrag(ByVal bFullDrag As Boolean)

    If Not (m_bFullDrag = bFullDrag) Then
        m_bFullDrag = bFullDrag
        If Not m_lHdrHandle = 0 Then
            pSetStyle HDS_FULLDRAG, bFullDrag
        End If
    End If
   
End Property

Friend Property Get GroupBoxHintText() As String
    GroupBoxHintText = m_sGroupBoxHintText
End Property

Friend Property Let GroupBoxHintText(ByVal sText As String)
    m_sGroupBoxHintText = sText
End Property

Friend Property Get GutterButtonColor() As Long
    GutterButtonColor = m_lGutterButtonColor
End Property

Friend Property Let GutterButtonColor(ByVal PropVal As Long)
    m_lGutterButtonColor = PropVal
End Property

Friend Property Get GutterButtonSkinned() As Boolean
    GutterButtonSkinned = m_bGutterButtonSkinned
End Property

Friend Property Let GutterButtonSkinned(ByVal PropVal As Boolean)
    m_bGutterButtonSkinned = PropVal
End Property

Friend Property Get HasButtons() As Boolean
    HasButtons = m_bButtons
End Property

Friend Property Let HasButtons(ByVal bHasButtons As Boolean)

    If Not (bHasButtons = m_bButtons) Then
        m_bButtons = bHasButtons
        If Not m_lHdrHandle = 0 Then
            pSetStyle HDS_BUTTONS, bHasButtons
        End If
    End If
   
End Property

Friend Property Get HeaderBackColor() As Long
    HeaderBackColor = m_lHeaderBgColor
End Property

Friend Property Let HeaderBackColor(PropVal As Long)
    m_lHeaderBgColor = PropVal
End Property

Friend Property Get HeaderColorAuto() As Boolean
    HeaderColorAuto = m_bHeaderClrAuto
End Property

Friend Property Let HeaderColorAuto(PropVal As Boolean)
    m_bHeaderClrAuto = PropVal
End Property

Friend Property Get HeaderSkinStyle() As Long
    HeaderSkinStyle = m_lHeaderSkinStyle
End Property

Friend Property Let HeaderSkinStyle(ByVal PropVal As Long)
    m_lHeaderSkinStyle = PropVal
End Property

Friend Property Get Height() As Long
    Height = m_tR.bottom - m_tR.top
End Property

Friend Property Let Height(ByVal lHeight As Long)

    m_tR.bottom = m_tR.top + lHeight
    pResize
   
End Property

Friend Property Get HideGroupingBox() As Boolean
    HideGroupingBox = m_bHideGroupingBox
End Property

Friend Property Let HideGroupingBox(ByVal Value As Boolean)

    If Not m_bHideGroupingBox = Value Then
        m_bHideGroupingBox = Value
        RaiseEvent Resize
    End If
   
End Property

Friend Property Get HotTrack() As Boolean
    HotTrack = m_bHotTrack
End Property

Friend Property Let HotTrack(ByVal bHotTrack As Boolean)

    If Not (m_bHotTrack = bHotTrack) Then
        m_bHotTrack = bHotTrack
        If Not (m_lHdrHandle = 0) Then
            pSetStyle HDS_HOTTRACK, bHotTrack
        End If
    End If
   
End Property

Private Property Get IdealHeaderHeight() As Long

Dim lr      As Long
Dim lHeight As Long
Dim cx      As Long
Dim cy      As Long
Dim tHDL    As HDLAYOUT
Dim rc      As RECT
Dim pwpos   As WINDOWPOS

    tHDL.lprc = VarPtr(rc)
    tHDL.lpwpos = VarPtr(pwpos)
    lr = SendMessageA(m_lHdrHandle, HDM_LAYOUT, 0&, tHDL)
    lHeight = pwpos.cy
    ImageList_GetIconSize m_hIml, cx, cy
    If (lHeight < cy + 4) Then
        lHeight = cy + 4
    End If
   
    IdealHeaderHeight = lHeight

End Property

Friend Property Get IdealHeight() As Long

Dim lHeight As Long
   
    lHeight = Height
    If (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
        If (m_iGroupedHeaderCount = 0) Then
            lHeight = lHeight + lHeight + 12 + 8
        Else
            lHeight = lHeight + lHeight + 12 + (m_iGroupedHeaderCount * 8)
        End If
    End If
   
    IdealHeight = lHeight
   
End Property

Public Property Get IHeader() As StdPicture
'/* header image
    Set IHeader = m_pHeader
End Property

Public Property Set IHeader(PropVal As StdPicture)
    Set m_pHeader = PropVal
End Property

Friend Property Get left() As Long
    left = m_tR.left
End Property

Friend Property Let left(ByVal lLeft As Long)

Dim lOrigLeft As Long

    With m_tR
        lOrigLeft = .left
        .right = lLeft + .right - .left
        .left = lLeft
        pResize Not (.left = lOrigLeft)
    End With
    
End Property

Friend Property Get HeaderLuminence() As Long
    HeaderLuminence = m_lHeaderLuminence
End Property

Friend Property Let HeaderLuminence(PropVal As Long)
    m_lHeaderLuminence = PropVal
End Property

Friend Property Get HeaderThemeColor() As Long
    HeaderThemeColor = m_lThemeColor
End Property

Friend Property Let HeaderThemeColor(ByVal PropVal As Long)
    m_lThemeColor = PropVal
End Property

Friend Property Get top() As Long
    top = m_tR.top
End Property

Friend Property Let top(ByVal lTop As Long)

    With m_tR
        .bottom = lTop + .bottom - .top
        .top = lTop
    End With
    pResize
    
End Property

Friend Property Get UseHeaderTheme() As Boolean
    UseHeaderTheme = m_bUseHeaderTheme
End Property

Friend Property Let UseHeaderTheme(ByVal PropVal As Boolean)
    m_bUseHeaderTheme = PropVal
End Property

Friend Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Friend Property Let Visible(ByVal bVisible As Boolean)

    m_bVisible = bVisible
    If Not m_lHdrHandle = 0 Then
        If bVisible Then
            ShowWindow m_lHdrHandle, SW_SHOW
        Else
            ShowWindow m_lHdrHandle, SW_HIDE
        End If
    End If

End Property

Friend Property Get Width() As Long
    Width = m_tR.right - m_tR.left
End Property

Friend Property Let Width(ByVal lWidth As Long)

    m_tR.right = m_tR.left + lWidth
    pResize
   
End Property

Friend Function AddColumn(ByVal sText As String, _
                          Optional ByVal lWidth As Long = 64, _
                          Optional ByVal eTextAlign As EHdrTextAlign = HdrTextALignLeft, _
                          Optional ByVal lExtraData As Long = 0, _
                          Optional ByVal lImage As Long = -1, _
                          Optional ByVal lInsertAfter As Long = -1) As Long

Dim b()     As Byte
Dim lr      As Long
Dim wP      As Long
Dim tHI     As HDITEMA
Dim tHIW    As HDITEMW

    With tHI
        .Mask = HDI_TEXT Or HDI_WIDTH Or HDI_FORMAT Or HDI_LPARAM
        .fmt = eTextAlign Or HDF_STRING
        If Not m_hIml = 0 Then
            If (pbValidImage(lImage)) Then
                .Mask = .Mask Or HDI_IMAGE
                .fmt = .fmt Or HDF_IMAGE
                .iImage = lImage
            End If
        End If
        .cxy = lWidth
        .lParam = lExtraData
        If Len(sText) = 0 Then
            .Mask = .Mask And Not HDI_TEXT
        Else
            .pszText = sText
            .cchTextMax = Len(sText)
        End If
    End With
    If (lInsertAfter < 0) Then
        wP = ColumnCount + 1
    Else
        wP = lInsertAfter
    End If
    
    If m_bIsNt Then
        '/* Copy to tHIW:
        With tHIW
            .Mask = tHI.Mask
            .fmt = tHI.fmt
            .iImage = tHI.iImage
            .cxy = tHI.cxy
            .lParam = tHI.lParam
            .cchTextMax = tHI.cchTextMax
            If (Len(sText) > 0) Then
                b = tHI.pszText
                ReDim Preserve b(0 To UBound(b) + 2) As Byte
            Else
                ReDim b(0 To 1) As Byte
            End If
            .pszText = VarPtr(b(0))
        End With
        m_lR = SendMessageW(m_lHdrHandle, HDM_INSERTITEMW, wP, tHIW)
    Else
        m_lR = SendMessageA(m_lHdrHandle, HDM_INSERTITEM, wP, tHI)
    End If
      
    If (ColumnCount > 0) Then
        ReDim Preserve m_sTag(0 To ColumnCount - 1) As String
    End If
           
End Function

Friend Sub ClearAllFilters()

    If m_bIsNt Then
        SendMessageLongW m_lHdrHandle, HDM_CLEARFILTER, -1, 0
    Else
        SendMessageLongA m_lHdrHandle, HDM_CLEARFILTER, -1, 0
    End If

End Sub

Friend Sub ClearDragImageList()

    If m_hImlDragImageList Then
        ImageList_Destroy m_hImlDragImageList
        m_hImlDragImageList = 0
    End If
   
End Sub

Friend Function ColumnHitTest(ByVal xPixels As Long, _
                              ByVal yPixels As Long) As Long

Dim i       As Long
Dim j       As Long
Dim lIndex  As Long
Dim yHeader As Long
Dim rc      As RECT

    lIndex = -1
    yHeader = yPixels
    If (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
        GetWindowRect m_lParentHwnd, rc
        yHeader = yPixels + rc.bottom - rc.top - Height
    End If
    For i = 0 To ColumnCount - 1
        If Not (ColumnIsGrouped(i)) Then
            If m_bIsNt Then
                SendMessageW m_lHdrHandle, HDM_GETITEMRECT, i, rc
            Else
                SendMessageA m_lHdrHandle, HDM_GETITEMRECT, i, rc
            End If
            If Not (PtInRect(rc, xPixels, yHeader) = 0) Then
                lIndex = i
                Exit For
            End If
        End If
    Next i
   
    If lIndex = -1 Then
        For i = 1 To m_iGroupedHeaderCount
            If Not (PtInRect(m_tGroupedCol(i).tR, xPixels, yPixels) = 0) Then
                lIndex = m_tGroupedCol(i).lColumn
                Exit For
            End If
        Next i
    End If
   
    ColumnHitTest = lIndex
   
End Function

Friend Sub ColumnStartFilterEdit(ByVal lColumn As Long)

    If m_bIsNt Then
        SendMessageLongW m_lHdrHandle, HDM_EDITFILTER, lColumn, 0
    Else
        SendMessageLongA m_lHdrHandle, HDM_EDITFILTER, lColumn, 0
    End If
   
End Sub

Private Sub HeaderThemeSettings()

    Select Case m_lHeaderLuminence
    Case 0
        m_sngLuminence = 0.3
    Case 1
        m_sngLuminence = 0.6
    Case 2
        m_sngLuminence = 1
    End Select
    
End Sub

Friend Sub Init(ByVal hWndParent As Long, ByVal bUserMode As Boolean)

    m_lParentHwnd = hWndParent
    m_bUserMode = bUserMode
    pCreateHeader
    
End Sub

Private Function LoadHeaderImage() As Boolean

On Error GoTo Handler

    Select Case m_lHeaderSkinStyle
    '/* flat
    Case 0
        Set IHeader = LoadResPicture("MTHEADER", vbResBitmap)
    '/* gloss
    Case 1
        Set IHeader = LoadResPicture("GSHEADER", vbResBitmap)
    '/* rounded
    Case 2
        Set IHeader = LoadResPicture("RDHEADER", vbResBitmap)
    End Select
    If m_bUseHeaderTheme Then
        HeaderThemeSettings
    End If
    LoadHeaderImage = True

On Error GoTo 0
Exit Function

Handler:

End Function

Friend Sub Move(ByVal lLeft As Long, _
                ByVal lTop As Long, _
                Optional ByVal lWidth As Long = -1, _
                Optional ByVal lHeight As Long = -1)

    With m_tR
        .right = lLeft + .right - .left
        .left = lLeft
        .bottom = lTop + .bottom - .top
        .top = lTop
        If (lWidth > -1) Then
            .right = .left + lWidth
        End If
        If (lHeight > -1) Then
            .bottom = .top + lHeight
        End If
    End With
    pResize
   
End Sub

Friend Sub PaintGroups(ByVal lHDC As Long, _
                       ByVal oGroupAreaColor As OLE_COLOR)

Dim lHeight     As Long
Dim i           As Long
Dim lIconIndex  As Long
Dim hPenDark    As Long
Dim hPenOld     As Long
Dim hBr         As Long
Dim lColor      As Long
Dim sText       As String
Dim tJunk       As POINTAPI
Dim tR          As RECT
Dim tCalcR      As RECT
Dim rc          As RECT

    If (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
        lHeight = Height
        GetClientRect m_lParentHwnd, rc
        hBr = CreateSolidBrush(TranslateColor(oGroupAreaColor))
        LSet tR = rc
        tR.bottom = tR.top + IdealHeight - Height
        FillRect lHDC, tR, hBr
        DeleteObject hBr
        If (m_iGroupedHeaderCount = 0) Then
            '/* Draw the GroupBoxDragHint text:
            With tR
                .left = 8
                .top = 8
                .bottom = .top + lHeight
                .right = rc.right - 8
            End With
            LSet tCalcR = tR
            If m_bIsNt Then
                DrawTextW lHDC, StrPtr(" " & m_sGroupBoxHintText & " "), -1, tCalcR, DT_CALCRECT Or DT_SINGLELINE
            Else
                DrawTextA lHDC, " " & m_sGroupBoxHintText & " ", -1, tCalcR, DT_CALCRECT Or DT_SINGLELINE
            End If
            tR.right = tR.left + tCalcR.right - tCalcR.left + 4
            '/* j. added optional skin
            If m_bGutterButtonSkinned Then
                SkinGutterButton lHDC, tR
                SetBkMode lHDC, Transparent
            Else
                If Not (m_lGutterButtonColor = vbButtonFace) Then
                    hBr = CreateSolidBrush(BlendColor(m_lGutterButtonColor, vbButtonFace))
                ElseIf (oGroupAreaColor = vbButtonShadow) Then
                    hBr = GetSysColorBrush(vbButtonFace And &H1F&)
                Else
                    hBr = CreateSolidBrush(BlendColor(oGroupAreaColor, vbButtonFace))
                End If
                FillRect lHDC, tR, hBr
                DeleteObject hBr
            
                SetBkMode lHDC, OPAQUE
                If (oGroupAreaColor = vbButtonShadow) Then
                    SetBkColor lHDC, GetSysColor(vbButtonFace And &H1F&)
                    SetTextColor lHDC, GetSysColor(vb3DShadow And &H1F&)
                Else
                    SetBkColor lHDC, BlendColor(oGroupAreaColor, vbButtonFace)
                    SetTextColor lHDC, BlendColor(oGroupAreaColor, vbWindowText)
                End If
            End If
            '/* draw text
            If m_bIsNt Then
                DrawTextW lHDC, StrPtr(" " & m_sGroupBoxHintText & " "), -1, tR, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
            Else
                DrawTextA lHDC, " " & m_sGroupBoxHintText & " ", -1, tR, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
            End If
        Else
            '/* Draw the grouped headers:
            With tR
                .left = 8
                .top = 8
                .bottom = .top + lHeight
                .right = rc.right - 8
            End With
            LSet tCalcR = tR
            For i = 1 To m_iGroupedHeaderCount
                '/* Calculate Text Size
                sText = ColumnHeader(m_tGroupedCol(i).lColumn)
                '/* draw text
                If m_bIsNt Then
                    DrawTextW lHDC, StrPtr(" " & sText & " "), -1, tCalcR, DT_SINGLELINE Or DT_CALCRECT
                Else
                    DrawTextA lHDC, " " & sText & " ", -1, tCalcR, DT_SINGLELINE Or DT_CALCRECT
                End If
                '/* draw the header
                tR.right = tR.left + (tCalcR.right - tCalcR.left) + 4
                lIconIndex = ColumnImage(m_tGroupedCol(i).lColumn)
                If (lIconIndex > -1) Then
                    tR.right = tR.right + 24
                End If
                '/* Store the rectangle:
                LSet m_tGroupedCol(i).tR = tR
                DrawGroupedHeaderItem lHDC, i, tR
                '/* Get the bound rectangle back again:
                LSet tR = m_tGroupedCol(i).tR
                If (m_lDragCandidateBefore = i) Then
                    '/* Draw candidate marks at the left edge
                    DrawDragCandidate lHDC, i, tR, True
                ElseIf (m_lDragCandidateAfter = i) Then
                    '/* Draw candidate marks at the right edge
                    DrawDragCandidate lHDC, i, tR, False
                End If
                If (i < m_iGroupedHeaderCount) Then
                    '/* Draw the connector:
                    hPenDark = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F))
                    hPenOld = SelectObject(lHDC, hPenDark)
                    MoveToEx lHDC, tR.right - 8, tR.bottom, tJunk
                    LineTo lHDC, tR.right - 8, tR.bottom + 4
                    LineTo lHDC, tR.right + 16, tR.bottom + 4
                    SelectObject lHDC, hPenOld
                    DeleteObject hPenDark
                End If
                '/* Offset for the next one:
                With tR
                    .left = .left + (.right - .left) + 8
                    .top = .top + 8
                    .bottom = .bottom + 8
                End With
            Next i
        End If
    End If

End Sub

Private Function pbValidImage(ByVal lImgIndex As Long) As Boolean

Dim iCount As Long

    If Not (m_hIml = 0) Then
        iCount = ImageList_GetImageCount(m_hIml)
        If (lImgIndex > -1) And (lImgIndex < iCount) Then
            pbValidImage = True
        End If
    End If
   
End Function

Friend Sub RemoveColumn(ByVal lColumn As Long)

Dim lr      As Long
Dim iCol    As Long

    If (ColumnIsGrouped(lColumn)) Then
        ColumnIsGrouped(lColumn) = False
    End If

    lr = SendMessageLongA(m_lHdrHandle, HDM_DELETEITEM, lColumn, 0)
    If Not (lr = 0) Then
        If (ColumnCount > 0) Then
            For iCol = lColumn To UBound(m_sTag) - 1
                m_sTag(iCol) = m_sTag(iCol + 1)
            Next iCol
            ReDim Preserve m_sTag(0 To ColumnCount - 1) As String
        Else
            Erase m_sTag
        End If
    End If
   
End Sub

Friend Sub SetFont(ByVal lHDC As Long, _
                  sFont As StdFont)

Dim hFnt As Long

    If Not (m_font Is sFont) Then
        Set m_font = sFont
        '/* Store a log font structure for this font:
        lHDC = CreateDc("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        pOLEFontToLogFont sFont, lHDC, m_tULF
        DeleteDC lHDC
        '/* Store old font handle:
        hFnt = m_hFnt
        '/* Create a new version of the font:
         m_hFnt = CreateFontIndirect(m_tULF)
        '/* Ensure the header has the correct font:
        If Not (m_lHdrHandle = 0) Then
            SendMessageA m_lHdrHandle, WM_SETFONT, m_hFnt, 1
        End If
        '/* Delete previous version, if we had one:
        If Not hFnt = 0 Then
            DeleteObject hFnt
        End If
        pResize
    End If
    
End Sub

Friend Function SetHotDivider(ByVal x As Long, _
                              ByVal y As Long) As Long

Dim rc As RECT

    GetClientRect m_lHdrHandle, rc
    If (y > -8) And (y < rc.bottom + 8) Then
        Dim lXY As Long
        lXY = (x And &HFFFF&)
        lXY = lXY Or (y And &H7FFF) * &H10000
        If (y And &H8000) = &H8000 Then
            lXY = lXY Or &H80000000
        End If
        If m_bIsNt Then
            SetHotDivider = SendMessageLongW(m_lHdrHandle, HDM_SETHOTDIVIDER, 1, lXY)
        Else
            SetHotDivider = SendMessageLongA(m_lHdrHandle, HDM_SETHOTDIVIDER, 1, lXY)
        End If
    Else
        SendMessageLongA m_lHdrHandle, HDM_SETHOTDIVIDER, 0&, -1
        SetHotDivider = -1
    End If

End Function

Friend Function SetHotDividerForCursorPos() As Long

Dim tP As POINTAPI

    GetCursorPos tP
    ScreenToClient m_lHdrHandle, tP
    SetHotDividerForCursorPos = SetHotDivider(tP.x, tP.y)
   
End Function

Friend Sub SetImageList(ByVal lHDC As Long, _
                        ByRef vImageList As Variant)

    If (VarType(vImageList) = vbLong) Then
        m_hIml = vImageList
        pSetImageList
    ElseIf (VarType(vImageList) = vbObject) Then
        On Error Resume Next
        '/* Ensure image list is initialised:
        vImageList.ListImages(1).Draw lHDC
        Err.Clear
        m_hIml = vImageList.hImagelist
        If Not Err.Number = 0 Then
            m_hIml = 0
            debugmsg "Error setting image list."
        Else
            If Not m_lHdrHandle = 0 Then
                pSetImageList
            End If
        End If
        On Error GoTo 0
    Else
        debugmsg "Error setting image list."
    End If

End Sub

Private Sub SkinGutterButton(ByVal lHDC As Long, _
                             ByRef tRect As RECT)
Dim hPenLight   As Long
Dim hPenOld     As Long
Dim tJunk       As POINTAPI

    '/* skin left/mid/right
    With tRect
        m_cRender.Stretch lHDC, .left, .top, 3, (.bottom - .top), m_lHdc(0), 0, 0, 3, m_lHeaderHeight, SRCCOPY
        m_cRender.Stretch lHDC, (.left + 3), .top, (.right - .left) - 6, (.bottom - .top), _
            m_lHdc(0), 3, 0, m_lHeaderWidth - 6, m_lHeaderHeight, SRCCOPY
        m_cRender.Stretch lHDC, (.right - .left) - 3, .top, 3, (.bottom - .top), _
            m_lHdc(0), m_lHeaderWidth - 3, 0, 3, m_lHeaderHeight, SRCCOPY
        '/* draw border
        hPenLight = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
        MoveToEx lHDC, .left, .bottom - 1, tJunk
        LineTo lHDC, .left, .top
        LineTo lHDC, (.right - 3), .top
        MoveToEx lHDC, (.right - 3), .top, tJunk
        LineTo lHDC, (.right - 3), .bottom - 1
        LineTo lHDC, .left, .bottom - 1
        DeleteObject hPenLight
    End With

End Sub

Private Sub SkinHeaders()

Dim tBmp            As BITMAP

    '/* skinned headers
    If m_bCustomHeader And Not m_bSkinLoaded Then
        If m_GXHeader Is Nothing Then Exit Sub
        If LoadHeaderImage Then
            Set m_cHeaderDc = New clsStoreDc
            With m_cHeaderDc
                .CreateFromPicture IHeader
                If m_bUseHeaderTheme Then
                    .ColorizeImage m_lThemeColor, m_sngLuminence
                End If
            End With
            '/* default sizes
            GetObject IHeader.handle, Len(tBmp), tBmp
            With tBmp
                m_lHeaderWidth = (.bmWidth / 4)
                m_lHeaderHeight = .bmHeight
            End With
            '/* render class
            Set m_cRender = New clsRender
            BackbufferImages
            If m_bHeaderClrAuto Then
                m_lHeaderBgColor = DefaultHeaderBg(m_cHeaderDc.hdc)
            End If
            SkinAttach
        End If
        m_bSkinLoaded = True
    End If

End Sub

Friend Sub StealthSetColumnWidth(ByVal lColumn As Long, _
                                 ByVal lWidthPixels As Long)

Dim i   As Long
Dim tHI As HDITEMA

    For i = 1 To m_iGroupedHeaderCount
        If (m_tGroupedCol(i).lColumn = lColumn) Then
            m_tGroupedCol(i).lOrigColSize = lWidthPixels
            Exit Sub
        End If
    Next i

    With tHI
        .Mask = HDI_WIDTH
        .cxy = lWidthPixels
    End With
    pbSetHeaderItemInfo lColumn, tHI
   
End Sub

Friend Sub UserControl_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, y As Single)

Dim i   As Long
Dim mX  As Long
Dim mY  As Long

    If (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
        If (m_iGroupedHeaderCount > 0) Then
            If (Button = vbLeftButton) Then
                '/* is the mouse down on a header?
                mX = x
                mY = y
                For i = 1 To m_iGroupedHeaderCount
                    If Not (PtInRect(m_tGroupedCol(i).tR, mX, mY) = 0) Then
                        m_tGroupedCol(i).bMouseOver = True
                        m_tGroupedCol(i).bPressed = True
                        m_tGroupedCol(i).sXMouseDown = x
                        m_tGroupedCol(i).sYMouseDown = y
                        RaiseEvent RePaint
                        Exit For
                    End If
                Next i
            End If
        End If
    End If

End Sub

Friend Sub UserControl_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)

Dim i       As Long
Dim mX      As Long
Dim mY      As Long
Dim iOver   As Long

    If (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
        If (m_iGroupedHeaderCount > 0) Then
            If (Button = vbLeftButton) Then
                mX = x
                mY = y
                '/* is the mouse down on a header?
                For i = 1 To m_iGroupedHeaderCount
                If (i > m_iGroupedHeaderCount) Then
                    Exit For
                End If
                If (m_tGroupedCol(i).bPressed) Then
                    If (PtInRect(m_tGroupedCol(i).tR, mX, mY) = 0) Then
                        '/* Mouse out of clicked item
                        If (m_tGroupedCol(i).bMouseOver) Then
                            m_tGroupedCol(i).bMouseOver = False
                            '/* repaint
                            RaiseEvent RePaint
                        End If
                    Else
                        '/* Mouse in clicked item
                        If Not (m_tGroupedCol(i).bMouseOver) Then
                            m_tGroupedCol(i).bMouseOver = True
                            '/* repaint
                            RaiseEvent RePaint
                        End If
                    End If
                    '/* Check for dragging:
                    If Abs(x - m_tGroupedCol(i).sXMouseDown) > 8 Or Abs(y - m_tGroupedCol(i).sYMouseDown) > 8 Then
                        m_iDragCol = m_tGroupedCol(i).lColumn
                        m_tGroupedCol(i).bMouseOver = False
                        m_tGroupedCol(i).bPressed = False
                        RaiseEvent OleDrag
                    End If
                End If
                Next i
            End If
        End If
    End If

End Sub

Friend Sub UserControl_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

Dim i   As Long
Dim mX  As Long
Dim mY  As Long

    If (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
        If (Button = vbRightButton) Then
            '/* right click
            RaiseEvent RightClick(x, y)
        ElseIf (Button = vbLeftButton) Then
            If (m_iGroupedHeaderCount > 0) Then
                mX = x
                mY = y
                '/* is the mouse released on the item that was pressed?
                For i = 1 To m_iGroupedHeaderCount
                    If Not (PtInRect(m_tGroupedCol(i).tR, mX, mY) = 0) Then
                        If (m_tGroupedCol(i).bPressed) Then
                            '/* This is a left click
                            RaiseEvent ColumnClick(m_tGroupedCol(i).lColumn)
                        End If
                        m_tGroupedCol(i).bMouseOver = True
                    Else
                        m_tGroupedCol(i).bMouseOver = False
                    End If
                    m_tGroupedCol(i).bPressed = False
                Next i
                RaiseEvent RePaint
            End If
        End If
    End If

End Sub

Friend Sub UserControl_OLECompleteDrag(Effect As Long)

    m_cDrag.CompleteDrag
    Enabled = True
    m_tmrDragDrop.Interval = 0
    m_lDragCandidateBefore = 0
    m_lDragCandidateAfter = 0
    SetHotDivider -100, -100
    ClearDragImageList
    RaiseEvent RePaint

End Sub

Friend Sub UserControl_OLEDragDrop(Data As DataObject, _
                                   Effect As Long, _
                                   Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   y As Single)

Dim b() As Byte
      
On Error Resume Next

    b = Data.GetData(&HFFFFB046)
    Dim s As String
    On Error GoTo 0
    s = b
    If (InStr(s, "SGrid2Header") = 1) Then
        m_cDrag.CompleteDrag
        ClearDragImageList
        Enabled = True
        m_tmrDragDrop.Interval = 0
        If (m_iDragCol = -1) Then
            Exit Sub
        End If
        Dim lIndex As Long
        lIndex = SetHotDividerForCursorPos()
        If (lIndex > -1) Then
            m_lDragCandidateBefore = 0
            m_lDragCandidateAfter = 0
            '/* Set the order of this item:
            If (ColumnIsGrouped(m_iDragCol)) Then
                '/* Adding the item back into the header:
                ColumnIsGroupedSub m_iDragCol, False, lIndex - 1
                RaiseEvent ColumnUnGroup(m_iDragCol)
            Else
                '/* Dragging with the header itself
                If (lIndex > ColumnIndex(m_iDragCol)) Then
                    ColumnIndex(m_iDragCol) = lIndex - 1
                    RaiseEvent ColumnEndDrag(m_iDragCol, lIndex - 1)
                Else
                    ColumnIndex(m_iDragCol) = lIndex
                    RaiseEvent ColumnEndDrag(m_iDragCol, lIndex)
                End If
            End If
        Else
            '/* Somewhere in the grouping area:
            Dim lOrder As Long
            If (ColumnIsGrouped(m_iDragCol)) Then
                If (m_lDragCandidateBefore = 0 And m_lDragCandidateAfter = 0) Then
                    '/* nothing to do
                    Exit Sub
                End If
            End If
            lOrder = m_iGroupedHeaderCount
            If (m_lDragCandidateBefore > 0) Then
                lOrder = m_lDragCandidateBefore
            ElseIf (m_lDragCandidateAfter > 0) Then
                lOrder = m_lDragCandidateAfter
            End If
            If (lOrder > 0) Then
                If Not (ColumnIsGrouped(m_iDragCol)) And (m_lDragCandidateAfter = m_iGroupedHeaderCount) Then
                '/* nothing
                Else
                    lOrder = lOrder - 1
                End If
            End If
            If (ColumnIsGrouped(m_iDragCol)) Then
                If (lOrder >= m_iGroupedHeaderCount - 1) Then
                    lOrder = m_iGroupedHeaderCount - 1
                End If
            End If
            m_lDragCandidateBefore = 0
            m_lDragCandidateAfter = 0
            ColumnGroupOrder(m_iDragCol) = lOrder
            RaiseEvent ColumnGroupChange(m_iDragCol)
        End If
        SetHotDivider -100, -100
    End If

End Sub

Friend Sub UserControl_OLEDragOver(Data As DataObject, _
                                   Effect As Long, _
                                   Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   y As Single, _
                                   State As Integer)

Dim i                   As Long
Dim lMinOffsetAfter     As Long
Dim lMinOffsetBefore    As Long
Dim lCandidateAfter     As Long
Dim lCandidateBefore    As Long
Dim lThisOffset         As Long
Dim lDragOrderIndex     As Long
            
    If (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
        Dim b() As Byte
        On Error Resume Next
        b = Data.GetData(&HFFFFB046)
        Dim s As String
        On Error GoTo 0
        s = b
        If (InStr(s, "SGrid2Header") = 1) Then
            Dim lhWndParent As Long
            lhWndParent = CLng(Mid(s, 14))
            If Not (lhWndParent = m_lParentHwnd) Then
                '/* Don't allow the header to be dropped onto
                '/* another SGrid control
                Effect = vbDropEffectNone
                Exit Sub
            End If
            Effect = vbDropEffectMove
            Dim tP As POINTAPI
            Dim rc As RECT
            GetCursorPos tP
            GetWindowRect m_lHdrHandle, rc
            If (tP.y >= rc.top - 8) Then
                If (tP.y >= rc.bottom + 16) Then
                    '/* Over the grid
                    Effect = vbDropEffectNone
                    Exit Sub
                End If
                If (m_lDragCandidateBefore > 0) Or (m_lDragCandidateAfter > 0) Then
                    m_lDragCandidateBefore = 0
                    m_lDragCandidateAfter = 0
                    '/* repaint
                    m_cDrag.HideDragImage True
                    RaiseEvent RePaint
                    m_cDrag.HideDragImage False
                End If
                Exit Sub
            End If
            GetWindowRect lhWndParent, rc
            InflateRect rc, 16, 16
            If PtInRect(rc, tP.x, tP.y) = 0 Then
                '/* Outside the owning control
                Effect = vbDropEffectNone
                Exit Sub
            End If
            
            For i = 1 To m_iGroupedHeaderCount
                If (m_tGroupedCol(i).lColumn = m_iDragCol) Then
                    lDragOrderIndex = i
                    Exit For
                End If
            Next i
         
            ScreenToClient m_lParentHwnd, tP
            lMinOffsetBefore = &H7FFFFFFF
            lMinOffsetAfter = &H7FFFFFFF
            For i = 1 To m_iGroupedHeaderCount
                '/* Candidate for replacement:
                If (i > lDragOrderIndex) Or (lDragOrderIndex = 0) Then
                    '/* Potentially the dragged item could go after this one:
                    lThisOffset = (tP.x - (m_tGroupedCol(i).tR.left + _
                        (m_tGroupedCol(i).tR.right - m_tGroupedCol(i).tR.left) \ 2))
                    If (lThisOffset > 0) Then
                        If (lThisOffset < lMinOffsetAfter) Then
                            lCandidateAfter = i
                            lMinOffsetAfter = lThisOffset
                        End If
                    End If
                End If
                If (i < lDragOrderIndex) Or (lDragOrderIndex = 0) Then
                    '/* Potentially the dragged item could go before this one:
                    lThisOffset = (m_tGroupedCol(i).tR.left + (m_tGroupedCol(i).tR.right - m_tGroupedCol(i).tR.left) \ 2) - tP.x
                    If (lThisOffset > 0) Then
                        If (lThisOffset < lMinOffsetBefore) Then
                            lCandidateBefore = i
                            lMinOffsetBefore = lThisOffset
                        End If
                    End If
                    End If
                Next i
         
                If (lCandidateBefore = m_lDragCandidateBefore) And (lCandidateAfter = m_lDragCandidateAfter) Then
                '/* nothing to do
                Else
                    m_lDragCandidateBefore = lCandidateBefore
                    m_lDragCandidateAfter = lCandidateAfter
                If (m_lDragCandidateBefore > 0) And (m_lDragCandidateAfter > 0) Then
                    If Abs(lMinOffsetBefore) < Abs(lMinOffsetAfter) Then
                        m_lDragCandidateAfter = 0
                    Else
                         m_lDragCandidateBefore = 0
                    End If
                End If
                m_cDrag.HideDragImage True
                RaiseEvent RePaint
                m_cDrag.HideDragImage False
            End If
        End If
    End If

End Sub

Friend Sub UserControl_OLEStartDrag(Data As DataObject, _
                                    AllowedEffects As Long)

Dim tP As POINTAPI

    m_cDrag.hImagelist = DragImageList(m_iDragCol)
    GetCursorPos tP
    If (ColumnIsGrouped(m_iDragCol)) Then
        ScreenToClient m_lParentHwnd, tP
        Dim i As Long
        For i = 1 To m_iGroupedHeaderCount
            If (m_tGroupedCol(i).lColumn = m_iDragCol) Then
                m_cDrag.StartDrag 0, tP.x - m_tGroupedCol(i).tR.left, tP.y - m_tGroupedCol(i).tR.top
                Exit For
            End If
        Next i
    Else
        ScreenToClient m_lHdrHandle, tP
        m_cDrag.StartDrag 0, tP.x - ColumnX(m_iDragCol), tP.y
    End If

    Dim b() As Byte
    b = "SGrid2Header:" & m_lParentHwnd
    Data.SetData b, &HFFFFB046
    Data.SetData ColumnHeader(m_iDragCol), vbCFText
    AllowedEffects = vbDropEffectMove
    m_tmrDragDrop.Interval = 25

End Sub

Private Sub DrawDragCandidate(ByVal lHDC As Long, _
                              ByVal lIndex As Long, _
                              tR As RECT, _
                              ByVal bBefore As Boolean)
                              
Dim hPenCandidate   As Long
Dim hPenOld         As Long
Dim tP              As POINTAPI
Dim tJunk           As POINTAPI

    hPenCandidate = CreatePen(PS_SOLID, 1, RGB(255, 0, 0))
    hPenOld = SelectObject(lHDC, hPenCandidate)
    If (bBefore) Then
        tP.x = tR.left
    Else
        tP.x = tR.right
    End If
    '/* Draw top candidate mark
    MoveToEx lHDC, tP.x, tR.top - 8, tJunk
    LineTo lHDC, tP.x, tR.top - 1
    MoveToEx lHDC, tP.x - 1, tR.top - 8, tJunk
    LineTo lHDC, tP.x - 1, tR.top - 2
    MoveToEx lHDC, tP.x + 1, tR.top - 8, tJunk
    LineTo lHDC, tP.x + 1, tR.top - 2
    MoveToEx lHDC, tP.x - 2, tR.top - 3, tJunk
    LineTo lHDC, tP.x + 3, tR.top - 3
    MoveToEx lHDC, tP.x - 3, tR.top - 4, tJunk
    LineTo lHDC, tP.x + 4, tR.top - 4
    '/* Draw bottom candidate mark
    MoveToEx lHDC, tP.x, tR.bottom + 8, tJunk
    LineTo lHDC, tP.x, tR.bottom + 1
    MoveToEx lHDC, tP.x - 1, tR.bottom + 8, tJunk
    LineTo lHDC, tP.x - 1, tR.bottom + 2
    MoveToEx lHDC, tP.x + 1, tR.bottom + 8, tJunk
    LineTo lHDC, tP.x + 1, tR.bottom + 2
    MoveToEx lHDC, tP.x - 2, tR.bottom + 3, tJunk
    LineTo lHDC, tP.x + 3, tR.bottom + 3
    MoveToEx lHDC, tP.x - 3, tR.bottom + 4, tJunk
    LineTo lHDC, tP.x + 4, tR.bottom + 4
    
    SelectObject lHDC, hPenOld
    DeleteObject hPenCandidate
   
End Sub

Private Sub ColumnIsGroupedSub(ByVal lColumn As Long, _
                               ByVal bState As Boolean, _
                               ByVal lNewOrder As Long)

Dim i               As Long
Dim lExistingIndex  As Long
Dim lOrigColSize    As Long
Dim lOrigColOrder   As Long
Dim tHI             As HDITEMA

    For i = 1 To m_iGroupedHeaderCount
        If (m_tGroupedCol(i).lColumn = lColumn) Then
            lOrigColSize = m_tGroupedCol(i).lOrigColSize
            lOrigColOrder = m_tGroupedCol(i).lOrigColOrder
            lExistingIndex = i
            Exit For
        End If
    Next i
   
    If bState Then
        If (lExistingIndex > 0) Then
            '/* nothing to do
            tHI.Mask = HDI_WIDTH
            tHI.cxy = 0
            pbSetHeaderItemInfo lColumn, tHI
        Else
            '/* Add to the end of the grouping:
            ColumnGroupOrder(lColumn) = m_iGroupedHeaderCount
        End If
    Else
        If (lExistingIndex = 0) Or (m_iGroupedHeaderCount = 0) Then
        Else
            '/* Return the column to the header:
            '/* Shift any existing items up:
            For i = lExistingIndex + 1 To m_iGroupedHeaderCount
                LSet m_tGroupedCol(i - 1) = m_tGroupedCol(i)
            Next i
            '/* Remove the item from the array:
            m_iGroupedHeaderCount = m_iGroupedHeaderCount - 1
            If (m_iGroupedHeaderCount > 0) Then
                ReDim Preserve m_tGroupedCol(1 To m_iGroupedHeaderCount) As tGroupCol
            Else
                Erase m_tGroupedCol
            End If
            '/* Make the column visible again at the end of the order
            If lNewOrder < 0 Then
                If lOrigColOrder < ColumnCount Then
                    ColumnIndex(lColumn) = lOrigColOrder
                Else
                    ColumnIndex(lColumn) = ColumnCount - 1
                End If
            Else
                ColumnIndex(lColumn) = lNewOrder
            End If
            ColumnWidth(lColumn) = lOrigColSize
            '/* Resize:
            pResize
            RaiseEvent RePaint
        End If
    End If

End Sub

Private Sub pSetStyle(ByVal lStyleFlags As Long, _
                      ByVal bState As Boolean)

Dim lStyle As Long

    If m_bIsNt Then
        lStyle = GetWindowLongW(m_lHdrHandle, GWL_STYLE)
    Else
        lStyle = GetWindowLong(m_lHdrHandle, GWL_STYLE)
    End If
    If bState Then
        lStyle = lStyle Or lStyleFlags
    Else
        lStyle = lStyle And Not lStyleFlags
    End If
    If m_bIsNt Then
        SetWindowLongW m_lHdrHandle, GWL_STYLE, lStyle
    Else
        SetWindowLong m_lHdrHandle, GWL_STYLE, lStyle
    End If
   
End Sub

Private Function pbGetHeaderItemInfo(ByVal lCol As Long, _
                                     tHI As HDITEMA, _
                                     Optional ByRef sFilter As String) As Boolean
   
Dim tHDTF As HDTEXTFILTER
   
    If m_bIsNt Then
        '/* Copy fields to tHIW structure
        Dim tHIW As HDITEMW
        Dim b() As Byte
        Dim lSize As Long
        Dim bFilter() As Byte
        With tHIW
            .cxy = tHI.cxy
            .fmt = tHI.fmt
            .hbm = tHI.hbm
            .iImage = tHI.iImage
            .iOrder = tHI.iOrder
            .lParam = tHI.lParam
            .Mask = tHI.Mask
            If ((.Mask And HDI_TEXT) = HDI_TEXT) Then
                ReDim b(0 To (tHI.cchTextMax * 2) - 1) As Byte
                .cchTextMax = tHI.cchTextMax
                .pszText = VarPtr(b(0))
            End If
            If ((.Mask And HDI_FILTER) = HDI_FILTER) Then
                ReDim bFilter(0 To 511) As Byte
                tHDTF.pszText = VarPtr(bFilter(0))
                tHDTF.cchTextMax = 255
                .pvFilter = VarPtr(tHDTF)
                .type = tHI.type
            End If
        End With
      
        '/* Send message as Unicode:
        If Not (SendMessageW(m_lHdrHandle, HDM_GETITEMW, lCol, tHIW) = 0) Then
            '/* Get the fields back into tHI
            With tHI
                .cchTextMax = tHIW.cchTextMax
                .cxy = tHIW.cxy
                .fmt = tHIW.fmt
                .hbm = tHIW.hbm
                .iImage = tHIW.iImage
                .iOrder = tHIW.iOrder
                .lParam = tHIW.lParam
                .Mask = tHIW.Mask
                If ((.Mask And HDI_TEXT) = HDI_TEXT) Then
                    '/* Trim nulls:
                    lSize = lstrlenW(VarPtr(b(0)))
                    If lSize > 0 Then
                        ReDim Preserve b(0 To (lSize * 2) - 1) As Byte
                        tHI.pszText = b
                    End If
                End If
                If ((.Mask And HDI_FILTER) = HDI_FILTER) Then
                    lSize = lstrlenW(VarPtr(bFilter(0)))
                    ReDim Preserve bFilter(0 To (lSize * 2) - 1) As Byte
                    sFilter = bFilter
                    .type = tHIW.type
                End If
            End With
            pbGetHeaderItemInfo = True
        End If
    Else
        If ((tHI.Mask And HDI_FILTER) = HDI_FILTER) Then
            ReDim bFilter(0 To 255) As Byte
            tHDTF.pszText = VarPtr(bFilter(0))
            tHDTF.cchTextMax = UBound(bFilter)
            tHI.pvFilter = VarPtr(tHDTF)
        End If
        If Not (SendMessageA(m_lHdrHandle, HDM_GETITEM, lCol, tHI) = 0) Then
            If ((tHI.Mask And HDI_FILTER) = HDI_FILTER) Then
                lSize = lstrlenA(VarPtr(bFilter(0)))
                ReDim Preserve bFilter(0 To lSize - 1) As Byte
                sFilter = StrConv(bFilter, vbUnicode)
            End If
            pbGetHeaderItemInfo = True
        End If
    End If

End Function

Private Function pbSetHeaderItemInfo(ByVal lCol As Long, _
                                     tHI As HDITEMA, _
                                      Optional sFilter As String) As Boolean
   
    If m_bIsNt Then
        Dim tHIW As HDITEMW
        Dim b() As Byte
        Dim bFilter() As Byte
        Dim lSize As Long
        Dim tHDTF As HDTEXTFILTER
        '/* Copy fields to tHIW structure
        With tHIW
            .cxy = tHI.cxy
            .fmt = tHI.fmt
            .hbm = tHI.hbm
            .iImage = tHI.iImage
            .iOrder = tHI.iOrder
            .lParam = tHI.lParam
            .Mask = tHI.Mask
            If ((.Mask And HDI_TEXT) = HDI_TEXT) Then
                If (Len(tHI.pszText) > 0) Then
                    b = tHI.pszText
                    ReDim Preserve b(0 To UBound(b) + 2) As Byte
                Else
                    ReDim b(0 To 1) As Byte
                End If
                .cchTextMax = (UBound(b) + 1) / 2
                .pszText = VarPtr(b(0))
            End If
            If ((.Mask And HDI_FILTER) = HDI_FILTER) Then
                If (Len(sFilter) > 0) Then
                    bFilter = sFilter
                    ReDim Preserve bFilter(0 To UBound(bFilter) + 2) As Byte
                Else
                    ReDim bFilter(0 To 1) As Byte
                End If
                tHDTF.pszText = VarPtr(bFilter(0))
                tHDTF.cchTextMax = (UBound(bFilter) + 1) / 2
                .pvFilter = VarPtr(tHDTF)
            End If
            .type = tHI.type
        End With
        '/* Send message as Unicode:
        If Not (SendMessageW(m_lHdrHandle, HDM_SETITEMW, lCol, tHIW) = 0) Then
            pbSetHeaderItemInfo = True
        End If
    Else
        If ((tHI.Mask And HDI_FILTER) = HDI_FILTER) Then
            If (Len(sFilter) > 0) Then
                bFilter = StrConv(sFilter, vbFromUnicode)
                ReDim Preserve bFilter(0 To UBound(bFilter) + 1) As Byte
            Else
                ReDim bFilter(0 To 0) As Byte
            End If
            tHDTF.cchTextMax = UBound(bFilter) + 1
            tHDTF.pszText = VarPtr(bFilter(0))
            tHI.pvFilter = VarPtr(tHDTF)
        End If
        If Not (SendMessageA(m_lHdrHandle, HDM_SETITEM, lCol, tHI) = 0) Then
            pbSetHeaderItemInfo = True
        End If
    End If

End Function

Private Sub pResize(Optional ByVal bForceUpdate As Boolean = False)

    If Not m_lHdrHandle = 0 Then
        m_bInitSize = True
        With m_tR
            MoveWindow m_lHdrHandle, .left, .top, .right - .left, .bottom - .top, 1
        End With
        'If Not (isXp) Or bForceUpdate Then
            InvalidateRect m_lHdrHandle, m_tR, 1
            UpdateWindow m_lHdrHandle
        'End If
    End If
   
End Sub

Private Sub pOLEFontToLogFont(fntThis As StdFont, _
                              hdc As Long, _
                              tLF As LOGFONT)

Dim sFont As String
Dim iChar As Integer

    '/* Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name
        '/* There is a quicker way involving StrConv and CopyMemory, but
        '/* this is simpler!:
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
        Next iChar
        '/* Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
    End With

End Sub

Private Sub pCreateHeader()

Dim wStyle      As Long
Dim lPtr        As Long
Dim rcParent    As RECT
Dim wP          As WINDOWPOS

    '/* create the header control.
    wStyle = WS_CHILD Or WS_VISIBLE Or HDS_HORZ
    If m_bHotTrack Then
        wStyle = wStyle Or HDS_HOTTRACK
    End If
    If m_bDragReorderColumns Then
        wStyle = wStyle Or HDS_DRAGDROP
    End If
    If m_bButtons Then
        wStyle = wStyle Or HDS_BUTTONS
    End If
    If m_bFullDrag Then
        wStyle = wStyle Or HDS_FULLDRAG
    End If
    If m_bFilterBar Then
        wStyle = wStyle Or HDS_FILTERBAR
    End If
   
    If Not m_bInitSize Then
        GetClientRect m_lParentHwnd, m_tR
        m_tR.bottom = m_tR.top + 20
        m_bInitSize = True
    End If
   
    With m_tR
        If m_bIsNt Then
            m_lHdrHandle = CreateWindowExW(0&, StrPtr(WC_HEADER), StrPtr(""), wStyle, 0&, 0&, _
                .right - .left, .bottom - .top, m_lParentHwnd, 0&, App.hInstance, 0&)
        Else
            m_lHdrHandle = CreateWindowEx(0&, WC_HEADER, "", wStyle, 0&, 0&, _
                .right - .left, .bottom - .top, m_lParentHwnd, 0&, App.hInstance, 0&)
        End If
    End With
    
    If Not m_lHdrHandle = 0 Then
        '/* Commence subclassing:
        SubclassAttach
        '/* Set the image list
        pSetImageList
        '/* If NT then we can have Unicode (thanks to Dana Seaman)
        If m_bIsNt Then
            SendMessageLongA m_lHdrHandle, CCM_SETUNICODEFORMAT, 1, 0&
        End If
        If Not (m_bUserMode) Then
            '/* why does the text not appear in design time?
            AddColumn "Header Control", 128
        End If
        Visible = m_bVisible
    End If

End Sub

Private Sub pSetImageList()

    If Not (m_hIml = 0) Then
        SendMessageLongA m_lHdrHandle, HDM_SETIMAGELIST, 0&, m_hIml
        If (m_hIml <> SendMessageLongA(m_lHdrHandle, HDM_GETIMAGELIST, 0&, 0&)) Then
          '  debugmsg "Error getting image list"
        End If
    End If
    
End Sub

Private Sub m_tmrDragDrop_ThatTime()

Dim rc  As RECT
Dim tP  As POINTAPI

    GetCursorPos tP
    GetWindowRect m_lParentHwnd, rc
    If (PtInRect(rc, tP.x, tP.y) = 0) Then
        m_cDrag.HideDragImage True
        If m_lLastDivider > -1 Then
            m_lLastDivider = SetHotDividerForCursorPos()
        End If
        m_bHidingDragImage = True
    Else
        GetWindowRect m_lHdrHandle, rc
        If (PtInRect(rc, tP.x, tP.y) = 0) Then
            '/* In control
            If (m_lLastDivider > -1) Then
                m_cDrag.HideDragImage True
                m_lLastDivider = SetHotDividerForCursorPos()
            ElseIf (m_bHidingDragImage) Then
            End If
            m_cDrag.HideDragImage False
        Else
            '/* In header:
            m_cDrag.HideDragImage True
            m_lLastDivider = SetHotDividerForCursorPos()
            m_cDrag.HideDragImage False
        End If
        m_bHidingDragImage = False
    End If

End Sub

Private Sub DrawGroupedHeaderItem(ByVal lHDC As Long, _
                                  ByVal lGroupIndex As Long, _
                                  tR As RECT)

Dim hBr         As Long
Dim hPenLight   As Long
Dim hPenDark    As Long
Dim lIconIndex  As Long
Dim lFmt        As Long
Dim sText       As String

    SetTextColor lHDC, GetSysColor(vbWindowText And &H1F&)
    '/* Fill background *** j. added button backcolor w/ optional skin
    If m_bGutterButtonSkinned Then
        SkinGutterButton lHDC, tR
    Else
        hPenLight = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
        hPenDark = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
        If Not m_lGutterButtonColor = vbButtonFace Then
            hBr = CreateSolidBrush(BlendColor(m_lGutterButtonColor, vbButtonFace))
        Else
            hBr = GetSysColorBrush(vbButtonFace And &H1F&)
        End If
        FillRect lHDC, tR, hBr
        DeleteObject hBr
        '/* Draw the border:
        Dim hPenTopLeft As Long
        Dim hPenRightBottom As Long
        Dim tJunk As POINTAPI
        Dim hPenOld As Long
        If (m_tGroupedCol(lGroupIndex).bPressed And m_tGroupedCol(lGroupIndex).bMouseOver) Then
            hPenTopLeft = hPenDark
            hPenRightBottom = hPenLight
        Else
            hPenRightBottom = hPenDark
            hPenTopLeft = hPenLight
        End If
        hPenOld = SelectObject(lHDC, hPenTopLeft)
        MoveToEx lHDC, tR.left, tR.bottom - 1, tJunk
        LineTo lHDC, tR.left, tR.top
        LineTo lHDC, tR.right - 1, tR.top
        SelectObject lHDC, hPenOld
        hPenOld = SelectObject(lHDC, hPenRightBottom)
        MoveToEx lHDC, tR.right - 1, tR.top, tJunk
        LineTo lHDC, tR.right - 1, tR.bottom - 1
        LineTo lHDC, tR.left, tR.bottom - 1
        SelectObject lHDC, hPenOld
    End If
    
    With tR
        .left = .left + 1
        .right = .right - 1
        .top = .top + 1
        .bottom = .bottom - 1
    End With
    
    '/* Get the text format:
    lFmt = DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
    If ((ColumnTextAlign(m_tGroupedCol(lGroupIndex).lColumn) And HdrTextALignCentre) = HdrTextALignCentre) Then
        lFmt = lFmt Or DT_CENTER
    ElseIf ((ColumnTextAlign(m_tGroupedCol(lGroupIndex).lColumn) And HdrTextALignRight) = HdrTextALignRight) Then
        lFmt = lFmt Or DT_RIGHT
    End If
    If ((ColumnTextAlign(m_tGroupedCol(lGroupIndex).lColumn) And HdrTextAlignRTLReading) = HdrTextAlignRTLReading) Then
        lFmt = lFmt Or DT_RTLREADING
    End If
   
    If (m_tGroupedCol(lGroupIndex).bPressed) And (m_tGroupedCol(lGroupIndex).bMouseOver) Then
        With tR
            .top = .top + 1
            .left = .left + 1
            .bottom = .bottom + 1
            .right = .right + 1
        End With
   End If
   
    '/* Draw icon if any:
    lIconIndex = ColumnImage(m_tGroupedCol(lGroupIndex).lColumn)
    With tR
        If (lIconIndex > -1) Then
            If (ColumnImageOnRight(m_tGroupedCol(lGroupIndex).lColumn)) Then
                ImageList_Draw m_hIml, lIconIndex, lHDC, .right - 20, .top + (.bottom - .top - 16) \ 2, ILD_TRANSPARENT
                .right = .right - 20
            Else
                ImageList_Draw m_hIml, lIconIndex, lHDC, .left + 4, .top + (.bottom - .top - 16) \ 2, ILD_TRANSPARENT
                .left = .left + 20
            End If
        End If
        .right = .right - 1
        .left = .left + 1
    End With
    
    sText = ColumnHeader(m_tGroupedCol(lGroupIndex).lColumn)
    '/* Draw text:
    SetBkMode lHDC, Transparent
    If (m_bIsNt) Then
        DrawTextW lHDC, StrPtr(" " & sText & " "), -1, tR, lFmt
    Else
        DrawTextA lHDC, " " & sText & " ", -1, tR, lFmt
    End If
    
    DeleteObject hPenLight
    DeleteObject hPenDark

End Sub

'////////////


Private Sub BackbufferImages()
'/* store state images

Dim lCt         As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long

    ReDim m_lHdc(3)
    ReDim m_lBmp(3)
    ReDim m_lBmpOld(3)

    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)

    For lCt = 0 To 3
        m_lHdc(lCt) = CreateCompatibleDC(lTmpDc)
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lHeaderWidth, m_lHeaderHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lHeaderWidth, m_lHeaderHeight, m_cHeaderDc.hdc, (m_lHeaderWidth * lCt), 0, SRCCOPY
    Next lCt

    ReleaseDC lHwnd, lTmpDc

End Sub

Private Sub BackbufferDestroy()
'/* destroy backbuffer

Dim lCt As Long

    For lCt = 0 To 3
        SelectObject m_lHdc(lCt), m_lBmpOld(lCt)
        DeleteObject m_lBmp(lCt)
        DeleteDC m_lHdc(lCt)
        m_lBmpOld(lCt) = 0
        m_lHdc(lCt) = 0
        m_lBmp(lCt) = 0
    Next lCt

End Sub

Private Sub ColumnRender(ByVal lOffset As Long, _
                         ByVal lColNum As Long)

'/* render columns

Dim bGrouped        As Boolean
Dim bHasIcon        As Boolean
Dim lCt             As Long
Dim lTmpDc          As Long
Dim lHwnd           As Long
Dim lDrawDc         As Long
Dim lBmp            As Long
Dim lBmpOld         As Long
Dim lx              As Long
Dim lHDC            As Long
Dim lHdrEnd         As Long
Dim tHdr            As RECT
Dim tTemp           As RECT
Dim tRect           As RECT

On Error GoTo Handler

    '/* get coordinates
    GetClientRect m_lHdrHandle, tHdr
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lDrawDc = CreateCompatibleDC(lTmpDc)
    lHDC = GetDC(m_lHdrHandle)
    LSet tRect = tHdr
    
    With tHdr
        OffsetRect tHdr, -.left, -.top
        lBmp = CreateCompatibleBitmap(lTmpDc, .right, .bottom)
    End With
    lBmpOld = SelectObject(lDrawDc, lBmp)

    lx = 0
    For lCt = 0 To ColumnCount - 1
        If Not ColumnIsGrouped(lCt) Then
            If m_bIsNt Then
                SendMessageW m_lHdrHandle, HDM_GETITEMRECT, lCt, tTemp
            Else
                SendMessageA m_lHdrHandle, HDM_GETITEMRECT, lCt, tTemp
            End If
            With tTemp
                If lCt = lColNum Then: lx = lOffset: Else: lx = 0
                '/* left side
                m_cRender.Stretch lDrawDc, .left, 0, 3, .bottom, m_lHdc(lx), 0, 0, 3, m_lHeaderHeight, SRCCOPY
                '/* center
                m_cRender.Stretch lDrawDc, .left + 3, 0, (.right - .left) - 3, .bottom, m_lHdc(lx), 3, 0, m_lHeaderWidth - 3, m_lHeaderHeight, SRCCOPY
            End With
            bHasIcon = DrawColumnIcon(lCt, lDrawDc, tTemp)
            If lx = 2 Then
                DrawColumnText lDrawDc, lCt, True, tTemp, bHasIcon
            Else
                DrawColumnText lDrawDc, lCt, False, tTemp, bHasIcon
            End If
        Else
            bGrouped = True
        End If
    Next lCt
    
    '/* draw header end piece
    If m_bIsNt Then
        SendMessageW m_lHdrHandle, HDM_GETITEMRECT, ColumnAtIndex(ColumnCount - 1), tTemp
    Else
        SendMessageA m_lHdrHandle, HDM_GETITEMRECT, ColumnAtIndex(ColumnCount - 1), tTemp
    End If
    lHdrEnd = tTemp.right
    
    With tRect
        m_cRender.Stretch lDrawDc, lHdrEnd, 0, (.right - (lHdrEnd / 2)), .bottom, m_lHdc(0), 3, 0, m_lHeaderWidth, m_lHeaderHeight, SRCCOPY
        '/* draw group preheader
        If bGrouped Then
            If m_bIsNt Then
                SendMessageW m_lHdrHandle, HDM_GETITEMRECT, ColumnAtIndex(0), tTemp
            Else
                SendMessageA m_lHdrHandle, HDM_GETITEMRECT, ColumnAtIndex(0), tTemp
            End If
            m_cRender.Stretch lDrawDc, 0, 0, tTemp.right, .bottom, m_lHdc(0), 0, 0, m_lHeaderWidth, m_lHeaderHeight, SRCCOPY
        End If
        m_cRender.Blit lHDC, .left, .top, .right, .bottom, lDrawDc, 0, 0, SRCCOPY
    End With

    SelectObject lDrawDc, lBmpOld
    DeleteObject lBmp
    DeleteDC lDrawDc
    ReleaseDC lHwnd, lTmpDc
    ReleaseDC m_lHdrHandle, lHDC

Handler:
    On Error GoTo 0

End Sub

Private Function DrawColumnIcon(ByVal lColNum As Long, _
                                ByVal lHDC As Long, _
                                ByRef tRect As RECT) As Boolean
'/* draw header icon

Dim lIconIndex         As Long

On Error Resume Next

    If lColNum = -1 Then Exit Function
    If m_hIml = 0 Then Exit Function
    lIconIndex = ColumnImage(lColNum)
    If (lIconIndex > -1) Then
        With tRect
            ImageList_Draw m_hIml, lIconIndex, lHDC, .left + 4, .top + (.bottom - .top - 16) \ 2, ILD_TRANSPARENT
        End With
        DrawColumnIcon = True
    End If
    
On Error GoTo 0

End Function

Private Sub DrawColumnText(ByVal lHDC As Long, _
                           ByVal lColNum As Long, _
                           ByVal bHighLite As Boolean, _
                           ByRef tRect As RECT, _
                           ByVal bOffset As Boolean)

'/* draw panel caption
Dim sText       As String
Dim lHFont      As Long
Dim lFontOld    As Long
Dim tPnt        As POINTAPI
Dim tRText      As RECT

On Error Resume Next

    sText = ColumnHeader(lColNum)
    If LenB(sText) = 0 Then Exit Sub
    lFontOld = SelectObject(lHDC, m_hFnt)
    SetBkMode lHDC, 1

    If bHighLite Then
        SetTextColor lHDC, m_lTextHighLite
    Else
        SetTextColor lHDC, m_lTextForeColor
    End If
    
    If bOffset Then
        Dim tShft As RECT
        CopyRect tShft, tRect
        '/ calculate text size
        If m_bIsNt Then
            GetTextExtentPoint32 lHDC, sText, lstrlenW(sText) + 2, tPnt
        Else
            GetTextExtentPoint32 lHDC, sText, lstrlenA(sText) + 2, tPnt
        End If
        With tRect
            tShft.left = .left + 20
            '/* test min size
            If (.right - .left) < (tPnt.x + 40) Then
                Dim sTmp As String
                sTmp = String(255, 0)
                sTmp = sText & vbNullChar
                '/* compact text
                PathCompactPath lHDC, sTmp, (.right - .left) - 40
            Else
                sTmp = sText
            End If
        End With
        LSet tRText = tRect
        tRText.left = tRText.left + 24
        If m_bIsNt Then
            DrawTextW lHDC, StrPtr(sTmp), -1, tRText, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        Else
            DrawTextA lHDC, StrPtr(sTmp), -1, tRText, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        End If
    Else
        LSet tRText = tRect
        tRText.left = tRText.left + 4
        If m_bIsNt Then
            DrawTextW lHDC, StrPtr(sText), -1, tRText, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        Else
            DrawTextA lHDC, StrPtr(sText), -1, tRText, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        End If
    End If

    SelectObject lHDC, lFontOld

On Error GoTo 0

End Sub

Private Function ColumnFocused(ByVal lHwnd As Long, _
                               ByVal lParam As Long) As Long

Dim sX  As Single
Dim sY  As Single
Dim tP  As POINTAPI

    GetCursorPos tP
    ScreenToClient m_lHdrHandle, tP
    sX = tP.x
    sY = tP.y + m_tR.top
    ColumnFocused = ColumnHitTest(sX, sY)

End Function

'
Private Function FunctionExported(ByVal sFunction As String, _
                                  ByVal sModule As String) As Boolean
'/* test for library support

Dim lModule         As Long
Dim bLoaded         As Boolean

    lModule = GetModuleHandleA(sModule)
    If lModule = 0 Then
        lModule = LoadLibraryA(sModule)
    End If
    If lModule Then
        bLoaded = True
    End If

    If lModule Then
        If GetProcAddress(lModule, sFunction) Then
            FunctionExported = True
        End If
    End If

    If bLoaded Then
        FreeLibrary lModule
    End If

End Function

Private Sub TrackMouse(ByVal lHwnd As Long)

Dim tme         As TRACKMOUSEEVENT_STRUCT

    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE
        .hwndTrack = lHwnd
    End With

    If m_bTrackUser32 Then
        TrackMouseEvent tme
    Else
        TrackMouseEventComCtl tme
    End If

End Sub

Private Function DefaultHeaderBg(lHDC As Long)

    DefaultHeaderBg = GetPixel(lHDC, 1, 1)

End Function

Friend Sub LoadSkin()

    If m_bCustomHeader Then
        SkinHeaders
    End If

End Sub

'/////////////

Private Sub SubclassAttach()

    If m_bUserMode Then
        If Not m_lHdrHandle = 0 Then
            If Not m_GXHeader Is Nothing Then
                SubclassDetach
            End If
        Set m_GXHeader = New GXMSubclass
        With m_GXHeader
            .Subclass m_lParentHwnd, Me
            .AddMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
            .Subclass m_lHdrHandle, Me
            .AddMessage m_lHdrHandle, UM_STARTDRAG, MSG_BEFORE
            .AddMessage m_lHdrHandle, UM_ENDDRAG, MSG_BEFORE
         End With
         m_bSubClass = True
      End If
   End If
    
End Sub

Public Sub SkinAttach()

    If Not m_GXHeader Is Nothing Then
        With m_GXHeader
            .AddMessage m_lHdrHandle, WM_PAINT, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_MOUSEMOVE, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_MOUSELEAVE, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_LBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_LBUTTONUP, MSG_BEFORE
            .AddMessage m_lHdrHandle, WM_ERASEBKGND, MSG_BEFORE
        End With
    End If
            
End Sub

Private Sub SubclassDetach()

    If Not m_lHdrHandle = 0 Then
        If m_bSubClass Then
            With m_GXHeader
                .DeleteMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
                .UnSubclass m_lParentHwnd
                .DeleteMessage m_lHdrHandle, UM_STARTDRAG, MSG_BEFORE
                .DeleteMessage m_lHdrHandle, UM_ENDDRAG, MSG_BEFORE
                .UnSubclass m_lHdrHandle
            End With
        End If
    End If
    
End Sub

Public Sub SkinDetach()

    If Not m_GXHeader Is Nothing Then
        With m_GXHeader
            .DeleteMessage m_lHdrHandle, WM_PAINT, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_MOUSEMOVE, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_MOUSELEAVE, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_LBUTTONDOWN, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_LBUTTONUP, MSG_BEFORE
            .DeleteMessage m_lHdrHandle, WM_ERASEBKGND, MSG_BEFORE
        End With
    End If
    
End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Dim bCancel             As Boolean
Dim lHDI()              As Long
Dim i                   As Long
Dim sCol                As String
Dim tNMH                As NMHDR
Dim tHDN                As NMHEADER
Dim tPAINT              As PAINTSTRUCT
Static lSelected        As Long
Static bSelect          As Boolean

    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            BeginPaint lHwnd, tPAINT
            Select Case m_lCurrState
            Case 1
                ColumnRender 1, lSelected
            Case 2
                ColumnRender 2, lSelected
            Case Else
                ColumnRender 0, 0
            End Select
            EndPaint lHwnd, tPAINT
            m_bPainting = False
            bHandled = True
        Else
            lReturn = m_GXHeader.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If
        
    Case WM_LBUTTONDOWN
        m_lCurrState = 1
        bSelect = True
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
        
    Case WM_LBUTTONUP
        m_lCurrState = 0
        bSelect = False
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
        
    Case WM_MOUSEMOVE
        If GetAsyncKeyState(vbLeftButton) = 0 Then
            lSelected = ColumnFocused(lHwnd, lParam)
            m_lCurrState = 2
            If Not bSelect Then
                SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
            End If
            TrackMouse lHwnd
        End If

    Case WM_MOUSELEAVE
        m_lCurrState = 0
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
    
    Case WM_ERASEBKGND
        lReturn = 1
        bHandled = True

    Case WM_NOTIFY
        CopyMemory tNMH, ByVal lParam, LenB(tNMH)
        Select Case tNMH.code
        Case HDN_BEGINTRACKA, HDN_BEGINTRACKW
            CopyMemory tHDN, ByVal lParam, Len(tHDN)
            ReDim lHDI(0 To 1) As Long
            CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
            '/* Check if this column is grouped:
            If m_bAllowGrouping Then
                If (ColumnIsGrouped(tHDN.iItem)) Then
                    '/* Don't allow it to be sized
                    lReturn = 1
                    bHandled = True
                    Exit Sub
                End If
            End If
            RaiseEvent StartColumnWidthChange(tHDN.iItem, lHDI(1), bCancel)
            If bCancel Then
                lReturn = 1
                bHandled = True
            End If
        
        '/* header size start tracking
        Case HDN_TRACKA, HDN_TRACKW
            CopyMemory tHDN, ByVal lParam, LenB(tHDN)
            ReDim lHDI(0 To 1) As Long
            CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
            RaiseEvent ColumnWidthChanging(tHDN.iItem, lHDI(1), bCancel)
            If bCancel Then
                lReturn = 1
                bHandled = True
            End If
        
        '/* header size end tracking
        Case HDN_ENDTRACKA, HDN_ENDTRACKW
            lReturn = m_GXHeader.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
            CopyMemory tHDN, ByVal lParam, LenB(tHDN)
            If m_bAllowGrouping Then
                If (ColumnIsGrouped(tHDN.iItem)) Then
                    '/* No events
                    'bHandled = True
                    Exit Sub
                End If
            End If
            ReDim lHDI(0 To 1) As Long
            CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
            m_lCol = tHDN.iItem
            m_lCXY = lHDI(1)
            RaiseEvent ColumnWidthChanged(tHDN.iItem, lHDI(1))
            m_lCol = -1
          
        Case HDN_DIVIDERDBLCLICKA, HDN_DIVIDERDBLCLICKW
            CopyMemory tHDN, ByVal lParam, LenB(tHDN)
            RaiseEvent DividerDblClick(tHDN.iItem)
            
        '/* header click
        Case HDN_ITEMCLICKA, HDN_ITEMCLICKW
            CopyMemory tHDN, ByVal lParam, LenB(tHDN)
            RaiseEvent ColumnClick(tHDN.iItem)
        
        '/* header doubleclick
        Case HDN_ITEMDBLCLICKA, HDN_ITEMDBLCLICKW
            CopyMemory tHDN, ByVal lParam, LenB(tHDN)
            RaiseEvent ColumnDblClick(tHDN.iItem)
        
        '/* header start drag
        Case HDN_BEGINDRAG
            If m_bPreventDrag Then
                lReturn = 1
                bHandled = True
            Else
                CopyMemory tHDN, ByVal lParam, LenB(tHDN)
                RaiseEvent ColumnBeginDrag(tHDN.iItem)
                Dim bForceManualDragDrop As Boolean
                RaiseEvent ColumnManualDragRequest(tHDN.iItem, bForceManualDragDrop)
                If bForceManualDragDrop Then
                    m_bPreventDrag = True
                    m_lCol = -1
                    lReturn = 1
                    bHandled = True
                ElseIf (m_bAllowGrouping And Not (m_bHideGroupingBox)) Then
                    m_iDragCol = tHDN.iItem
                    PostMessage m_lHdrHandle, UM_STARTDRAG, 0&, 0&
                    lReturn = 1
                    bHandled = True
                End If
            End If
        
        '/* header end drag
        Case HDN_ENDDRAG
            lReturn = m_GXHeader.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
            CopyMemory tHDN, ByVal lParam, LenB(tHDN)
            ReDim lHDI(0 To 8) As Long
            CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 36
            '/* check for cancel
            If (lHDI(8) > -1) Then
                m_lCol = -1
                PostMessage m_lHdrHandle, UM_ENDDRAG, tHDN.iItem, lHDI(8)
            End If
            m_lCol = -1
            
        Case HDN_FILTERBTNCLICK
            '/* Get NMHDFILTERBTNCLICK.  Only require NMHDR + 1
            CopyMemory tHDN, ByVal lParam, 16
            RaiseEvent ColumnFilterClick(tHDN.iItem)
         
        Case HDN_FILTERCHANGE
            '/* Filter change.  Don't get details of the item in
            '/* this event, so need to call GetItem..
            CopyMemory tHDN, ByVal lParam, LenB(tHDN)
            Dim sFilter As String
            sFilter = ColumnFilter(tHDN.iItem)
            RaiseEvent ColumnFilterChange(tHDN.iItem, sFilter)
        
        '/* right click header
        Case NM_RCLICK
            '/* Right click in control
            Dim sX As Single, sY As Single
            Dim tP As POINTAPI
            GetCursorPos tP
            ScreenToClient m_lHdrHandle, tP
            sX = tP.x
            sY = tP.y + m_tR.top
            RaiseEvent RightClick(sX, sY)
        
        '/* stop dragging mode
        Case NM_RELEASEDCAPTURE
            m_bPreventDrag = False
        End Select
        bHandled = True
    '/* start dragging
    Case UM_STARTDRAG
        RaiseEvent OleDrag
        m_lCol = -1
    
    '/* end dragging
    Case UM_ENDDRAG
        RaiseEvent ColumnEndDrag(wParam, lParam)
    End Select
   
End Sub

Public Sub CleanUp()

    If Not m_lHdrHandle = 0 Then
        If m_bCustomHeader Then
            SkinDetach
        End If
        SubclassDetach
        ShowWindow m_lHdrHandle, SW_HIDE
        SetParent m_lHdrHandle, 0
        DestroyWindow m_lHdrHandle
        If Not m_hFnt = 0 Then DeleteObject m_hFnt
        If m_bSkinLoaded Then
            BackbufferDestroy
            If Not m_cRender Is Nothing Then Set m_cRender = Nothing
            If Not m_cHeaderDc Is Nothing Then Set m_cHeaderDc = Nothing
            If Not IHeader Is Nothing Then Set IHeader = Nothing
            Erase m_lHdc
            Erase m_lBmp
            Erase m_lBmpOld
            m_lHdrHandle = 0
            m_lImlHandle = 0
            m_lParentHwnd = 0
            m_lHdrHandle = 0
        End If
        If Not m_GXHeader Is Nothing Then Set m_GXHeader = Nothing
    End If
   
End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub
