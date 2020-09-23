VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'/* just a demo, not sure if it will be finished (if not, just use a manifest, -though manifest
'/* has text display bugs when using unicode, so test it)..
'/* todo* change tool window to a createwindow'd picturebox, and write in the hit testing
'/* -or, figure out a way to change toolwindow font..

Implements GXISubclass


Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700
Private Const LOGPIXELSY                        As Long = 90
Private Const DT_LEFT                           As Long = &H0
Private Const DT_CENTER                         As Long = &H1
Private Const DT_RIGHT                          As Long = &H2
Private Const DT_VCENTER                        As Long = &H4
Private Const DT_SINGLELINE                     As Long = &H20
Private Const DT_CALCRECT                       As Long = &H400
Private Const DT_END_ELLIPSIS                   As Long = &H8000&
Private Const DT_RTLREADING                     As Long = &H20000

Private Const ILD_TRANSPARENT                   As Long = &H1

Private Const HD_CALGN                          As Long = &H3

Private Const HDI_WIDTH                         As Long = &H1
Private Const HDI_TEXT                          As Long = &H2
Private Const HDI_FORMAT                        As Long = &H4
Private Const HDI_IMAGE                         As Long = &H20
Private Const HDI_ORDER                         As Long = &H80

Private Const HDF_LEFT                          As Long = &H0
Private Const HDF_RIGHT                         As Long = &H1
Private Const HDF_CENTER                        As Long = &H2
Private Const HDF_RTLREADING                    As Long = &H4
Private Const HDF_SORTDOWN                      As Long = &H200
Private Const HDF_SORTUP                        As Long = &H400
Private Const HDF_IMAGE                         As Long = &H800
Private Const HDF_STRING                        As Long = &H4000
Private Const HDF_BITMAP_ON_RIGHT               As Long = &H1000

Private Const HDM_FIRST                         As Long = &H1200
Private Const HDM_GETITEMCOUNT                  As Long = (HDM_FIRST + 0)
Private Const HDM_GETITEMA                      As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEMA                      As Long = (HDM_FIRST + 4)
Private Const HDM_GETITEMRECT                   As Long = (HDM_FIRST + 7)
Private Const HDM_GETITEMW                      As Long = (HDM_FIRST + 11)
Private Const HDM_SETITEMW                      As Long = (HDM_FIRST + 12)

Private Const ILC_MASK                          As Long = &H1
Private Const ILC_COLOR32                       As Long = &H20

Private Const LVCF_FMT                          As Long = &H1
Private Const LVCF_WIDTH                        As Long = &H2
Private Const LVCF_TEXT                         As Long = &H4
Private Const LVCF_SUBITEM                      As Long = &H8
Private Const LVCF_IMAGE                        As Long = &H10
Private Const LVCF_ORDER                        As Long = &H20

Private Const LVM_FIRST                         As Long = &H1000
Private Const LVM_GETCOLUMNA                    As Long = (LVM_FIRST + 25)
Private Const LVM_GETCOLUMNWIDTH                As Long = (LVM_FIRST + 29)
Private Const LVM_GETHEADER                     As Long = (LVM_FIRST + 31)
Private Const LVM_GETCOLUMNW                    As Long = (LVM_FIRST + 95)

Private Const SM_CXVSCROLL                      As Long = 2

Private Const SS_LEFT                           As Long = &H0
Private Const SS_CENTER                         As Long = &H1
Private Const SS_RIGHT                          As Long = &H2
Private Const SS_ICON                           As Long = &H3
Private Const SS_BLACKRECT                      As Long = &H4
Private Const SS_GRAYRECT                       As Long = &H5
Private Const SS_WHITERECT                      As Long = &H6
Private Const SS_BLACKFRAME                     As Long = &H7
Private Const SS_GRAYFRAME                      As Long = &H8
Private Const SS_WHITEFRAME                     As Long = &H9
Private Const SS_USERITEM                       As Long = &HA
Private Const SS_SIMPLE                         As Long = &HB
Private Const SS_LEFTNOWORDWRAP                 As Long = &HC
Private Const SS_OWNERDRAW                      As Long = &HD
Private Const SS_BITMAP                         As Long = &HE
Private Const SS_ENHMETAFILE                    As Long = &HF
Private Const SS_ETCHEDHORZ                     As Long = &H10
Private Const SS_ETCHEDVERT                     As Long = &H11
Private Const SS_ETCHEDFRAME                    As Long = &H12
Private Const SS_TYPEMASK                       As Long = &H1F
Private Const SS_NOPREFIX                       As Long = &H80
Private Const SS_NOTIFY                         As Long = &H100
Private Const SS_CENTERIMAGE                    As Long = &H200
Private Const SS_RIGHTJUST                      As Long = &H400
Private Const SS_REALSIZEIMAGE                  As Long = &H800
Private Const SS_SUNKEN                         As Long = &H1000
Private Const SS_ENDELLIPSIS                    As Long = &H4000
Private Const SS_PATHELLIPSIS                   As Long = &H8000
Private Const SS_WORDELLIPSIS                   As Long = &HC000
Private Const SS_ELLIPSISMASK                   As Long = &HC000

Private Const SW_HIDE                           As Long = &H0
Private Const SW_NORMAL                         As Long = &H1

Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_NOZORDER                      As Long = &H4

Private Const WM_USER                           As Long = &H400

Private Const WM_GETFONT                        As Long = &H31
Private Const WM_SETFONT                        As Long = &H30

Private Const WS_CHILD                          As Long = &H40000000
Private Const WS_EX_TOOLWINDOW                  As Long = &H80&
Private Const WS_EX_CLIENTEDGE                  As Long = &H200
Private Const WS_VSCROLL                        As Long = &H200000
Private Const GWL_STYLE                         As Long = -16

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum


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

Private Type HDITEMA
    Mask                                        As Long
    cxy                                         As Long
    pszText                                     As String
    hbm                                         As Long
    cchTextMax                                  As Long
    fmt                                         As Long
    lParam                                      As Long
    iImage                                      As Long
    iOrder                                      As Long
    type                                        As Long
    pvFilter                                    As Long
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
    type                                        As Long
    pvFilter                                    As Long
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

Private Type PAINTSTRUCT
    hdc                                         As Long
    fErase                                      As Boolean
    rcPaint                                     As RECT
    fRestore                                    As Boolean
    fIncUpdate                                  As Boolean
    rgbReserved(32)                             As Byte
End Type

Private Type BITMAP
    bmType                                      As Long
    bmWidth                                     As Long
    bmHeight                                    As Long
    bmWidthBytes                                As Long
    bmPlanes                                    As Integer
    bmBitsPixel                                 As Integer
    bmBits                                      As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                      As Long
    dwFlags                                     As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
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

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
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

Private Declare Function CreateDc Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 lpDeviceName As Any, _
                                                                 lpOutput As Any, _
                                                                 lpInitData As Any) As Long

Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

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

Private Declare Function GetTextExtentPoint32A Lib "gdi32" (ByVal hdc As Long, _
                                                            ByVal lpsz As String, _
                                                            ByVal cbString As Long, _
                                                            lpSize As POINTAPI) As Long

Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As Long, _
                                                            ByVal lpsz As Long, _
                                                            ByVal cbString As Long, _
                                                            lpSize As POINTAPI) As Long

Private Declare Function PathCompactPathA Lib "shlwapi.dll" (ByVal hdc As Long, _
                                                             ByVal pszPath As String, _
                                                             ByVal dX As Long) As Long

Private Declare Function PathCompactPathW Lib "shlwapi.dll" (ByVal hdc As Long, _
                                                             ByVal pszPath As String, _
                                                             ByVal dX As Long) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, _
                                                  lpString2 As Any) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal Length As Long)

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long

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

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()

Private Declare Function ImageList_Draw Lib "Comctl32" (ByVal hIml As Long, _
                                                        ByVal i As Long, _
                                                        ByVal hdcDst As Long, _
                                                        ByVal X As Long, _
                                                        ByVal Y As Long, _
                                                        ByVal fStyle As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImagelist As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_AddIcon Lib "Comctl32" (ByVal hImagelist As Long, _
                                                           ByVal hIcon As Long) As Long

Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImagelist As Long) As Long


Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long


Private m_bPainting                             As Boolean
Private m_bTrackUser32                          As Boolean
Private m_bSkinLoaded                           As Boolean
Private m_bUseHeaderTheme                       As Boolean
Private m_bHeaderActive                         As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bFontCreated                          As Boolean
Private m_lTextHighLite                         As Long
Private m_lTextForeColor                        As Long
Private m_lTextPressed                          As Long
Private m_lLVHwnd                               As Long
Private m_lHdrHwnd                              As Long
Private m_lParentHwnd                           As Long
Private m_lHeaderMask                           As Long
Private m_lHeaderWidth                          As Long
Private m_lHeaderHeight                         As Long
Private m_lImlHeader                            As Long
Private m_ImageWidth                            As Long
Private m_lCurrState                            As Long
Private m_lhFnt                                 As Long
Private m_lCol                                  As Long
Private m_lColOrder                             As Long
Private m_tULF                                  As LOGFONT
Private m_tR                                    As RECT
Private m_oFont                                 As StdFont
Private m_cXPStyle                              As clsXPStyle
Private m_GXPHeader                             As GXMSubclass


Private Sub Class_Initialize()

    InitCommonControls
    m_bTrackUser32 = FunctionExported("TrackMouseEvent", "User32")

End Sub


'> Properties
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Property Get ColumnAlign(ByVal lColumn As Long) As EHdrTextAlign
'*/ retieve a columns text alignment

Dim uLVC        As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    uLVC.Mask = LVCF_FMT
    If m_bUseUnicode Then
        SendMessageW m_lLVHwnd, LVM_GETCOLUMNW, lColumn, uLVC
    Else
        SendMessageA m_lLVHwnd, LVM_GETCOLUMNA, lColumn, uLVC
    End If
    ColumnAlign = (HD_CALGN And uLVC.fmt)
    
End Property


Private Property Get ColumnAtIndex(ByVal lIndex As Long) As Long

Dim i As Long

    For i = 0 To ColumnCount - 1
        If ColumnIndex(i) = lIndex Then
            ColumnAtIndex = i
            Exit For
        End If
    Next i

End Property

Private Property Get ColumnCount() As Long
'*/ retieve column count
    ColumnCount = SendMessageLongA(m_lHdrHwnd, HDM_GETITEMCOUNT, 0&, 0&)
End Property

Public Property Get ColumnIcon(ByVal lColumn As Long) As Long
'*/ retieve header icon index

Dim uHDI    As HDITEMA
Dim uLVC    As LVCOLUMN

    If m_lHdrHwnd = 0 Then Exit Property
    uHDI.Mask = HDI_FORMAT
    If m_bUseUnicode Then
        SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, uHDI
    Else
        SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, uHDI
    End If
    If (uHDI.fmt And HDF_IMAGE) = HDF_IMAGE Then
        uLVC.Mask = LVCF_IMAGE
        If m_bUseUnicode Then
            SendMessageW m_lLVHwnd, LVM_GETCOLUMNW, lColumn, uLVC
        Else
            SendMessageA m_lLVHwnd, LVM_GETCOLUMNA, lColumn, uLVC
        End If
        ColumnIcon = uLVC.iImage
    Else
        ColumnIcon = -1
    End If

End Property

Private Property Get ColumnIndex(ByVal lColumn As Long) As Long

Dim tHI     As HDITEMA
Dim tHW     As HDITEMW

    If Not (lColumn = m_lCol) Then
        If m_bUseUnicode Then
            tHW.Mask = HDI_ORDER
            If Not (SendMessageW(m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW) = 0) Then
                ColumnIndex = tHW.iOrder
            End If
        Else
            tHI.Mask = HDI_ORDER
            If Not (SendMessageA(m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI) = 0) Then
                ColumnIndex = tHI.iOrder
            End If
        End If
    Else
        ColumnIndex = m_lColOrder
    End If

End Property

Private Property Get ColumnText(ByVal lColumn As Long) As String
'*/ get a columns heading

Dim aText(261)  As Byte
Dim lLen        As Long
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

Public Property Get ColumnWidth(ByVal lColumn As Long) As Long
'*/ retrieve a columns length

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    ColumnWidth = SendMessageLongA(m_lLVHwnd, LVM_GETCOLUMNWIDTH, lColumn, 0&)

End Property

Friend Property Get HeaderActive() As Boolean
    HeaderActive = m_bHeaderActive
End Property

Friend Property Let HeaderActive(ByVal PropVal As Boolean)
    m_bHeaderActive = PropVal
End Property

Friend Property Get HeaderForeColor() As OLE_COLOR
    HeaderForeColor = m_lTextForeColor
End Property

Friend Property Let HeaderForeColor(PropVal As OLE_COLOR)
    m_lTextForeColor = PropVal
End Property

Friend Property Get HeaderHighLite() As OLE_COLOR
    HeaderHighLite = m_lTextHighLite
End Property

Friend Property Let HeaderHighLite(PropVal As OLE_COLOR)
    m_lTextHighLite = PropVal
End Property

Friend Property Get HeaderPressed() As OLE_COLOR
    HeaderPressed = m_lTextPressed
End Property

Friend Property Let HeaderPressed(PropVal As OLE_COLOR)
    m_lTextPressed = PropVal
End Property

Friend Property Get UseUnicode() As Boolean
    UseUnicode = m_bUseUnicode
End Property

Friend Property Let UseUnicode(ByVal PropVal As Boolean)
    m_bUseUnicode = PropVal
End Property


'> Procedures
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Function ColumnFocused() As Long

Dim sX As Single
Dim sY As Single
Dim tP As POINTAPI

    GetCursorPos tP
    ScreenToClient m_lHdrHwnd, tP
    sX = tP.X
    sY = tP.Y + m_tR.top
    ColumnFocused = ColumnHitTest(sX, sY)

End Function

Private Function ColumnHitTest(ByVal xPixels As Long, _
                               ByVal yPixels As Long) As Long

Dim i           As Long
Dim lY          As Long
Dim lIndex      As Long
Dim rc          As RECT

    lIndex = -1
    lY = yPixels
    For i = 0 To ColumnCount - 1
        If m_bUseUnicode Then
            SendMessageW m_lHdrHwnd, HDM_GETITEMRECT, i, rc
        Else
            SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, i, rc
        End If
        If Not (PtInRect(rc, xPixels, lY) = 0) Then
            lIndex = i
            Exit For
        End If
    Next i
    ColumnHitTest = lIndex

End Function

Private Sub ColumnRender(ByVal lOffset As Long, _
                         ByVal lColumn As Long)

'/* render columns

Dim lHdc        As Long
Dim lCt         As Long
Dim lXMax       As Long
Dim lColCnt     As Long
Dim tRect       As RECT
Dim tHdr        As RECT

On Error GoTo Handler

    lHdc = GetDC(m_lHeaderMask)
    lColCnt = ColumnCount - 1
    
    For lCt = 0 To lColCnt
        '/* get coordinates
        If m_bUseUnicode Then
            SendMessageW m_lHdrHwnd, HDM_GETITEMRECT, lCt, tHdr
        Else
            SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, lCt, tHdr
        End If
        
        GetClientRect m_lLVHwnd, tRect
        lXMax = tRect.right
        If tHdr.right > lXMax Then
            tHdr.right = lXMax
        Else
            If lCt = lColCnt Then
                If tHdr.right < lXMax Then
                    tHdr.right = lXMax
                End If
            End If
        End If
        
        With m_cXPStyle
            .hdc = lHdc
            .ParentHwnd = m_lHeaderMask
            .top = 0
            .left = tHdr.left
            .Height = 19
            .Width = (tHdr.right - tHdr.left)
            .PartID = 1
            If lCt = lColumn Then
                .State = lOffset
            Else
                .State = 1
            End If
            .UxDrawPart
        End With
        DrawColumnText lCt
        'DrawColumnIcon lCt
    Next lCt
        
    ReleaseDC m_lHeaderMask, lHdc
    
Handler:
    On Error GoTo 0

End Sub

Private Function CreateHeader()

Dim bOnDesktop As Boolean
    bOnDesktop = (m_lParentHwnd = GetDesktopWindow())
    If m_lHeaderMask = 0 Then
        m_lHeaderMask = CreateWindowEx(&H20, "Static", "", WS_CHILD, 0&, 0&, 0&, 0&, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
        'm_lHeaderMask = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_lParentHwnd, 0, App.hInstance, ByVal 0)
        Mask_Show m_lHeaderMask, False
    End If
    
End Function

Private Sub DefaultFont(ByRef oFont As StdFont)

    Set oFont = New StdFont
    With oFont
        .Charset = 3
        .Name = "MS Sans Serif"
        .Weight = 400
        .SIZE = 9
    End With

End Sub

Private Function DestroyFont() As Boolean
'*/ font cleanup

On Error GoTo Handler

    If Not m_lhFnt = 0 Then
        If DeleteObject(m_lhFnt) Then
            DestroyFont = True
            m_lhFnt = 0
        End If
    End If

On Error GoTo 0
Exit Function

Handler:

End Function

Private Function DestroyImlHeader() As Boolean
'*/ destroy small icons image list

On Error GoTo Handler

    If m_lImlHeader = 0 Then Exit Function
    If ImageList_Destroy(m_lImlHeader) Then
        DestroyImlHeader = True
        m_lImlHeader = 0
    End If

On Error GoTo 0
Exit Function

Handler:

End Function

Private Sub DrawColumnText(ByVal lColumn As Long)
'/* draw column caption

Dim lAlign  As Long
Dim sText   As String

On Error Resume Next

    sText = ColumnText(lColumn)
    If LenB(sText) = 0 Then Exit Sub
    'lHDC = GetWindowDC(m_lHeaderMask)
    'SetBkMode lHDC, 1
    'lFntOld = SelectObject(m_lHeaderMask, m_lhFnt)
    
    '/* text alignment
    lAlign = ColumnAlign(lColumn)
    Select Case lAlign
    '/* left
    Case 0
        With m_cXPStyle
            .Text = sText
            .UxDrawText sText, 4, UxDT_VCENTER Or UxDT_LEFT
        End With
    '/* right
    Case 1
        With m_cXPStyle
            .Text = sText
            .UxDrawText sText, -12, UxDT_VCENTER Or UxDT_RIGHT
        End With
    '/* center
    Case 2
        With m_cXPStyle
            .Text = sText
            .UxDrawText sText, 0, UxDT_VCENTER Or UxDT_CENTER
        End With
    '/* left
    Case 3
        With m_cXPStyle
            .Text = sText
            .UxDrawText sText, 4, UxDT_VCENTER Or UxDT_LEFT
        End With
    End Select

    'SelectObject lHDC, lFntOld
    'ReleaseDC m_lHdrHwnd, lHDC
    
On Error GoTo 0

End Sub

Private Function DrawColumnIcon(ByVal lColumn As Long) As Boolean
'/* draw header icon < requires 32 bit icon

Dim lIndex         As Long

On Error Resume Next

    If lColumn = -1 Then Exit Function
    lIndex = ColumnIcon(lColumn)
    If lIndex > -1 Then
        With m_cXPStyle
            .Class = "Header"
            .ImlHwnd = m_lImlHeader
            .PartID = HP_HEADERSORTARROW
            .State = lIndex + 1
            .top = 2
            .Width = 12
            .Height = 12
            .left = 40
            .UxDrawIcon 1
        End With
        DrawColumnIcon = True
    End If
    
On Error GoTo 0

End Function

Private Function FunctionExported(ByVal sFunction As String, _
                                  ByVal sModule As String) As Boolean
'/* test for library support

Dim lModule As Long
Dim bLoaded As Boolean

    If m_bUseUnicode Then
        lModule = GetModuleHandleW(StrPtr(sModule))
        If lModule = 0 Then
            lModule = LoadLibraryW(StrPtr(sModule))
        End If
    Else
        lModule = GetModuleHandleA(sModule)
        If lModule = 0 Then
            lModule = LoadLibraryA(sModule)
        End If
    End If
    If Not lModule = 0 Then
        bLoaded = True
    End If
    If lModule Then
        If GetProcAddress(lModule, StrPtr(sFunction)) Then
            FunctionExported = True
        End If
    End If
    If bLoaded Then
        FreeLibrary lModule
    End If

End Function

Private Function HasVertical() As Boolean

Dim lStyle  As Long

    lStyle = GetWindowLong(m_lLVHwnd, GWL_STYLE)
    HasVertical = (lStyle And WS_VSCROLL) <> 0

End Function

Private Function HeaderHwnd() As Long
'*/ return the column header handle

    If m_lLVHwnd = 0 Then Exit Function
    HeaderHwnd = SendMessageLongA(m_lLVHwnd, LVM_GETHEADER, 0&, 0&)
    
End Function

Public Function InitImlHeader() As Boolean
'*/ initialize xp header image list

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlHeader
    m_lImlHeader = ImageList_Create(16&, 16&, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    InitImlHeader = (Not m_lImlHeader = 0)
    
On Error GoTo 0
Exit Function

Handler:

End Function

Friend Sub LoadXpSkin(ByVal lListHwnd As Long, _
                      ByVal lParHnd As Long)
'/* entry point

    If Not lParHnd = 0 Then
        m_lParentHwnd = lParHnd
        m_lLVHwnd = lListHwnd
        InitImlHeader
        SkinHeaders lListHwnd
    End If

End Sub

Private Sub Mask_Show(ByVal lHwnd As Long, _
                      ByVal bVisible As Boolean)
    
    If lHwnd = 0 Then Exit Sub
    If bVisible Then
        ShowWindow lHwnd, SW_NORMAL
        m_bHeaderActive = True
    Else
        ShowWindow lHwnd, SW_HIDE
        m_bHeaderActive = False
    End If
    
End Sub

Private Function PointerToString(ByVal lpString As Long) As String
'/* get string from pointer

Dim lLen    As Long

On Error GoTo Handler

    If lpString Then
        lLen = lstrlenW(ByVal lpString)
        If lLen Then
            '/* allocate string with nLen chars
            PointerToString = String$(lLen, 0)
            lstrcpyW ByVal StrPtr(PointerToString), ByVal lpString
        End If
    End If

On Error GoTo 0

Handler:

End Function

Friend Function ResetHeaderSkin() As Boolean

    m_bHeaderActive = False
    If Not m_oFont Is Nothing Then
        Set m_oFont = Nothing
    End If
    ThemeDetach

End Function

Friend Sub SetFont(ByVal oFont As StdFont)
'*/ change list font

Dim lChar   As Long
Dim lHdc    As Long
Dim uLF     As LOGFONT

On Error GoTo Handler

    DestroyFont
    If oFont Is Nothing Then
        DefaultFont oFont
    End If
    
    lHdc = CreateDc("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    With uLF
         For lChar = 1 To Len(oFont.Name)
             .lfFaceName(lChar - 1) = CByte(Asc(Mid$(oFont.Name, lChar, 1)))
         Next lChar
         .lfHeight = -MulDiv(oFont.SIZE, GetDeviceCaps(lHdc, LOGPIXELSY), 72)
         .lfItalic = oFont.Italic
         .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
         .lfUnderline = oFont.Underline
         .lfStrikeOut = oFont.Strikethrough
         .lfCharSet = 4
    End With
    DeleteDC lHdc
    
    If m_bUseUnicode Then
        m_lhFnt = CreateFontIndirectW(uLF)
    Else
        m_lhFnt = CreateFontIndirectA(uLF)
    End If

On Error GoTo 0
Exit Sub

Handler:

End Sub

Private Sub SetHeaderFont()
'/* store header font handle

    If Not (m_lHdrHwnd = 0) Then '<- can't change a windows title font this way.. works for most everything else though
        If m_bUseUnicode Then
            m_lhFnt = SendMessageLongW(m_lHdrHwnd, WM_GETFONT, 0&, 0&)
            If m_lhFnt = 0 Then
                SendMessageLongW m_lHeaderMask, WM_SETFONT, m_lhFnt, True
            End If
        Else
            m_lhFnt = SendMessageLongA(m_lHdrHwnd, WM_GETFONT, 0&, 0&)
            If m_lhFnt = 0 Then
                SendMessageLongA m_lHeaderMask, WM_SETFONT, m_lhFnt, True
            End If
        End If
    End If

    If m_lhFnt = 0 Then
        '/ create new
        SetFont m_oFont
    Else
        m_bFontCreated = False
    End If

End Sub

Private Function SetHeaderPos()

Dim lScrOffset  As Long
Dim uPt         As POINTAPI
Dim tOnRt       As RECT

    GetWindowRect m_lHdrHwnd, tOnRt
    CopyMemory uPt, tOnRt, Len(uPt)
    ScreenToClient m_lParentHwnd, uPt
    
    With tOnRt
        OffsetRect tOnRt, -.left, -.top
    End With
    
    If HasVertical Then
        lScrOffset = GetSystemMetrics(SM_CXVSCROLL)
    End If
    
    With tOnRt
        If Not (m_lHeaderMask = 0) Then
            SetWindowPos m_lHeaderMask, 0, uPt.X, uPt.Y, (.right - .left) - lScrOffset, (.bottom - .top), SWP_SHOWWINDOW
        End If
    End With
    Mask_Show m_lHeaderMask, True
    
End Function

Private Sub SkinHeaders(ByVal lListHnd As Long)

    If lListHnd = 0 Then Exit Sub
    m_lLVHwnd = lListHnd
    m_lHdrHwnd = HeaderHwnd
    If m_lHdrHwnd = 0 Then Exit Sub
    Set m_GXPHeader = New GXMSubclass
    Set m_cXPStyle = New clsXPStyle
    With m_cXPStyle
        .Class = "Header"
        .PartID = 1
    End With
    CreateHeader
    SetHeaderPos
    SetHeaderFont
    ThemeAttach

End Sub

Private Sub TrackMouse(ByVal lHwnd As Long)

Dim tme     As TRACKMOUSEEVENT_STRUCT

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

Friend Sub UnLoadXpSkin()

    CleanUp

End Sub


'> Subclass
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ThemeAttach()

    If Not m_GXPHeader Is Nothing Then
        With m_GXPHeader
            .Subclass m_lHdrHwnd, Me
            .AddMessage m_lHdrHwnd, WM_PAINT, MSG_BEFORE
            .AddMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
            .AddMessage m_lHdrHwnd, WM_MOUSELEAVE, MSG_BEFORE
            .AddMessage m_lHdrHwnd, WM_LBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
            .AddMessage m_lHdrHwnd, WM_ERASEBKGND, MSG_BEFORE
            .AddMessage m_lHdrHwnd, WM_SIZE, MSG_BEFORE
            .AddMessage m_lHdrHwnd, WM_STYLECHANGED, MSG_BEFORE
        End With
        m_bSkinLoaded = True
    End If
            
End Sub

Private Sub ThemeDetach()

    If Not m_GXPHeader Is Nothing Then
        With m_GXPHeader
            .DeleteMessage m_lHdrHwnd, WM_PAINT, MSG_BEFORE
            .DeleteMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
            .DeleteMessage m_lHdrHwnd, WM_MOUSELEAVE, MSG_BEFORE
            .DeleteMessage m_lHdrHwnd, WM_LBUTTONDOWN, MSG_BEFORE
            .DeleteMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
            .DeleteMessage m_lHdrHwnd, WM_ERASEBKGND, MSG_BEFORE
            .DeleteMessage m_lHdrHwnd, WM_SIZE, MSG_BEFORE
            .DeleteMessage m_lHdrHwnd, WM_STYLECHANGED, MSG_BEFORE
            .UnSubclass m_lHdrHwnd
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

Dim tPAINT              As PAINTSTRUCT
Static lSelected        As Long
Static bSelect          As Boolean

    If Not m_bHeaderActive Then
        Exit Sub
    End If
    
    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            BeginPaint lHwnd, tPAINT
            Select Case m_lCurrState
            Case 1
                ColumnRender 3, lSelected
            Case 2
                ColumnRender 2, lSelected
            Case Else
                ColumnRender 1, lSelected
            End Select
            EndPaint lHwnd, tPAINT
            m_bPainting = False
        Else
            lReturn = m_GXPHeader.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If
        bHandled = True
        
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
            lSelected = ColumnFocused
            m_lCurrState = 2
            If Not bSelect Then
                SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
            End If
            TrackMouse lHwnd
        End If

    Case WM_MOUSELEAVE
        m_lCurrState = 0
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
    
    Case WM_STYLECHANGED
        SetHeaderPos
        
    Case WM_SIZE
        SetHeaderPos
    
    Case WM_ERASEBKGND
        lReturn = 1
        bHandled = True
    End Select
   
End Sub


'> Cleanup
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub CleanUp()

    If Not m_lHdrHwnd = 0 Then
        If m_bSkinLoaded Then
            ThemeDetach
            DestroyImlHeader
            DestroyWindow m_lHeaderMask
            m_lHeaderMask = 0
            m_lHdrHwnd = 0
            m_lImlHeader = 0
            m_lParentHwnd = 0
            m_lHdrHwnd = 0
            m_bSkinLoaded = False
        End If
        If Not m_oFont Is Nothing Then
            Set m_oFont = Nothing
        End If
        If Not m_cXPStyle Is Nothing Then Set m_cXPStyle = Nothing
        If Not m_GXPHeader Is Nothing Then Set m_GXPHeader = Nothing
    End If
   
End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub
