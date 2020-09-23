VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkinHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements GXISubclass


Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

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

Private Const ILD_TRANSPARENT                   As Long = &H1

Private Const LF_DEFAULT_QUALITY                As Long = 0
Private Const LF_DRAFT_QUALITY                  As Long = 1
Private Const LF_PROOF_QUALITY                  As Long = 2
Private Const LF_NONANTIALIASED_QUALITY         As Long = 3
Private Const LF_ANTIALIASED_QUALITY            As Long = 4
Private Const LF_CLEARTYPE_QUALITY              As Long = 5

Private Const LVM_FIRST                         As Long = &H1000
Private Const LVM_GETCOLUMNWIDTH                As Long = (LVM_FIRST + 29)
Private Const LVM_GETHEADER                     As Long = (LVM_FIRST + 31)

Private Const LOGPIXELSY                        As Long = 90

Private Const VER_PLATFORM_WIN32_NT             As Long = 2

Private Const WM_GETFONT                        As Long = &H31
Private Const WM_SETFONT                        As Long = &H30


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

Public Enum EHdrTextEffect
    HdrTextNormal = 0
    HdrTextEmbossed = 1
    HdrTextEngraved = 2
End Enum

Private Enum EColAlignFlags
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

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type


Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
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

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

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
                                                             ByVal pszPath As Long, _
                                                             ByVal dX As Long) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, _
                                                  ByVal lpString2 As Long) As Long


Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function CreateDc Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 lpDeviceName As Any, _
                                                                 lpOutput As Any, _
                                                                 lpInitData As Any) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

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

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "Comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Long, _
                                                      ByVal bErase As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long


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

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long


Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long) As Long

Private m_bPainting                             As Boolean
Private m_bTrackUser32                          As Boolean
Private m_bSkinLoaded                           As Boolean
Private m_bUseHeaderTheme                       As Boolean
Private m_bHeaderActive                         As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bDragState                            As Boolean
Private m_bIsNt                                 As Boolean
Private m_bIsXp                                 As Boolean
Private m_bColumnVerticalText                   As Boolean
Private m_bColumnSizeChange                     As Boolean
Private m_bSortDescending                       As Boolean
Private m_lHeaderTextEffect                     As Long
Private m_lColumnSpace                          As Long
Private m_lColumnSorted                         As Long
Private m_lvFntDc                               As Long
Private m_lHeaderDc                             As Long
Private m_lHeaderLuminence                      As Long
Private m_lThemeColor                           As Long
Private m_lHeaderSkinStyle                      As Long
Private m_lTextForeColor                        As Long
Private m_lTextHighLite                         As Long
Private m_lTextPressed                          As Long
Private m_lHGHwnd                               As Long
Private m_lHdrHwnd                              As Long
Private m_lHeaderBmpWidth                       As Long
Private m_lHeaderBmpHeight                      As Long
Private m_lHeaderHeight                         As Long
Private m_lImlHandle                            As Long
Private m_ImageWidth                            As Long
Private m_lHdc()                                As Long
Private m_lBmp()                                As Long
Private m_lBmpOld()                             As Long
Private m_lCurrState                            As Long
Private m_lhFnt                                 As Long
Private m_lImlHdHndl                            As Long
Private m_sngLuminence                          As Single
Private m_tR                                    As RECT
Private m_oFont                                 As StdFont
Private m_oVtFont                               As StdFont
Private m_pHeader                               As StdPicture
Private m_cRender                               As clsRender
Private m_cHeaderDc                             As clsStoreDc
Private m_GXHeader                              As GXMSubclass


Private Sub Class_Initialize()
    
    InitCommonControls
    VersionCheck
    m_bTrackUser32 = FunctionExported("TrackMouseEvent", "User32")
    '/* default colors
    m_lTextForeColor = &H111111
    m_lTextHighLite = &H676767
    m_lTextPressed = &HDEDEDE

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


'> Properties
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Property Get ColumnAlign(ByVal lColumn As Long) As EHdrTextAlign
'*/ retieve a columns text alignment

Dim tHI     As HDITEMA
Dim tHW     As HDITEMW

    If m_lHGHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    If m_bUseUnicode Then
        tHW.Mask = HDI_FORMAT
        SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
        ColumnAlign = (HD_CALGN And tHW.fmt)
    Else
        tHI.Mask = HDI_FORMAT
        SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
        ColumnAlign = (HD_CALGN And tHI.fmt)
    End If

End Property

Friend Property Get ColumnAtIndex(ByVal lIndex As Long) As Long

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

Private Property Get ColumnIcon(ByVal lColumn As Long) As Long
'*/ retieve header icon index

Dim tHI     As HDITEMA
Dim tHW     As HDITEMW

    If m_lHdrHwnd = 0 Then Exit Property
    ColumnIcon = -1
    If m_bUseUnicode Then
        tHW.Mask = HDI_FORMAT
        SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
        If (tHW.fmt And HDF_IMAGE) = HDF_IMAGE Then
            tHW.Mask = HDI_IMAGE
            SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
            ColumnIcon = tHW.iImage
        End If
    Else
        tHI.Mask = HDI_FORMAT
        SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
        If (tHI.fmt And HDF_IMAGE) = HDF_IMAGE Then
            tHI.Mask = HDI_IMAGE
            SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
            ColumnIcon = tHI.iImage
        End If
    End If

End Property

Private Property Let ColumnIcon(ByVal lColumn As Long, _
                               ByVal lIcon As Long)
'*/ change header icon

Dim lAlign      As Long
Dim uHDI        As HDITEMA

    If (m_lHGHwnd = 0) Or (m_lHdrHwnd = 0) Then Exit Property
    With uHDI
        .Mask = HDI_FORMAT
        If m_bUseUnicode Then
            SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, uHDI
        Else
            SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, uHDI
        End If
        lAlign = HD_CALGN And .fmt
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

Friend Property Get ColumnIndex(ByVal lColumn As Long) As Long

Dim tHI     As HDITEMA
Dim tHW     As HDITEMW

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

End Property

Friend Property Get ColumnSizeChange() As Boolean
    ColumnSizeChange = m_bColumnSizeChange
End Property

Friend Property Let ColumnSizeChange(ByVal PropVal As Boolean)
    m_bColumnSizeChange = PropVal
End Property

Friend Property Get ColumnSorted() As Long
    ColumnSorted = m_lColumnSorted
End Property

Friend Property Let ColumnSorted(ByVal PropVal As Long)
    m_lColumnSorted = PropVal
End Property

Private Property Get ColumnText(ByVal lColumn As Long) As String
'*/ get a columns heading

Dim aText(261)  As Byte
Dim lLen        As Long
Dim sTemp       As String
Dim tHI         As HDITEMA
Dim tHW         As HDITEMW

    If m_lHGHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    If m_bUseUnicode Then
        With tHW
            .pszText = VarPtr(aText(0))
            .cchTextMax = UBound(aText) + 1
            .Mask = HDI_TEXT
        End With
        SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
        ColumnText = PointerToString(tHW.pszText)
    Else
        With tHI
            sTemp = String(260, Chr$(0))
            .pszText = sTemp
            .cchTextMax = 261
            .Mask = HDI_TEXT
        End With
        SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
        ColumnText = tHI.pszText
        lLen = InStr(ColumnText, vbNullChar)
        If lLen Then
            ColumnText = left$(ColumnText, lLen - 1)
        End If
    End If

End Property

Private Property Get ColumnWidth(ByVal lColumn As Long) As Long
'*/ retrieve a columns width

Dim tHI As HDITEMA

    If m_lHdrHwnd = 0 Then Exit Property
    tHI.Mask = HDI_WIDTH
    SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
    ColumnWidth = tHI.cxy

End Property

Friend Property Get DragState() As Boolean
    DragState = m_bDragState
End Property

Friend Property Let DragState(ByVal PropVal As Boolean)
    m_bDragState = PropVal
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

Friend Property Get HeaderIml() As Long
    HeaderIml = m_lImlHdHndl
End Property

Friend Property Let HeaderIml(ByVal PropVal As Long)
    m_lImlHdHndl = PropVal
End Property

Friend Property Get HeaderLuminence() As Long
    HeaderLuminence = m_lHeaderLuminence
End Property

Friend Property Let HeaderLuminence(ByVal PropVal As Long)
    m_lHeaderLuminence = PropVal
End Property

Friend Property Get HeaderPressed() As OLE_COLOR
    HeaderPressed = m_lTextPressed
End Property

Friend Property Let HeaderPressed(PropVal As OLE_COLOR)
    m_lTextPressed = PropVal
End Property

Friend Property Get HeaderSkinStyle() As Long
    HeaderSkinStyle = m_lHeaderSkinStyle
End Property

Friend Property Let HeaderSkinStyle(ByVal PropVal As Long)
    m_lHeaderSkinStyle = PropVal
End Property

Friend Property Get HeaderTextEffect() As EHdrTextEffect
    HeaderTextEffect = m_lHeaderTextEffect
End Property

Friend Property Let HeaderTextEffect(ByVal PropVal As EHdrTextEffect)
    m_lHeaderTextEffect = PropVal
End Property

Friend Property Get HeaderThemeColor() As Long
    HeaderThemeColor = m_lThemeColor
End Property

Friend Property Let HeaderThemeColor(ByVal PropVal As Long)
    m_lThemeColor = PropVal
End Property

Private Property Get IHeader() As StdPicture
'/* header image
    Set IHeader = m_pHeader
End Property

Private Property Set IHeader(PropVal As StdPicture)
    Set m_pHeader = PropVal
End Property

Friend Property Get UseHeaderTheme() As Boolean
    UseHeaderTheme = m_bUseHeaderTheme
End Property

Friend Property Let UseHeaderTheme(ByVal PropVal As Boolean)
    m_bUseHeaderTheme = PropVal
End Property

Friend Property Get UseUnicode() As Boolean
    UseUnicode = m_bUseUnicode
End Property

Friend Property Let UseUnicode(ByVal PropVal As Boolean)
    m_bUseUnicode = PropVal
End Property


'> Procedures
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

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

Private Sub BackbufferImages()
'/* store state images

Dim lCt    As Long
Dim lTmpDc As Long
Dim lHwnd  As Long

    ReDim m_lHdc(3)
    ReDim m_lBmp(3)
    ReDim m_lBmpOld(3)

    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)

    For lCt = 0 To 3
        m_lHdc(lCt) = CreateCompatibleDC(lTmpDc)
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lHeaderBmpWidth, m_lHeaderBmpHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lHeaderBmpWidth, m_lHeaderBmpHeight, m_cHeaderDc.hdc, (m_lHeaderBmpWidth * lCt), 0, SRCCOPY
    Next lCt

    ReleaseDC lHwnd, lTmpDc

End Sub

Friend Function ColumnFocused() As Long

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

Friend Property Get ColumnVerticalText() As Boolean
    ColumnVerticalText = m_bColumnVerticalText
End Property

Friend Property Let ColumnVerticalText(ByVal PropVal As Boolean)

    If m_bHeaderActive Then
        If PropVal Then
            SetVerticalFont
        Else
            DestroyVericalFont
        End If
    End If
    m_bColumnVerticalText = PropVal
    
End Property

Friend Property Get ColumnSortDescending() As Boolean
    ColumnSortDescending = m_bSortDescending
End Property

Friend Property Let ColumnSortDescending(ByVal PropVal As Boolean)
    m_bSortDescending = PropVal
End Property

Private Sub DrawSortArrow(ByVal lHdc As Long, _
                          ByRef tRect As RECT)

Dim lhPen       As Long
Dim lhPenOld    As Long
Dim lArwClr     As Long
Dim lVtCtr      As Long
Dim lhBrush     As Long
Dim lOldBrush   As Long
Dim tPnt        As POINTAPI
    
    If m_lColumnSpace > tRect.right - 26 Then
        Exit Sub
    End If
    '/* vertical center
    lVtCtr = (tRect.bottom - tRect.top) / 2
    '/* arrow fill color
    lArwClr = GetPixelColor(lHdc, 0, 0)
    '/* outline pen
    lhPen = CreatePen(0&, 1&, &H808080)
    lhPenOld = SelectObject(lHdc, lhPen)
    '/* fill brush
    lhBrush = CreateSolidBrush(lArwClr)
    lOldBrush = SelectObject(lHdc, lhBrush)
    '/* track fill path
    BeginPath lHdc
    '/* draw outline
    With tRect
        If Not m_bSortDescending Then
            MoveToEx lHdc, .right - 24, lVtCtr + 3, tPnt
            LineTo lHdc, .right - 17, lVtCtr - 4
            LineTo lHdc, .right - 10, lVtCtr + 3
            LineTo lHdc, .right - 24, lVtCtr + 3
        Else
            MoveToEx lHdc, .right - 24, lVtCtr - 4, tPnt
            LineTo lHdc, .right - 17, lVtCtr + 3
            LineTo lHdc, .right - 10, lVtCtr - 4
            LineTo lHdc, .right - 24, lVtCtr - 4
        End If
    End With
    '/* finished tracking
    EndPath lHdc
    '/* fill the shape
    StrokeAndFillPath lHdc
    '/* cleanup
    SelectObject lHdc, lOldBrush
    DeleteObject lhBrush
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen

End Sub

Private Function GetPixelColor(ByVal lHdc As Long, _
                               ByVal lX As Long, _
                               ByVal lY As Long) As Long

    GetPixelColor = GetPixel(lHdc, lX, lY)

End Function

Friend Property Get HeaderDc() As Long

    If m_lHdrHwnd = 0 Then Exit Function
    m_lHeaderDc = GetDC(m_lHdrHwnd)
    ReleaseDC m_lHdrHwnd, m_lHeaderDc
    HeaderDc = m_lHeaderDc
    
End Property

Private Sub TextEngraved(ByVal lHdc As Long, _
                         ByVal sText As String, _
                         ByVal lFlags As Long, _
                         ByRef tRect As RECT)
'/* embossed text

Dim lFrClr As Long

    If m_lCurrState = 1 Then
        lFrClr = m_lTextPressed
    ElseIf m_lCurrState = 2 Then
        lFrClr = m_lTextHighLite
    Else
        lFrClr = m_lTextForeColor
    End If
    
    If m_bUseUnicode Then
        SetTextColor lHdc, &HEFEFEF
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        InflateRect tRect, 1, 1
        SetTextColor lHdc, lFrClr
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
    Else
        SetTextColor lHdc, vbWhite
        DrawTextA lHdc, sText, -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextA lHdc, sText, -1, tRect, lFlags
        InflateRect tRect, 1, 1
        SetTextColor lHdc, lFrClr
        DrawTextA lHdc, sText, -1, tRect, lFlags
    End If

End Sub

Private Sub TextEmbossed(ByVal lHdc As Long, _
                         ByVal sText As String, _
                         ByVal lFlags As Long, _
                         ByRef tRect As RECT)
'/* engraved text

Dim lFrClr As Long

    If m_lCurrState = 1 Then
        lFrClr = m_lTextPressed
    ElseIf m_lCurrState = 2 Then
        lFrClr = m_lTextHighLite
    Else
        lFrClr = m_lTextForeColor
    End If
    
    If m_bUseUnicode Then
        SetTextColor lHdc, vbWhite
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        OffsetRect tRect, 1, 0
        SetTextColor lHdc, lFrClr
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
    Else
        SetTextColor lHdc, &HCCCCCC
        DrawTextA lHdc, sText, -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextA lHdc, sText, -1, tRect, lFlags
        OffsetRect tRect, 1, 0
        SetTextColor lHdc, lFrClr
        DrawTextA lHdc, sText, -1, tRect, lFlags
    End If

End Sub

Private Sub ColumnRender(ByVal lOffset As Long, _
                         ByVal lColumn As Long)

'/* render columns

Dim lCt         As Long
Dim lDrawDc     As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lX          As Long
Dim lHdc        As Long
Dim lHdrEnd     As Long
Dim tHdr        As RECT
Dim tTemp       As RECT
Dim tRect       As RECT
Dim tRCtr       As RECT

On Error GoTo Handler

    '/* get coordinates
    GetClientRect m_lHdrHwnd, tHdr
    lHdc = GetDC(m_lHdrHwnd)
    lDrawDc = CreateCompatibleDC(lHdc)
    CopyRect tRect, tHdr

    With tHdr
        OffsetRect tHdr, -.left, -.top
        lBmp = CreateCompatibleBitmap(lHdc, .right, .bottom)
    End With
    lBmpOld = SelectObject(lDrawDc, lBmp)

    lX = 0
    For lCt = 0 To ColumnCount - 1
        If m_bUseUnicode Then
            SendMessageW m_lHdrHwnd, HDM_GETITEMRECT, lCt, tTemp
        Else
            SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, lCt, tTemp
        End If
        With tTemp
            If lCt = lColumn Then
                lX = lOffset
            Else
                lX = 0
            End If
            If m_bDragState Then
                lX = 0
            End If
            m_lHeaderHeight = (.bottom - .top)
            '/* left side
            m_cRender.Stretch lDrawDc, .left, 3, 3, (.bottom - 6), m_lHdc(lX), 0, 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
            '/* right side
            m_cRender.Stretch lDrawDc, (.right - 3), (.top + 3), 3, (.bottom - 6), m_lHdc(lX), (m_lHeaderBmpWidth - 3), 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
            '/* top left corner
            m_cRender.Stretch lDrawDc, .left, 0, 3, 3, m_lHdc(lX), 0, 0, 3, 3, SRCCOPY
            '/* top
            m_cRender.Stretch lDrawDc, (.left + 3), 0, ((.right - .left) - 3), 3, m_lHdc(lX), 3, 0, (m_lHeaderBmpWidth - 3), 3, SRCCOPY
            '/* bottom
            m_cRender.Stretch lDrawDc, .left, (.bottom - 3), (.right - .left), 3, m_lHdc(lX), 0, (m_lHeaderBmpHeight - 3), m_lHeaderBmpWidth, 3, SRCCOPY
            '/* center
            m_cRender.Stretch lDrawDc, (.left + 3), 3, ((.right - .left) - 6), (.bottom - 6), m_lHdc(lX), 3, 3, (m_lHeaderBmpWidth - 6), (m_lHeaderBmpHeight - 6), SRCCOPY
        End With
        DrawColumnText lDrawDc, lCt, lX, tTemp
        DrawColumnIcon lCt, lDrawDc, tTemp
    Next lCt

    '/* draw header end piece
    If m_bUseUnicode Then
        SendMessageW m_lHdrHwnd, HDM_GETITEMRECT, ColumnAtIndex(ColumnCount - 1), tTemp
    Else
        SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, ColumnAtIndex(ColumnCount - 1), tTemp
    End If
    lHdrEnd = tTemp.right

    With tRect
        '/* left side
        m_cRender.Stretch lDrawDc, lHdrEnd, 3, 3, (.bottom - 6), m_lHdc(0), 0, 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
        '/* right side
        m_cRender.Stretch lDrawDc, (.right - 3), (.top + 3), 3, (.bottom - 6), m_lHdc(0), (m_lHeaderBmpWidth - 3), 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
        '/* top
        m_cRender.Stretch lDrawDc, lHdrEnd, 0, (.right - .left), 3, m_lHdc(0), 0, 0, m_lHeaderBmpWidth, 3, SRCCOPY
        '/* bottom
        m_cRender.Stretch lDrawDc, lHdrEnd, (.bottom - 3), .right, 3, m_lHdc(0), 0, (m_lHeaderBmpHeight - 3), m_lHeaderBmpWidth, 3, SRCCOPY
        '/* center
        m_cRender.Stretch lDrawDc, (lHdrEnd + 3), 3, ((.right - .left) - 6), (.bottom - 6), m_lHdc(0), 3, 3, (m_lHeaderBmpWidth - 6), (m_lHeaderBmpHeight - 6), SRCCOPY
        '/* draw to screen
        m_cRender.Blit lHdc, .left, .top, .right, .bottom, lDrawDc, 0, 0, SRCCOPY
    End With

    SelectObject lDrawDc, lBmpOld
    DeleteObject lBmp
    lBmpOld = 0
    DeleteDC lDrawDc
    ReleaseDC m_lHdrHwnd, lHdc

Handler:
    On Error GoTo 0

End Sub

Private Function DrawColumnIcon(ByVal lColumn As Long, _
                                ByVal lHdc As Long, _
                                ByRef tRect As RECT) As Boolean
'/* draw header icon

Dim lIndex         As Long

On Error Resume Next

    If lColumn = -1 Then Exit Function
    If lColumn = m_lColumnSorted Then
    '    DrawSortArrow lHDC, tRect '<- from grid control
    End If
    If m_lImlHdHndl = 0 Then Exit Function
    lIndex = ColumnIcon(lColumn)
    If lIndex > -1 Then
        With tRect
            ImageList_Draw m_lImlHdHndl, lIndex, lHdc, (.left + 4), (.top + ((.bottom - .top - 16) \ 2)), ILD_TRANSPARENT
        End With
        DrawColumnIcon = True
    End If

On Error GoTo 0

End Function

Private Sub DrawColumnText(ByVal lHdc As Long, _
                           ByVal lColumn As Long, _
                           ByVal lState As Long, _
                           ByRef tRect As RECT)

'/* draw column caption
Dim bIcon       As Boolean
Dim lFontOld    As Long
Dim lAlign      As Long
Dim lWidth      As Long
Dim lOffset     As Long
Dim lFlags      As Long
Dim sText       As String
Dim sTmp        As String
Dim tPnt        As POINTAPI
Dim tRText      As RECT
Dim tShft       As RECT

On Error GoTo Handler

    sText = ColumnText(lColumn)
    If LenB(sText) = 0 Then Exit Sub
    lFontOld = SelectObject(lHdc, m_lhFnt)
    SetBkMode lHdc, 1

    '/* pressed
    If lState = 1 Then
        SetTextColor lHdc, m_lTextPressed
        '/* over
    ElseIf lState = 2 Then
        SetTextColor lHdc, m_lTextHighLite
    Else
        SetTextColor lHdc, m_lTextForeColor
    End If

    '/* calculate text ellipses
    CopyRect tShft, tRect
    '/ calculate text size
    If m_bUseUnicode Then
        GetTextExtentPoint32W lHdc, StrPtr(sText), lstrlenW(StrPtr(sText)), tPnt
    Else
        GetTextExtentPoint32A lHdc, sText, Len(sText), tPnt
    End If
    
    '/* has icon
    bIcon = (ColumnIcon(lColumn) > -1)
    '/* text alignment
    lAlign = ColumnAlign(lColumn)

    LSet tRText = tRect
    lWidth = ColumnWidth(lColumn)
        
    If bIcon Then
        lOffset = m_ImageWidth + 2
    End If

    Select Case lAlign
    '/* left
    Case 0
        lOffset = lOffset + 4
        With tRText
            .left = (.left + lOffset)
        End With
    '/* right
    Case 1
        With tRText
            .left = (.left + (lWidth - tPnt.X) - 4)
        End With
    '/* center
    Case 2
        With tRText
            .left = (.left + (lWidth - tPnt.X) / 2)
        End With
    '/* left
    Case 3
        lOffset = lOffset + 4
        With tRText
            .left = (.left + lOffset)
        End With
    End Select

    With tRText
        '/* minimum x right align
        If lAlign = 1 Then
            If (.left + lOffset) < (tRect.left + lOffset + tPnt.X) Then
                .left = (tRect.left + lOffset + 8)
            End If
        '/* minimum x center align
        ElseIf lAlign = 2 Then
            If (.left + lOffset) < (tRect.left + lOffset + tPnt.X) Then
                .left = (tRect.left + lOffset) + 8
            End If
        '/* minimum x left align
        Else
            If (.left - 2) < tRect.left Then
                .left = (tRect.left)
            End If
        End If
    End With
    
    '/* draw elipses
    If m_bColumnVerticalText Then
        If ((tPnt.X + lOffset + 14) > lWidth) Then
            If ((tPnt.X + lOffset + 14) < m_lHeaderHeight) Then
                SelectObject lHdc, lFontOld
                lFontOld = 0
                lFontOld = SelectObject(lHdc, m_lvFntDc)
                lFlags = DT_LEFT Or DT_BOTTOM Or DT_SINGLELINE
                OffsetRect tRText, -3, 5
            Else
                tRText.right = tRText.right - 4
                lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
            End If
        Else
            lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        End If
    Else
        With tRText
            '/* test min size
            If ((tPnt.X + lOffset + 14) > lWidth) Then
                .right = .right - 4
                lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
            Else
                lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
            End If
        End With
    End If
    
    m_lColumnSpace = tRText.left + tPnt.X + lOffset

    If m_lHeaderTextEffect = 1 Then
        TextEmbossed lHdc, sText, lFlags, tRText
    ElseIf m_lHeaderTextEffect = 2 Then
        TextEngraved lHdc, sText, lFlags, tRText
    Else
        '/* draw normal text
        If m_bUseUnicode Then
            DrawTextW lHdc, StrPtr(sText), -1, tRText, lFlags
        Else
            DrawTextA lHdc, sText, -1, tRText, lFlags
        End If
    End If
    
Handler:
    '/* delete font
    SelectObject lHdc, lFontOld
    
On Error GoTo 0

End Sub

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

Private Function HeaderHwnd() As Long
'*/ return the column header handle

    If m_lHGHwnd = 0 Then Exit Function
    HeaderHwnd = SendMessageLongA(m_lHGHwnd, LVM_GETHEADER, 0&, 0&)
    
End Function

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

Private Function LoadHeaderImage() As Boolean

On Error GoTo Handler

    Select Case m_lHeaderSkinStyle
        '/* classic
    Case 0
        Set IHeader = LoadResPicture("CLASSIC-HEADER", vbResBitmap)
        '/* eclipse
    Case 1
        Set IHeader = LoadResPicture("ECLIPSE-HEADER", vbResBitmap)
        '/* lime
    Case 2
        Set IHeader = LoadResPicture("LIME-HEADER", vbResBitmap)
        '/* metallic
    Case 3
        Set IHeader = LoadResPicture("METALLIC-HEADER", vbResBitmap)
        '/* Gloss
    Case 4
        Set IHeader = LoadResPicture("GLOSS-HEADER", vbResBitmap)
    '/* XP
    Case 5
        Set IHeader = LoadResPicture("XP-HEADER", vbResBitmap)
    End Select
    If m_bUseHeaderTheme Then
        HeaderThemeSettings
    End If
    LoadHeaderImage = True

On Error GoTo 0

Handler:

End Function

Friend Sub LoadSkin(ByVal lHgHwnd As Long)

Dim lX  As Long
Dim lY  As Long

    SkinHeaders lHgHwnd
    '/* icon size
    If Not HeaderIml = 0 Then
        ImageList_GetIconSize HeaderIml, lX, lY
        If lX = 0 Then
            m_ImageWidth = 16
        End If
        m_ImageWidth = lX
    Else
        m_ImageWidth = 16
    End If
    If m_bColumnVerticalText Then
        SetVerticalFont
    Else
        DestroyVericalFont
    End If
    m_bHeaderActive = True
    
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
            lstrcpyW StrPtr(PointerToString), lpString
        End If
    End If

On Error GoTo 0

Handler:

End Function

Friend Function ResetHeaderSkin() As Boolean

    m_bHeaderActive = False
    If Not m_cHeaderDc Is Nothing Then Set m_cHeaderDc = Nothing
    If Not IHeader Is Nothing Then Set IHeader = Nothing
    If Not m_oFont Is Nothing Then Set m_oFont = Nothing
    If Not m_oVtFont Is Nothing Then Set m_oVtFont = Nothing
    m_bColumnVerticalText = False
    BackbufferDestroy
    'SkinDetach

End Function

Friend Sub SetFont(ByVal oFont As StdFont)

    DestroyFont
    m_lhFnt = CreateFont(oFont)
    
End Sub

Private Sub SetVerticalFont()

    DestroyVericalFont
    Set m_oVtFont = New StdFont
    m_oVtFont.Name = "Arial"
    m_oVtFont.SIZE = 9
    m_lvFntDc = CreateFont(m_oVtFont, True)

End Sub

Private Function CreateFont(ByVal oFont As StdFont, _
                            Optional ByVal bVertical As Boolean) As Long
'*/ change list font

Dim lChar   As Long
Dim lHdc    As Long
Dim uLF     As LOGFONT

On Error GoTo Handler

    If oFont Is Nothing Then
        DefaultFont
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
         If m_bIsXp Then
            .lfQuality = LF_CLEARTYPE_QUALITY
         Else
            .lfQuality = LF_ANTIALIASED_QUALITY
         End If
         If bVertical Then
            .lfEscapement = 900
        End If
    End With
    DeleteDC lHdc
    
    If m_bUseUnicode Then
        CreateFont = CreateFontIndirectW(uLF)
    Else
        CreateFont = CreateFontIndirectA(uLF)
    End If

On Error GoTo 0
Exit Function

Handler:

End Function

Private Sub DefaultFont()

    Set m_oFont = New StdFont
    With m_oFont
        .Charset = 3
        .Name = "MS Sans Serif"
        .Weight = 400
        .SIZE = 9
    End With

End Sub

Private Sub DestroyFont()

    If Not m_lhFnt = 0 Then
        DeleteObject m_lhFnt
        If Not m_oFont Is Nothing Then
            Set m_oFont = Nothing
        End If
        m_lhFnt = 0
    End If

End Sub

Private Sub DestroyVericalFont()

    If Not m_lvFntDc = 0 Then
        DeleteObject m_lvFntDc
        If Not m_oVtFont Is Nothing Then
            Set m_oVtFont = Nothing
        End If
        m_lvFntDc = 0
    End If

End Sub

Friend Sub Refresh(ByVal lColumn As Long)

Dim lCt As Long

    If lColumn = -1 Then
        For lCt = 0 To ColumnCount - 1
            InvalidateRect m_lHdrHwnd, ByVal 0&, 1&
            UpdateWindow m_lHdrHwnd
        Next lCt
    Else
        m_lCurrState = lColumn
        SendMessageLongA m_lHdrHwnd, WM_PAINT, 0&, 0&
    End If

End Sub

Private Sub SkinHeaders(ByVal lHgHwnd As Long)

Dim tBmp    As BITMAP

    If lHgHwnd = 0 Then Exit Sub
    m_lHGHwnd = lHgHwnd
    m_lHdrHwnd = HeaderHwnd
    If m_lHdrHwnd = 0 Then Exit Sub
    If m_bSkinLoaded Then
        If LoadHeaderImage Then
            Set m_cHeaderDc = New clsStoreDc
            With m_cHeaderDc
                .CreateFromPicture IHeader
                If m_bUseHeaderTheme Then
                    .ColorizeImage m_lThemeColor, m_sngLuminence
                End If
            End With
            '/* default sizes
            GetObject IHeader.Handle, Len(tBmp), tBmp
            With tBmp
                m_lHeaderBmpWidth = (.bmWidth / 4)
                m_lHeaderBmpHeight = .bmHeight
            End With
            '/* render class
            Set m_cRender = New clsRender
            BackbufferImages
        End If
    Else
        Set m_GXHeader = New GXMSubclass
        If LoadHeaderImage Then
            Set m_cHeaderDc = New clsStoreDc
            With m_cHeaderDc
                .CreateFromPicture IHeader
                If m_bUseHeaderTheme Then
                    .ColorizeImage m_lThemeColor, m_sngLuminence
                End If
            End With
            '/* default sizes
            GetObject IHeader.Handle, Len(tBmp), tBmp
            With tBmp
                m_lHeaderBmpWidth = (.bmWidth / 4)
                m_lHeaderBmpHeight = .bmHeight
            End With
            '/* render class
            Set m_cRender = New clsRender
            BackbufferImages
            SkinAttach
            m_bSkinLoaded = True
        End If
    End If

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


'> Subclass
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub SkinAttach()

    If m_lHdrHwnd = 0 Then Exit Sub
    If Not m_GXHeader Is Nothing Then
        If Not m_bSkinLoaded Then
            With m_GXHeader
                .Subclass m_lHdrHwnd, Me
                .AddMessage m_lHdrHwnd, WM_PAINT, MSG_BEFORE
                .AddMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
                .AddMessage m_lHdrHwnd, WM_MOUSELEAVE, MSG_BEFORE
                .AddMessage m_lHdrHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
                .AddMessage m_lHdrHwnd, WM_ERASEBKGND, MSG_BEFORE
                m_bSkinLoaded = True
            End With
        End If
    End If

End Sub

Private Sub SkinDetach()

    If m_lHdrHwnd = 0 Then Exit Sub
    If Not m_GXHeader Is Nothing Then
        If m_bSkinLoaded Then
            With m_GXHeader
               .DeleteMessage m_lHdrHwnd, WM_PAINT, MSG_BEFORE
               .DeleteMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
               .DeleteMessage m_lHdrHwnd, WM_MOUSELEAVE, MSG_BEFORE
               .DeleteMessage m_lHdrHwnd, WM_LBUTTONDOWN, MSG_BEFORE
               .DeleteMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
               .DeleteMessage m_lHdrHwnd, WM_ERASEBKGND, MSG_BEFORE
               .UnSubclass m_lHdrHwnd
               m_bSkinLoaded = False
            End With
        End If
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

Dim tPAINT          As PAINTSTRUCT
Static lSelected    As Long
Static bSelect      As Boolean

    If Not m_bHeaderActive Then Exit Sub

    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            BeginPaint lHwnd, tPAINT
            If m_bColumnSizeChange Then
                m_lCurrState = 0
            End If
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

    Case WM_ERASEBKGND
        lReturn = 1
        bHandled = True
    End Select

End Sub


'> Cleanup
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub CleanUp()

    If Not m_lHdrHwnd = 0 Then
        SkinDetach
        If Not m_GXHeader Is Nothing Then Set m_GXHeader = Nothing
        DestroyFont
        DestroyVericalFont
        BackbufferDestroy
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_oFont Is Nothing Then Set m_oFont = Nothing
        If Not m_oVtFont Is Nothing Then Set m_oFont = Nothing
        If Not m_cHeaderDc Is Nothing Then Set m_cHeaderDc = Nothing
        If Not IHeader Is Nothing Then Set IHeader = Nothing
        Erase m_lHdc
        Erase m_lBmp
        Erase m_lBmpOld
        m_lHdrHwnd = 0
        m_lImlHandle = 0
        m_lhFnt = 0
        m_lvFntDc = 0
        m_lHGHwnd = 0
        m_lHdrHwnd = 0
    End If

End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub


