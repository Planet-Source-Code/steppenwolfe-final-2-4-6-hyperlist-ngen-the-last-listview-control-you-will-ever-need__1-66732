VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkinScrollbars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' what a pain in the ass this was ;o(

Implements GXISubclass


Private Const WM_VSCROLL                        As Long = &H115
Private Const WM_HSCROLL                        As Long = &H114
Private Const SPI_GETWHEELSCROLLLINES           As Long = &H68
Private Const FSB_FLAT_MODE                     As Long = 2
Private Const FSB_ENCARTA_MODE                  As Long = 1
Private Const FSB_REGULAR_MODE                  As Long = 0
Private Const SIF_RANGE                         As Long = &H1
Private Const SIF_PAGE                          As Long = &H2
Private Const SIF_POS                           As Long = &H4
Private Const SIF_ALL                           As Double = (SIF_RANGE Or SIF_PAGE Or SIF_POS)
Private Const OBJID_HSCROLL                     As Long = &HFFFFFFFA
Private Const OBJID_VSCROLL                     As Long = &HFFFFFFFB
Private Const SS_CUSTOMDRAW                     As Long = &HD
Private Const GWL_STYLE                         As Long = -16
Private Const GWL_EXSTYLE                       As Long = -20
Private Const WS_EX_TOOLWINDOW                  As Long = &H80&
Private Const WS_EX_CLIENTEDGE                  As Long = &H200
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_CHILD                          As Long = &H40000000
Private Const WS_HSCROLL                        As Long = &H100000
Private Const WS_VSCROLL                        As Long = &H200000
Private Const SB_LINEDOWN                       As Long = 1
Private Const SB_LINELEFT                       As Long = 0
Private Const SB_LINERIGHT                      As Long = 1
Private Const SB_LINEUP                         As Long = 0
Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SW_HIDE                           As Long = &H0
Private Const SW_NORMAL                         As Long = &H1


Private Enum SCR_STYLE
    SB_HORZ = 0
    SB_VERT = 1
    SB_BOTH = 3
    SB_SZR = 4
End Enum

'/* system metrics
Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CYSMCAPTION = 51
    SM_CXMINIMIZED = 57
    SM_CYMINIMIZED = 58
    SM_CXMAXTRACK = 59
    SM_CYMAXTRACK = 60
    SM_CXMAXIMIZED = 61
    SM_CYMAXIMIZED = 62
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Enum EFSStyleConstants
    efsRegular = FSB_REGULAR_MODE
    efsEncarta = FSB_ENCARTA_MODE
    efsFlat = FSB_FLAT_MODE
End Enum

Public Enum EFSScrollBarConstants
    efsHorizontal = SB_HORZ
    efsVertical = SB_VERT
End Enum

Public Enum EFSOrientationConstants
    efsoHorizontal
    efsoVertical
    efsoBoth
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

Private Type BITMAP
    bmType                                      As Long
    bmWidth                                     As Long
    bmHeight                                    As Long
    bmWidthBytes                                As Long
    bmPlanes                                    As Integer
    bmBitsPixel                                 As Integer
    bmBits                                      As Long
End Type

Private Type SCROLLINFO
    cbSize                                      As Long
    fMask                                       As Long
    nMin                                        As Long
    nMax                                        As Long
    nPage                                       As Long
    nPos                                        As Long
    nTrackPos                                   As Long
End Type

Private Type SCROLLBARINFO
    cbSize                                      As Long
    rcScrollBar                                 As RECT
    dxyLineButton                               As Long
    xyThumbTop                                  As Long
    xyThumbBottom                               As Long
    reserved                                    As Long
    rgstate(0 To 5)                             As Long
End Type


Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function CreateWindowEx Lib "USER32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
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

Private Declare Function SystemParametersInfo Lib "USER32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByRef lpvParam As Long, _
                                                                                          ByVal fuWinIni As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long

Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function CopyRect Lib "USER32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function OffsetRect Lib "USER32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long

Private Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetScrollBarInfo Lib "USER32" (ByVal hwnd As Long, _
                                                        ByVal idObject As Long, _
                                                        psbi As SCROLLBARINFO) As Long

Private Declare Function GetScrollInfo Lib "USER32" (ByVal hwnd As Long, _
                                                     ByVal n As Long, _
                                                     lpScrollInfo As SCROLLINFO) As Long

Private Declare Function DestroyWindow Lib "USER32" (ByVal hwnd As Long) As Long

Private Declare Sub SetWindowPos Lib "USER32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function GetParent Lib "USER32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "USER32" () As Long

Private Declare Function ScreenToClient Lib "USER32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal Length As Long)

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetWindowDC Lib "USER32" (ByVal hwnd As Long) As Long

Private Declare Function ShowWindow Lib "USER32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Function IsWindowVisible Lib "USER32" (ByVal hwnd As Long) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long


Private m_bInitialised                          As Boolean
Private m_bHasVertical                          As Boolean
Private m_bHasHorizontal                        As Boolean
Private m_bHasSizer                             As Boolean
Private m_bSkinScrollBar                        As Boolean
Private m_bUseScrollBarTheme                    As Boolean
Private m_bScrollBarActive                      As Boolean
Private m_bVisible                              As Boolean
Private m_bSelect                               As Boolean
Private m_lParentHnd                            As Long
Private m_lScrollLuminence                      As Long
Private m_lThemeColor                           As Long
Private m_lSmallChangeHorz                      As Long
Private m_lSmallChangeVert                      As Long
Private m_lWheelScrollLines                     As Long
Private m_lScrollSkinStyle                      As Long
Private m_lHBtnWidth                            As Long
Private m_lVBtnWidth                            As Long
Private m_hWndParent                            As Long
Private m_lVXScroll                             As Long
Private m_lVYScroll                             As Long
Private m_lHXScroll                             As Long
Private m_lHYScroll                             As Long
Private m_lVYThumb                              As Long
Private m_lHXThumb                              As Long
Private m_lCtrlHnd                              As Long
Private m_lVScrollHnd                           As Long
Private m_lHScrollHnd                           As Long
Private m_lSizerHnd                             As Long
Private m_lLastVCode                            As Long
Private m_lLastHCode                            As Long
Private m_sngLuminence                          As Single
Private m_eOrientation                          As EFSOrientationConstants
Private m_eStyle                                As EFSStyleConstants
Private m_tVertRect                             As RECT
Private m_tVScrInfo                             As SCROLLINFO
Private m_tVScrBar                              As SCROLLBARINFO
Private m_pVTrack                               As StdPicture
Private m_pVThumb                               As StdPicture
Private m_pVBtDwn                               As StdPicture
Private m_pVBtUp                                As StdPicture
Private m_cVTrackDc                             As clsStoreDc
Private m_cVThumbDc                             As clsStoreDc
Private m_cBtDwnDc                              As clsStoreDc
Private m_cBtUpDc                               As clsStoreDc
Private m_tHorzRect                             As RECT
Private m_tHScrInfo                             As SCROLLINFO
Private m_tHScrBar                              As SCROLLBARINFO
Private m_pHTrack                               As StdPicture
Private m_pHThumb                               As StdPicture
Private m_pHBtLft                               As StdPicture
Private m_pHBtRgt                               As StdPicture
Private m_pSizer                                As StdPicture
Private m_cHTrackDc                             As clsStoreDc
Private m_cHThumbDc                             As clsStoreDc
Private m_cHBtLftDc                             As clsStoreDc
Private m_cHBtRgtDc                             As clsStoreDc
Private m_cSzrDc                                As clsStoreDc
Private m_cRender                               As clsRender
Private m_GXScrollbar                           As GXMSubclass

Private Sub Class_Initialize()

    m_lSmallChangeHorz = 1
    m_lSmallChangeVert = 1
    m_eStyle = efsRegular
    m_eOrientation = efsoBoth

End Sub



Friend Property Get ScrollBarActive() As Boolean
    ScrollBarActive = m_bScrollBarActive
End Property

Friend Property Let ScrollBarActive(ByVal PropVal As Boolean)
    m_bScrollBarActive = PropVal
End Property

Private Property Get IHBtLft() As StdPicture
'/* scrollbar horizontal button left
    Set IHBtLft = m_pHBtLft
End Property

Private Property Set IHBtLft(PropVal As StdPicture)
    Set m_pHBtLft = PropVal
End Property

Private Property Get IHBtRgt() As StdPicture
'/* scrollbar horizontal button right
    Set IHBtRgt = m_pHBtRgt
End Property

Private Property Set IHBtRgt(PropVal As StdPicture)
    Set m_pHBtRgt = PropVal
End Property

Private Property Get IHThumb() As StdPicture
'/* scrollbar horizontal thumb
    Set IHThumb = m_pHThumb
End Property

Private Property Set IHThumb(PropVal As StdPicture)
    Set m_pHThumb = PropVal
End Property

Private Property Get IHTrack() As StdPicture
'/* scrollbar horizontal track
    Set IHTrack = m_pHTrack
End Property

Private Property Set IHTrack(PropVal As StdPicture)
    Set m_pHTrack = PropVal
End Property

Private Property Get ISizer() As StdPicture
'/* scrollbar sizer
    Set ISizer = m_pSizer
End Property

Private Property Set ISizer(PropVal As StdPicture)
    Set m_pSizer = PropVal
End Property

Private Property Get IVBtDwn() As StdPicture
'/* scrollbar vertical button down
    Set IVBtDwn = m_pVBtDwn
End Property

Private Property Set IVBtDwn(PropVal As StdPicture)
    Set m_pVBtDwn = PropVal
End Property

Private Property Get IVBtUp() As StdPicture
'/* scrollbar vertical button up
    Set IVBtUp = m_pVBtUp
End Property

Private Property Set IVBtUp(PropVal As StdPicture)
    Set m_pVBtUp = PropVal
End Property

Private Property Get IVThumb() As StdPicture
'/* scrollbar vertical thumb
    Set IVThumb = m_pVThumb
End Property

Private Property Set IVThumb(PropVal As StdPicture)
    Set m_pVThumb = PropVal
End Property

Private Property Get IVTrack() As StdPicture
'/* scrollbar vertical track
    Set IVTrack = m_pVTrack
End Property

Private Property Set IVTrack(PropVal As StdPicture)
    Set m_pVTrack = PropVal
End Property

Friend Property Get ScrollBarSkinStyle() As Long
    ScrollBarSkinStyle = m_lScrollSkinStyle
End Property

Friend Property Let ScrollBarSkinStyle(ByVal PropVal As Long)
    m_lScrollSkinStyle = PropVal
End Property

Friend Property Get ScrollLuminence() As Long
    ScrollLuminence = m_lScrollLuminence
End Property

Friend Property Let ScrollLuminence(PropVal As Long)
    m_lScrollLuminence = PropVal
End Property

Friend Property Get ScrollThemeColor() As Long
    ScrollThemeColor = m_lThemeColor
End Property

Friend Property Let ScrollThemeColor(ByVal PropVal As Long)
    m_lThemeColor = PropVal
End Property

Friend Property Get SkinScrollBar() As Boolean
    SkinScrollBar = m_bSkinScrollBar
End Property

Friend Property Let SkinScrollBar(ByVal PropVal As Boolean)
    m_bSkinScrollBar = PropVal
End Property

Friend Property Get UseScrollBarTheme() As Boolean
    UseScrollBarTheme = m_bUseScrollBarTheme
End Property

Friend Property Let UseScrollBarTheme(ByVal PropVal As Boolean)
    m_bUseScrollBarTheme = PropVal
End Property


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'> Functions
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Friend Function ResetScrollBarSkin()

    m_bScrollBarActive = False
    CleanupSkin
    
End Function

Private Sub Build_Bars()

    Store_Horizontal m_lCtrlHnd
    Create_Scrollbar SB_HORZ
    Init_HImages

    Store_Vertical m_lCtrlHnd
    Create_Scrollbar SB_VERT
    Init_VImages

    Create_Scrollbar SB_SZR
    Init_SZImage

End Sub

Private Function BuildImageList() As Boolean

On Error GoTo Handler

    Select Case m_lScrollSkinStyle
    '*/ classic
    Case 0
        '* / horz
        Set IHBtLft = LoadResPicture("CLASSIC-HZSCLFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("CLASSIC-HZSCRT", vbResBitmap)
        Set IHThumb = LoadResPicture("CLASSIC-HZSCTB", vbResBitmap)
        Set IHTrack = LoadResPicture("CLASSIC-HZSCTY", vbResBitmap)
        '/* vert
        Set IVBtDwn = LoadResPicture("CLASSIC-VTSCDN", vbResBitmap)
        Set IVBtUp = LoadResPicture("CLASSIC-VTBTUP", vbResBitmap)
        Set IVThumb = LoadResPicture("CLASSIC-VTBTTB", vbResBitmap)
        Set IVTrack = LoadResPicture("CLASSIC-VTSCTY", vbResBitmap)
        '/* sizer
        Set ISizer = LoadResPicture("CLASSIC-HZSCSZ", vbResBitmap)
    '/* eclipse
    Case 1
        Set IHBtLft = LoadResPicture("ECLIPSE-HZSCLFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("ECLIPSE-HZSCRT", vbResBitmap)
        Set IHThumb = LoadResPicture("ECLIPSE-HZSCTB", vbResBitmap)
        Set IHTrack = LoadResPicture("ECLIPSE-HZSCTY", vbResBitmap)
        Set IVBtDwn = LoadResPicture("ECLIPSE-VTSCDN", vbResBitmap)
        Set IVBtUp = LoadResPicture("ECLIPSE-VTBTUP", vbResBitmap)
        Set IVThumb = LoadResPicture("ECLIPSE-VTBTTB", vbResBitmap)
        Set IVTrack = LoadResPicture("ECLIPSE-VTSCTY", vbResBitmap)
        Set ISizer = LoadResPicture("ECLIPSE-HZSCSZ", vbResBitmap)
    
    '/* lime
    Case 2
        Set IHBtLft = LoadResPicture("LIME-HZSCLFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("LIME-HZSCRT", vbResBitmap)
        Set IHThumb = LoadResPicture("LIME-HZSCTB", vbResBitmap)
        Set IHTrack = LoadResPicture("LIME-HZSCTY", vbResBitmap)
        Set IVBtDwn = LoadResPicture("LIME-VTSCDN", vbResBitmap)
        Set IVBtUp = LoadResPicture("LIME-VTBTUP", vbResBitmap)
        Set IVThumb = LoadResPicture("LIME-VTBTTB", vbResBitmap)
        Set IVTrack = LoadResPicture("LIME-VTSCTY", vbResBitmap)
        Set ISizer = LoadResPicture("LIME-HZSCSZ", vbResBitmap)
        
    '/* metallic
    Case 3
        Set IHBtLft = LoadResPicture("METALLIC-HZSCLFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("METALLIC-HZSCRT", vbResBitmap)
        Set IHThumb = LoadResPicture("METALLIC-HZSCTB", vbResBitmap)
        Set IHTrack = LoadResPicture("METALLIC-HZSCTY", vbResBitmap)
        Set IVBtDwn = LoadResPicture("METALLIC-VTSCDN", vbResBitmap)
        Set IVBtUp = LoadResPicture("METALLIC-VTBTUP", vbResBitmap)
        Set IVThumb = LoadResPicture("METALLIC-VTBTTB", vbResBitmap)
        Set IVTrack = LoadResPicture("METALLIC-VTSCTY", vbResBitmap)
        Set ISizer = LoadResPicture("METALLIC-HZSCSZ", vbResBitmap)
        
    '/* gloss
    Case 4
        Set IHBtLft = LoadResPicture("GLOSS-HZSCLFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("GLOSS-HZSCRT", vbResBitmap)
        Set IHThumb = LoadResPicture("GLOSS-HZSCTB", vbResBitmap)
        Set IHTrack = LoadResPicture("GLOSS-HZSCTY", vbResBitmap)
        Set IVBtDwn = LoadResPicture("GLOSS-VTSCDN", vbResBitmap)
        Set IVBtUp = LoadResPicture("GLOSS-VTBTUP", vbResBitmap)
        Set IVThumb = LoadResPicture("GLOSS-VTBTTB", vbResBitmap)
        Set IVTrack = LoadResPicture("GLOSS-VTSCTY", vbResBitmap)
        Set ISizer = LoadResPicture("GLOSS-HZSCSZ", vbResBitmap)
    End Select
    
    If m_bUseScrollBarTheme Then
        ScrollThemeSettings
    End If
    BuildImageList = True

Handler:
    On Error GoTo 0
    
End Function

Private Sub ScrollThemeSettings()

    Select Case m_lScrollLuminence
    Case 0
        m_sngLuminence = 0.2
    Case 1
        m_sngLuminence = 0.5
    Case 2
        m_sngLuminence = 1
    End Select
    
End Sub

Friend Sub LoadSkin(ByVal lCtrlHnd As Long, ByVal lParHnd As Long)

    m_lCtrlHnd = lCtrlHnd
    m_lParentHnd = lParHnd
    If m_bSkinScrollBar Then
        If BuildImageList Then
            SkinAttach
            ScrollbarAttach
        End If
    End If
    
End Sub

Private Function Create_Scrollbar(eStyle As SCR_STYLE) As Long

Dim bOnDesktop As Boolean

    m_hWndParent = GetParent(m_lParentHnd)
    bOnDesktop = (m_hWndParent = GetDesktopWindow())

    Select Case eStyle
    Case SB_HORZ
        If m_lHScrollHnd = 0 Then
            m_lHScrollHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
            Mask_Show m_lHScrollHnd, False
        End If
    Case SB_VERT
        If m_lVScrollHnd = 0 Then
            m_lVScrollHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
            Mask_Show m_lVScrollHnd, False
        End If
    Case SB_SZR
        If m_lSizerHnd = 0 Then
            m_lSizerHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
            Mask_Show m_lSizerHnd, False
        End If
    End Select

End Function

Private Function Destroy_Scrollbar(eStyle As SCR_STYLE) As Long

    Select Case eStyle
    Case SB_HORZ
        If Not m_lHScrollHnd = 0 Then
            DestroyWindow m_lHScrollHnd
            m_lHScrollHnd = 0
        End If
    Case SB_VERT
        If Not m_lVScrollHnd = 0 Then
            DestroyWindow m_lVScrollHnd
            m_lVScrollHnd = 0
        End If
    Case SB_BOTH
        If Not m_lHScrollHnd = 0 Then
            DestroyWindow m_lHScrollHnd
            m_lHScrollHnd = 0
        End If
        If Not m_lVScrollHnd = 0 Then
            DestroyWindow m_lVScrollHnd
            m_lHScrollHnd = 0
        End If
        If Not m_lSizerHnd = 0 Then
            DestroyWindow m_lSizerHnd
            m_lSizerHnd = 0
        End If
    Case SB_SZR
        If Not m_lSizerHnd = 0 Then
            DestroyWindow m_lSizerHnd
            m_lSizerHnd = 0
        End If
    End Select
    
End Function

Private Sub Init_HImages()
'/* init hz image classes

Dim tBmp            As BITMAP

    '/* image size
    GetObject IHBtLft.Handle, Len(tBmp), tBmp
    With tBmp
        m_lHBtnWidth = (.bmWidth / 2)
    End With

    '/* create the image classes
    Set m_cHTrackDc = New clsStoreDc
    With m_cHTrackDc
        .CreateFromPicture IHTrack
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cHThumbDc = New clsStoreDc
    With m_cHThumbDc
        .CreateFromPicture IHThumb
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cHBtLftDc = New clsStoreDc
    With m_cHBtLftDc
        .CreateFromPicture IHBtLft
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cHBtRgtDc = New clsStoreDc
    With m_cHBtRgtDc
        .CreateFromPicture IHBtRgt
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
End Sub

Private Sub Init_SZImage()
'/* sizer
    
    Set m_cSzrDc = New clsStoreDc
    With m_cSzrDc
        .CreateFromPicture ISizer
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
End Sub

Private Sub Init_VImages()
'/* init vt image classes

Dim tBmp            As BITMAP

    GetObject IVBtDwn.Handle, Len(tBmp), tBmp
    With tBmp
        m_lVBtnWidth = (.bmWidth / 2)
    End With

    '/* create the image classes
    Set m_cVTrackDc = New clsStoreDc
    With m_cVTrackDc
        .CreateFromPicture IVTrack
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cVThumbDc = New clsStoreDc
    With m_cVThumbDc
        .CreateFromPicture IVThumb
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cBtDwnDc = New clsStoreDc
    With m_cBtDwnDc
        .CreateFromPicture IVBtDwn
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cBtUpDc = New clsStoreDc
    With m_cBtUpDc
        .CreateFromPicture IVBtUp
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
End Sub

Private Sub Mask_Show(ByVal lHwnd As Long, _
                      ByVal bVisible As Boolean)
    
    If lHwnd = 0 Then Exit Sub
    If bVisible Then
        ShowWindow lHwnd, SW_NORMAL
    Else
        ShowWindow lHwnd, SW_HIDE
    End If
    
End Sub

Private Sub Reset_SCPos()

Dim lHscr       As Long
Dim uPt         As POINTAPI
Dim tOnRt       As RECT

    GetWindowRect m_lCtrlHnd, tOnRt
    CopyMemory uPt, tOnRt, Len(uPt)
    ScreenToClient m_hWndParent, uPt
    BarVisible m_lCtrlHnd
    
    With tOnRt
        OffsetRect tOnRt, -.left, -.top
    End With

    With tOnRt
        If Not (m_lVScrollHnd = 0) And HasVertical Then
            SetWindowPos m_lVScrollHnd, 0, uPt.X + (.right - m_lVXScroll), uPt.Y, m_lVXScroll, (.bottom - .top), SWP_SHOWWINDOW
        End If
        If Not (m_lHScrollHnd = 0) And HasHorizontal Then
            SetWindowPos m_lHScrollHnd, 0, uPt.X, (uPt.Y + .bottom) - m_lHYScroll, (.right - lHscr), m_lVYScroll, SWP_SHOWWINDOW
        End If
        If Not (m_lSizerHnd = 0) And HasSizer Then
            SetWindowPos m_lSizerHnd, 0, uPt.X + (.right - m_lVXScroll), (uPt.Y + .bottom) - m_lHYScroll, m_lVXScroll, m_lVYScroll, SWP_SHOWWINDOW
        End If
    End With
    
End Sub

Public Sub ScrollbarAttach()

    If Not m_lCtrlHnd = 0 Then
        '/* store system sizes
        Scrollbar_Metrics
        Build_Bars
        BarVisible m_lCtrlHnd
        '/* load image classes
        Create_Scrollbar SB_VERT
        Init_VImages
        Create_Scrollbar SB_HORZ
        Init_HImages
        Create_Scrollbar SB_SZR
        Init_SZImage
        m_bScrollBarActive = True
        Set m_cRender = New clsRender
    End If
    
End Sub

Private Sub Scrollbar_Metrics()
'/* store scrollbar metrics

    m_lVXScroll = GetSystemMetrics(SM_CXVSCROLL)
    m_lVYScroll = GetSystemMetrics(SM_CYVSCROLL)
    m_lHXScroll = GetSystemMetrics(SM_CXHSCROLL)
    m_lHYScroll = GetSystemMetrics(SM_CYHSCROLL)
    m_lVYThumb = GetSystemMetrics(SM_CYVTHUMB)
    m_lHXThumb = GetSystemMetrics(SM_CXHTHUMB)
    
End Sub

Private Function Store_Horizontal(ByVal lHwnd As Long) As Boolean

Dim lState As Long

On Error GoTo Handler

    With m_tHScrInfo
        .cbSize = Len(m_tHScrInfo)
        .fMask = SIF_ALL
    End With
    '/* copy vertical scrollbar structure
    lState = GetScrollInfo(lHwnd, SB_HORZ, m_tHScrInfo)
    '/* no scrollbar
    If lState = 0 Then
        Exit Function
    End If
    '/* get bar params
    m_tHScrBar.cbSize = Len(m_tHScrBar)
    GetScrollBarInfo lHwnd, OBJID_HSCROLL, m_tHScrBar
    '/* copy size to working rect
    CopyRect m_tHorzRect, m_tHScrBar.rcScrollBar
    Store_Horizontal = True
    
Handler:
    On Error GoTo 0

End Function

Private Function Store_Vertical(ByVal lHwnd As Long) As Boolean

Dim lState As Long

On Error GoTo Handler

    With m_tVScrInfo
        .cbSize = Len(m_tVScrInfo)
        .fMask = SIF_ALL
    End With
    '/* copy vertical scrollbar structure
    lState = GetScrollInfo(lHwnd, SB_VERT, m_tVScrInfo)
    '/* no scrollbar
    If lState = 0 Then
        Exit Function
    End If
    '/* get bar params
    m_tVScrBar.cbSize = Len(m_tVScrBar)
    GetScrollBarInfo lHwnd, OBJID_VSCROLL, m_tVScrBar
    '/* copy size to working rect
    CopyRect m_tVertRect, m_tVScrBar.rcScrollBar
    '/* apply client offsets
    Store_Vertical = True

Handler:
    On Error GoTo 0
    
End Function

Private Sub HBar_Paint(ByVal lOffset As Long)

'/* paint horz scrollbar
Dim lHDC        As Long
Dim lDiff       As Long
Dim lHOffset    As Long
Dim lLOffset    As Long
Dim lXOffset    As Long
Dim lSclHdc     As Long
Dim lSzrDc      As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long
Dim tRect       As RECT
Dim tSzr        As RECT
Dim tPar        As RECT

On Error GoTo Handler

    '/* validate position
    GetWindowRect m_lHScrollHnd, tRect
    GetWindowRect m_lCtrlHnd, tPar
    If tRect.top < tPar.top Then
        Exit Sub
    End If
    
    '/* thumb size
    With m_tHScrBar
        lDiff = .xyThumbBottom - .xyThumbTop
    End With
    '/* backbuffer the image
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lHDC = CreateCompatibleDC(lTmpDc)
    With m_tHorzRect
        OffsetRect m_tHorzRect, -.left, -.top
        lBmp = CreateCompatibleBitmap(lTmpDc, .right, .bottom)
    End With
    lBmpOld = SelectObject(lHDC, lBmp)
    lSclHdc = GetDC(m_lHScrollHnd)

    '/* apply offset
    If lOffset = 0 Then
        lLOffset = 1
    ElseIf lOffset = 1 Then
        lHOffset = 1
    ElseIf lOffset = 5 Then
        lXOffset = 1
    ElseIf lOffset = -1 Then
        If m_lLastHCode = 0 Then
            lLOffset = 1
        ElseIf m_lLastHCode = 1 Then
            lHOffset = 1
        ElseIf m_lLastHCode = 5 Then
            lXOffset = 1
        End If
    End If
    m_lLastHCode = lOffset
    
    With m_tHorzRect
        '/* build to temp dc
        m_cRender.Stretch lHDC, 0, 0, .right, m_lHXScroll, m_cHTrackDc.hdc, 0, 0, m_cHTrackDc.Width, m_cHTrackDc.Height, SRCCOPY
        m_cRender.Stretch lHDC, 0, 0, m_lHXScroll, m_lHYScroll, m_cHBtRgtDc.hdc, m_lHBtnWidth * lLOffset, 0, m_cHBtRgtDc.Width / 2, m_cHBtRgtDc.Height, SRCCOPY
        m_cRender.Stretch lHDC, .right - (m_lHYScroll + 1), .top, m_lHXScroll, m_lHYScroll, m_cHBtLftDc.hdc, m_lHBtnWidth * lHOffset, 0, m_cHBtLftDc.Width / 2, m_cHBtLftDc.Height, SRCCOPY
        With m_cHThumbDc
            If lDiff > .Width Then
                m_cRender.Stretch lHDC, m_tHScrBar.xyThumbTop, 0, 4, m_lHYScroll + 1, .hdc, ((.Width / 2) * lXOffset), 0, 4, .Height, SRCCOPY
                m_cRender.Stretch lHDC, m_tHScrBar.xyThumbTop + 4, 0, lDiff - 8, m_lHYScroll + 1, .hdc, ((.Width / 2) * lXOffset) + 4, 0, (.Width / 2) - 8, .Height, SRCCOPY
                m_cRender.Stretch lHDC, m_tHScrBar.xyThumbBottom - 4, 0, 4, m_lHYScroll + 1, .hdc, ((.Width / 2) * (lXOffset + 1)) - 4, 0, 4, .Height, SRCCOPY
            Else
                m_cRender.Stretch lHDC, m_tHScrBar.xyThumbTop, 0, lDiff, m_lHYScroll, .hdc, (.Width * lXOffset), 0, (.Width / 2), .Height, SRCCOPY
            End If
        End With
        '/* copy to dest
        m_cRender.Blit lSclHdc, 0, 0, .right, .bottom, lHDC, 0, 0, SRCCOPY
    End With
    
    '/* sizer
    If m_bHasSizer Then
        GetWindowRect m_lSizerHnd, tRect
        GetWindowRect m_lParentHnd, tSzr
        If tRect.left > tSzr.left Then
            lSzrDc = GetDC(m_lSizerHnd)
            m_cRender.Stretch lSzrDc, 0, 0, m_lHYScroll, m_lVXScroll, m_cSzrDc.hdc, 0, 0, m_cSzrDc.Width, m_cSzrDc.Height, SRCCOPY
            ReleaseDC m_lSizerHnd, lSzrDc
        End If
    End If
    
    '/* cleanup
    SelectObject lHDC, lBmpOld
    DeleteObject lBmp
    DeleteDC lHDC
    ReleaseDC lHwnd, lTmpDc
    ReleaseDC m_lVScrollHnd, lSclHdc
    
Handler:
    On Error GoTo 0

End Sub

Private Sub VBar_Paint(ByVal lOffset As Long)

'/* paint vert scrollbar
Dim lHDC        As Long
Dim lDiff       As Long
Dim lHOffset    As Long
Dim lLOffset    As Long
Dim lXOffset    As Long
Dim lSclHdc     As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim lHwnd       As Long
Dim tRect       As RECT
Dim tPar        As RECT

On Error GoTo Handler

    '/* validate position
    GetWindowRect m_lVScrollHnd, tRect
    GetWindowRect m_lCtrlHnd, tPar
    If tRect.left < tPar.left Then
        Exit Sub
    End If
    '/* thumb size
    With m_tVScrBar
        lDiff = .xyThumbBottom - .xyThumbTop
    End With
    '/* backbuffer the image
    lHwnd = GetDesktopWindow
    lTmpDc = GetWindowDC(lHwnd)
    lHDC = CreateCompatibleDC(lTmpDc)
    With m_tVertRect
        OffsetRect m_tVertRect, -.left, -.top
        lBmp = CreateCompatibleBitmap(lTmpDc, .right, .bottom)
    End With
    lBmpOld = SelectObject(lHDC, lBmp)
    lSclHdc = GetDC(m_lVScrollHnd)
    
    '/* apply offset
    If lOffset = 0 Then
        lLOffset = 1
    ElseIf lOffset = 1 Then
        lHOffset = 1
    ElseIf lOffset = 5 Then
        lXOffset = 1
    ElseIf lOffset = -1 Then
        If m_lLastVCode = 0 Then
            lLOffset = 1
        ElseIf m_lLastVCode = 1 Then
            lHOffset = 1
        ElseIf m_lLastVCode = 5 Then
            lXOffset = 1
        End If
    End If
    m_lLastVCode = lOffset
    
    With m_tVertRect
        '/* build to temp dc
        m_cRender.Stretch lHDC, 0, 0, m_lVXScroll, .bottom, m_cVTrackDc.hdc, 0, 0, m_cVTrackDc.Width, m_cVTrackDc.Height, SRCCOPY
        m_cRender.Stretch lHDC, 0, 0, m_lVXScroll, m_lVYScroll, m_cBtUpDc.hdc, m_lVBtnWidth * lLOffset, 0, m_cBtUpDc.Width / 2, m_cBtUpDc.Height, SRCCOPY
        m_cRender.Stretch lHDC, 0, .bottom - m_lVYScroll, m_lVXScroll, m_lVYScroll, m_cBtDwnDc.hdc, m_lVBtnWidth * lHOffset, 0, m_cBtUpDc.Width / 2, m_cBtUpDc.Height, SRCCOPY
        With m_cVThumbDc
            If lDiff > .Height Then
                '/* top :2 state thumb - (width / 2) x offset, offset * lxoffset
                m_cRender.Stretch lHDC, 0, m_tVScrBar.xyThumbBottom - lDiff, m_lVXScroll + 1, 4, .hdc, (.Width / 2) * lXOffset, 0, (.Width / 2), 4, SRCCOPY
                '/* center
                m_cRender.Stretch lHDC, 0, (m_tVScrBar.xyThumbBottom - (lDiff - 4)), m_lVXScroll + 1, lDiff - 8, .hdc, (.Width / 2) * lXOffset, 4, (.Width / 2), .Height - 8, SRCCOPY
                '/* bottom
                m_cRender.Stretch lHDC, 0, m_tVScrBar.xyThumbBottom - 4, m_lVXScroll + 1, 4, .hdc, (.Width / 2) * lXOffset, .Height - 4, (.Width / 2), 4, SRCCOPY
            Else
                m_cRender.Stretch lHDC, 0, m_tVScrBar.xyThumbBottom - lDiff, m_lVXScroll, lDiff, .hdc, (.Width / 2) * lXOffset, 0, (.Width / 2), .Height, SRCCOPY
            End If
        End With
        '/* copy to dest
        m_cRender.Blit lSclHdc, 0, 0, .right, .bottom, lHDC, 0, 0, SRCCOPY
    End With
    
    '/* cleanup
    SelectObject lHDC, lBmpOld
    DeleteObject lBmp
    DeleteDC lHDC
    ReleaseDC lHwnd, lTmpDc
    ReleaseDC m_lVScrollHnd, lSclHdc

Handler:
    On Error GoTo 0
    
End Sub

Public Sub Resize()
    Reset_SCPos
End Sub

Friend Function ScrollVertical(ByVal bDown As Boolean)

    If bDown Then
        SendMessageLong m_lCtrlHnd, WM_VSCROLL, SB_LINEDOWN, 0
    Else
        SendMessageLong m_lCtrlHnd, WM_VSCROLL, SB_LINEUP, 0
    End If
    
End Function

Friend Function ScrollHorizontal(ByVal bRight As Boolean)

    If bRight Then
        SendMessageLong m_lCtrlHnd, WM_HSCROLL, SB_LINERIGHT, 0
    Else
        SendMessageLong m_lCtrlHnd, WM_HSCROLL, SB_LINELEFT, 0
    End If
    
End Function

Friend Function HasHorizontal() As Boolean

Dim lStyle   As Long

    lStyle = GetWindowLong(m_lCtrlHnd, GWL_STYLE)
    HasHorizontal = (lStyle And WS_HSCROLL) <> 0

End Function

Friend Function HasVertical() As Boolean

Dim lStyle  As Long

    lStyle = GetWindowLong(m_lCtrlHnd, GWL_STYLE)
    HasVertical = (lStyle And WS_VSCROLL) <> 0

End Function

Private Function HasSizer() As Boolean

Dim bHorizontal As Boolean
Dim bVertical   As Boolean
Dim lStyle      As Long

    lStyle = GetWindowLong(m_lCtrlHnd, GWL_STYLE)
    bHorizontal = (lStyle And WS_HSCROLL) <> 0
    bVertical = (lStyle And WS_VSCROLL) <> 0
    HasSizer = (bHorizontal And bVertical)
    
End Function

Private Sub BarVisible(ByVal lHwnd As Long)
'/* bar status

Dim lStyle   As Long

    lStyle = GetWindowLong(lHwnd, GWL_STYLE)
    m_bHasHorizontal = (lStyle And WS_HSCROLL) <> 0
    m_bHasVertical = (lStyle And WS_VSCROLL) <> 0
    m_bHasSizer = (m_bHasHorizontal And m_bHasVertical)
    ScrollbarStatus

End Sub

Friend Function Refresh()

    m_bSelect = False
    SendMessageLong m_lCtrlHnd, WM_PAINT, 0&, 0&
    
End Function

Private Sub ScrollbarStatus()

    If IsWindowVisible(m_lCtrlHnd) = 0 Then
        Mask_Show m_lHScrollHnd, False
        Mask_Show m_lVScrollHnd, False
        Mask_Show m_lSizerHnd, False
        Exit Sub
    End If
    
    If HasHorizontal Then
        Mask_Show m_lHScrollHnd, True
    Else
        Mask_Show m_lHScrollHnd, False
    End If
    
    If HasVertical Then
        Mask_Show m_lVScrollHnd, True
    Else
        Mask_Show m_lVScrollHnd, False
    End If
    
    If HasSizer Then
        Mask_Show m_lSizerHnd, True
    Else
        Mask_Show m_lSizerHnd, False
    End If

End Sub

Public Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Public Property Let Visible(PropVal As Boolean)

    If PropVal Then
        If Not m_bVisible Then
            ScrollbarStatus
            m_bScrollBarActive = True
            SendMessageLong m_lCtrlHnd, WM_PAINT, 0&, 0&
        End If
    Else
        Mask_Show m_lHScrollHnd, False
        Mask_Show m_lVScrollHnd, False
        Mask_Show m_lSizerHnd, False
        m_bScrollBarActive = False
    End If
    m_bVisible = PropVal

End Property

Private Sub SkinAttach()

    If m_GXScrollbar Is Nothing Then
        Set m_GXScrollbar = New GXMSubclass
        With m_GXScrollbar
            .Subclass m_lCtrlHnd, Me
            .AddMessage m_lCtrlHnd, WM_ERASEBKGND, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_PAINT, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_STYLECHANGED, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_VSCROLL, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_HSCROLL, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_SIZE, MSG_BEFORE
        End With
    End If
    
End Sub

Private Sub SkinDetach()

    If Not m_GXScrollbar Is Nothing Then
        With m_GXScrollbar
            .DeleteMessage m_lCtrlHnd, WM_ERASEBKGND, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_PAINT, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_STYLECHANGED, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_VSCROLL, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_HSCROLL, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_SIZE, MSG_BEFORE
            .UnSubclass m_lCtrlHnd
        End With
        Set m_GXScrollbar = Nothing
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

Dim lScrollCode     As Long
Dim wStyle          As Long
Static lStyle       As Long

    If IsWindowVisible(m_lCtrlHnd) = 0 Then
        BarVisible m_lCtrlHnd
        Exit Sub
    Else
        BarVisible m_lCtrlHnd
    End If
    
    If Not m_bScrollBarActive Then
        Exit Sub
    End If

    Select Case uMsg
    '/* client paint
    Case WM_PAINT
        If Not m_bSelect Then
            If HasHorizontal Then
                Store_Horizontal lHwnd
                HBar_Paint -1
            End If
            If HasVertical Then
                Store_Vertical lHwnd
                VBar_Paint -1
            End If
        Else
            m_bSelect = False
        End If
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        bHandled = True
    
    '/* horizontal scroll
    Case WM_HSCROLL
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Store_Horizontal lHwnd
        lScrollCode = (wParam And &HFFFF&)
        HBar_Paint lScrollCode
        m_bSelect = True
        bHandled = True
    
    '/* vertical Scroll
    Case WM_VSCROLL
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Store_Vertical lHwnd
        lScrollCode = (wParam And &HFFFF&)
        VBar_Paint lScrollCode
        m_bSelect = True
        bHandled = True

    Case WM_WINDOWPOSCHANGED, WM_STYLECHANGED
        ScrollbarStatus
    
    Case WM_SIZE
        wStyle = GetWindowLong(lHwnd, GWL_STYLE)
        If Not wStyle = lStyle Then
            Reset_SCPos
            lStyle = wStyle
        End If

    Case WM_ERASEBKGND
        If HasHorizontal Then
            Store_Horizontal lHwnd
            HBar_Paint -1
        End If
        If HasVertical Then
            Store_Vertical lHwnd
            VBar_Paint -1
        End If
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        bHandled = True
    End Select

End Sub

Private Sub CleanUp()

    If Not m_lCtrlHnd = 0 Then
        If m_bSkinScrollBar Then
            CleanupSkin
        End If
        m_lCtrlHnd = 0
        m_bInitialised = False
    End If
    
End Sub

Private Sub CleanupSkin()

    If Not m_lCtrlHnd = 0 Then
        SkinDetach
        Destroy_Scrollbar SB_BOTH
        If Not m_cVTrackDc Is Nothing Then Set m_cVTrackDc = Nothing
        If Not m_cVThumbDc Is Nothing Then Set m_cVThumbDc = Nothing
        If Not m_cBtDwnDc Is Nothing Then Set m_cBtDwnDc = Nothing
        If Not m_cBtUpDc Is Nothing Then Set m_cBtUpDc = Nothing
        If Not m_cHTrackDc Is Nothing Then Set m_cHTrackDc = Nothing
        If Not m_cHThumbDc Is Nothing Then Set m_cHThumbDc = Nothing
        If Not m_cHBtLftDc Is Nothing Then Set m_cHBtLftDc = Nothing
        If Not m_cHBtRgtDc Is Nothing Then Set m_cHBtRgtDc = Nothing
        If Not m_cSzrDc Is Nothing Then Set m_cSzrDc = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_pVTrack Is Nothing Then Set m_pVTrack = Nothing
        If Not m_pVThumb Is Nothing Then Set m_pVThumb = Nothing
        If Not m_pVBtDwn Is Nothing Then Set m_pVBtDwn = Nothing
        If Not m_pVBtUp Is Nothing Then Set m_pVBtUp = Nothing
        If Not m_pHTrack Is Nothing Then Set m_pHTrack = Nothing
        If Not m_pHThumb Is Nothing Then Set m_pHThumb = Nothing
        If Not m_pHBtLft Is Nothing Then Set m_pHBtLft = Nothing
        If Not m_pHBtRgt Is Nothing Then Set m_pHBtRgt = Nothing
        If Not m_pSizer Is Nothing Then Set m_pSizer = Nothing
        m_lHScrollHnd = 0
        m_lVScrollHnd = 0
        m_lSizerHnd = 0
    End If

End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub
