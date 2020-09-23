VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bCellCustom                   As Boolean
Private m_bXPColors                     As Boolean
Private m_lIndex                        As Long
Private m_lIcon                         As Long
Private m_lSubItemCount                 As Long
Private m_lBackColor                    As Long
Private m_lForeColor                    As Long
Private m_lSubItemIcon()                As Long
Private m_lCellBackColor()              As Long
Private m_lCellForeColor()              As Long
Private m_sKey                          As String
Private m_sText                         As String
Private m_sSubItemText()                As String
Private m_oRowFont                      As StdFont


Private Sub Class_Initialize()
    m_lSubItemCount = -1
    m_lBackColor = -1
    m_lForeColor = -1
End Sub

Private Function ArrayCheck(ByRef vArray As Variant) As Boolean
'/* validity test

On Error Resume Next

    '/* an array
    If Not IsArray(vArray) Then
        GoTo Handler
    '/* dimensioned
    ElseIf IsError(UBound(vArray)) Then
        GoTo Handler
    ElseIf UBound(vArray) = -1 Then
        GoTo Handler
    End If
    ArrayCheck = True

Handler:
    On Error GoTo 0

End Function

Public Sub Init()
    
    If Not m_lSubItemCount = -1 Then
        ReDim m_lSubItemIcon(1 To (m_lSubItemCount))
        ReDim m_sSubItemText(1 To (m_lSubItemCount))
    End If
    
End Sub

Private Sub ResizeArray()

    ReDim Preserve m_lSubItemIcon(1 To (m_lSubItemCount))
    ReDim Preserve m_sSubItemText(1 To (m_lSubItemCount))

End Sub

Public Sub Add(ByVal lIndex As Long, _
               ByVal sKey As String, _
               ByVal sText As String, _
               ByVal lIcon As Long, _
               ByVal lSmallIcon As Long)

    Index = lIndex
    Key = sKey
    Text = sText
    Icon = lIcon
    lSmallIcon = lSmallIcon

End Sub

Private Function CellTest(ByVal lCell As Long) As Boolean

    If Not ArrayCheck(m_lCellBackColor) Then
        CellCustom = False
    ElseIf Not ArrayCheck(m_lCellForeColor) Then
        CellCustom = False
    Else
        If UBound(m_lCellBackColor) < lCell Then
            CellCustom = False
        ElseIf UBound(m_lCellForeColor) < lCell Then
            CellCustom = False
        Else
            CellTest = True
        End If
    End If
    
End Function

Public Sub SubItem(ByVal lIndex As Long, _
                   ByVal sText As String)

On Error GoTo Handler

    If m_lSubItemCount = -1 Then
        m_lSubItemCount = lIndex
        Init
    ElseIf lIndex > m_lSubItemCount Then
        m_lSubItemCount = lIndex
        ResizeArray
    End If
    m_sSubItemText(lIndex) = sText

Handler:
    On Error GoTo 0

End Sub

Public Sub SubIcon(ByVal lIndex As Long, _
                   ByVal lIcon As Long)

On Error GoTo Handler

    If lIndex > m_lSubItemCount Then Exit Sub
    m_lSubItemIcon(lIndex) = lIcon

Handler:
    On Error GoTo 0

End Sub

Private Function XPShift(ByVal lColor As Long, _
                         Optional ByVal Base As Long = &HB0) As Long

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

Public Property Get BackColor() As Long
    BackColor = m_lBackColor
End Property

Public Property Let BackColor(ByVal PropVal As Long)

    m_lBackColor = PropVal
    If m_bXPColors Then
        m_lBackColor = XPShift(m_lBackColor)
    End If
    
End Property

Public Property Get CellBackColor(ByVal lCell As Long) As Long

On Error GoTo Handler

    CellBackColor = m_lCellBackColor(lCell)

Handler:
    On Error GoTo 0

End Property

Public Property Let CellBackColor(ByVal lCell As Long, _
                                  ByVal PropVal As Long)

On Error GoTo Handler

    ReDim Preserve m_lCellBackColor(0 To lCell)
    m_lCellBackColor(lCell) = PropVal
    If m_bXPColors Then
        m_lCellBackColor(lCell) = XPShift(PropVal)
    End If

Handler:
    On Error GoTo 0

End Property

Public Property Get CellCustom() As Boolean
    CellCustom = m_bCellCustom
End Property

Public Property Let CellCustom(ByVal PropVal As Boolean)
    m_bCellCustom = PropVal
End Property

Public Property Get CellForeColor(ByVal lCell As Long) As Long

On Error GoTo Handler

    CellForeColor = m_lCellForeColor(lCell)

Handler:
    On Error GoTo 0

End Property

Public Property Let CellForeColor(ByVal lCell As Long, _
                                  ByVal PropVal As Long)

On Error GoTo Handler

    ReDim Preserve m_lCellForeColor(0 To lCell)
    m_lCellForeColor(lCell) = PropVal

Handler:
    On Error GoTo 0

End Property

Public Property Get Font() As StdFont

On Error GoTo Handler

    Set Font = m_oRowFont

Handler:
    On Error GoTo 0

End Property

Public Property Set Font(ByVal PropVal As StdFont)

On Error GoTo Handler

    Set m_oRowFont = PropVal

Handler:
    On Error GoTo 0

End Property

Public Property Get ForeColor() As Long
    ForeColor = m_lForeColor
End Property

Public Property Let ForeColor(ByVal PropVal As Long)
    m_lForeColor = PropVal
End Property

Public Property Get Key() As String
    Key = m_sKey
End Property

Public Property Let Key(ByVal PropVal As String)
    m_sKey = PropVal
End Property

Public Property Get Icon() As Long
    Icon = m_lIcon
End Property

Public Property Let Icon(ByVal PropVal As Long)
    m_lIcon = PropVal
End Property

Public Property Get Index() As Long
    Index = m_lIndex
End Property

Public Property Let Index(ByVal PropVal As Long)
    m_lIndex = PropVal
End Property

Public Property Get SubItemCount() As Long
    SubItemCount = m_lSubItemCount
End Property

Public Property Let SubItemCount(ByVal PropVal As Long)
    m_lSubItemCount = PropVal
End Property

Public Property Get Text() As String
    Text = m_sText
End Property

Public Property Let Text(ByVal PropVal As String)
    m_sText = PropVal
End Property

Public Property Get SubItemIcon(ByVal lIndex As Long) As Long

On Error GoTo Handler

    SubItemIcon = m_lSubItemIcon(lIndex)

Handler:
    On Error GoTo 0

End Property

Public Property Let SubItemIcon(ByVal lIndex As Long, _
                                ByVal lIcon As Long)
    
On Error GoTo Handler

    If lIndex > m_lSubItemCount Then
        m_lSubItemCount = lIndex
    End If
    If UBound(m_lSubItemIcon) < lIndex Then
        ReDim Preserve m_lSubItemIcon(LBound(m_lSubItemIcon) To lIndex)
    ElseIf UBound(m_lSubItemIcon) = -1 Then
        ReDim m_lSubItemIcon(lIndex)
    End If
    m_lSubItemIcon(lIndex) = lIcon

Exit Property

Handler:
    On Error GoTo 0

End Property

Public Property Get SubItemText(ByVal lIndex As Long) As String

On Error GoTo Handler

    SubItemText = m_sSubItemText(lIndex)

Handler:
    On Error GoTo 0

End Property

Public Property Let SubItemText(ByVal lIndex As Long, _
                                ByVal sText As String)

On Error GoTo Handler

    If lIndex > m_lSubItemCount Then
        m_lSubItemCount = lIndex
    End If
    If UBound(m_sSubItemText) < lIndex Then
        ReDim Preserve m_sSubItemText(LBound(m_sSubItemText) To lIndex)
    ElseIf UBound(m_sSubItemText) = -1 Then
        ReDim m_sSubItemText(lIndex)
    End If
    m_sSubItemText(lIndex) = sText

Handler:
    On Error GoTo 0

End Property

Public Property Get XPColors() As Boolean
'/* [get] use xp colors
    XPColors = m_bXPColors
End Property

Public Property Let XPColors(PropVal As Boolean)
'/* [let] use xp colors
    m_bXPColors = PropVal
End Property

Private Sub Class_Terminate()

    If IsArray(m_lSubItemIcon) Then Erase m_lSubItemIcon
    If IsArray(m_sSubItemText) Then Erase m_sSubItemText

End Sub
