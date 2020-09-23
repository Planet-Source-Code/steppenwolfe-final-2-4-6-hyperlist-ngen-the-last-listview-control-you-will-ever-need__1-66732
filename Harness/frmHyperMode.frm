VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\prjHyperList.vbp"
Begin VB.Form frmHyperMode 
   Caption         =   "HyperList 4.0 - Hyper Mode Example"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin HyperListUC.ucHyperListNG ucHyperListNG1 
      Height          =   4605
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   8123
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaBarTransparency=   70
      ForeColor       =   0
      HeaderColor     =   13160660
      HeaderDragDrop  =   0   'False
      HeaderForeColor =   0
      HeaderHighLite  =   0
      HeaderPressed   =   0
      ListMode        =   2
      ThemeColor      =   -1
      UseThemeColors  =   0   'False
      ViewMode        =   0
      XPColors        =   -1  'True
   End
   Begin VB.CommandButton cmdRemDup 
      Caption         =   "Remove Duplicates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6390
      TabIndex        =   6
      Top             =   5895
      Width           =   1905
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6390
      TabIndex        =   5
      Top             =   5580
      Width           =   1905
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Load"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6390
      TabIndex        =   4
      Top             =   5310
      Width           =   1905
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6390
      TabIndex        =   3
      Top             =   4995
      Width           =   1905
   End
   Begin VB.ComboBox cbStyles 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   5175
      Width           =   1680
   End
   Begin VB.CommandButton cmdBench 
      Caption         =   "HyperList Bench: Add 100,000 Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   180
      TabIndex        =   0
      Top             =   5130
      Width           =   3570
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   7785
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":57F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":AFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":124E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":14C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":1AF32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":1B24C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":1B56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":1B890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":1BBAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":1E35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":20B0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   7785
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":26870
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":26B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":26EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":271BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":274D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":29C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":29FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":2C756
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":2CA70
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":2CD8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":2D0A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHyperMode.frx":2D3BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   210
      Left            =   180
      TabIndex        =   1
      Top             =   5625
      Width           =   375
   End
End
Attribute VB_Name = "frmHyperMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'~ HyperList 4.0 Test Harness by John Underhill (Steppenwolfe)


Private Type HLISubItm
    lIcon()                         As Long
    Text()                          As String
End Type

Private Type HLIStc
    Item()                          As String
    lIcon()                         As Long
    SubItem()                       As HLISubItm
End Type

Private Declare Sub CopyMemBv Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal pDest As Any, _
                                                                    ByVal pSrc As Any, _
                                                                    ByVal lByteLen As Long)


Private Declare Sub CopyMemBr Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                  pSrc As Any, _
                                                                  ByVal lByteLen As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long


Private m_bReportMode               As Boolean
Private m_lItemCount                As Long
Private m_lPointer                  As Long
Private m_aData2()                  As String
Private m_HLIStc()                  As HLIStc
Private cTiming                     As clsTiming


'> Properties
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Property Let Pointer(ByVal PropVal As Long)
    m_lPointer = PropVal
End Property

Private Property Get Pointer() As Long
    Pointer = m_lPointer
End Property

Private Property Let ItemCount(ByVal PropVal As Long)
    m_lItemCount = PropVal
End Property

Private Property Get ItemCount() As Long
    ItemCount = m_lItemCount
End Property

Private Sub cmdFind_Click()
    frmFind.FindParent = 1
    frmFind.Show vbModeless, Me
End Sub

Private Sub cmdRemDup_Click()
    ucHyperListNG1.RemoveDuplicates
End Sub

Private Sub Form_Resize()

On Error Resume Next

    With ucHyperListNG1
        .Left = 135
        .Width = Me.ScaleWidth - 270
        .Top = 90
    End With

End Sub

'> Hyperlist Events
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ucHyperListNG1_eHColumnClick(ByVal Column As Long)
    Debug.Print "Column: " & Column & " clicked."
End Sub

Private Sub ucHyperListNG1_eHErrCond(ByVal sRtn As String, ByVal lErr As Long)
    Debug.Print "Error " & lErr & " in routine: " & sRtn
End Sub

Private Sub ucHyperListNG1_eHItemCheck(ByVal lItem As Long)
    Debug.Print "Item: " & lItem & " checked."
End Sub

Private Function RandomNum(ByVal lBase As Long, ByVal lSpan As Long) As Long
    RandomNum = Int(Rnd() * lSpan) + lBase
End Function

Private Sub cmdSave_Click(Index As Integer)

    Select Case Index
    Case 0
        ucHyperListNG1.SaveToFile App.Path & "test.dat"
        cmdSave(1).Enabled = True
    Case 1
        ucHyperListNG1.ClearList
        ucHyperListNG1.LoadFromFile App.Path & "test.dat"
    End Select
    
End Sub

Public Sub Form_Load()

Dim lX  As Long
Dim lC  As Long

    Randomize
    '/* create a portable structure
    ReDim m_HLIStc(0)

    '/* instance list and timer
    Set cTiming = New clsTiming
    
    '/* apply list settings
    With ucHyperListNG1
        lX = .Width / Screen.TwipsPerPixelX
        .ListMode = eHyperList
        '/* large icons
        .InitImlLarge
        For lC = 1 To 12
            .ImlLargeAddIcon iml32.ListImages.Item(lC).Picture
        Next lC
        '/* header icons
        .InitImlHeader
        .ImlHeaderAddIcon iml16.ListImages.Item(1).Picture
        .ImlHeaderAddIcon iml16.ListImages.Item(2).Picture
        .FullRowSelect = True
        '/* small images
        .InitImlSmall ' 32, 32
        For lC = 1 To 12
            .ImlSmallAddIcon iml16.ListImages.Item(lC).Picture
        Next lC
        '/* enable checkboxes
        .Checkboxes = True
        .ForeColor = &H454D40
        '/* set viewmode
        .ViewMode = StyleReport
        '/* add columns
        .ColumnAdd 0, "Item", lX / 3, [ColumnLeft], -1, SortAuto
        .ColumnAdd 1, "Sub 2", lX / 3, [ColumnLeft], -1, SortAuto
        .ColumnAdd 2, "Sub 3", lX / 3, [ColumnLeft], -1, SortAuto
        .ItemsSorted = True
        .SkinCheckBox CheckBoxMetallic, False
        .SkinHeaders HeaderMetallic, &HE1E8E7, &H524944, &H3A574D, False
        .SkinScrollBars ScrollMetallic, False
    End With

    '/* dimension the list
    ItemCount = 100000
    
    ItemsArray ItemCount
    cmdBench.Caption = "HyperList Bench: Add " & ItemCount & " Items"
    '/* list styles
    With cbStyles
        .Text = "Styles"
        .AddItem "0 - Icon"
        .AddItem "1 - Report"
        .AddItem "2 - SmallIcon"
        .AddItem "3 - List"
        .ListIndex = 1
    End With

End Sub

Private Sub cmdBench_Click()

    cTiming.Reset
    PutArray
    lblResults.Caption = ItemCount & " items added to HyperList in: " & _
        Format$(cTiming.Elapsed / 1000, "0.0000") & "s"

End Sub

Private Sub cbStyles_Click()

    With ucHyperListNG1
        m_bReportMode = (cbStyles.ListIndex = 1)
        .ViewMode = cbStyles.ListIndex
        .ListRefresh
    End With

End Sub


'> Array Management
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub ItemsArray(ByVal lCount As Long)

Dim lPos            As Long
Dim lOffset         As Long
Dim lLb             As Long
Dim lUb             As Long
Dim aData1()        As String
Dim aData3()        As String
Const SI1           As String = "SubItem1 - "
Const SI2           As String = "SubItem2 - "

    lCount = lCount - 1
    '/* init temp array
    ReDim aData1(0 To 9)
    '/* base array
    aData1(0) = LoadResString(102) & "|" & LoadResString(122)
    aData1(1) = LoadResString(103) & "|" & LoadResString(123)
    aData1(2) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)
    aData1(3) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)
    aData1(4) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)
    aData1(5) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)
    aData1(6) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)
    aData1(7) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)
    aData1(8) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)
    aData1(9) = LoadResString(101 + 10 * Rnd) & "|" & LoadResString(121 + 10 * Rnd)

    '/* generate items array
    ReDim m_aData2(0 To lCount)
    ReDim m_HLIStc(0).Item(0 To lCount)
    '/* merge arrays to size
    With m_HLIStc(0)
        For lPos = 0 To (lCount) Step 10
            '/* create a 'scratch array' to avoid pointer duplication
            aData3 = aData1
            For lOffset = 0 To 9
                '/* copy the pointer to the dest array
                CopyMemBv VarPtr(.Item(lOffset + lPos)), VarPtr(aData3(lOffset)), 4&
                '/* deallocate the string
                CopyMemBr ByVal VarPtr(aData3(lOffset)), 0&, 4&
            Next lOffset
        Next lPos
    End With
    
    '/* generate subitems array
    lLb = LBound(m_aData2)
    lUb = UBound(m_aData2)
    lPos = 0
    With m_HLIStc(0)
        ReDim .SubItem(lLb To lUb)
        ReDim .lIcon(lLb To lUb)
        Do
            .lIcon(lPos) = RandomNum(2, 7)
            ReDim .SubItem(lPos).Text(1 To 2)
            ReDim .SubItem(lPos).lIcon(1 To 2)
            .SubItem(lPos).Text(1) = SI1 & lPos
            .SubItem(lPos).lIcon(1) = RandomNum(2, 7)
            .SubItem(lPos).Text(2) = SI2 & lPos
            .SubItem(lPos).lIcon(2) = RandomNum(2, 7)
            lPos = lPos + 1
        Loop Until lPos > lUb
    End With

End Sub

Private Sub PutArray()
'/* forward struct pointer into library

On Error GoTo Handler

    If Not ArrayCheck(m_HLIStc(0).Item) Then
        ItemsArray ItemCount
    End If
    '/* copy struct pointer into list control
    If Not m_lPointer = 0 Then
        DestroyItems
        ResetTempArray
    End If
    CopyMemBr m_lPointer, ByVal VarPtrArray(m_HLIStc), 4&
    ucHyperListNG1.StructPtr = m_lPointer
    '/* load the data struct
    ucHyperListNG1.LoadArray
    '/* set the item count, this will fire the callback
    '/* and populate the list
    ucHyperListNG1.SetItemCount UBound(m_HLIStc(0).Item) + 1
    
Handler:
    On Error GoTo 0

End Sub

Private Sub ResetTempArray()

    If ArrayCheck(m_aData2) Then
        Erase m_aData2
    End If
    ItemsArray ItemCount

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

Private Function DestroyItems() As Boolean

On Error GoTo Handler

    CopyMemBr ByVal VarPtrArray(m_HLIStc), 0&, 4&
    Erase m_HLIStc
    m_lPointer = 0
    DestroyItems = True

Handler:
    On Error GoTo 0

End Function


'> Timing
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub Reset()
'/* reset timer
    cTiming.Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
'/* destroy the pointer

    DestroyItems
    Set cTiming = Nothing

End Sub

