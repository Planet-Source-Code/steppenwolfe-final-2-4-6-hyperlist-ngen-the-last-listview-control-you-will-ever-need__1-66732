VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStandard 
   Caption         =   "Standard Listview Compare Load Speed"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBench 
      Caption         =   " Listview: Add 100,000 Items"
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
      TabIndex        =   1
      Top             =   5355
      Width           =   3165
   End
   Begin MSComctlLib.ListView lvwTest 
      Height          =   4965
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8758
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "iml16"
      SmallIcons      =   "iml16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   8460
      Top             =   1440
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
            Picture         =   "frmStandard.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":57F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":AFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":124E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":14C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":1AF32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":1B24C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":1B56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":1B890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":1BBAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":1E35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStandard.frx":20B0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   210
      Left            =   180
      TabIndex        =   2
      Top             =   5940
      Width           =   375
   End
End
Attribute VB_Name = "frmStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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


Private m_HLIStc()                  As HLIStc
Private cTiming                     As clsTiming


Private Sub Form_Load()

    With lvwTest
        .View = lvwReport
        .Checkboxes = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Item", .Width / 3
        .ColumnHeaders.Add 2, , "SubItem 1", .Width / 3
        .ColumnHeaders.Add 3, , "SubItem 2", (.Width / 3) - 100
    End With
    
    Set cTiming = New clsTiming
    ReDim m_HLIStc(0)
    ItemsArray 100000
    
End Sub

Private Function RandomNum(ByVal lBase As Long, ByVal lSpan As Long) As Long
    RandomNum = Int(Rnd() * lSpan) + lBase
End Function

Private Sub cmdBench_Click()

    cTiming.Reset
    PutStandard
    lblResults.Caption = "100000 items added to Standard List in: " & _
        Format$(cTiming.Elapsed / 1000, "0.0000") & "s"

End Sub

Private Sub PutStandard()

Dim lC      As Long
Dim cItem   As ListItem

On Error Resume Next

    For lC = 0 To UBound(m_HLIStc(0).Item)
        With m_HLIStc(0)
            Set cItem = lvwTest.ListItems.Add(lC, , .Item(lC))
            cItem.SubItems(1) = .SubItem(lC).Text(1)
            cItem.SubItems(2) = .SubItem(lC).Text(2)
        End With
    Next lC
    
On Error GoTo 0

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

Private Function DestroyItems() As Boolean

On Error GoTo Handler

    CopyMemBr ByVal VarPtrArray(m_HLIStc), 0&, 4&
    Erase m_HLIStc
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
    lvwTest.ListItems.Clear
    Set cTiming = Nothing

End Sub
