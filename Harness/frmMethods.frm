VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\prjHyperList.vbp"
Begin VB.Form frmMethods 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HyperList - Methods Demo"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin HyperListUC.ucHyperListNG ucHyperListNG1 
      Height          =   3570
      Left            =   135
      TabIndex        =   14
      Top             =   180
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   6297
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
      ListMode        =   0
      ThemeColor      =   -1
      UseThemeColors  =   0   'False
      ViewMode        =   0
      XPColors        =   -1  'True
   End
   Begin VB.Frame frmDemo 
      Caption         =   "Demo"
      Height          =   3660
      Index           =   0
      Left            =   4545
      TabIndex        =   0
      Top             =   90
      Width           =   3300
      Begin VB.CommandButton cmdItems 
         Caption         =   "Get Selected"
         Height          =   285
         Index           =   8
         Left            =   135
         TabIndex        =   12
         Top             =   3195
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "Get Checked"
         Height          =   285
         Index           =   7
         Left            =   135
         TabIndex        =   11
         Top             =   2835
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "UnCheck All"
         Height          =   285
         Index           =   6
         Left            =   135
         TabIndex        =   10
         Top             =   2475
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "Check All"
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   9
         Top             =   2115
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "SubItem 0 Text"
         Height          =   285
         Index           =   4
         Left            =   135
         TabIndex        =   5
         Top             =   1755
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "Item 0 Icon "
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   4
         Top             =   1395
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "Item 0 Text"
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   1035
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "Remove Item"
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   675
         Width           =   1320
      End
      Begin VB.CommandButton cmdItems 
         Caption         =   "Add Item"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label lblData 
         Height          =   195
         Index           =   3
         Left            =   1575
         TabIndex        =   13
         Top             =   3240
         Width           =   690
      End
      Begin VB.Label lblData 
         Height          =   195
         Index           =   2
         Left            =   1575
         TabIndex        =   8
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblData 
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   7
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label lblData 
         Height          =   195
         Index           =   0
         Left            =   1575
         TabIndex        =   6
         Top             =   1080
         Width           =   690
      End
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmMethods.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":57F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":AFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":124E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":14C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":1AF32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":1B24C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":1B56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":1B890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":1BBAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":1E35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMethods.frx":20B0E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

Dim lX  As Long
Dim lC  As Long

    With ucHyperListNG1
        lX = .Width / Screen.TwipsPerPixelX
        '/* set the list mode (rem to test custom draw mode)
        .ListMode = eHyperList
        '/* initialize header image list
        .InitImlHeader
        '/* add header icons
        .ImlHeaderAddIcon iml16.ListImages.Item(1).Picture
        .ImlHeaderAddIcon iml16.ListImages.Item(2).Picture
        '/* initialize imagelist
        .InitImlSmall
        '/* add small icons
        For lC = 1 To 11
            .ImlSmallAddIcon iml16.ListImages.Item(lC).Picture
        Next lC
        'Stop
        '/* use checkboxes
        .Checkboxes = True
        .InitImlState
        '.ImlStateAddIcon iml16.ListImages.Item(1).Picture
        '.ImlStateAddIcon iml16.ListImages.Item(2).Picture
        .LoadStateImageList
        '/* set viewmode
        .ViewMode = StyleReport
        '/* add columns
        .ColumnAdd 0, "Item", lX / 3, [ColumnLeft]
        .ColumnAdd 1, "Sub 2", lX / 3, [ColumnLeft]
        .ColumnAdd 2, "Sub 3", lX / 3, [ColumnLeft]
        '/* use sorting
        .ItemsSorted = True
        
        '/* skin checkboxes
        .SkinCheckBox CheckBoxClassic, False
        '/* skin headers
        .SkinHeaders HeaderClassic, &H753F17, &H94511F, &HC46A28, False
        '/* skin scrollbars
        .SkinScrollBars ScrollClassic, False
        '/* manually initialize list
        .InitList 0, 2 '<- 0 items, 2 subitem columns
    End With
    
    '/* load a start item
    With ucHyperListNG1
        .ItemAdd 0, "", "TestItem 0", 3, 3
        .SubItemsAdd 0, 1, "subitem 1"
        .SubItemsAdd 0, 2, "subitem 2"
    End With

End Sub

Private Sub cmdItems_Click(Index As Integer)

Dim lCt As Long

    With ucHyperListNG1
        lCt = .Count
        Select Case Index
        '/* add an item
        Case 0
            .ItemAdd lCt, "", "TestItem " & lCt, 6, 3
            .SubItemsAdd lCt, 1, "subitem 1"
            .SubItemsAdd lCt, 2, "subitem 2"
        '/* remove item
        Case 1
            .ItemRemove (lCt - 1)
        '/* item text
        Case 2
            lblData(0).Caption = .ItemText(0)
        '/* item icon index
        Case 3
            lblData(1).Caption = .ItemIcon(0)
        '/* subitem text
        Case 4
            lblData(2).Caption = .SubItemText(0, 1)
        '/* check all
        Case 5
            .CheckAll
        '/* uncheck all
        Case 6
            .UnCheckAll
        '/* get checked
        Case 7
            For lCt = 0 To lCt - 1
                If .Checked(lCt) Then
                    Debug.Print "Item " & lCt & " checked"
                End If
            Next lCt
        '/* selected
        Case 8
            For lCt = 0 To lCt - 1
                If .ItemSelected(lCt) Then
                    Debug.Print "Item " & lCt & " selected"
                End If
            Next lCt
        End Select
    End With

End Sub
