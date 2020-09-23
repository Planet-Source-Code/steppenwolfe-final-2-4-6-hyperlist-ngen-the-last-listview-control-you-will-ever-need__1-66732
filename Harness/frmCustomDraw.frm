VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\prjHyperList.vbp"
Begin VB.Form frmCustomDraw 
   BackColor       =   &H00FFFFFF&
   Caption         =   "HyperList NGEN -Custom Draw Mode Example"
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
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
   ScaleHeight     =   10920
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin HyperListUC.ucHyperListNG ucHyperListNG1 
      Height          =   7935
      Left            =   135
      TabIndex        =   74
      Top             =   180
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   13996
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
      XPColors        =   -1  'True
   End
   Begin VB.PictureBox picBar 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   14520
      TabIndex        =   38
      Top             =   10635
      Width           =   14580
   End
   Begin VB.CommandButton cmdBench 
      Caption         =   "HyperList Bench: Add 25,000 Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   135
      TabIndex        =   32
      Top             =   8325
      Width           =   3570
   End
   Begin VB.Frame fmFunctions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Functions (40 more..)"
      Height          =   1770
      Left            =   10125
      TabIndex        =   4
      Top             =   8820
      Width           =   4335
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   45
         ScaleHeight     =   1500
         ScaleWidth      =   4245
         TabIndex        =   30
         Top             =   180
         Width           =   4245
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Methods"
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
            Left            =   2250
            TabIndex        =   72
            Top             =   450
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
            Left            =   2250
            TabIndex        =   71
            Top             =   90
            Width           =   1905
         End
         Begin VB.CommandButton cmdFunct 
            Caption         =   "Clear List"
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
            Index           =   3
            Left            =   135
            TabIndex        =   43
            Top             =   1170
            Width           =   1950
         End
         Begin VB.CommandButton cmdFunct 
            Caption         =   "UnCheck All"
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
            Index           =   2
            Left            =   135
            TabIndex        =   42
            Top             =   810
            Width           =   1950
         End
         Begin VB.CommandButton cmdFunct 
            Caption         =   "Check All"
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
            Index           =   1
            Left            =   135
            TabIndex        =   41
            Top             =   450
            Width           =   1950
         End
         Begin VB.CommandButton cmdFunct 
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
            Index           =   0
            Left            =   135
            TabIndex        =   39
            Top             =   90
            Width           =   1950
         End
      End
   End
   Begin VB.Frame fmModes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modes"
      Height          =   1770
      Left            =   6165
      TabIndex        =   3
      Top             =   8820
      Width           =   3840
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   45
         ScaleHeight     =   1455
         ScaleWidth      =   3705
         TabIndex        =   29
         Top             =   225
         Width           =   3705
         Begin VB.CommandButton cmdBench 
            Caption         =   "DataBase Mode Add 100 Million Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   180
            TabIndex        =   34
            Top             =   540
            Width           =   3390
         End
         Begin VB.CommandButton cmdBench 
            Caption         =   "Hyper Mode: Add 100k Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   33
            Top             =   90
            Width           =   3390
         End
         Begin VB.CommandButton cmdBench 
            Caption         =   "Standard Listview: Add 100k Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   31
            Top             =   990
            Width           =   3390
         End
      End
   End
   Begin VB.Frame fmExtended 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Extended Styles"
      Height          =   1770
      Left            =   135
      TabIndex        =   2
      Top             =   8820
      Width           =   5910
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   45
         ScaleHeight     =   1455
         ScaleWidth      =   5775
         TabIndex        =   28
         Top             =   225
         Width           =   5775
         Begin VB.CheckBox chkSkin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Colorize"
            Height          =   210
            Index           =   3
            Left            =   4455
            TabIndex        =   60
            Top             =   540
            Width           =   915
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Left            =   0
            ScaleHeight     =   690
            ScaleWidth      =   5730
            TabIndex        =   53
            Top             =   765
            Width           =   5730
            Begin VB.CommandButton cmdReset 
               Caption         =   "Reset"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4545
               TabIndex        =   59
               Top             =   270
               Width           =   1095
            End
            Begin VB.TextBox txtDepth 
               Height          =   285
               Left            =   2475
               TabIndex        =   58
               Text            =   "2"
               Top             =   315
               Width           =   240
            End
            Begin VB.OptionButton optRowDec 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Linear"
               Height          =   210
               Index           =   1
               Left            =   1125
               TabIndex        =   56
               Top             =   360
               Width           =   825
            End
            Begin VB.OptionButton optRowDec 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Checker"
               Height          =   210
               Index           =   0
               Left            =   90
               TabIndex        =   54
               Top             =   360
               Value           =   -1  'True
               Width           =   960
            End
            Begin VB.Label lblStyles 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Depth (ordinal)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   2475
               TabIndex        =   57
               Top             =   90
               Width           =   1215
            End
            Begin VB.Label lblStyles 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Row Decoration"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   90
               TabIndex        =   55
               Top             =   90
               Width           =   1290
            End
         End
         Begin VB.CheckBox chkSkin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skin Checkboxes"
            Height          =   210
            Index           =   2
            Left            =   2835
            TabIndex        =   52
            Top             =   540
            Value           =   1  'Checked
            Width           =   1545
         End
         Begin VB.CheckBox chkSkin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skin Scrollbars"
            Height          =   210
            Index           =   1
            Left            =   1350
            TabIndex        =   51
            Top             =   540
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkSkin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skin Header"
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   50
            Top             =   540
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Gloss"
            Height          =   210
            Index           =   4
            Left            =   3780
            TabIndex        =   49
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Metallic"
            Height          =   210
            Index           =   3
            Left            =   2835
            TabIndex        =   48
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lime"
            Height          =   210
            Index           =   2
            Left            =   2025
            TabIndex        =   47
            Top             =   270
            Width           =   690
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Eclipse"
            Height          =   210
            Index           =   1
            Left            =   1035
            TabIndex        =   46
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Classic"
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   45
            Top             =   270
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.Label lblStyles 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skin Styles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   44
            Top             =   0
            Width           =   915
         End
      End
   End
   Begin VB.Frame fmProp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Properties (there are more then 70 in all..)"
      Height          =   8655
      Left            =   8865
      TabIndex        =   1
      Top             =   90
      Width           =   5595
      Begin VB.PictureBox picBg 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8340
         Left            =   90
         ScaleHeight     =   8340
         ScaleWidth      =   5370
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   5370
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Row Drop in List"
            Height          =   240
            Index           =   21
            Left            =   2790
            TabIndex        =   73
            Top             =   3735
            Width           =   1860
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "XP Headers (Demo)"
            Height          =   240
            Index           =   29
            Left            =   2790
            TabIndex        =   70
            Top             =   3375
            Width           =   2310
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable Unicode (NT/2K/XP)"
            Height          =   240
            Index           =   28
            Left            =   2790
            TabIndex        =   69
            Top             =   3015
            Width           =   2400
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Per Cell Colors"
            Height          =   240
            Index           =   27
            Left            =   2790
            TabIndex        =   68
            Top             =   2655
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Per Cell Fonts"
            Height          =   240
            Index           =   26
            Left            =   2790
            TabIndex        =   67
            Top             =   2295
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BackGround Picture"
            Height          =   240
            Index           =   30
            Left            =   2790
            TabIndex        =   61
            Top             =   4230
            Width           =   1860
         End
         Begin VB.OptionButton optBg 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Center"
            Height          =   195
            Index           =   1
            Left            =   3465
            TabIndex        =   64
            Top             =   4500
            Width           =   825
         End
         Begin VB.OptionButton optBg 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tile"
            Height          =   195
            Index           =   0
            Left            =   2790
            TabIndex        =   63
            Top             =   4500
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SubItem Edit"
            Height          =   240
            Index           =   25
            Left            =   2790
            TabIndex        =   62
            Top             =   1935
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Backcolor (custom draw off)"
            Height          =   240
            Index           =   24
            Left            =   2790
            TabIndex        =   40
            Top             =   1575
            Value           =   1  'Checked
            Width           =   2490
         End
         Begin VB.ComboBox cbBorder 
            Height          =   330
            Left            =   2790
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   4950
            Width           =   1695
         End
         Begin VB.ComboBox cbStyles 
            Height          =   330
            ItemData        =   "frmCustomDraw.frx":0000
            Left            =   2790
            List            =   "frmCustomDraw.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   5490
            Width           =   1695
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "AlphaBar Active"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   180
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Checkboxes"
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   540
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Column Align"
            Height          =   240
            Index           =   2
            Left            =   180
            TabIndex        =   25
            Top             =   885
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Custom Draw"
            Height          =   240
            Index           =   3
            Left            =   180
            TabIndex        =   24
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enabled"
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   23
            Top             =   1590
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ForeColor"
            Height          =   240
            Index           =   5
            Left            =   180
            TabIndex        =   22
            Top             =   1935
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Full Row Select"
            Height          =   240
            Index           =   6
            Left            =   180
            TabIndex        =   21
            Top             =   2295
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Grid Lines"
            Height          =   240
            Index           =   7
            Left            =   180
            TabIndex        =   20
            Top             =   2640
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Header Drag Drop"
            Height          =   240
            Index           =   8
            Left            =   180
            TabIndex        =   19
            Top             =   3000
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Header Fixed Width"
            Height          =   240
            Index           =   9
            Left            =   180
            TabIndex        =   18
            Top             =   3345
            Width           =   1770
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Header Flat"
            Height          =   240
            Index           =   10
            Left            =   180
            TabIndex        =   17
            Top             =   3705
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Header Hide"
            Height          =   240
            Index           =   11
            Left            =   180
            TabIndex        =   16
            Top             =   4050
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Info Tips"
            Height          =   240
            Index           =   12
            Left            =   180
            TabIndex        =   15
            Top             =   4410
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item Border (Icon)"
            Height          =   240
            Index           =   13
            Left            =   180
            TabIndex        =   14
            Top             =   4755
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item Indent (Icon mode)"
            Height          =   240
            Index           =   14
            Left            =   180
            TabIndex        =   13
            Top             =   5115
            Width           =   2040
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label Edit"
            Height          =   240
            Index           =   15
            Left            =   180
            TabIndex        =   12
            Top             =   5445
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label Tips"
            Height          =   240
            Index           =   16
            Left            =   180
            TabIndex        =   11
            Top             =   5760
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Multi Select"
            Height          =   240
            Index           =   17
            Left            =   180
            TabIndex        =   10
            Top             =   6105
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Scroll Bar Flat"
            Height          =   240
            Index           =   18
            Left            =   2790
            TabIndex        =   9
            Top             =   180
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SubItem Images"
            Height          =   240
            Index           =   19
            Left            =   2790
            TabIndex        =   8
            Top             =   510
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Track Selected"
            Height          =   240
            Index           =   20
            Left            =   2790
            TabIndex        =   7
            Top             =   870
            Width           =   1635
         End
         Begin VB.CheckBox chkProperties 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Visible"
            Height          =   240
            Index           =   22
            Left            =   2790
            TabIndex        =   6
            Top             =   1215
            Width           =   1635
         End
         Begin MSComctlLib.ListView lvwOLETest 
            Height          =   1275
            Left            =   0
            TabIndex        =   66
            ToolTipText     =   "Drag an item from the list and place it here"
            Top             =   6840
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   2249
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            OLEDropMode     =   1
            NumItems        =   0
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OLE Drag and Drop Demo:"
            Height          =   210
            Index           =   2
            Left            =   45
            TabIndex        =   65
            Top             =   6615
            Width           =   1890
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "View:"
            Height          =   210
            Index           =   1
            Left            =   2790
            TabIndex        =   0
            Top             =   5310
            Width           =   435
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Border:"
            Height          =   210
            Index           =   0
            Left            =   2790
            TabIndex        =   37
            Top             =   4770
            Width           =   540
         End
      End
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   13050
      Top             =   7920
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
            Picture         =   "frmCustomDraw.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":57F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":AFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":124EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":14C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":1AF36
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":1B250
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":1B572
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":1B894
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":1BBAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":1E360
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":20B12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   13050
      Top             =   7335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":26874
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":26B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":26EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":271C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":274DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":29C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":29FA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":2C75A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":2CA74
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":2CD8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":2D0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":2D3C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomDraw.frx":2D6DC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCustomDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'~ HyperList 2.4 Test Harness by John Underhill (Steppenwolfe)



Private Declare Sub CopyMemBr Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                    pSrc As Any, _
                                                                    ByVal lByteLen As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long


Private m_bLoading                  As Boolean
Private m_bReportMode               As Boolean
Private m_lItemCount                As Long
Private m_lPointer                  As Long
Private m_lCt                       As Long
Private m_oForeColor                As OLE_COLOR
Private m_oHighlite                 As OLE_COLOR
Private m_oPressed                  As OLE_COLOR
Private m_oBaseClr                  As OLE_COLOR
Private m_oOffset                   As OLE_COLOR
Private m_aData2()                  As String
Private m_oListItem                 As ListItem
Private m_cListItems()              As clsListItem
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
    frmFind.FindParent = 0
    frmFind.Show vbModeless, Me
End Sub

Private Sub cmdMethods_Click()
    frmMethods.Show
End Sub

'> Events
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

Private Sub ucHyperListNG1_eHItemClick(ByVal lItem As Long)
    Debug.Print "Item " & lItem & " clicked"
End Sub


Public Sub Form_Load()

Dim lX  As Long
Dim lC  As Long

    m_bLoading = True
    Randomize
    '/* default colors
    m_oForeColor = &H683815
    m_oBaseClr = &HF6E1D2
    m_oHighlite = &H94511F
    m_oPressed = &HC46A28
    
    '/* instance list and timer
    Set cTiming = New clsTiming
    '/* apply list settings
    With ucHyperListNG1
        If Not (.IsWinNT) Then
            chkProperties(28).Value = 2
            chkProperties(28).Enabled = False
        End If
        lX = .Width / Screen.TwipsPerPixelX
        '/* add columns
        .ColumnAdd 0, "Item", (lX / 3), ColumnLeft, -1, SortAuto
        .ColumnAdd 1, "Sub 2", (lX / 3), ColumnCenter, -1, SortAuto
        .ColumnAdd 2, "Sub 3", (lX / 3), ColumnLeft, -1, SortAuto
        '/* set viewmode
        .ViewMode = StyleReport
        '/* large icons
        .InitImlLarge
        For lC = 1 To 12
            .ImlLargeAddIcon iml32.ListImages(lC).Picture
        Next lC
        '/* header icons
        .InitImlHeader
        .ImlHeaderAddIcon iml16.ListImages.Item(1).Picture
        .ImlHeaderAddIcon iml16.ListImages.Item(2).Picture
        '/* small images
        .InitImlSmall '32, 32
        For lC = 1 To 12
            .ImlSmallAddIcon iml16.ListImages(lC).Picture
        Next lC
        '/* enable checkboxes
        .Checkboxes = True
        .FullRowSelect = True
        .ItemsSorted = True
        .ThemeColor = &H9C541F
        .ThemeLuminence = ThemeSoft
        .SkinCheckBox CheckBoxClassic, False
        .SkinHeaders HeaderClassic, &H753F17, &H94511F, &HC46A28, False
        .SkinScrollBars ScrollClassic, False
        .AlphaSelectorBar 120, True, True
        .AlphaBarActive = True
        .BackColor = &HF6E1D2
        .ForeColor = &H9C541F
        .RowDecoration RowChecker, &H9C541F, &HDF9660, True, 2
    End With

    '/* dimension the list
    ItemCount = 10000
    ItemsArray ItemCount
    cmdBench(0).Caption = "HyperList Bench: Add " & ItemCount & " Items"
    
    '/* list styles
    With cbStyles
        .AddItem "0 - Icon"
        .AddItem "1 - Report"
        .AddItem "2 - SmallIcon"
        .AddItem "3 - List"
        .ListIndex = 1
    End With
    '/* border
    With cbBorder
        .AddItem "0 - None"
        .AddItem "1 - Thin"
        .AddItem "2 - Thick"
        .ListIndex = 1
    End With
    m_bLoading = False
    
    With lvwOLETest
        .Appearance = ccFlat
        .View = lvwReport
        .ColumnHeaders.Add 1, Text:="Item", Width:=(.Width / 2)
        .ColumnHeaders.Add 2, Text:="Col 1", Width:=(.Width / 4)
        .ColumnHeaders.Add 3, Text:="Col 2", Width:=(.Width / 4)
        Set .SmallIcons = iml16
        Set m_oListItem = .ListItems.Add(1, , "drag item here", 0, 3)
        m_oListItem.SubItems(1) = "sub 1"
        m_oListItem.SubItems(2) = "sub 2"
    End With

End Sub

Private Sub Form_Resize()

On Error Resume Next

    With ucHyperListNG1
        fmProp.Left = (Me.ScaleWidth - (fmProp.Width + 100))
        .Width = Me.ScaleWidth - (fmProp.Width + 700)
    End With
    
End Sub

Private Sub cmdBench_Click(Index As Integer)

    cTiming.Reset

    Select Case Index
    '/* hyperlist
    Case 0
        PutArray
        picBar.Print " " & ItemCount & " items added to HyperList in: " & _
            Format$(cTiming.Elapsed / 1000, "0.0000") & "s"
        cmdBench(1).Enabled = True
        cmdBench(0).Enabled = False
    '/* standard listview
    Case 1
        frmStandard.Show
    '/* hyper mode
    Case 2
        frmHyperMode.Show
    '/* database mode
    Case 3
        frmDataBaseMode.Show
    End Select

End Sub

Private Sub chkProperties_Click(Index As Integer)

    With ucHyperListNG1
        Select Case Index
        Case 0
            .AlphaBarActive = (Not .AlphaBarActive)
        Case 1
            If .Checkboxes = False Then
                .Checkboxes = True
                If chkSkin(2).Value = 1 Then
                    .SkinCheckBox m_lCt, False
                End If
            Else
                .Checkboxes = False
            End If
        Case 2
            If .ColumnAlign(1) = Columnright Then
                .ColumnAlign(1) = ColumnLeft
            Else
                .ColumnAlign(1) = Columnright
            End If
        Case 3
            .CustomDraw = (Not .CustomDraw)
        Case 4
            .Enabled = (Not .Enabled)
        Case 5
            If .ForeColor = m_oForeColor Then
                .ForeColor = &H0
            Else
                .ForeColor = m_oForeColor
            End If
        Case 6
            .FullRowSelect = (Not .FullRowSelect)
        Case 7
            .GridLines = (Not .GridLines)
        Case 8
            .HeaderDragDrop = (Not .HeaderDragDrop)
        Case 9
            .HeaderFixedWidth = (Not .HeaderFixedWidth)
        Case 10
            If chkProperties(10).Value = 1 Then
                .UnSkinHeaders
            Else
                .SkinHeaders m_lCt, m_oForeColor, m_oHighlite, m_oPressed, False
            End If
            .HeaderFlat = (Not .HeaderFlat)
        Case 11
            .HeaderHide = (Not .HeaderHide)
        Case 12
            .InfoTips = (Not .InfoTips)
        Case 13
            If .ViewMode = StyleIcon Then
                .ItemBorderSelect = (Not .ItemBorderSelect)
            End If
        Case 14
            If .ViewMode = StyleIcon Then
                If Not (.IconSpaceX = 200) Then
                    .IconSpaceX = 200
                    .Checkboxes = False
                Else
                    .IconSpaceX = 75
                    .Checkboxes = CBool(chkProperties(1).Value)
                End If
                .Visible = False
                .Visible = True
            End If
        Case 15
            .LabelEdit = (Not .LabelEdit)
        Case 16
            .LabelTips = (Not .LabelTips)
        Case 17
            If chkProperties(17).Value = 1 Then
                .AlphaBarActive = False
            Else
                .AlphaBarActive = CBool(chkProperties(0).Value)
            End If
            .MultiSelect = (Not .MultiSelect)
        Case 18
            If chkProperties(18).Value = 1 Then
                .UnSkinScrollBars
            Else
                .SkinScrollBars m_lCt, False
            End If
            .ScrollBarFlat = (Not .ScrollBarFlat)
        Case 19
            .SubItemImages = (Not .SubItemImages)
            .ListRefresh 1
        Case 20
            .TrackSelected = (Not .TrackSelected)
        Case 21
            If chkProperties(21).Value = 1 Then
                .OLEDropMode = vbOLEDropManual
            Else
                .OLEDropMode = vbOLEDropNone
            End If
        Case 22
            .Visible = (Not .Visible)
            If .Visible Then
            End If
        Case 23
            .XPColors = (Not .XPColors)
        Case 24
            If chkProperties(24).Value = 1 Then
                .CustomDraw = False
                .BackColor = m_oBaseClr
            Else
                .CustomDraw = CBool(chkProperties(3).Value = 1)
            End If
        Case 25
            If chkProperties(25).Value = 1 Then
                If .LabelEdit = False Then
                    .LabelEdit = True
                End If
                .SubItemsEdit = True
            Else
                .SubItemsEdit = CBool(chkProperties(25).Value)
                .LabelEdit = False
            End If
        Case 26
            .UseCellFont = (Not .UseCellFont)
        Case 27
            .UseCellColor = (Not .UseCellColor)
        Case 28
            If (.IsWinNT) Then
                .UseUnicode = CBool(chkProperties(28).Value = 1)
            End If
        Case 29
            If (.IsWinXP) Then
                If chkProperties(29).Value = 1 Then
                    If chkSkin(1).Value = 1 Then
                        .UnSkinHeaders
                    End If
                    .SkinXPHeader
                Else
                    .UnSkinXPHeader
                    .ListRefresh 1
                    If chkSkin(1).Value = 1 Then
                        .SkinHeaders m_lCt, m_oForeColor, m_oHighlite, m_oPressed, False
                        .ListRefresh
                    End If
                End If
            End If
        Case 30
            If chkProperties(30).Value Then
                If .CustomDraw Then
                    .CustomDraw = False
                End If
                If optBg(0).Value Then
                    .BackgroundPicture App.Path & "\Images\tile.gif", BgTile
                Else
                    .BackColor = &HFFFFFF
                    .BackgroundPicture App.Path & "\Images\bg.gif", BgNormal, True
                End If
            Else
                .BackColor = m_oBaseClr
                .BackgroundPicture App.Path & "\Images\bg.gif", BgNone
                .CustomDraw = CBool(chkProperties(3).Value = 1)
            End If
        End Select
    End With
    
End Sub

Private Sub cmdFunct_Click(Index As Integer)

    With ucHyperListNG1
        Select Case Index
        Case 0
            .RemoveDuplicates
        Case 1
            .CheckAll
        Case 2
            .UnCheckAll
        Case 3
            .ClearList
            cmdBench(0).Enabled = True
        End Select
    End With
    
End Sub

Private Sub cmdReset_Click()

Dim bCustom     As Boolean
Dim lDepth      As Long
Dim oFont       As StdFont

    With txtDepth
        If Not IsNumeric(.Text) Then
            .Text = 0
        ElseIf .Text > 4 Then
            .Text = 4
        End If
        lDepth = CLng(.Text)
    End With
    For m_lCt = 0 To 4
        If optStyles(m_lCt).Value Then
            Exit For
        End If
    Next m_lCt

    '/* create our alternating test font
    Set oFont = New StdFont
    With ucHyperListNG1
        .UseCellFont = False
        .UseCellColor = False
        .XPColors = True
        Select Case m_lCt
        Case 0
            m_oForeColor = &H753F17
            m_oHighlite = &H94511F
            m_oPressed = &HC46A28
            m_oBaseClr = &H9C541F
            m_oOffset = &HDF9660
            With oFont
                .Name = "Times New Roman"
                .Size = 8
            End With
            .BackColor = &HF6E1D2
            .ForeColor = &H535353
        Case 1
            m_oForeColor = &HF8CBB6
            m_oHighlite = &HFCE9E0
            m_oPressed = &HD8C6BA
            m_oBaseClr = &H966D55
            m_oOffset = &HC1A593
            With oFont
                .Name = "Verdana"
                .Size = 8
            End With
            .BackColor = &HC1A593
            .ForeColor = &H9C541F
        Case 2
            m_oForeColor = &H225E39
            m_oHighlite = &H47BC73
            m_oPressed = &HE3F4EA
            m_oBaseClr = &H7DCC9D
            m_oOffset = &HAADDBE
            With oFont
                .Name = "Comic Sans MS"
                .Size = 8
            End With
            .BackColor = &HAADDBE
            .ForeColor = &H163F26
        Case 3
            m_oForeColor = &H3A4136
            m_oHighlite = &H768E6C
            m_oPressed = &HD7DED3
            m_oBaseClr = &H6B7C63
            m_oOffset = &H96A68E
            With oFont
                .Name = "Ariel"
                .Size = 8
            End With
            .BackColor = &HC7CFC3
            .ForeColor = &H454D40
        Case 4
            m_oForeColor = &HE1E8E7
            m_oHighlite = &H524944
            m_oPressed = &H3A574D
            m_oBaseClr = &H81A095
            m_oOffset = &HB0C4BD
            With oFont
                .Name = "Tahoma"
                .Size = 8
            End With
            .BackColor = &HD5DFDC
            .ForeColor = &H454D40
        End Select

        Set .Font = oFont
        .ThemeColor = m_oBaseClr
        .ThemeLuminence = ThemeSoft
        If chkSkin(3).Value = 1 Then
            bCustom = True
        End If
        If chkSkin(0).Value = 1 Then
            .SkinHeaders m_lCt, m_oForeColor, m_oHighlite, m_oPressed, bCustom
        Else
            .UnSkinHeaders
            .HeaderFlat = True
            .ListRefresh True
        End If
        If chkSkin(1).Value = 1 Then
            .SkinScrollBars m_lCt, bCustom
        Else
            .UnSkinScrollBars
        End If
        If chkSkin(2).Value = 1 Then
            .SkinCheckBox m_lCt, bCustom
        Else
            .UnSkinCheckBox
        End If
        If optRowDec(0).Value Then
            .RowDecoration RowChecker, m_oBaseClr, m_oOffset, True, lDepth
        Else
            .RowDecoration RowLine, m_oBaseClr, m_oOffset, True, lDepth
        End If
        .AlphaSelectorBar 120, True, True
        .ListRefresh
    End With
    
End Sub

Private Sub cbStyles_Click()

    With ucHyperListNG1
        m_bReportMode = (cbStyles.ListIndex = 1)
        If Not m_bLoading Then
            .ViewMode = cbStyles.ListIndex
            .ListRefresh
        End If
    End With

End Sub

Private Sub cbBorder_Click()

    With ucHyperListNG1
        .BorderStyle = cbBorder.ListIndex
    End With
    
End Sub

Private Sub lvwOLETest_OLEDragDrop(Data As MSComctlLib.DataObject, _
                                   Effect As Long, _
                                   Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   y As Single)

Dim bData() As Byte
Dim lIcon   As Long
Dim sData   As String
Dim aText() As String

On Error GoTo Handler

    bData = Data.GetData(vbCFText)
    sData = bData
    Debug.Print sData

    aText = Split(sData, vbCrLf)
    If IsNumeric(aText(0)) Then
        lIcon = CLng(aText(0)) + 1
    Else
        lIcon = -1
    End If
    If UBound(aText) = 3 Then
        Set m_oListItem = lvwOLETest.ListItems.Add((lvwOLETest.ListItems.Count + 1), , aText(1), , lIcon)
        m_oListItem.SubItems(1) = aText(2)
        m_oListItem.SubItems(2) = aText(3)
    End If

Handler:

End Sub

Private Sub optBg_Click(Index As Integer)

    If chkProperties(30).Value = 1 Then
        chkProperties(30).Value = 0
        chkProperties(30).Value = 1
    End If

End Sub


'> Array Management
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Function RandomNum(ByVal lBase As Long, ByVal lSpan As Long) As Long
    RandomNum = Int(Rnd() * lSpan) + lBase
End Function

Private Sub InitList(ByVal lCount As Long, _
                     Optional ByVal lSubItemCt As Long = -1)

Dim lCt As Long

    ReDim m_cListItems(0 To lCount) As clsListItem
    For lCt = 0 To lCount
        Set m_cListItems(lCt) = New clsListItem
        If Not lSubItemCt = -1 Then
            m_cListItems(lCt).SubItemCount = lSubItemCt
            m_cListItems(lCt).Init
        End If
    Next lCt

End Sub

Private Sub ItemsArray(ByVal lCount As Long)

Dim lPos            As Long
Dim aData1()        As String
Dim oFont           As StdFont

    lCount = lCount - 1
    '/* init temp array
    ReDim aData1(0 To 9)
    
    '/* unicode: base array
    If ucHyperListNG1.IsWinNT Then
        aData1(0) = LoadResString(101) & "|" & LoadResString(121)
        aData1(1) = LoadResString(102) & "|" & LoadResString(122)
        aData1(2) = LoadResString(103) & "|" & LoadResString(123)
        aData1(3) = LoadResString(104) & "|" & LoadResString(124)
        aData1(4) = LoadResString(105) & "|" & LoadResString(125)
        aData1(5) = LoadResString(113) & "|" & LoadResString(123)
        aData1(6) = LoadResString(107) & "|" & LoadResString(127)
        aData1(7) = LoadResString(108) & "|" & LoadResString(128)
        aData1(8) = LoadResString(109) & "|" & LoadResString(129)
        aData1(9) = LoadResString(110) & "|" & LoadResString(130)
    '/* legacy:
    Else
        aData1(0) = "When label editing begins,"
        aData1(1) = "an edit control is created"
        aData1(2) = "positioned, and initialized."
        aData1(3) = "Before it is displayed, "
        aData1(4) = "the list-view control sends"
        aData1(5) = "its parent window an "
        aData1(6) = "LVN_BEGINLABELEDIT notification message. "
        aData1(7) = "To customize label editing, "
        aData1(8) = "implement a handler for LVN_BEGINLABELEDIT"
        aData1(9) = "and have it send an LVM_GETEDITCONTROL"
    End If
    
    '/* generate items array
    ReDim m_aData2(0 To lCount)
    '/* subitems are ordinal
    InitList lCount, 2
    
    '/* create our alternating test font
    Set oFont = New StdFont
    With oFont
        .Name = "Times New Roman"
        .Size = 10
        .Charset = 3
        .Italic = True
    End With

    '/* merge arrays to size
    For lPos = 0 To (lCount)
        With m_cListItems(lPos)
            .Add lPos, "", aData1(RandomNum(0, 9)), 0, 0
            .SubItemCount = 2
            .SubItem 1, Now - RandomNum(1, 100)
            .SubItem 2, (lPos)
            .SubIcon 1, RandomNum(2, 7)
            .SubIcon 2, RandomNum(2, 7)
            .Icon = RandomNum(2, 7)
            '/* per-cell fore/backcolor and fonts are definable
            If ((lPos) Mod 2) Then
                '/* use custom setting with this row
                .CellCustom = True
                '/* per-row font selection (could be expanded to per cell)
                Set .Font = oFont
                '/* cell colors
                .XPColors = True
                .CellBackColor(0) = RGB(50, 150, 250)
                .CellForeColor(0) = &H222222
                .CellBackColor(1) = RGB(50, 200, 200)
                .CellForeColor(1) = &H444444
                .CellBackColor(2) = RGB(50, 250, 150)
                .CellForeColor(2) = &H666666
                '/* adjust for entire row
                '.BackColor = RGB(175, 200, 150)
                '.ForeColor = &H333333
            End If
        End With
    Next lPos

End Sub

Private Sub PutArray()
'/* forward struct pointer into library

On Error GoTo Handler

    If Not ArrayCheck(m_cListItems) Then
        ItemsArray ItemCount
    End If
    '/* copy struct pointer into list control
    If Not m_lPointer = 0 Then
        DestroyItems
        ResetTempArray
    End If
    CopyMemBr m_lPointer, ByVal VarPtrArray(m_cListItems), 4&
    ucHyperListNG1.StructPtr = m_lPointer
    '/* load the data struct
    ucHyperListNG1.LoadArray
    '/* set the item count, this will fire the callback
    '/* and populate the list
    ucHyperListNG1.SetItemCount UBound(m_cListItems) + 1
    
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

    CopyMemBr ByVal VarPtrArray(m_cListItems), 0&, 4&
    Erase m_cListItems
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
