VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\prjHyperList.vbp"
Begin VB.Form frmDataBaseMode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HyperList 4.0 - DataBase Mode Example"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin HyperListUC.ucHyperListNG ucHyperListNG1 
      Height          =   4695
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8281
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
      ThemeColor      =   -1
      UseThemeColors  =   0   'False
      ViewMode        =   0
      XPColors        =   -1  'True
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
      Left            =   6570
      TabIndex        =   3
      Top             =   5130
      Width           =   1320
   End
   Begin VB.ComboBox cbStyles 
      Height          =   315
      Left            =   4500
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   5130
      Width           =   1680
   End
   Begin VB.CommandButton cmdBench 
      Caption         =   "HyperList Bench: Add 100 Million Items"
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
      Left            =   135
      TabIndex        =   0
      Top             =   5085
      Width           =   3570
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   7740
      Top             =   3915
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
            Picture         =   "frmDataBaseMode.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":57F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":AFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":124E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":14C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":1AF32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":1B24C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":1B56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":1B890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":1BBAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":1E35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":20B0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   7740
      Top             =   3375
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
            Picture         =   "frmDataBaseMode.frx":26870
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":26B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":26EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":271BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":274D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":29C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":29FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":2C756
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":2CA70
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":2CD8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":2D0A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataBaseMode.frx":2D3BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblResults 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   210
      Left            =   135
      TabIndex        =   1
      Top             =   5580
      Width           =   375
   End
End
Attribute VB_Name = "frmDataBaseMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'~ HyperList 2.4 Test Harness by John Underhill (Steppenwolfe)


Private m_bReportMode               As Boolean
Private m_lItemCount                As Long
Private cTiming                     As clsTiming


Private Property Let ItemCount(ByVal PropVal As Long)
    m_lItemCount = PropVal
End Property

Private Property Get ItemCount() As Long
    ItemCount = m_lItemCount
End Property


Private Function RandomNum(ByVal lBase As Long, ByVal lSpan As Long) As Long
    RandomNum = Int(Rnd() * lSpan) + lBase
End Function

Private Sub cmdFind_Click()
    frmFind.FindParent = 2
    frmFind.Show vbModeless, Me
End Sub

Private Sub ucHyperListNG1_eHIndirect(ByVal iItem As Long, _
                                      ByVal iSubItem As Long, _
                                      ByVal fMask As Long, _
                                      sText As String, _
                                      iImage As Long)

'/* Indirect callback method:
'/* if using an external database, this
'/* is where reecords would be passed
'/* by index into the list
'/* iItem would map to the record number
'/* iSubItem 0 is the list item, and subsequent
'/* would map to record fields intended for subitems
'/* the sText is byref, pass the item text with this var
'/* iImage could correspond to a field maintaining icon indece

'* Note to all the 'database gurus' who can't figure out a callback..
'* example of usage would be something like this:

'Dim i As Long

'    rs.Move iItem, 1
'    i = 0
'    Dim f As Variant
'    For Each f In rs.Fields
'        If i = iSubItem Then
'            sText = f.Value
'            Exit For
'        End If
'        i = i + 1
'    Next

    If m_bReportMode Then
        Select Case iSubItem
        '/* item
        Case 0
            sText = "Item " & Format$(iItem, "#,###,##0")
        '/* subitems
        Case 1
            sText = "Subitem 1"
        Case 2
            sText = "Subitem 2"
        End Select
    Else
        sText = "Item " & Format$(iItem, "#,###,##0")
    End If
    
    iImage = RandomNum(3, 5) '<- fetch the icon index from a field in table

End Sub

Public Sub Form_Load()

Dim lX  As Long
Dim lC  As Long

    Randomize
    Set cTiming = New clsTiming
    '/* instance list and timer
    Set cTiming = New clsTiming
    
    '/* apply list settings
    With ucHyperListNG1
        lX = .Width / Screen.TwipsPerPixelX
        .ListMode = eDatabase
        '/* large icons
        .InitImlLarge
        For lC = 1 To 11
            .ImlLargeAddIcon iml32.ListImages.Item(lC).Picture
        Next lC
        '/* header icons
        .InitImlHeader
        .ImlHeaderAddIcon iml16.ListImages.Item(1).Picture
        .ImlHeaderAddIcon iml16.ListImages.Item(2).Picture
        '/* small images
        .InitImlSmall ' 32, 32
        For lC = 1 To 11
            .ImlSmallAddIcon iml16.ListImages.Item(lC).Picture
        Next lC
        '/* set viewmode
        .ViewMode = StyleReport
        '/* add columns
        .ColumnAdd 0, "Item", lX / 3, [ColumnLeft]
        .ColumnAdd 1, "Sub 2", lX / 3, [ColumnLeft]
        .ColumnAdd 2, "Sub 3", lX / 3, [ColumnLeft]
        .ItemsSorted = True
        .SkinHeaders HeaderXP, &H753F17, &H94511F, &HC46A28, False
        .SkinScrollBars ScrollClassic, False
    End With
    
    '/* dimension the list
    ItemCount = 100000000
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
    ucHyperListNG1.SetItemCount m_lItemCount
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


'> Timing
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub Reset()
'/* reset timer
    cTiming.Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
'/* destroy the pointer

    Set cTiming = Nothing

End Sub
