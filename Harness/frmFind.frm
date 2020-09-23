VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Column"
      Height          =   645
      Left            =   3465
      TabIndex        =   9
      Top             =   630
      Width           =   1680
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         ScaleHeight     =   330
         ScaleWidth      =   1545
         TabIndex        =   10
         Top             =   270
         Width           =   1545
         Begin VB.OptionButton optColumn 
            Caption         =   "2"
            Height          =   195
            Index           =   2
            Left            =   1035
            TabIndex        =   13
            Top             =   90
            Width           =   420
         End
         Begin VB.OptionButton optColumn 
            Caption         =   "0"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   12
            Top             =   90
            Value           =   -1  'True
            Width           =   420
         End
         Begin VB.OptionButton optColumn 
            Caption         =   "1"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   11
            Top             =   90
            Width           =   420
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      Height          =   645
      Left            =   1665
      TabIndex        =   5
      Top             =   630
      Width           =   1680
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         ScaleHeight     =   330
         ScaleWidth      =   1545
         TabIndex        =   6
         Top             =   270
         Width           =   1545
         Begin VB.OptionButton optDirection 
            Caption         =   "Down"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   8
            Top             =   90
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optDirection 
            Caption         =   "Up"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Top             =   90
            Width           =   600
         End
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4185
      TabIndex        =   4
      Top             =   90
      Width           =   1050
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Exact Match"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1035
      Width           =   1275
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Match Case"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   2
      Top             =   765
      Width           =   1275
   End
   Begin VB.TextBox txtFind 
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Top             =   135
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find What:"
      Height          =   210
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bNext     As Boolean
Private m_lParent   As Long


Public Property Get FindParent() As Long
    FindParent = m_lParent
End Property

Public Property Let FindParent(PropVal As Long)
    m_lParent = PropVal
End Property


Private Sub cmdFind_Click()

Dim lCol As Long

    If optColumn(0).Value Then
        lCol = 0
    ElseIf optColumn(1).Value Then
        lCol = 1
    ElseIf optColumn(2).Value Then
        lCol = 2
    End If
    
    If Len(txtFind.Text) = 0 Then Exit Sub
    Select Case m_lParent
    '/* cd
    Case 0
        frmCustomDraw.ucHyperListNG1.Find txtFind.Text, lCol, CBool(chkMatch(0).Value), _
            CBool(chkMatch(1).Value), optDirection(0).Value, m_bNext, True
    '/* hl
    Case 1
        frmHyperMode.ucHyperListNG1.Find txtFind.Text, lCol, CBool(chkMatch(0).Value), _
            CBool(chkMatch(1).Value), optDirection(0).Value, m_bNext, True
    '/* db
    Case 2
        frmDataBaseMode.ucHyperListNG1.Find txtFind.Text, lCol, CBool(chkMatch(0).Value), _
            CBool(chkMatch(1).Value), optDirection(0).Value, m_bNext, True
    End Select
    m_bNext = True

End Sub

Private Sub optColumn_Click(Index As Integer)
    m_bNext = False
End Sub

Private Sub optDirection_Click(Index As Integer)
    m_bNext = False
End Sub
