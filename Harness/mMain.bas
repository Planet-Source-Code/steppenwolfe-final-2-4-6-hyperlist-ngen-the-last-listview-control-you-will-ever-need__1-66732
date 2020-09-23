Attribute VB_Name = "mMain"
Option Explicit

Private Const ICC_LISTVIEW_CLASSES              As Long = &H1


Private Type tagINITCOMMONCONTROLSEX
    dwSize As Long
    dwICC As Long
End Type


Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub Main()

    InitComctl32
    frmCustomDraw.Show
    
End Sub


Private Function InitComctl32() As Boolean

Dim icc As tagINITCOMMONCONTROLSEX

On Error GoTo Handler
  
  icc.dwSize = Len(icc)
  icc.dwICC = ICC_LISTVIEW_CLASSES
  InitComctl32 = InitCommonControlsEx(icc)

On Error GoTo 0
Exit Function

Handler:
  InitCommonControls

End Function

