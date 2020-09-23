Attribute VB_Name = "mData"
Option Explicit

Public Type HLISubItm
    lIcon       As Long
    Text()      As String
End Type

Public Type HLIStc
    Item()      As String
    lIcon()     As Long
    SubItem()   As HLISubItm
End Type
