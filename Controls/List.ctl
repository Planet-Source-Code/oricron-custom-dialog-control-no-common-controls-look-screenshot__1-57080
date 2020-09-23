VERSION 5.00
Begin VB.UserControl oList 
   BackColor       =   &H007E5229&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   870
End
Attribute VB_Name = "oList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ListCount As Long
Dim Item() As String


Public Sub AddItem(Text As String)
If ListCount > 0 Then ReDim Preserve Item(UBound(Item) + 1)
Item(ListCount) = Text
ListCount = UBound(Item) + 1

End Sub

Public Function GetItem(index As Long) As String
If index < ListCount And index >= 0 Then
    GetItem = Item(index)
End If

End Function

Private Sub UserControl_Initialize()
ReDim Item(0)
ListCount = 0

End Sub

Private Sub UserControl_Resize()
UserControl.Width = 32 * Screen.TwipsPerPixelX
UserControl.Height = 32 * Screen.TwipsPerPixelY


End Sub

Public Sub Clear()
ReDim Item(0)
ListCount = 0

End Sub
