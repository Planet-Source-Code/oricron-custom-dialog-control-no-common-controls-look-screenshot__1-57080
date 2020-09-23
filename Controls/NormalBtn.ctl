VERSION 5.00
Begin VB.UserControl Button 
   BackColor       =   &H007E5229&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   Begin VB.Timer tmrToolTip 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   1080
   End
   Begin VB.Image imgP 
      Height          =   255
      Left            =   2040
      Top             =   2160
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgF 
      Height          =   255
      Left            =   1320
      Top             =   2160
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image ImgN 
      Height          =   255
      Left            =   600
      Top             =   2160
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------
'In and Out Mouse function is modiffied from a project
'called "A smartButton" found on PSC
'There is no author information in that project, so i
'do not know the author:-/
'-----------------------------------------------------

Option Explicit
Public hwnd As Long
Public TopRegion As Integer

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private ret As Long
Public Clicked As Boolean
Private FlagInside As Boolean

Dim ToolTipCnt As Integer
Public cEnabled As Boolean

Public Event EnableStateChange(Enabled As Boolean)

Public Event EnterFocus()
Public Event ExitFocus()
Public Event Click()
Public Event RightClick(Tag As String)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
Public Event MouseOut()
Public Event TotalMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Tag As String)
Public Event RaiseToolTip(Tag As String)
Public Event HideToolTip()

Public Tip As Integer
Public sTag As String
Public Položaj As Integer

Dim DisplayingTooltip As Boolean

Dim mode As String
Dim iTag(2) As String

Public MenuMode As Boolean

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long

Dim m_CurrentRGN As Long

Public Property Let RegionFile(FileName As String)
On Error GoTo handle
'tHIS PART IS USED TO CHANGE THE BUTTON SHAPE...
If FileName = "" Then Exit Property

Dim mBinary() As Byte, nCount As Long, nF
nCount = FileLen(FileName)

If nCount > 0 Then
    ReDim mBinary(nCount - 1)
    
    nF = FreeFile
    
    Open FileName For Binary As #nF
        Get #nF, 1, mBinary
    Close #nF
    
    m_CurrentRGN = ExtCreateRegion(ByVal 0&, nCount, mBinary(0))
End If

If Not m_CurrentRGN = 0 Then SetWindowRgn UserControl.hwnd, m_CurrentRGN, True

Exit Property

handle:

End Property

Public Function UserControl_MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FlagInside = False
    If X < 0 Or X > UserControl.Width / Screen.TwipsPerPixelX Or Y < Me.TopRegion Or Y > UserControl.Height / Screen.TwipsPerPixelY Then
        FlagInside = False ' set mouse is outside

        ToolTipCnt = 0
        DisplayingTooltip = False
        tmrToolTip.Enabled = False
        
        ret = ReleaseCapture()
        'Cls

        RaiseEvent HideToolTip

        If Clicked = False Then
            If mode <> "N" Then
                UserControl.Picture = ImgN.Picture
                mode = "N"
            End If
        End If
        RaiseEvent MouseOut
    Else
        If FlagInside = False Then
            FlagInside = True ' set mouse is inside
            ret = SetCapture(UserControl.hwnd)
            'DrawButtonUp
            If Clicked = False Then
                If mode <> "F" Then
                    UserControl.Picture = imgF.Picture
                    mode = "F"
                End If
            End If

            
        End If
    End If
End Function

Private Sub Image1_Click()

End Sub

Private Sub tmrToolTip_Timer()

If FlagInside = True And Clicked = False Then
    ToolTipCnt = ToolTipCnt + 1
    If ToolTipCnt > 15 And DisplayingTooltip = False Then
        DisplayingTooltip = True
        If Položaj = Tip Then
            RaiseEvent RaiseToolTip(iTag(0))
        Else
            RaiseEvent RaiseToolTip(iTag(Položaj + 1))
        End If
        tmrToolTip.Enabled = False
    End If
End If
End Sub

Private Sub UserControl_Click()
If cEnabled = True Then RaiseEvent Click

End Sub

Private Sub UserControl_EnterFocus()
RaiseEvent EnterFocus

End Sub

Private Sub UserControl_ExitFocus()
RaiseEvent ExitFocus

End Sub

Private Sub UserControl_GotFocus()
        If FlagInside = False Then
            FlagInside = True ' set mouse is inside
            ret = SetCapture(UserControl.hwnd)
            'DrawButtonUp
            If Clicked = False Then
                If mode <> "F" Then
                    UserControl.Picture = imgF.Picture
                    mode = "F"
                End If
            End If

            
        End If
End Sub

Private Sub UserControl_Initialize()
hwnd = UserControl.hwnd
cEnabled = True

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
UserControl.SetFocus
UserControl.Picture = imgP.Picture
Clicked = True
If DisplayingTooltip = True Then RaiseEvent HideToolTip
If MenuMode = True Then
    If cEnabled = True Then RaiseEvent RightClick(iTag(0))
    UserControl.Picture = ImgN.Picture
    Clicked = False
    FlagInside = False
    ret = ReleaseCapture()
    mode = "N"
Else
    If Button = vbRightButton Then
        If Tip > 0 Then
            If Položaj < Tip Then
                If cEnabled = True Then RaiseEvent RightClick(iTag(Položaj + 1))
            Else
                If cEnabled = True Then RaiseEvent RightClick(iTag(0))
            End If
        Else
            If cEnabled = True Then RaiseEvent RightClick(iTag(0))
        End If
        UserControl_MouseUp Button, Shift, X, Y
    End If
        If cEnabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y, iTag(Položaj))
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cEnabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y, iTag(Položaj))

If Button = 0 Then
    UserControl_MouseOut Button, Shift, X, Y
    
    If DisplayingTooltip = False Then tmrToolTip.Enabled = True
    
Else
    If X >= 0 And X <= UserControl.Width / Screen.TwipsPerPixelX And Y >= 0 And Y <= UserControl.Height / Screen.TwipsPerPixelY Then
        If mode <> "P" Then
            UserControl.Picture = imgP.Picture
            mode = "P"
        End If
    Else
        If mode <> "F" Then
            UserControl.Picture = imgF.Picture
            mode = "F"
        End If
    End If
    
End If

End Sub

Public Property Set NormalImage(Image As Picture)
Set ImgN.Picture = Image

UserControl.Width = ImgN.Width * Screen.TwipsPerPixelX
UserControl.Height = ImgN.Height * Screen.TwipsPerPixelY

UserControl.Picture = ImgN.Picture

End Property

Public Property Set FocusedImage(Image As Picture)
Set imgF.Picture = Image

End Property

Public Property Set PressedImage(Image As Picture)
Set imgP.Picture = Image

End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Clicked = False
mode = ""
If X >= 0 And X <= UserControl.Width / Screen.TwipsPerPixelX And Y >= 0 And Y <= UserControl.Height / Screen.TwipsPerPixelY Then
    If Button = vbLeftButton Then
        If Položaj < Tip Then
            Položaj = Položaj + 1
        Else
            Položaj = 0
        End If
    End If
    
    If cEnabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y, iTag(Položaj))

    sTag = iTag(Položaj)
    
    If mode <> "F" Then
        UserControl.Picture = imgF.Picture
        mode = "F"
    End If
    
    If Button = vbLeftButton Then
        
    End If
Else
    If mode <> "N" Then
        UserControl.Picture = ImgN.Picture
        mode = "N"
    End If
End If
If cEnabled = True Then RaiseEvent TotalMouseUp(Button, Shift, X, Y, iTag(Položaj))
    UserControl_MouseOut Button, Shift, X, Y

End Sub

Private Sub UserControl_Resize()
'UserControl.Width = ImgN.Width * Screen.TwipsPerPixelX
'UserControl.Height = ImgN.Height * Screen.TwipsPerPixelY

UserControl.Picture = ImgN.Picture
End Sub

Public Sub SetTag(Tag As String)
iTag(Položaj) = Tag
sTag = iTag(Položaj)

End Sub

Public Sub SetPoložaj(index)
If index <= Tip And index >= 0 Then
    Položaj = index
    UserControl.Picture = ImgN.Picture
    sTag = iTag(Položaj)
End If


End Sub

Public Function GetTag(index As Integer) As String
If index >= 0 And index <= Tip Then
    GetTag = iTag(index)
End If

End Function

Public Sub SetState(state As String)
If state = "N" Then
    UserControl.Picture = ImgN.Picture
    Clicked = False

ElseIf state = "H" Then
    UserControl.Picture = imgF.Picture
    Clicked = False
ElseIf state = "P" Then
    UserControl.Picture = imgP.Picture
    Clicked = True

End If
End Sub

Public Property Let Enabled(value As Boolean)
cEnabled = value
RaiseEvent EnableStateChange(value)

End Property

Public Sub ResetSize()
UserControl.Width = ImgN.Width * Screen.TwipsPerPixelX
UserControl.Height = ImgN.Height * Screen.TwipsPerPixelY

End Sub

