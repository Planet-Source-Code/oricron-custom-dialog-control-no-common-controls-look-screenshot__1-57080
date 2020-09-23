VERSION 5.00
Object = "*\A..\Dialog.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin OricronDialog.Dialog Dialog1 
      Left            =   120
      Top             =   1080
      _ExtentX        =   1164
      _ExtentY        =   847
   End
   Begin VB.CheckBox chk3 
      Caption         =   "Sizable"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CheckBox chk2 
      Caption         =   "Remember current path (puts you back into the same folder on reopening the dialog)"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   6495
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Multiselect (show open only!)"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "Dialog Title"
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   9495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Save"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Open"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Selected files:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ALL CAPTIONS in the dialog can be changed, just look into the code - I mean all! Including menus!"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Show dialog to open/select files
InitDialog
Dialog1.ShowOpen

Dim iCnt As Integer
Dim spl() As String

Me.List1.Clear

If Dialog1.FileName <> "" Then
    spl() = Split(Dialog1.FileName, """" & " " & """")

    For iCnt = 0 To UBound(spl)
        If Left(spl(iCnt), 1) = """" Then spl(iCnt) = Mid$(spl(iCnt), 2)
        If Right(spl(iCnt), 1) = """" Then spl(iCnt) = Left(spl(iCnt), Len(spl(iCnt)) - 1)
        
        Me.List1.AddItem spl(iCnt)
    Next iCnt
End If


End Sub

Private Sub Command2_Click()
'open dialog to save files
'just an exampel string - do not use spaces between definitions and always use "name|*.extention" captions!
InitDialog
Dialog1.DialogFilter = "m3u playlist|*.m3u|mp3 file|*.mp3"
Dialog1.ShowSave

End Sub

Private Sub Dialog1_SaveFileAs(FileName As String)
'This sub fires, when ShowSave is fired, and a filename is specified
MsgBox "Saving file: " & FileName

End Sub

Private Sub InitDialog()
'This sub initializes the variables for the dialog

'look at the captions in form main to see what hapends... Or just tr it:p
Dialog1.DialogTitle = Text1
Dialog1.MultiSelect = chk1.Value
Dialog1.RememberCurrentPath = chk2.Value
Dialog1.Sizable = chk3.Value

'Remember, all captions in dialog (menues, details names, etc. can be changed!

'just try, for exampel (uncoment the line below, and try to see the details of an empty drive)
'Dialog1.CAPTION_DEVICEUNAVAILABLE = "Dude, this drive doesn't work!"
'etc...

End Sub
