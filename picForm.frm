VERSION 5.00
Begin VB.Form frmPicForm 
   Caption         =   "Picture - Not loaded"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6660
   ScaleWidth      =   10050
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPage 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Choose page to jump to and hit Go. Title bar shows total pages"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "Go"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   195
      Width           =   420
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Unload main image"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "--->"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1080
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Next page of thumbs"
      Top             =   265
      Width           =   735
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<---"
      Enabled         =   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Previous page of thumbs"
      Top             =   265
      Width           =   735
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "Path"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Change folder path"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtMask 
      Height          =   285
      Left            =   7920
      TabIndex        =   2
      Text            =   "*.*"
      ToolTipText     =   "File mask for loading thumbnails"
      Top             =   240
      Width           =   735
   End
   Begin VB.FileListBox filImgList 
      Height          =   1260
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdCloseAll 
      Caption         =   "Close All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Unload thumbnails and main image"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "Load thumbnails from folder"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape shpHighlight 
      BorderColor     =   &H80000012&
      BorderWidth     =   2
      Height          =   825
      Left            =   75
      Shape           =   1  'Square
      Top             =   795
      Width           =   825
   End
   Begin VB.Label lblPage 
      Caption         =   "Page"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   11
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Top             =   5640
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   10
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   5640
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   9
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   8
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   7
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   6
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   5
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   4
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   3
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   2
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblMask 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      Height          =   195
      Left            =   8040
      TabIndex        =   11
      Top             =   0
      Width           =   390
   End
   Begin VB.Image imgThumb 
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   840
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   1920
      X2              =   9960
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   9960
      X2              =   9960
      Y1              =   600
      Y2              =   6600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   1920
      X2              =   1920
      Y1              =   600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   1920
      X2              =   9960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image imgPicture 
      Height          =   5775
      Left            =   2040
      OLEDropMode     =   1  'Manual
      Top             =   720
      Width           =   7815
   End
End
Attribute VB_Name = "frmPicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' picForm Picture Viewer
' Authored by: LJetH
' email: ljeth@angelfire.com
' www: http://www.angelfire.com/pop/ljh/
'
' July 1, 2000
'
'
'Mouse Control:
'
'Drag and drop files or folders onto form.
'Right click on main image to clear it.
'
'Keyboard shortcuts
'
'L or Ctrl-L    : load thumbs
'P or Ctrl-P    : change path
'ESC:             Unload Image / thumbs
'A or up-arrow  : prev pic
'Z or dn-arrow  : next pic
'S or pg-up     : prev page of thumbs
'X or pg-dn     : next page of thumbs
'Alt-Q, Alt-F4  : Exit

'******************************************************************************************

'Form constants & variables

Option Explicit

Const cAppCaption = "picForm v1.0"
Const cAppCaptionNotLoaded = " - Not loaded"

'stores current path
Public gPath As String

'height/width of thumbs
Public gThumbDimension As Integer

'no. of thumbs allowed on the form ( = one page)
Public gNumTotalThumbs As Integer

'keeps track of the current page of thumbs
' goes in multiples of gNumTotalThumbs
Public gThumbIndex As Integer

'index to keep track of highlight
Dim thumbNum As Byte

'keeps track of original height and width of
' image represented by thumbs
Dim sizeArray(11) As sizeType

'used to store state of form (state of buttons etc)
Dim formStatusRec As formCloseStatusType

Dim oldMask As String

'keeps track of how many times ESC is pressed.
' Press ESC once to clear main image, press
' twice to clear thumbs.
Dim escFlag As Byte

Dim numPages As Integer

'indicates whether keyboard shortcuts can be used.
' Shortcuts are disabled when textboxes are in use.
Dim blnKeyTrap As Boolean

'stores max dimensions of the main image window.
' Main image is resized to fit within this window.
Dim MaxImgHeight As Long, MaxImgWidth As Long

Private Sub cmdClose_Click()

  'clear the main image
  imgPicture.Picture = LoadPicture("")
  cmdClose.Enabled = False
  
  'set state of the ESC key
  escFlag = 1
  
End Sub

Private Sub cmdCloseAll_Click()

  Dim i As Integer
  
  'clear the image and thumbs
  imgPicture.Picture = LoadPicture("")
  For i = 0 To gNumTotalThumbs - 1
    imgThumb(i).Enabled = False
    imgThumb(i).Visible = False
  Next i
  
  'store current status of form
  With formStatusRec
    .blnCloseAllFlag = True
    .blncmdPrevEnabled = cmdPrev.Enabled
    .blncmdNextEnabled = cmdNext.Enabled
    .sfrmPicFormCaption = frmPicForm.Caption
    .blnlblPageEnabled = lblPage.Enabled
    .stxtPageText = txtPage.Text
    .blntxtPageEnabled = txtPage.Enabled
    .blncmdPageEnabled = cmdPage.Enabled
    .blncmdCloseAllEnabled = cmdCloseAll.Enabled
    .bescFlag = escFlag
    .iThumbIndex = gThumbIndex
  End With
  
  'change form status
  cmdPrev.Enabled = False
  cmdNext.Enabled = False
  frmPicForm.Caption = cAppCaption & cAppCaptionNotLoaded
  lblPage.Enabled = False
  txtPage.Text = ""
  txtPage.Enabled = False
  cmdPage.Enabled = False
  cmdCloseAll.Enabled = False
  escFlag = 0
  thumbNum = 0
  highlightThumb thumbNum
  
End Sub

Private Sub cmdExit_Click()

  Unload Me
  
End Sub

Private Sub cmdLoad_Click()

'The load button either loads a new set of thumbs or unhides
' temporarily hidden thumbs.
'
'I have implemented the ESC key to hide the thumbs/image from view.
' blnCloseAllFlag, which belongs to the formCloseStatusType type,
' indicates whether the thumbs are temporarily hidden
' or whether they have been unloaded. True indicates they are hidden.
'
'When the ESC key is pressed twice, all thumbs are
'  hidden, i.e. they are made invisible but not unloaded,
'  and the state of the form (buttons/text etc.) is changed
'  to give the appearance of a hidden state. When the user
'  presses the Load button, the form unhides the thumbs if
'  they are hidden, or else it loads thumbs from disk.
'
'I made this feature so you can quickly hide the images
' you're viewing if someone unexpectedly intrudes and you
' don't want them to see the images. It's like a panic
' button or a boss key and quite handy in certain situations.
  
  
  Dim i As Integer
  
  'If thumbs are loaded, unload them before loading a new set of thumbs.
  If Not formStatusRec.blnCloseAllFlag Then
    For i = 0 To gNumTotalThumbs - 1
      imgThumb(i).Picture = LoadPicture("")
    Next i
  End If
  
  'get list of all files in current dir
  filImgList.Refresh

  'set indicator to first page of thumbs
  gThumbIndex = gNumTotalThumbs
  
  'if thumbs are hidden unhide them
  ' and restore the unhidden state of the form
  If formStatusRec.blnCloseAllFlag Then
  
    MousePointer = vbHourglass
    
    'restore form state
    With formStatusRec
      cmdPrev.Enabled = .blncmdPrevEnabled
      cmdNext.Enabled = .blncmdNextEnabled
      frmPicForm.Caption = .sfrmPicFormCaption
      lblPage.Enabled = .blnlblPageEnabled
      txtPage.Text = .stxtPageText
      txtPage.Enabled = .blntxtPageEnabled
      cmdPage.Enabled = .blncmdPageEnabled
      cmdCloseAll.Enabled = .blncmdCloseAllEnabled
      escFlag = .bescFlag
      gThumbIndex = .iThumbIndex
    End With
    
    'make thumbs visible
    For i = 0 To gNumTotalThumbs - 1
      imgThumb(i).Visible = True
      imgThumb(i).Enabled = True
    Next i
    
    MousePointer = vbDefault
    
  Else
    'load fresh thumbs from disk
  
    'the previous and next buttons call the
    ' thumb-loading sub. They have to be
    ' enabled to work.
    cmdPrev.Enabled = True
    cmdPrev_Click
  
  End If
  
  'disable the 'next' button if there is only one page
  If filImgList.ListCount < gNumTotalThumbs Then
    cmdNext.Enabled = False
  End If
  
  'enable other form elements and reset variables
  ' to starting values
  lblPage.Enabled = True
  txtPage.Enabled = True
  txtPage.Text = (gThumbIndex \ 12) + 1
  cmdPage.Enabled = True
  cmdCloseAll.Enabled = True
  escFlag = 1
  
  'get total no. of pages
  ' and display it in the form caption
  i = (filImgList.ListCount \ gNumTotalThumbs)
  If (filImgList.ListCount Mod gNumTotalThumbs) <> 0 Then i = i + 1
  frmPicForm.Caption = cAppCaption & " - " & i & " pages"
  
  'reset other variables
  formStatusRec.blnCloseAllFlag = False
  numPages = i
  blnKeyTrap = True
  thumbNum = 0
  
  'set highlight to first thumb
  highlightThumb thumbNum
  
End Sub

Private Sub cmdNext_Click()

  'exit if disabled
  If Not cmdNext.Enabled Then Exit Sub

  MousePointer = vbHourglass
  
  'flip to next page of thumbs
  gThumbIndex = gThumbIndex + gNumTotalThumbs
  txtPage.Text = (gThumbIndex \ 12) + 1
  
  'disable 'next' button if on the last page
  If (gThumbIndex + gNumTotalThumbs) >= filImgList.ListCount Then
    cmdNext.Enabled = False
  End If
  
  'call the thumb-loading sub
  Call loadThumbs(gThumbIndex, sizeArray(), gPath)
  
  'set button and highlight status
  cmdPrev.Enabled = True
  thumbNum = 0
  highlightThumb thumbNum
  
  MousePointer = vbDefault
  
End Sub

Private Sub cmdPage_Click()

  Dim i As Integer, j As Integer
  
  MousePointer = vbHourglass
  
  'get page no. to jump to
  i = Val(txtPage.Text)
  
  'validate it to see if it is an
  ' acceptable value.
  If i >= 1 And i <= numPages Then
  
    'get the no. of the first thumb on the
    ' page to jump to
    j = (i - 1) * gNumTotalThumbs
    
    'if already at page do nothing
    If gThumbIndex = j Then
      GoTo endSub
    Else
      'set page indicator
      gThumbIndex = j
    End If
    
    'set state of buttons
    ' if on 1st page, enable only 'next'
    ' if in between, enable both
    ' else enable only 'prev'
    If i = 1 Then
      cmdPrev.Enabled = False
      If numPages > 1 Then cmdNext.Enabled = True
    ElseIf i = numPages Then
      cmdNext.Enabled = False
      If numPages > 1 Then cmdPrev.Enabled = True
    Else
      cmdPrev.Enabled = True
      cmdNext.Enabled = True
    End If
    
    'call the thumb loader
    Call loadThumbs(gThumbIndex, sizeArray(), gPath)
  End If
  
  'show page no. in text box
  txtPage.Text = (gThumbIndex \ 12) + 1
  cmdPage.SetFocus
  
endSub:
  
  MousePointer = vbDefault
  
End Sub

Private Sub cmdPath_Click()

'At the time of writing this utility I didn't stop to
' think if I should invoke the windows "browse for folder"
' dialog to let the user choose a path.
' Instead, I created a work-around using my own path browser
' form and global variables to pass the path information
' back to this main form. As it turns out, however, my
' path browser is better suited as it remembers
' the current path, whereas the windows path browser
' forces you to start from the "My Computer" root
' every time and then navigate to your desired path.
' I should have written a path browser class to encapsulate
' my path browser form, which makes the code cleaner
' and is the correct way to do this, but I was busy with
' other stuff and didn't really have the time.

  
  'load path browser form
  Load frmPath
  frmPath.Show 1
  
End Sub

Private Sub cmdPrev_Click()

'This is quite similar to the cmdNext_Click sub

  Dim i As Integer

  'exit if not enabled
  If Not cmdPrev.Enabled Then Exit Sub
  
  MousePointer = vbHourglass
  
  'flip to prev page of thumbs and update text box
  gThumbIndex = gThumbIndex - gNumTotalThumbs
  txtPage.Text = (gThumbIndex \ 12) + 1
  
  'disable 'prev' button if on the first page
  If gThumbIndex < gNumTotalThumbs Then
    cmdPrev.Enabled = False
  End If
  
  'load thumbs
  Call loadThumbs(gThumbIndex, sizeArray(), gPath)
  
  'set button state and highlight the first thumb
  cmdNext.Enabled = True
  thumbNum = 0
  highlightThumb thumbNum
  
  MousePointer = vbDefault
  
End Sub

Private Sub Form_Click()

  cmdLoad.SetFocus
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  'call the keyboard shortcut sub to take action
  myFormKeyDown KeyCode, Shift, thumbNum
  
End Sub

Private Sub Form_Load()

'This sub mainly sets startup variables and
' properties
  
  'set some form variables/properties
  MaxImgHeight = imgPicture.height
  MaxImgWidth = imgPicture.width
  Line1.X2 = width - 165
  Line2.Y2 = height - 465
  Line3.X1 = Line1.X2
  Line3.X2 = Line1.X2
  Line3.Y2 = Line2.Y2
  Line4.X2 = Line1.X2
  Line4.Y1 = Line2.Y2
  Line4.Y2 = Line2.Y2
  
  'get current path
  gPath = CurDir
  filImgList.path = gPath
  
  'set other variables
  frmPicForm.Caption = cAppCaption & cAppCaptionNotLoaded
  gThumbDimension = imgThumb(0).height  'thumb-holders are square
  gNumTotalThumbs = 12
  oldMask = "*.*"
  escFlag = 0
  blnKeyTrap = True
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'call the mouse_down sub
  imgPicture_MouseDown Button, Shift, X, Y
  
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'call the dragdrop sub
  imgPicture_OLEDragDrop Data, Effect, Button, Shift, X, Y
  
End Sub

Private Sub Form_Resize()

'This sub resizes the image and the image window
' borders when the window is resized.

  Dim newX As Long, newY As Long
  
  'exit if form is being minimized
  If frmPicForm.WindowState = 1 Then GoTo endSub
  
  'set properties for the image borders
  Line1.X2 = width - 165
  Line2.Y2 = height - 465
  Line3.X1 = Line1.X2
  Line3.X2 = Line1.X2
  Line3.Y2 = Line2.Y2
  Line4.X2 = Line1.X2
  Line4.Y1 = Line2.Y2
  Line4.Y2 = Line2.Y2
  
  'keep minimum size fixed
  If height < 7065 Then   'manually measured value
    height = 7065
  End If
  
  If width < 10170 Then   'manually measured value
    width = 10170
  End If
  
  MaxImgHeight = height - 1290  'manually measured value
  MaxImgWidth = width - 2355    'manually measured value
  
  
  'get new window size
  newY = imgPicture.height
  newX = imgPicture.width
  
  'call the resize sub to get the new size
  ' for the main image
  Call AutoSize(MaxImgHeight, MaxImgWidth, newY, newX)
  
  'set the new size for the image
  imgPicture.height = newY
  imgPicture.width = newX

endSub:

End Sub

Private Sub imgPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'clear image if right mouse button clicked
  If Button = vbKeyRButton Then
    cmdClose_Click
  End If

End Sub

Private Sub imgPicture_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Provides ole dragdrop support so that the following
' objects can be dropped on the form (usually from
' windows file explorer):
'
'An image file
'A folder
'

  Dim newY As Long, newX As Long
  
  On Error GoTo errorHandler
    
  'clear the hide flag and main image
  formStatusRec.blnCloseAllFlag = False
  imgPicture.Visible = False
  
  'check if a file or folder was dropped
  If Data.GetFormat(vbCFFiles) Then
  
    'if folder was dropped
    If isFolder(Data.Files(1)) Then
      gPath = Data.Files(1)
      filImgList.path = gPath
      cmdLoad_Click
      GoTo endSub
    End If
  
    'file was dropped...
    
    'load the picture
    imgPicture.Stretch = False
    imgPicture.Picture = LoadPicture(Data.Files(1))
    
    'get the size of pic
    newY = imgPicture.height
    newX = imgPicture.width
    
    'if the pic doesn't fit within the image
    ' window resize it
    If newY > MaxImgHeight Or newX > MaxImgWidth Then
      Call AutoSize(MaxImgHeight, MaxImgWidth, newY, newX)
      imgPicture.Stretch = True
      imgPicture.height = newY
      imgPicture.width = newX
    End If
    
    'set necessary properties
    imgPicture.ToolTipText = Data.Files(1)
    cmdClose.Enabled = True
    escFlag = 2
  
  End If

errorHandler:
  imgPicture.Visible = True
  
endSub:

End Sub

Private Sub imgThumb_Click(index As Integer)

'This sub is called when a thumb is clicked
' but it's called from other places as well.
' Hence it checks to make sure that certain
' conditions are satisfied before being executed.

  Dim newX As Long, newY As Long
  Dim X As Long, Y As Long
  
  'make sure index is valid (when called from another sub)
  If index > gNumTotalThumbs - 1 Or index < 0 Then Exit Sub
  thumbNum = index
  
  'exit if the user tries to move to an empty thumb
  ' with the keyboard
  If Not imgThumb(index).Enabled Then Exit Sub
  
  'get the image info so that image can be resized
  ' to fit within the main image window
  newY = sizeArray(index).height
  newX = sizeArray(index).width
  
  'resize to fit in window
  Call AutoSize(MaxImgHeight, MaxImgWidth, newY, newX)
  
  'set the appropriate properties to display image
  ' and highlight the thumb
  imgPicture.Picture = imgThumb(index).Picture
  imgPicture.Stretch = True
  imgPicture.height = newY
  imgPicture.width = newX
  imgPicture.Visible = True
  imgPicture.ToolTipText = imgThumb(index).ToolTipText
  cmdClose.Enabled = True
  escFlag = 2
  highlightThumb index
  cmdLoad.SetFocus
  
End Sub

Private Sub imgThumb_OLEDragDrop(index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'call the dragdrop sub
  imgPicture_OLEDragDrop Data, Effect, Button, Shift, X, Y

End Sub

Private Sub txtMask_GotFocus()

  'select the text
  txtMask.SelStart = 0
  txtMask.SelLength = Len(txtMask.Text)
  
  'disable keyboard shortcut handling
  blnKeyTrap = False
  
End Sub

Private Sub txtMask_KeyPress(KeyAscii As Integer)

  Select Case KeyAscii
  
  Case 13
    GoTo start
    
  Case 27
    'press ESC
    blnKeyTrap = True
    cmdLoad.SetFocus
    
  Case Else
    GoTo endSub
    
  End Select
  
start:

  On Error GoTo setOldMask
  filImgList.Pattern = txtMask.Text
  oldMask = txtMask.Text
  
  'if mask is changed then disable prev & next buttons
  cmdPrev.Enabled = False
  cmdNext.Enabled = False
  
  cmdLoad.SetFocus
  GoTo endSub
  
setOldMask:
  'set original contents
  txtMask.Text = ""
  txtMask.SelText = oldMask

endSub:

End Sub

Private Sub txtMask_LostFocus()
  
  'start trapping keyboard events
  blnKeyTrap = True
  
End Sub

Private Sub txtPage_GotFocus()

  'select the text
  txtPage.SelStart = 0
  txtPage.SelLength = Len(txtPage.Text)
  
  'disable keyboard shortcut handling
  blnKeyTrap = False
  
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
  
  'click go button on enter
  Select Case KeyAscii
  Case 13
    Call cmdPage_Click
  End Select
  
End Sub

Private Sub txtPage_LostFocus()

  'start trapping keyboard events
  blnKeyTrap = True

End Sub

Private Sub myFormKeyDown(KeyCode As Integer, Shift As Integer, ByVal thumbNum As Byte)

'The purpose of this sub is to handle recursive calls
' to itself. This is required in a folder containing
' both image and non-image files. The image thumbs are
' displayed whereas the non-images are shown as blank
' disabled thumbs. If the user moves to a disabled thumb
' with the keyboard, this sub calls itself recursively
' until it encounters a valid thumb and then displays it.
'
'This scenario occurs if the file mask is set to include
' non-image files, eg: *.*
' To avoid it set a file mask to include only image files eg: *.jpg;*.gif;*.bmp


  Dim ShiftKey As Integer
  
  'ignore keyboard events if this
  ' flag is not set - used to
  ' disable key-events within
  ' textboxes.
  If Not blnKeyTrap Then Exit Sub
  
  'get only the first 3 bits
  ShiftKey = Shift And 7
  
  'check key pressed
  Select Case ShiftKey
    
    'Shift key pressed if first bit is set
    Case 1
    
    'Ctrl key pressed if second bit is set
    Case 2
      Select Case KeyCode
      Case vbKeyL
        cmdLoad_Click
      Case vbKeyP
        cmdPath_Click
      End Select
    
    'Alt keys pressed if third bit is set
    Case 4
      If KeyCode = vbKeyF4 Or KeyCode = vbKeyQ Then
        cmdExit_Click
      End If
    
    Case Else
      Select Case KeyCode
      
      'ESC was pressed
      Case 27
        Select Case escFlag
        Case 1
          Call cmdCloseAll_Click
          escFlag = 0
        Case 2
          cmdClose_Click
          escFlag = 1
        End Select
        
      'pg up or S
      Case vbKeyPageUp, vbKeyS
        cmdPrev_Click
        
      'pg dn or X
      Case vbKeyPageDown, vbKeyX
        cmdNext_Click
        
      'L
      Case vbKeyL
        cmdLoad_Click
      
      'P
      Case vbKeyP
        cmdPath_Click
        
      'A or up arrow
      ' This routine handles the
      ' highlight moving aspect.
      'thumbNum keeps track of the
      ' highlight's current position.
      'Same principle used below
      ' for the Z key.
      Case vbKeyA, vbKeyUp
        'if at the first thumb
        ' of the page
        If thumbNum = 0 Then
          ' and if a prev page exists
          If cmdPrev.Enabled Then
            ' go to prev page
            cmdPrev_Click
            imgThumb_Click (gNumTotalThumbs - 1)
          End If
        Else
          'if prev thumb is empty or disabled
          ' move another step back by calling
          ' this sub as if the 'A' key was
          ' pressed again. This is where
          ' recursion is useful.
          'Keep moving back until a valid thumb
          ' is encountered.
          If IsEmpty(imgThumb(thumbNum - 1)) Or Not imgThumb(thumbNum - 1).Enabled Then
            myFormKeyDown 65, Shift, thumbNum - 1
          Else
            'Display the valid thumb
            imgThumb_Click (thumbNum - 1)
          End If
        End If
        
      'Z or down arrow
      'same principle applied
      ' as the 'A' key
      Case vbKeyZ, vbKeyDown
        If thumbNum = gNumTotalThumbs - 1 Then
          If cmdNext.Enabled Then
            cmdNext_Click
            imgThumb_Click (0)
          End If
        Else
          If IsEmpty(imgThumb(thumbNum + 1)) Or Not imgThumb(thumbNum + 1).Enabled Then
            myFormKeyDown 90, Shift, thumbNum + 1
          Else
            imgThumb_Click (thumbNum + 1)
          End If
        End If
      
      End Select
      
  End Select

End Sub

Private Sub highlightThumb(ByVal index As Integer)

  'move the highlight to the required thumb
  shpHighlight.Visible = False
  shpHighlight.Top = imgThumb(index).Top - 45
  shpHighlight.Left = imgThumb(index).Left - 45
  shpHighlight.Visible = True

End Sub
