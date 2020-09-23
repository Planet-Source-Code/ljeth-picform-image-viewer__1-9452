Attribute VB_Name = "functions"

'This type is used to store the state of the form
Type formCloseStatusType
  blnCloseAllFlag As Boolean
  blncmdPrevEnabled As Boolean
  blncmdNextEnabled As Boolean
  sfrmPicFormCaption As String
  blnlblPageEnabled As Boolean
  stxtPageText As String
  blntxtPageEnabled As Boolean
  blncmdPageEnabled As Boolean
  blncmdCloseAllEnabled As Boolean
  bescFlag As Byte
  iThumbIndex As Integer
End Type

'This type is for storing original
' image information
Type sizeType
  height As Long
  width As Long
End Type

Sub AutoSize(ByVal maxHeight As Integer, ByVal maxWidth As Integer, _
             newY As Long, newX As Long)
             
'This sub resizes images to desired dimensions,
'  maintaining the aspect ratio, and returns the new
'  sizes. maxHeight and maxWidth contain the desired
'  dimensions, newY and newX contain the image height
'  and width. The new sizes are returned using the
'  newY and newX vars themselves.
 
  Dim H2W_Ratio As Single, W2H_Ratio As Single
  
  'get the needed ratios
  H2W_Ratio = newY / newX
  W2H_Ratio = newX / newY
  
  'if image height exceeds desired height resize it
  If newX > maxWidth Then
    newX = maxWidth
    newY = Round(newX * H2W_Ratio)
  End If
  
  'if image width exceeds desired width resize it
  If newY > maxHeight Then
    newY = maxHeight
    newX = Round(newY * W2H_Ratio)
  End If
  
End Sub

Sub loadThumbs(thumbIndex As Integer, sizeArray() As sizeType, ByVal path As String)
  
'This sub loads thumbs from a folder.
  
  Dim i As Integer, j As Integer
  Dim newX As Long, newY As Long
  Dim fileName As String
  
  For i = 0 To (frmPicForm.gNumTotalThumbs - 1)
  
    j = i   'j used in next loop
    
    fileName = path & "\" & _
        frmPicForm.filImgList.List(thumbIndex + i)
    frmPicForm.imgThumb(i).Visible = False
    frmPicForm.imgThumb(i).Stretch = False
    
    On Error GoTo skipLoad
    
    'first load the image that's to be
    ' resized to a thumb
    frmPicForm.imgThumb(i).Picture = LoadPicture(fileName)
    
    'get image's original size
    newY = frmPicForm.imgThumb(i).height
    newX = frmPicForm.imgThumb(i).width
    
    'store image's original size for
    ' future use
    sizeArray(i).height = newY
    sizeArray(i).width = newX
    
    'resize image to thumb dimensions
    Call AutoSize(frmPicForm.gThumbDimension, frmPicForm.gThumbDimension, _
        newY, newX)
    
    'set the required image control properties
    frmPicForm.imgThumb(i).Stretch = True
    frmPicForm.imgThumb(i).height = newY
    frmPicForm.imgThumb(i).width = newX
    frmPicForm.imgThumb(i).Visible = True
    frmPicForm.imgThumb(i).Enabled = True
    frmPicForm.imgThumb(i).ToolTipText = fileName
  
skipLoad2:
    
    'don't go past the total no. of files in the folder
    If (thumbIndex + i) >= frmPicForm.filImgList.ListCount Then
      Exit For
    End If
    
  Next i
  
  'if no. of thumbs are less than what the page allows
  ' set remaining image controls to blank and disable them
  If j < (frmPicForm.gNumTotalThumbs - 1) Then
    For i = j To (frmPicForm.gNumTotalThumbs - 1)
      frmPicForm.imgThumb(i).Picture = LoadPicture("")
      frmPicForm.imgThumb(i).Enabled = False
    Next i
  End If
  
  GoTo endSub
    
skipLoad:
    'error 481 - invalid picture file
    'This means a corrupted image file or
    ' non-image file was encountered. Set
    ' the thumb to blank and disable it.
    If Err.Number = 481 Then
      frmPicForm.imgThumb(i).Picture = LoadPicture("")
      frmPicForm.imgThumb(i).Enabled = False
    End If
    
    'error has been handled so
    ' continue with the For loop.
    Resume skipLoad2
    
endSub:

End Sub

Function isFolder(folderName As String) As Boolean

'checks if a file name is an existing folder

  Dim fs
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  isFolder = fs.FolderExists(folderName)
  
End Function
