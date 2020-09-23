picForm version 1.0 source code

Thank you for downloading picform version 1.0. This program and its source code have been released to the public domain and may be used or modified in any way free of charge. All I ask is if you do modify it please mention that I originally created it and that it was originally called picForm. Kindly send me an email if and when you decide to release your modified version, including the address of your website so I may link to it.

The story of picForm is a simple one. I got bored and decided to write my very own image viewing utility. The reasons for it were, first, that I did not want to download a shareware version of someone else's complicated utility and then complain about it not having the features I wanted, and second, I wanted to know what it would take to write a program such as this. I decided to keep it simple. I named it picForm because that was the name I first thought of for the main form of the application, and also because I'm not a very good application namer.

I have decided to release picForm as freeware open source so that others may use it and improve upon it, and also due to my own lack of time, for which reason I have not been able to add some features. I have tried to comment it without making it too messy, and I hope the comments help. As a side note, while the interface is a simple one it's not a sight for sore eyes. I wanted to allow the maximum area to the main image and the thumbs, with just enough space for the most basic buttons. Hopefully someone should be able to come up with a more pleasing visual interface. A crucial aspect of picForm, however, and one of the main reasons for my writing it, is its convenience of use particularly with the keyboard shortcuts. I control it with only two fingers of my left hand (after initially dragging a folder onto it - see tips below). I have tried to implement a parallel set of shortcuts for the right hand using the arrow keys, but I have not yet succeeded due to a stubborn refusal on the part of VB to comply with my instructions.

Unfortunately, one thing I haven't been able to include is a help file. I confess I still don't know how to make one, although I would say picform hardly qualifies for it. If you would like to help me in this regard I would appreciate it. I've heard of a couple of help building utilities, but I wonder if its possible to do without any of those. Who knows, I might even end up writing one. You are welcome to email me any useful suggestions and/or comments. I may consider implementing them, if I have the time and motivation, in a future version of picForm. 

Well, I guess that's about it from me. I've included further information about picForm below. Have fun!


LJetH, author of picform

email: ljeth@angelfire.com
www: http://www.angelfire.com/pop/ljh/
----------------------------------------------------------------

Information about picForm

Features:

- Shows images previewed as thumbs side by side.
- Makes any image fit within the main view window.
- Drag and drop support for individual files or folders.
- Includes a panic button feature (press ESC twice).
- Includes tooltips and displays the number of "thumb pages"
  in the title bar.
- Simple, convenient interface (see tips in readme.txt file).


Mouse Control:

- Drag and drop files or folders onto the form.
- Right click on main image to clear it.


Keyboard shortcuts:

 L or Ctrl-L    : load thumbs
 P or Ctrl-P    : change path
 ESC            : Unload Image / thumbs
 A or up-arrow  : prev pic
 Z or dn-arrow  : next pic
 S or pg-up     : prev page of thumbs
 X or pg-dn     : next page of thumbs
 Alt-Q, Alt-F4  : Exit


Tips:

- It helps to have windows file explorer open
  next to the form, so you can easily drag and
  drop folders containing the images. I use the
  ALT+Tab windows shortcut. Click and hold down
  the mouse over a folder in explorer, and
  move the mouse just a bit to let windows 
  know you're going to drag it. Then hit 
  ALT+Tab, and drop the folder on picForm.

- Keep the second and third fingers of the left
  hand on the A and Z keys to move between
  thumbs, and similarly use the S and X keys
  to move between pages. With a little practice
  you should be able to do this comfortably
  without looking at the keyboard.

- Hit Tab, type a page number, and hit Enter to 
  jump to the selected page. The total number
  of pages is displayed in the title bar if the
  thumbs are loaded.

- Hit ESC once to clear the main image, and once
  more to clear the thumbs. Hit L to restore the
  hidden thumbs. This feature may be used as
  a panic button or boss key.


Known bugs:

- The up-arrow and down-arrow keys are not trapped
  by the Form_KeyDown event. I couldn't figure it
  out at the time of releasing this source. The result
  is that the keyboard shortcuts for those two keys
  don't work.

- Takes some time to load the thumbs. Although this
  is not a bug in itself, it may be improved by 
  writing a (complicated?) routine to display a 
  low-quality version of the image as a thumb 
  rather than using VB's image controls.

- Using this app to view all the images stored
  in the Internet Explorer cache folders may
  cause the system to hang. Although windows
  explorer displays IE's cache files, not all
  of them are "real" files. Some may be 
  special objects that are references to 
  other files. picForm tries to treat them 
  like regular files and probably runs into 
  a brick wall.


Some additional features I've thought of including:

- Show image info (height, width, size, name etc) in tooltip.
- Pre-load images to make the thumbs load faster.
- Add scroll bars as an option to scroll images.

----------------------------------------------------------------
