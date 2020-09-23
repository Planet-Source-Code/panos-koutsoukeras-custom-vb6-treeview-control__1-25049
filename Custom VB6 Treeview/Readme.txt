Treeview.ocx Demo
 
A UserControl made with a VB6 Treeview control (MSCOMCTL.OCX - 6.00.8862)
with customized Background (Color, GradientRectHor, GradientRectHor
GradientTri, tiled Picture), Backcolor, Forecolor, Buttons, Tooltips etc.


Copyright © 2001 Panos Koutsoukeras
Company:  Inspired Creations
Web:      http://globalinspired.com
Mail:     software@globalinspired.com
Date:     16 July 2001


Credits:
Ben Baird, http://www.vbthunder.com
Brad Martinez, http://www.mvps.org
http://vbaccelerator.com/ (SsubTimer)


Limitations:
The UserControl does not repaint correctly if the ClipControls property of the Form1 has been set to False

Some of the Background options behave strangely, because of the Treeview control, but this is only a demo, showing some ways to draw on a VB6 Treeview background.

If Background = fvGrdntRectVer or fvGrdntTri or fvPicturedTiled (m_BackScroll = True) then the control
does not show Tooltips and Scrolls one at a time
