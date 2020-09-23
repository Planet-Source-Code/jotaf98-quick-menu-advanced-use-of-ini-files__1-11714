
(QuickMenu by Jotaf98)
 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
(E-mail: jotaf98@hotmail.com)
 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯


(Info)
 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
QuickMenu is basically a tray icon that, when you
click on it, pops up a handy menu with some
shortcuts.

You don't need to modify the menu in the code, it's
all kept in an .ini file. You can even modify the
tray icon if you want!

When you run it trough the .exe, the Options dialog
will appear first. Read the "Options Dialog"
section for more details on it.

To run it without the Options dialog, you'll need
to include "-NoOptions" in the command line. I've 
made a .bat file that does exactly that, try
running it and you'll notice that only the tray
icon will appear. By creating a shortcut to it in
"Start Menu -> Programs -> Startup", QuickMenu
will load everytime you start Windows.

Note that the .scf files only work with Windows 98
or later (a friend told me that, I'm not 100%
sure). So "Show/Hide Desktop" will only work in
those versions of Windows.

Also, don't forget to go to
http://www.planet-source-code.com and vote for my
code, if you find it useful! It's in the Visual
Basic section, with the name "Quick Menu".



(Options Dialog)
 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
The Options dialog has two sections: one where you
can choose another icon for Quick Menu and another
one where you add, delete, modify and change the
order of your shortcuts.

To change the icon, click "Browse..." and choose an
icon. Remember, it will be resized to 16x16 pixels
and will have its colors reduced to 16 (at least it
happened to me, I have no clue as to solve this
problem). Clicking "Use Default" will restore the
default icon. There is a "preview" image so you can
see how it will look like in the System Tray.

You can edit a shortcut by selecting it from the
list. Select "Name" or "Command" from the combo box
to edit the selected shortcut's name or command
(see the "Commands" section to know what the
available commands are). You can then type directly
into the text box below the list and press "Enter"
to accept the changes.

To add a new shortcut, click the button with a plus
sign (+). To delete the selected shortcut, click
the button with the minus sign (-). To change the
selected shortcut's order in the menu, click the
buttons with the up/down arrows.



(Commands)
 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
This is a list with the commands you can use and
some examples of how you can use them.


[] Simple .exe file names (as long as they're in
   the Windows directory)
	(Eg.: notepad.exe)


[] Complete paths to files/programs
	(Eg.: C:\MyDocuments\ToDoList.txt)
	(Eg.: C:\Games\SomeGame\Play.exe)


[] .exe file names like described above, with
   command line arguments
	(Eg.: C:\VisualBasic\vb6.exe /sdi


[] Complete paths to folders
	(Eg.: C:\Games)


[] Internet pages
	(Eg.: http://www.planet-source-code.com)


[] E-mails (you'll need to include "mailto:" before
   the e-mail address, with ***no spaces*** )
	(Eg.: mailto:jotaf98@hotmail.com)

		^- See? No spaces between "mailto:"
		   and "jotaf98@hotmail.com" :)
