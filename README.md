# NotepadPP-PythonScript-RestartNPP

Allow restarting Notepad++, preserving the current session state (with PythonScript plugin installed)

Read carefully the warnings at the beginning of the main script file, and in the dialog when executing the script


Tested with Notepad++ 7.8.2 64 bits, with PythonScript plugin 1.5.2,

on Windows 8.1 64 bits (NOT tested with Notepad++ 32 Bits but should be compatible)


Features :
  * restart Notepad++ (assuming the executable is still named 'Notepad++.exe' otherwise the option must be set to another filename at the beginning of the script
  * remember opened files, caret position and text selection
  
Warnings :
  * modified workspaces must be manually saved before restarting, otherwise the previous Notepad++ instance will not close
  * 'Preferences' changed during this session will NOT be remembered after the restart
  * window position, and recent files list changed during this session will NOT be remembered after the restart
  * undo/redo history will NOT be remembered after the restart
  
# Why ?

* If you are writing and running a Python script 'MyScript.py' from Notepad++, and this script has an 'Import MyLib.py', and you have edited and saved 'MyLib.py'. The changes made in 'MyLib.py' will not be taken into account until you restart Notepad++ (probably because the Import(s) are cached)
* If you have set a hook or a callback via a Python script, you may be unable to remove it without restarting Notepad++
* Probably other reasons...

# Install :

This script can be set as a restart button in the toolbar (folder below is for a local installation) : 

* copy the main Perso_RestartNPP .py script file (and the needed library file) in :

C:\Users\[username]\AppData\Roaming\Notepad++\plugins\config\PythonScript\scripts

* In Menu > Plugins > PythonScript > Configuration : with 'User Scripts' > add this script as 'Toolbar icons'

(you can choose an icon : which, for me, required to be a .bmp file, 16x16 being better, .ico files did not work)

Possibly restart Notepad++ to have the script icon appear on the toolbar

# Versions :

Perso_RestartNPP_v1_0.py, (requires Perso__Lib_Window.py included in the same folder)
