# RestartNPP

Allow restarting Notepad++, and restore the current session state (with PythonScript plugin installed)

Read the warnings in the confirmation dialog when executing the script

Tested with Notepad++ 7.8.2 64 bits, on Windows 8.1 64 bits (NOT tested with Notepad++ 32 Bits but should be compatible)

using PythonScript plugin 1.5.2 from https://github.com/bruderstein/PythonScript/releases/ (based on python 2.7).

Features :
  * auto-save, or ask where to save new files, before restarting Notepad++
  * restart Notepad++, and restore the session (all opened files, text selection, cursor position)
  * remember opened files, caret position and text selection

Warnings :
  * modified workspaces must be manually saved before restarting, otherwise the previous Notepad++ instance will not close
  * undo/redo history will NOT be remembered after the restart

# Why ?

* If you are writing and running a Python script 'MyScript.py' from Notepad++, and this script has an 'Import MyLib.py', and you have edited and saved 'MyLib.py'. The changes made in 'MyLib.py' will not be taken into account until you restart Notepad++ (probably because the Import(s) are cached)
* If you have set a hook or a callback via a Python script, you may be unable to remove it without restarting Notepad++
* Possibly other reasons...

# Install :

This script can be set as a restart button in the toolbar (folder below is for a local installation) : 

* copy the main FP_RestartNPP .py script file (and the needed library file, and starting from v2_0 the VBS file) in :

C:\Users\[username]\AppData\Roaming\Notepad++\plugins\config\PythonScript\scripts

* In Menu > Plugins > PythonScript > Configuration : with 'User Scripts' > add this script as 'Toolbar icons'

(you can choose an icon : which, for me, required to be a .bmp file, 16x16 being better, .ico files did not work)

Possibly restart Notepad++ to have the script icon appear on the toolbar

# Versions :

FP_RestartNPP_v1_0.py
  * requires the library FP__Lib_Window.py included in the same folder
  * 2 Warnings below : for version v1_0 only, this has been solved in version v2_0
  * Warnings : 'Preferences' changed during this session will NOT be remembered after the restart
  * Warnings : window position, and recent files list changed during this session will NOT be remembered after the restart

FP_RestartNPP_v2_0.py
changes :
  * now requires the VBS file : FP_RestartNPP.vbs included in the same folder
  * now remembers 'Preferences', window position, recent files list and other configuration settings across restart
  * uses a VBS script to wait for the previous instance to have closed, before starting the new instance
  * safer and easier handling of the restart operation

FP_RestartNPP_v2_1.py
changes :
  * script name changed from Perso_RestartNPP to FP_RestartNPP for easier identification
  * minor update
