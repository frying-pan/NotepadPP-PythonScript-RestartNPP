# PythonScript that restarts Notepad++ and restore the session (all opened files, caret, text selection)
# Tested with Notepad++ 7.8.2 64 bits, with PythonScript plugin 1.5.2
# on Windows 8.1 64 bits (NOT tested with Notepad++ 32 Bits but should be compatible)
# /!\ this file uses TABS for indent /!\ (better read as 4 chars wide tabs)

# /!\ WARNING : issues caused by the fact that the new Notepad++ instance starts before the current has closed
# * if some Preferences have been changed during the session, these changes will be lost
# * if the Notepad++ window has been resized/moved during the session, these changes will be lost
# * if the recent files list has changed during the session, these changes will be lost
# /!\ WARNING : the current instance will not close automatically if some modified workspaces have not be saved
# (i) Info : as for any Notepad++ restart the Undo/Redo history will be forgotten

# technical notice : this script will create a temporary session file (OS/user dependant location),
# and disable the current Notepad++ window for a few seconds before closing its process,
# so that it can delete the temporary session file *after* it has been read by the new Notepad++ instance

# *options values* : set to other STRING/INTEGER values if needed
s_npp_exe_filename			= "Notepad++.exe"	# Notepad++ executable name (without path) : in case it was renamed
i_s_wait_new_npp_process	= 5					# seconds to wait for the new instance to have started and read
												# the temporary session file : increase if computer is slow
# end of *options values*

# script declarations ******************************************************************************
from Npp import *

s_script_name = "Restart Notepad++"

class C_Restart_NPP():
	# class constructor
	def __init__(self, s_npp_exe_filename, i_s_wait_new_npp_process):
		import os, tempfile, subprocess, time
		self.os			= os
		self.tempfile	= tempfile
		self.subprocess	= subprocess
		self.time		= time

		from Perso__Lib_Window import C_Block_UnBlock_Input
		self.o_block_unblock_input_nppscintilla_wins = C_Block_UnBlock_Input()

		self.npp_exe_filename		= s_npp_exe_filename
		self.s_wait_new_npp_process	= i_s_wait_new_npp_process
		self.session_filefull = ""

	def Restart(self):
		def Try_Start():
			s_multinst		= "-multiInst"
			s_opensession	= "-openSession"
			s_temp_prefix	= "RestartNPP_Session_"
			s_temp_suffix	= ".tmp.npps"
			s_restartabort	= "Restart aborted"

			i_response = notepad.messageBox( \
							"Check if all 'Workspaces' are saved before restarting Notepad++" + "\n\n" + \
							"As a precautionary mesure, all files will now be saved, and then closed" + "\n" + \
							"(and you will be asked to save the changes if you want)" + "\n\n" + \
							"/!\\ WARNING : if you answer 'No' to a save file dialog" + "\n" + \
							"this(these) file(s) will be closed WITHOUT being saved /!\\" + "\n\n" + \
							"/!\\ WARNING : changes made during this session" + "\n" + \
							"to your 'Preferences', window position, and recent files list" + "\n" + \
							"will NOT be remembered after the restart" + "\n\n" + \
							"/!\\ WARNING : the current instance will NOT close automatically" + "\n" + \
							"if some modified workspaces have not be saved beforehand...", \
							s_script_name, MESSAGEBOXFLAGS.ICONEXCLAMATION + MESSAGEBOXFLAGS.OKCANCEL)
			if i_response != MESSAGEBOXFLAGS.RESULTOK:
				return False
			notepad.saveAllFiles()

			# create a temporary session file that will be cleaned by the current instance before closing
			b_has_error = True
			try:
				t_temp = self.tempfile.mkstemp(s_temp_suffix, s_temp_prefix, None)
			except:
				dummy = 0
			else:
				if not(t_temp is None):
					fd_session			= t_temp[0]
					s_session_filefull	= t_temp[1]
					if (not(fd_session is None) and not(s_session_filefull is None)):
						if not(s_session_filefull == ""): b_has_error = False
			if b_has_error:
				notepad.messageBox( \
					"Error creating temporary session file !" + "\n" + s_restartabort, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False

			self.session_filefull = s_session_filefull

			try:
				notepad.saveCurrentSession(s_session_filefull)
				self.os.close(fd_session)
			except:
				notepad.messageBox( \
					"Error saving temporary session file !" + "\n" + s_restartabort, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False

			# create a dummy document to avoid 'Exit on close the last tab' option to close Notepad++ with a closeAll command
			notepad.new()
			id_buffer = notepad.getCurrentBufferID()

			notepad.closeAllButCurrent()
			# do not allow restart if some modified documents are still open, especially in case the 'Session snapshot' option is enabled
			if len(notepad.getFiles()) > 1:
				# close dummy document
				notepad.activateBufferID(id_buffer)
				notepad.close()

				notepad.messageBox( \
					"Some files were not closed !" + "\n" + s_restartabort, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			if (editor.getModify() or editor1.getModify() or editor2.getModify()):
				# should not happen, since it would mean the dummy document has been modified...
				notepad.messageBox( \
					"Some files were not saved !" + "\n" + s_restartabort, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False

			s_npp_folfull = notepad.getNppDir()
			if ((s_npp_folfull is None) or s_npp_folfull == ""):
				notepad.messageBox( \
					"Notepad++ program folder can NOT be found !" + "\n" + s_restartabort, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			s_npp_filefull = s_npp_folfull.replace("\\", "\\") + "\\" + self.npp_exe_filename

			# start the new Notepad++ instance, and give the new instance some time to read the temporary session file
			try:
				self.subprocess.Popen([s_npp_filefull, s_multinst, s_opensession, s_session_filefull])
			except:
				notepad.loadSession(s_session_filefull)
				notepad.messageBox( \
					"Notepad++ restart failed !" + "\n" + s_npp_filefull + "\n\n" + s_restartabort, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False

			s_quitmsg = "This Notepad++ instance will auto-close after " + str(self.s_wait_new_npp_process) + " seconds..."
			print "\t" + ("*" * len(s_quitmsg))
			print "\t" + s_quitmsg
			print "\t" + ("*" * len(s_quitmsg))
			self.time.sleep(self.s_wait_new_npp_process)

			return True

		def Try_Re_Start():
			b_canclose = Try_Start()
			if not(self.session_filefull == ""):
				try:
					self.os.remove(self.session_filefull)
				except:
					dummy = 0
			if not(b_canclose):
				return False

			notepad.menuCommand(MENUCOMMAND.FILE_EXIT)			# -> IMPORTANT : should exit the process NOW if everything went OK
			return True

		self.o_block_unblock_input_nppscintilla_wins.GetWinsHandle()

		self.o_block_unblock_input_nppscintilla_wins.block()	# -> IMPORTANT : block Notepad++ window keyboard and mouse input
		b_restarted = Try_Re_Start()
		# if some modified workspaces were not saved, or anything went wrong, un-block the current instance
		self.o_block_unblock_input_nppscintilla_wins.unblock()	# -> IMPORTANT : un-block Notepad++ window keyboard and mouse input

		return b_restarted
# end of class

# script code **************************************************************************************
print "[" + s_script_name + " starts]"

# create an instance of the restart class, and populate object properties with script variables
o_restart_npp = C_Restart_NPP(s_npp_exe_filename, i_s_wait_new_npp_process)

b_restarted = o_restart_npp.Restart()
if not(b_restarted):
	print "\t" + "Restart aborted"
