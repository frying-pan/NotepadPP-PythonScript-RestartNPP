# PythonScript that restarts Notepad++ and restore the session (all opened files, caret, text selection)

# Tested with Notepad++ 7.8.2 64 bits, with PythonScript plugin 1.5.2
# on Windows 8.1 64 bits (NOT tested with Notepad++ 32 Bits but should be compatible)
# /!\ this file uses TABS for indent /!\ (better read as 4 chars wide tabs)

# technical notice :
# (helper VBS script file must have the same name as this script, with a .vbs extension instead of .py)
# 1. the PY script will create a temporary session file (OS/user dependant location),
# and will then run the VBS script before closing the current Notepad++ instance
# 2. once the current has closed, the VBS script will start a new Notepad++ instance,
# and then clean-up the temporary session file after a 'i_s_wait_npp_start' seconds delay
# 3. in the VBS script, a 1 second delay is given between current instance closing and new instance start
# (*might* be needed for a proper saving of the config.xml, containing Notepad++ 'Preferences')

# *options values* *********************************************************************************
# set options to other INTEGER > 0 if needed (default : i_s_wait_npp_close=10, i_s_wait_npp_start=30)
i_s_wait_npp_close	= 10	# maximum seconds to wait for the current instance to have closed (before starting new instance)
i_s_wait_npp_start	= 30	# seconds to wait for the new instance to have read the temporary session file (before deleting this file)
							# increase these delays if your computer is (very) slow to close/restart Notepad++

# script declarations ******************************************************************************
from Npp import *

s_script_name = "Restart Notepad++"

s_controleditscript = "(hold CTRL while starting the script to edit its options in Notepad++)"

class C_Restart_NPP():
	# class constructor
	def __init__(self, i_s_wait_npp_close, i_s_wait_npp_start):
		import ctypes
		import sys, os, platform, tempfile, subprocess
		self.sys		= sys
		self.os			= os
		self.platform	= platform
		self.tempfile	= tempfile
		self.subprocess	= subprocess

		self.GetCurrentProcessId = ctypes.windll.kernel32.GetCurrentProcessId

		from Perso__Lib_Window import C_Block_UnBlock_Input
		self.o_block_unblock_input_nppscintilla_wins = C_Block_UnBlock_Input()

		self.s_wait_npp_close	= i_s_wait_npp_close
		self.s_wait_npp_start	= i_s_wait_npp_start
		self.rest_session_filefull = None

	def Restart(self):
		def Try_ReStart():
			s_restartaborted	= "Restart aborted"
			s_multinst		= "-multiInst"
			s_opensession	= "-openSession"
			s_temp_prefix	= "RestartNPP_Session_"
			s_temp_suffix	= ".tmp.npps"
			s_os_type_win			= "Windows"
			s_wscript_exe_filefull	= "C:\\Windows\\System32\\WScript.exe"
			s_vbs_ext				= ".vbs"

			if self.platform.system().lower() != s_os_type_win.lower():
				notepad.messageBox("This script is only intended for Windows systems" + "\n\n" + s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONEXCLAMATION)
				return False

			i_response = notepad.messageBox( \
				"Check if all 'Workspaces' are saved before restarting Notepad++" + "\n\n" + \
				"As a precaution all files will now be SAVED, and then CLOSED" + "\n" + \
				"(and you will be asked to save the changes if you want)" + "\n\n" + \
				"/!\\ WARNING :" + "\n" + \
				"if you answer 'No' to a file save dialog, this file will be closed" + "\n" + \
				"WITHOUT being saved, even with 'Session Snapshot' enabled...", \
				s_script_name, MESSAGEBOXFLAGS.ICONEXCLAMATION + MESSAGEBOXFLAGS.OKCANCEL)
			if i_response != MESSAGEBOXFLAGS.RESULTOK:
				return False
			notepad.saveAllFiles()

			# create a temporary session file that will be cleaned by the VBS script
			b_has_error = True
			try:
				t_temp = self.tempfile.mkstemp(s_temp_suffix, s_temp_prefix, None)
			except:
				pass
			else:
				if not(t_temp is None):
					fd_session_file		= t_temp[0]
					s_session_filefull	= t_temp[1]
					if (not(fd_session_file is None) and not(s_session_filefull is None)):
						if s_session_filefull != "": b_has_error = False
			if b_has_error:
				notepad.messageBox("Error creating temporary session file" + "\n\n" + s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False

			# save current session (opened files) to the temporary session file
			try:
				notepad.saveCurrentSession(s_session_filefull)
				self.os.close(fd_session_file)
			except:
				try:
					self.os.close(fd_session_file)
					self.os.remove(s_session_filefull)
				except:
					pass
				notepad.messageBox( \
					"Error saving temporary session file" + "\n" + s_session_filefull + "\n\n" + \
					s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			self.rest_session_filefull = s_session_filefull

			# create a dummy document to avoid 'Exit on close the last tab' option to close NPP with a closeAll command
			notepad.new()
			bufferid_dummy = notepad.getCurrentBufferID()
			notepad.closeAllButCurrent()

			# can happen if the user canceled a file save dialog
			if len(notepad.getFiles()) > 1:
				# close dummy document
				notepad.activateBufferID(bufferid_dummy)
				notepad.close()

				notepad.messageBox("Some files were not closed" + "\n\n" + s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			# should not happen, since it would mean the dummy document has been modified...
			if (editor.getModify() or editor1.getModify() or editor2.getModify()):
				notepad.messageBox("Some files were not saved" + "\n\n" + s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False

			# get NPP and current script file full paths
			# self.sys.argv[0] is a relative file path that can be just the file name
			s_npp_filerel = str(self.sys.argv[0])
			s_npp_filename = s_npp_filerel[s_npp_filerel.rfind("\\") + 1:]
			if s_npp_filename == "":
				notepad.messageBox("Notepad++ file name can NOT be identified" + "\n\n" + s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			# notepad.getNppDir() should be a full folder path
			s_nppdir = notepad.getNppDir()
			if ((s_nppdir is None) or s_nppdir == ""):
				notepad.messageBox( \
					"Notepad++ program folder can NOT be located" + "\n" + s_npp_filename + "\n\n" + \
					s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			s_npp_filefull = s_nppdir + "\\" + s_npp_filename

			# self.os.path.realpath(__file__) should be a full file path
			s_script_filefull = self.os.path.realpath(__file__)
			if ((s_script_filefull is None) or s_script_filefull == ""):
				notepad.messageBox("Current script can NOT be located" + "\n\n" + s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False

			# for s_cmdline only : within the VBS script ' will be replaced by ", and then * by '
			s_vbs_filefull		= s_script_filefull[:s_script_filefull.rfind(".")] + s_vbs_ext
			i_cur_pid			= self.GetCurrentProcessId()
			s_npp_procname		= s_npp_filename
			s_cmdline			= \
				"'" + s_npp_filefull.replace("'", "*") + "'" + " " + \
				s_multinst + " " + s_opensession + " " + "'" + s_session_filefull.replace("'", "*") + "'"
			i_s_wait_npp_close	= self.s_wait_npp_close
			i_s_wait_npp_start	= self.s_wait_npp_start

			if not(self.os.path.isfile(s_vbs_filefull)):
				notepad.messageBox("VBS script file NOT found" + "\n" + s_vbs_filefull + "\n\n" + s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			# run the VBS script that will start the new Notepad++ instance
			try:
				# Popen automatically puts double-quotes around a parameter if it contains some spaces
				self.subprocess.Popen( \
					[s_wscript_exe_filefull, s_vbs_filefull, \
					str(i_cur_pid), s_npp_procname, s_cmdline, s_session_filefull, \
					str(i_s_wait_npp_close), str(i_s_wait_npp_start)])
			except:
				notepad.messageBox( \
					"VBS script run failed" + "\n" + \
					chr(34) + s_wscript_exe_filefull + chr(34) + " " + chr(34) + s_vbs_filefull + chr(34) + "\n\n" + \
					s_restartaborted, \
					s_script_name, MESSAGEBOXFLAGS.ICONSTOP)
				return False
			self.rest_session_filefull = None

			return True

		self.o_block_unblock_input_nppscintilla_wins.GetWinsHandle()

		# prepare restart by blocking current instance input
		self.o_block_unblock_input_nppscintilla_wins.block()	# -> IMPORTANT : block NPP window keyboard and mouse input
		b_restarted = Try_ReStart()
		if b_restarted:
			s_quitmsg = "This Notepad++ instance should auto-close now"
			print "\t" + ("*" * len(s_quitmsg))
			print "\t" + s_quitmsg
			print "\t" + ("*" * len(s_quitmsg))

			notepad.menuCommand(MENUCOMMAND.FILE_EXIT)			# -> IMPORTANT : should exit the process NOW if everything went OK
		else:
			print "\t" + "Restart aborted"
			if not(self.rest_session_filefull is None):
				notepad.loadSession(self.rest_session_filefull)	# otherwise restore opened files if they were closed
				try:
					self.os.remove(self.rest_session_filefull)	# and clean-up the temporary session file
				except:
					pass
		# if NPP did not close for any reason, un-block current instance input
		self.o_block_unblock_input_nppscintilla_wins.unblock()	# -> IMPORTANT : un-block NPP window keyboard and mouse input

		return b_restarted
# end of class

# Main() code **************************************************************************************
def Main():
	print "[" + s_script_name + " starts] " + s_controleditscript

	# create an instance of the restart class, and populate object properties with script variables
	o_restart_npp = C_Restart_NPP(i_s_wait_npp_close, i_s_wait_npp_start)

	o_restart_npp.Restart()

Main()
