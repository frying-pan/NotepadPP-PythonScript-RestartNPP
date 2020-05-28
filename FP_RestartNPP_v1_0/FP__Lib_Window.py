# ##################################################################################################
# Perso Library : Window
# ##################################################################################################
from Npp import *

class C_Get_NPPScintilla_Wins():
	# class constructor
	def __init__(self):
		import ctypes
		from ctypes import wintypes
		self.ctypes		= ctypes
		self.wintypes	= wintypes

		self.EnumWindowsProc			= ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
		self.EnumWindows				= ctypes.windll.user32.EnumWindows
		self.EnumChildWindows			= ctypes.windll.user32.EnumChildWindows
		self.RealGetWindowClass			= ctypes.windll.user32.RealGetWindowClassW
		self.GetParent					= ctypes.windll.user32.GetParent
		self.GetWindowThreadProcessId	= ctypes.windll.user32.GetWindowThreadProcessId
		self.GetCurrentProcessId		= ctypes.windll.kernel32.GetCurrentProcessId

	def GetWinsInfos(self, i_max_num):
		def CB_Enum_Window_Hwnd(hwnd, lparam):
			lst_enum_window_hwnd.append(hwnd)
			return True

		s_npp_class		= u"Notepad++"
		s_scint_class	= u"Scintilla"

		lst_scint_hwnd = []

		i_cur_pid = self.GetCurrentProcessId()
		npp_win_hwnd = None

		lst_enum_window_hwnd = []
		self.EnumWindows(self.EnumWindowsProc(CB_Enum_Window_Hwnd), 0)
		# get all class="Notepad++" top_level NPP windows handle (+1 for disambiguation, +1 for null terminator)
		buff = self.ctypes.create_unicode_buffer(len(s_npp_class) + 1 + 1)
		for win_hwnd in lst_enum_window_hwnd:
			self.RealGetWindowClass(win_hwnd, buff, len(s_npp_class) + 1 + 1)
			if buff.value == s_npp_class:
				# check if each NPP window handle is owned by current process, stop at the first found
				ci_win_pid = self.wintypes.DWORD(0)
				self.GetWindowThreadProcessId(win_hwnd, self.ctypes.pointer(ci_win_pid))
				if i_cur_pid == ci_win_pid.value:
					npp_win_hwnd = win_hwnd
					break
		if npp_win_hwnd is None:
			return (None, (), s_npp_class, s_scint_class)

		del lst_enum_window_hwnd[:]
		self.EnumChildWindows(npp_win_hwnd, self.EnumWindowsProc(CB_Enum_Window_Hwnd), 0)
		# get at most i_max_num first class="Scintilla" window handles, childs of npp_win_hwnd
		# i_max_num <= 0, or is None, will get all Scintilla child windows
		# (if i_max_num = 2 : *should* get primary and secondary views/editors window)
		buff = self.ctypes.create_unicode_buffer(len(s_scint_class) + 1 + 1)
		for win_hwnd in lst_enum_window_hwnd:
			self.RealGetWindowClass(win_hwnd, buff, len(s_scint_class) + 1 + 1)
			if buff.value == s_scint_class:
				lst_scint_hwnd.append(win_hwnd)
				if len(lst_scint_hwnd) == i_max_num: break

		return (npp_win_hwnd, lst_scint_hwnd, s_npp_class, s_scint_class)
# end of class

class C_Block_UnBlock_Input():
	# class constructor
	def __init__(self):
		import platform
		import ctypes
		from ctypes import wintypes

		self.GWL_Style = 0xFFFFFFF0
		self.WS_Disabled = 0x08000000
		self.SW_Hide	= 0x0
		self.SW_Show	= 0x5

		LONG_PTR = wintypes.LPARAM

		b_x86 = (platform.architecture()[0] == "32bit")
		if b_x86:
			self.GetWindowLong = ctypes.windll.user32.GetWindowLongW
			self.GetWindowLong.restype = wintypes.LONG
			self.GetWindowLong.argtypes = [wintypes.HWND, wintypes.INT]

			self.SetWindowLong = ctypes.windll.user32.SetWindowLongW
			self.SetWindowLong.restype = wintypes.LONG
			self.SetWindowLong.argtypes = [wintypes.HWND, wintypes.INT, wintypes.LONG]
		else:
			self.GetWindowLong = ctypes.windll.user32.GetWindowLongPtrW
			self.GetWindowLong.restype = LONG_PTR
			self.GetWindowLong.argtypes = [wintypes.HWND, wintypes.INT]

			self.SetWindowLong = ctypes.windll.user32.SetWindowLongPtrW
			self.SetWindowLong.restype = LONG_PTR
			self.SetWindowLong.argtypes = [wintypes.HWND, wintypes.INT, LONG_PTR]

		self.ShowWindowAsync	= ctypes.windll.user32.ShowWindowAsync
		self.SetLastError		= ctypes.windll.kernel32.SetLastError
		self.GetLastError		= ctypes.windll.kernel32.GetLastError

	def GetWinsHandle(self):
		o_get_nppscintilla_wins = C_Get_NPPScintilla_Wins()
		self.npp_win_hwnd, self.lst_win_hwnd, s_npp_class, s_scint_class = o_get_nppscintilla_wins.GetWinsInfos(None)
		self.lst_win_hwnd.append(self.npp_win_hwnd)

	def block(self):
		self.block_unblock(True)
	def unblock(self):
		self.block_unblock(False)
	def block_unblock(self, b_block):
		if self.npp_win_hwnd is None:
			return False

		for win_hwnd in self.lst_win_hwnd:
			self.SetLastError(0)
			i_winstyle_old = self.GetWindowLong(win_hwnd, self.GWL_Style)
			i_apierr = self.GetLastError()
			if i_apierr == 0:
				i_winstyle_new = i_winstyle_old
				if b_block:
					if ((i_winstyle_old & self.WS_Disabled) != self.WS_Disabled):
						i_winstyle_new = (i_winstyle_old | self.WS_Disabled)
				else:
					if ((i_winstyle_old & self.WS_Disabled) == self.WS_Disabled):
						i_winstyle_new = (i_winstyle_old - self.WS_Disabled)
				if i_winstyle_new != i_winstyle_old:
					self.SetWindowLong(win_hwnd, self.GWL_Style, i_winstyle_new)

		if b_block:
			self.ShowWindowAsync(self.npp_win_hwnd, self.SW_Hide)
			self.ShowWindowAsync(self.npp_win_hwnd, self.SW_Show)

		return True
# end of class
