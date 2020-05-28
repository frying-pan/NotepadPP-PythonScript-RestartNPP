Option Explicit
Public S_CURVBS
S_CURVBS = WScript.ScriptName & " [" & Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1) & "]"

MAIN

Sub MAIN()
	Dim i_npp_proc_pid, s_npp_procname, s_cmdline, s_session_filefull, i_s_wait_close, i_s_wait_start
	Dim s_u_session_filefull, s_u_folfull_bsl_1, s_u_folfull_bsl_2, s_u_folfull_bsl_3, s_u_folfull_bsl_4
	Dim b_npp_closed, b_file_exists, i_errnum, s_errdes
	Dim o_WshShell, o_FileSysObj, o_WMIServiceObj, o_FileObj
	Const s_restartfailed	= "Notepad++ restart failed"
	Const s_system_env		= "System"
	Const s_user_env		= "User"
	Const s_env_var_tmp		= "Tmp"
	Const s_env_var_temp	= "Temp"
	Const i_ms_savecfg_gracedelay = 1000

	On Error Resume Next
		Set o_WshShell		= WScript.CreateObject("WScript.Shell")
		Set o_FileSysObj	= WScript.CreateObject("Scripting.FileSystemObject")
		i_errnum = Err.Number
	On Error GoTo 0
	If i_errnum <> 0 Then
		MsgBox "Unable to create some WScript objects" & vbCr & vbCr & s_restartfailed, vbCritical, S_CURVBS
		WScript.Quit
	End If

	On Error Resume Next
		i_npp_proc_pid		= Int(	WScript.Arguments.Item(0))
		s_npp_procname		=		WScript.Arguments.Item(1)
		s_cmdline			=		WScript.Arguments.Item(2)
		s_session_filefull	=		WScript.Arguments.Item(3)
		i_s_wait_close		= Int(	WScript.Arguments.Item(4))
		i_s_wait_start		= Int(	WScript.Arguments.Item(5))
		i_errnum = Err.Number
		s_errdes = Err.Description
	On Error GoTo 0
	If i_errnum <> 0 Then
		MsgBox "Incorrect input parameters" & vbCr & vbCr & s_errdes & vbCr & vbCr & s_restartfailed, vbCritical, S_CURVBS
		WScript.Quit
	End If
	s_cmdline = Replace(s_cmdline, "'", Chr(34))
	s_cmdline = Replace(s_cmdline, "*", "'")

	On Error Resume Next
		Set o_WMIServiceObj = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
		i_errnum = Err.Number
	On Error GoTo 0
	If i_errnum <> 0 Then
		MsgBox "Unable to connect to the WMI service" & vbCr & vbCr & s_restartfailed, _
			vbExclamation, S_CURVBS
		Exit Sub
	End If

	b_npp_closed = WAIT_PREV_NPP_CLOSED(o_WMIServiceObj, i_npp_proc_pid, s_npp_procname, i_s_wait_close)
	If IsEmpty(b_npp_closed) Then
		MsgBox "Unable to check if previous Notepad++ instance has closed" & vbCr & vbCr & s_restartfailed, _
			vbExclamation, S_CURVBS
		Exit Sub
	End If

	If Not(b_npp_closed) Then
		MsgBox _
			"Previous Notepad++ instance did not close in time" & vbCr & vbCr & _
			"(Time-out : i_s_wait_npp_close = "  & i_s_wait_close & " seconds)" & vbCr & vbCr & _
			s_restartfailed, _
			vbExclamation, S_CURVBS
		Exit Sub
	End If

	WScript.Sleep i_ms_savecfg_gracedelay
	On Error Resume Next
		o_WshShell.Run s_cmdline, 1, False
		i_errnum = Err.Number
	On Error GoTo 0
	If i_errnum <> 0 Then
		MsgBox "Unable to start the new Notepad++ instance" & vbCr & s_cmdline & vbCr & vbCr & s_restartfailed, _
			vbExclamation, S_CURVBS
		Exit Sub
	End If

	s_u_session_filefull = UCase(GET_SHORT_PATH(o_FileSysObj, s_session_filefull, True))
	If IsEmpty(s_u_session_filefull) Then
		MsgBox _
			"Unable to locate the temporary session file to clean-up" & vbCr & s_session_filefull, _
			vbExclamation, S_CURVBS
		Exit Sub
	End If

	On Error Resume Next
		s_u_folfull_bsl_1 = o_WshShell.ExpandEnvironmentStrings(o_WshShell.Environment(s_user_env)(s_env_var_tmp)) & "\"
		s_u_folfull_bsl_2 = o_WshShell.ExpandEnvironmentStrings(o_WshShell.Environment(s_user_env)(s_env_var_temp)) & "\"
		s_u_folfull_bsl_3 = o_WshShell.ExpandEnvironmentStrings(o_WshShell.Environment(s_system_env)(s_env_var_tmp)) & "\"
		s_u_folfull_bsl_4 = o_WshShell.ExpandEnvironmentStrings(o_WshShell.Environment(s_system_env)(s_env_var_temp)) & "\"
	On Error GoTo 0
	s_u_folfull_bsl_1 = UCase(GET_SHORT_PATH(o_FileSysObj, s_u_folfull_bsl_1, False))
	s_u_folfull_bsl_2 = UCase(GET_SHORT_PATH(o_FileSysObj, s_u_folfull_bsl_2, False))
	s_u_folfull_bsl_3 = UCase(GET_SHORT_PATH(o_FileSysObj, s_u_folfull_bsl_3, False))
	s_u_folfull_bsl_4 = UCase(GET_SHORT_PATH(o_FileSysObj, s_u_folfull_bsl_4, False))
	If ( _
			(Len(s_u_folfull_bsl_1) = 0 Or Left(s_u_session_filefull, Len(s_u_folfull_bsl_1)) <> s_u_folfull_bsl_1) And _
			(Len(s_u_folfull_bsl_2) = 0 Or Left(s_u_session_filefull, Len(s_u_folfull_bsl_2)) <> s_u_folfull_bsl_2) And _
			(Len(s_u_folfull_bsl_3) = 0 Or Left(s_u_session_filefull, Len(s_u_folfull_bsl_3)) <> s_u_folfull_bsl_3) And _
			(Len(s_u_folfull_bsl_4) = 0 Or Left(s_u_session_filefull, Len(s_u_folfull_bsl_4)) <> s_u_folfull_bsl_4)) Then
		MsgBox _
			"Clean-up of the temporary session file NOT done" & vbCr & s_session_filefull & vbCr & vbCr & _
			"(This file was NOT safely located in a temporary folder)", _
			vbExclamation, S_CURVBS
		Exit Sub
	End If

	WScript.Sleep (i_s_wait_start * 1000)
	On Error Resume Next
		Set o_FileObj = o_FileSysObj.GetFile(s_session_filefull)
		o_FileObj.Delete False
		b_file_exists = o_FileSysObj.FileExists(s_session_filefull)
		i_errnum = Err.Number
	On Error GoTo 0
	If (i_errnum <> 0 Or b_file_exists) Then
		MsgBox _
			"Unable to clean-up the temporary session file" & vbCr & _
			"(This file was not found, read-only or locked)" & vbCr & vbCr & _
			s_session_filefull, _
			vbExclamation, S_CURVBS
		Exit Sub
	End If
End Sub

Function GET_SHORT_PATH(ByVal o_FileSysObj, ByVal s_filefolfull, ByVal b_file_or_fol)
	Dim i_errnum
	Dim o_FileFolObj

	On Error Resume Next
		If b_file_or_fol Then
			Set o_FileFolObj = o_FileSysObj.GetFile(s_filefolfull)
		Else
			Set o_FileFolObj = o_FileSysObj.GetFolder(s_filefolfull)
		End If
		s_filefolfull = o_FileFolObj.ShortPath
		i_errnum = Err.Number
	On Error GoTo 0
	If i_errnum <> 0 Then
		GET_SHORT_PATH = Empty
		Exit Function
	End If

	GET_SHORT_PATH = s_filefolfull
End Function

Function WAIT_PREV_NPP_CLOSED(ByVal o_WMIServiceObj, ByVal i_npp_proc_pid, ByVal s_npp_procname, ByVal i_s_wait_close)
	Dim s_procname, b_npp_closed, i_ms_waited, i_errnum
	Dim o_ProcsColl, o_ProcObj
	Const i_ms_waitloop = 250

	WAIT_PREV_NPP_CLOSED = Empty

	i_ms_waited = 0
	Do
		b_npp_closed = True
		On Error Resume Next
			Set o_ProcsColl = o_WMIServiceObj.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId=" & i_npp_proc_pid)
			For Each o_ProcObj In o_ProcsColl
				s_procname = o_ProcObj.Name
				If UCase(s_procname) = UCase(s_npp_procname) Then b_npp_closed = False
			Next
			i_errnum = Err.Number
		On Error GoTo 0
		If i_errnum = 0 Then
			If b_npp_closed Then Exit Do
		Else
			b_npp_closed = Empty
		End If

		If i_ms_waited >= (i_s_wait_close * 1000) Then Exit Do
		i_ms_waited = i_ms_waited + i_ms_waitloop
		WScript.Sleep i_ms_waitloop
	Loop

	WAIT_PREV_NPP_CLOSED = b_npp_closed
End Function
