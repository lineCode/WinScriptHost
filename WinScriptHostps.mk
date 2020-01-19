
WinScriptHostps.dll: dlldata.obj WinScriptHost_p.obj WinScriptHost_i.obj
	link /dll /out:WinScriptHostps.dll /def:WinScriptHostps.def /entry:DllMain dlldata.obj WinScriptHost_p.obj WinScriptHost_i.obj \
		mtxih.lib mtx.lib mtxguid.lib \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \
		ole32.lib advapi32.lib 

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		/MD \
		$<

clean:
	@del WinScriptHostps.dll
	@del WinScriptHostps.lib
	@del WinScriptHostps.exp
	@del dlldata.obj
	@del WinScriptHost_p.obj
	@del WinScriptHost_i.obj
