
OutlookAddinps.dll: dlldata.obj OutlookAddin_p.obj OutlookAddin_i.obj
	link /dll /out:OutlookAddinps.dll /def:OutlookAddinps.def /entry:DllMain dlldata.obj OutlookAddin_p.obj OutlookAddin_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del OutlookAddinps.dll
	@del OutlookAddinps.lib
	@del OutlookAddinps.exp
	@del dlldata.obj
	@del OutlookAddin_p.obj
	@del OutlookAddin_i.obj
