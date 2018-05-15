@echo off

for /r %%i in (tests\*.vbs) do (
	echo running test %%i ...
	cscript.exe /Nologo tests.wsf %%i
	echo tests done!
)
