cd /d %~dp0

set OOoScriptPath=%USERPROFILE%\AppData\Roaming\OpenOffice\4\user\Scripts\python
set OOoPath="%ProgramFiles%\OpenOffice 4\program"

IF NOT EXIST %OOoPath% (

   set OOoPath="%ProgramFiles(x86)%\OpenOffice 4\program"
   IF NOT EXIST %OOoPath% (
	set OOoPath="%ProgramFiles%\OpenOffice.org 3\program"
	IF NOT EXIST %OOoPath% (
		set OOoPath="%ProgramFiles(x86)%\OpenOffice.org 3\program"
	)
   )
)

IF NOT EXIST %OOoScriptPath% (
	set OOoScriptPath=%USERPROFILE%\AppData\Roaming\OpenOffice.org\3\user\Scripts\python
)

del %OOoScriptPath%\OOoCalcRTC.py
del %OOoScriptPath%\OOoDrawRTC.py
del %OOoScriptPath%\OOoWriterRTC.py
del %OOoScriptPath%\OOoBaseRTC.py
del %OOoScriptPath%\OOoImpressRTC.py



del %OOoPath%\rtc.conf
rd /s %OOoPath%\OOoRTC


set OOoCD=%CD%

cd %OOoPath%
unopkg remove -v OOoCalcControlRTC.oxt
unopkg remove -v OOoDrawControlRTC.oxt
unopkg remove -v OOoWriterControlRTC.oxt
unopkg remove -v OOoBaseControlRTC.oxt
unopkg remove -v OOoImpressControlRTC.oxt

pause

