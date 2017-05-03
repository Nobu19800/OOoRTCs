cd /d %~dp0

cmd /c OOoRTC\CalcIDL\idlcompile.bat
cmd /c OOoRTC\BaseIDL\idlcompile.bat
cmd /c OOoRTC\WriterIDL\idlcompile.bat

set OOoScriptPath=%USERPROFILE%\AppData\Roaming\OpenOffice\4\user\Scripts\python
set OOoPath="%ProgramFiles%\OpenOffice 4\program"

IF NOT EXIST %OOoPath% (
	set OOoPath="%ProgramFiles(x86)%\OpenOffice 4\program"
)

IF NOT EXIST %OOoPath% (
	set OOoPath="%ProgramFiles%\OpenOffice.org 3\program"
)

IF NOT EXIST %OOoPath% (
	set OOoPath="%ProgramFiles(x86)%\OpenOffice.org 3\program"
)

IF NOT EXIST %OOoScriptPath% (
	set OOoScriptPath=%USERPROFILE%\AppData\Roaming\OpenOffice.org\3\user\Scripts\python
)

mkdir %OOoScriptPath%

copy OOoCalcRTC.py %OOoScriptPath%\OOoCalcRTC.py
copy OOoDrawRTC.py %OOoScriptPath%\OOoDrawRTC.py
copy OOoWriterRTC.py %OOoScriptPath%\OOoWriterRTC.py
copy OOoBaseRTC.py %OOoScriptPath%\OOoBaseRTC.py
copy OOoImpressRTC.py %OOoScriptPath%\OOoImpressRTC.py



copy rtc.conf %OOoPath%\rtc.conf
xcopy OOoRTC /e %OOoPath%\OOoRTC


set OOoCD=%CD%

cd %OOoPath%
unopkg add -v %OOoCD%\OOoCalcControlRTC.oxt
unopkg add -v %OOoCD%\OOoDrawControlRTC.oxt
unopkg add -v %OOoCD%\OOoWriterControlRTC.oxt
unopkg add -v %OOoCD%\OOoBaseControlRTC.oxt
unopkg add -v %OOoCD%\OOoImpressControlRTC.oxt

pause

