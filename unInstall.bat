cd /d %~dp0

set OOoScriptPath=%USERPROFILE%\AppData\Roaming\OpenOffice.org\3\user\Scripts\python
set OOoPath="%ProgramFiles%\OpenOffice.org 3\program"

IF NOT EXIST %OOoPath% (

   set OOoPath="%ProgramFiles(x86)%\OpenOffice.org 3\program"
)

del %OOoScriptPath%\OOoCalcRTC.py
del %OOoScriptPath%\OOoDrawRTC.py
del %OOoScriptPath%\OOoWriterRTC.py
del %OOoScriptPath%\OOoBaseRTC.py
del %OOoScriptPath%\OOoImpressRTC.py


del %OOoPath%\OOoRTC.py
del %OOoPath%\rtc.conf
rd /s %OOoPath%\OOoRTC


set OOoCD=%CD%

cd %OOoPath%
unopkg remove -v %OOoCD%\OOoCalcControlRTC.oxt
unopkg remove -v %OOoCD%\OOoDrawControlRTC.oxt
unopkg remove -v %OOoCD%\OOoWriterControlRTC.oxt
unopkg remove -v %OOoCD%\OOoBaseControlRTC.oxt
unopkg remove -v %OOoCD%\OOoImpressControlRTC.oxt

pause

