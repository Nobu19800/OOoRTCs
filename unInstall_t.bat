cd /d %~dp0

set OOoScriptPath=%USERPROFILE%\AppData\Roaming\OpenOffice.org\3\user\Scripts\python
set OOoPath="%ProgramFiles%\OpenOffice.org 3\program"

IF NOT EXIST %OOoPath% (

   set OOoPath="%ProgramFiles(x86)%\OpenOffice.org 3\program"
)




del %OOoPath%\OOoRTC.py

rd /s %OOoPath%\BaseIDL
rd /s %OOoPath%\CalcIDL
rd /s %OOoPath%\WriterIDL



pause

