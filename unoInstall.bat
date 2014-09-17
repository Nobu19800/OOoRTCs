cd /d %~dp0


set OOoPath="%ProgramFiles%\OpenOffice.org 3\program"

IF NOT EXIST %OOoPath% (

   set OOoPath="%ProgramFiles(x86)%\OpenOffice.org 3\program"
)


set OOoCD=%CD%

cd %OOoPath%

unopkg remove -v OOoCalcControlRTC.oxt
unopkg remove -v OOoDrawControlRTC.oxt
unopkg remove -v OOoWriterControlRTC.oxt
unopkg remove -v OOoBaseControlRTC.oxt
unopkg remove -v OOoImpressControlRTC.oxt

unopkg add -v "%OOoCD%\OOoCalcControlRTC.oxt"
unopkg add -v "%OOoCD%\OOoDrawControlRTC.oxt"
unopkg add -v "%OOoCD%\OOoWriterControlRTC.oxt"
unopkg add -v "%OOoCD%\OOoBaseControlRTC.oxt"
unopkg add -v "%OOoCD%\OOoImpressControlRTC.oxt"


