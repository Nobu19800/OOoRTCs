#-adsh_path_var PATH,OOoScriptPath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python







mkdir ~/.openoffice.org/3/user/Scripts
mkdir ${OOoScriptPath}

cp OOoCalcRTC.py ${OOoScriptPath}/OOoCalcRTC.py
cp OOoDrawRTC.py ${OOoScriptPath}/OOoDrawRTC.py
cp OOoWriterRTC.py ${OOoScriptPath}/OOoWriterRTC.py
cp OOoBaseRTC.py ${OOoScriptPath}/OOoBaseRTC.py
cp OOoImpressRTC.py ${OOoScriptPath}/OOoImpressRTC.py




cp rtc.conf ~/rtc.conf
cp OOoRTC -r ~/OOoRTC




unopkg add -v OOoCalcControlRTC.oxt
unopkg add -v OOoDrawControlRTC.oxt
unopkg add -v OOoWriterControlRTC.oxt
unopkg add -v OOoBaseControlRTC.oxt
unopkg add -v OOoImpressControlRTC.oxt