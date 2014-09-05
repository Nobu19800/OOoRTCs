#-adsh_path_var PATH,OOoScriptPath,OpenOfficePath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python



OpenOfficePath=/opt/openoffice3/program



mkdir ~/.openoffice.org/3/user/Scripts
mkdir ${OOoScriptPath}

cp OOoCalcRTC.py ${OOoScriptPath}/OOoCalcRTC.py
cp OOoDrawRTC.py ${OOoScriptPath}/OOoDrawRTC.py
cp OOoWriterRTC.py ${OOoScriptPath}/OOoWriterRTC.py
cp OOoBaseRTC.py ${OOoScriptPath}/OOoBaseRTC.py
cp OOoImpressRTC.py ${OOoScriptPath}/OOoImpressRTC.py




cp rtc.conf ~/rtc.conf
cp OOoRTC -r ~/OOoRTC



if [ -e $OpenOfficePath ]; then
	${OpenOfficePath}/unopkg add -v OOoCalcControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoDrawControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoWriterControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoBaseControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoImpressControlRTC.oxt
else
	unopkg add -v OOoCalcControlRTC.oxt
	unopkg add -v OOoDrawControlRTC.oxt
	unopkg add -v OOoWriterControlRTC.oxt
	unopkg add -v OOoBaseControlRTC.oxt
	unopkg add -v OOoImpressControlRTC.oxt
fi



