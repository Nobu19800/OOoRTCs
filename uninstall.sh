#-adsh_path_var PATH,OOoScriptPath,OOoPath,OpenOfficePath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python

OOoPath=/opt/openoffice3/basis-link/program

OpenOfficePath=/opt/openoffice3/program

if [ ! -e $OOoPath ]; then
	OOoPath=/usr/lib/openoffice/basis-link/program
fi



rm ${OOoScriptPath}/OOoCalcRTC.py
rm ${OOoScriptPath}/OOoDrawRTC.py
rm ${OOoScriptPath}/OOoWriterRTC.py
rm ${OOoScriptPath}/OOoBaseRTC.py
rm ${OOoScriptPath}/OOoImpressRTC.py






rm ~/rtc.conf

rm -rf ~/OOoRTC


if [ -e $OpenOfficePath ]; then
	${OpenOfficePath}/unopkg remove -v OOoCalcControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoDrawControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoWriterControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoBaseControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoImpressControlRTC.oxt
else
	unopkg remove -v OOoCalcControlRTC.oxt
	unopkg remove -v OOoDrawControlRTC.oxt
	unopkg remove -v OOoWriterControlRTC.oxt
	unopkg remove -v OOoBaseControlRTC.oxt
	unopkg remove -v OOoImpressControlRTC.oxt
fi




