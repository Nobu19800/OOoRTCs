#-adsh_path_var PATH,OOoScriptPath,OpenOfficePath

OOoScriptPath=~/.config/libreoffice/3/user

OpenOfficePath=/usr/lib/libreoffice/program

if [ -e $OOoScriptPath ]; then
	OOoScriptPath=~/.config/libreoffice/3/user
elif [ -e $OOoScriptPath ]; then
	OOoScriptPath=~/.config/libreoffice/4/user
else
	OOoScriptPath=~/.openoffice.org/3/user
fi

rm ${OOoScriptPath}/Scripts/python/OOoCalcRTC.py
rm ${OOoScriptPath}/Scripts/python/OOoDrawRTC.py
rm ${OOoScriptPath}/Scripts/python/OOoWriterRTC.py
rm ${OOoScriptPath}/Scripts/python/OOoBaseRTC.py
rm ${OOoScriptPath}/Scripts/python/OOoImpressRTC.py






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

