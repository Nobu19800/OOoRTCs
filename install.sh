#-adsh_path_var PATH,OOoScriptPath,OpenOfficePath

OOoScriptPath=~/.config/libreoffice/3/user


OpenOfficePath=/usr/lib/libreoffice/program


if [ -e $OOoScriptPath ]; then
	OOoScriptPath=~/.config/libreoffice/3/user
else
	OOoScriptPath=~/.openoffice.org/3/user
fi


mkdir $OOoScriptPath/Scripts
mkdir $OOoScriptPath/Scripts/python

cp OOoCalcRTC.py ${OOoScriptPath}/Scripts/python/OOoCalcRTC.py
cp OOoDrawRTC.py ${OOoScriptPath}/Scripts/python/OOoDrawRTC.py
cp OOoWriterRTC.py ${OOoScriptPath}/Scripts/python/OOoWriterRTC.py
cp OOoBaseRTC.py ${OOoScriptPath}/Scripts/python/OOoBaseRTC.py
cp OOoImpressRTC.py ${OOoScriptPath}/Scripts/python/OOoImpressRTC.py





rm -rf ~/OOoRTC
cp OOoRTC -r ~/OOoRTC




if [ -e $OpenOfficePath ]; then
	${OpenOfficePath}/unopkg remove -v OOoCalcControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoDrawControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoWriterControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoBaseControlRTC.oxt
	${OpenOfficePath}/unopkg remove -v OOoImpressControlRTC.oxt

	${OpenOfficePath}/unopkg add -v OOoCalcControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoDrawControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoWriterControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoBaseControlRTC.oxt
	${OpenOfficePath}/unopkg add -v OOoImpressControlRTC.oxt
else
	unopkg remove -v OOoCalcControlRTC.oxt
	unopkg remove -v OOoDrawControlRTC.oxt
	unopkg remove -v OOoWriterControlRTC.oxt
	unopkg remove -v OOoBaseControlRTC.oxt
	unopkg remove -v OOoImpressControlRTC.oxt

	unopkg add -v OOoCalcControlRTC.oxt
	unopkg add -v OOoDrawControlRTC.oxt
	unopkg add -v OOoWriterControlRTC.oxt
	unopkg add -v OOoBaseControlRTC.oxt
	unopkg add -v OOoImpressControlRTC.oxt
fi
