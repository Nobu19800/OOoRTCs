#-adsh_path_var PATH,OOoScriptPath,OOoPath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python

OOoPath=/opt/openoffice3/basis-link/program

if [ ! -e $OOoPath ]; then
	OOoPath=/usr/lib/openoffice/basis-link/program
fi



rm ${OOoScriptPath}/OOoCalcRTC.py
rm ${OOoScriptPath}/OOoDrawRTC.py
rm ${OOoScriptPath}/OOoWriterRTC.py
rm ${OOoScriptPath}/OOoBaseRTC.py
rm ${OOoScriptPath}/OOoImpressRTC.py


rm -rf ~/OOoRTC



sudo rm ${OOoPath}/rtc.conf

unopkg remove -v OOoCalcControlRTC.oxt
unopkg remove -v OOoDrawControlRTC.oxt
unopkg remove -v OOoWriterControlRTC.oxt
unopkg remove -v OOoBaseControlRTC.oxt
unopkg remove -v OOoImpressControlRTC.oxt




