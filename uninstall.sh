#-adsh_path_var PATH,OOoScriptPath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python



rm ${OOoScriptPath}/OOoCalcRTC.py
rm ${OOoScriptPath}/OOoDrawRTC.py
rm ${OOoScriptPath}/OOoWriterRTC.py
rm ${OOoScriptPath}/OOoBaseRTC.py
rm ${OOoScriptPath}/OOoImpressRTC.py






rm ~/rtc.conf

rm -rf ~/OOoRTC



unopkg remove -v OOoCalcControlRTC.oxt
unopkg remove -v OOoDrawControlRTC.oxt
unopkg remove -v OOoWriterControlRTC.oxt
unopkg remove -v OOoBaseControlRTC.oxt
unopkg remove -v OOoImpressControlRTC.oxt

