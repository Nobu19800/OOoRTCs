#-adsh_path_var PATH,OOoScriptPath,OOoPath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python
OOoPath=/usr/lib/openoffice/basis-link/program


cp OOoCalcRTC.py ${OOoScriptPath}/OOoCalcRTC.py
cp OOoDrawRTC.py ${OOoScriptPath}/OOoDrawRTC.py
cp OOoWriterRTC.py ${OOoScriptPath}/OOoWriterRTC.py
cp OOoBaseRTC.py ${OOoScriptPath}/OOoBaseRTC.py
cp OOoImpressRTC.py ${OOoScriptPath}/OOoImpressRTC.py



cp OOoRTC.py ${OOoPath}/OOoRTC.py
cp rtc.conf ${OOoPath}/rtc.conf
cp BaseIDL -r ${OOoPath}/BaseIDL
cp CalcIDL -r ${OOoPath}/CalcIDL

unopkg add -v OOoCalcControlRTC.oxt
unopkg add -v OOoDrawControlRTC.oxt
unopkg add -v OOoWriterControlRTC.oxt
unopkg add -v OOoBaseControlRTC.oxt
unopkg add -v OOoImpressControlRTC.oxt
