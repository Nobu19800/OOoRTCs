#-adsh_path_var PATH,OOoScriptPath,OOoPath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python

OOoPath=/opt/openoffice3/basis-link/program

if [ ! -e $OOoPath ]; then
	OOoPath=/usr/lib/openoffice/basis-link/program
fi

mkdir ~/.openoffice.org/3/user/Scripts
mkdir ${OOoScriptPath}

cp OOoCalcRTC.py ${OOoScriptPath}/OOoCalcRTC.py
cp OOoDrawRTC.py ${OOoScriptPath}/OOoDrawRTC.py
cp OOoWriterRTC.py ${OOoScriptPath}/OOoWriterRTC.py
cp OOoBaseRTC.py ${OOoScriptPath}/OOoBaseRTC.py
cp OOoImpressRTC.py ${OOoScriptPath}/OOoImpressRTC.py



cp OOoRTC -r ~/OOoRTC


sudo cp rtc.conf ${OOoPath}/rtc.conf

unopkg add -v OOoCalcControlRTC.oxt
unopkg add -v OOoDrawControlRTC.oxt
unopkg add -v OOoWriterControlRTC.oxt
unopkg add -v OOoBaseControlRTC.oxt
unopkg add -v OOoImpressControlRTC.oxt



