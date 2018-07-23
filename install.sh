#!/bin/sh
cd `dirname $0`

sh OOoRTC/CalcIDL/idlcompile.sh
sh OOoRTC/BaseIDL/idlcompile.sh
sh OOoRTC/WriterIDL/idlcompile.sh

OOoScriptPath1=~/.config/libreoffice/3/user
OOoScriptPath2=~/.openoffice.org/3/user
OOoScriptPath3=~/.config/libreoffice/4/user

OpenOfficePath=/usr/lib/libreoffice/program

if [ -e $OOoScriptPath1 ]; then
	OOoScriptPath=$OOoScriptPath1
elif [ -e $OOoScriptPath2 ]; then
	OOoScriptPath=$OOoScriptPath2
else
	OOoScriptPath=$OOoScriptPath3
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
