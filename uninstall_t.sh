#-adsh_path_var PATH,OOoScriptPath,OOoPath

OOoScriptPath=~/.openoffice.org/3/user/Scripts/python

OOoPath=/opt/openoffice3/basis-link/program

if [ ! -e $OOoPath ]; then
	OOoPath=/usr/lib/openoffice/basis-link/program
fi



rm -rf ${OOoPath}/BaseIDL
rm -rf ${OOoPath}/CalcIDL
rm -rf ${OOoPath}/WriterIDL
rm -rf ${OOoPath}/OOoRTC.py



