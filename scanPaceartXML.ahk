/*	Scan paceart\done XML files
	Get Providers/PatientProvider/Provider and /ProviderType for each XML
	Output to CSV or XLSX
*/

#Requires AutoHotkey v1.1
#Include includes

path := "paceart\done"
fcount:=ComObjCreate("Scripting.FileSystemObject").GetFolder(path).Files.Count

progress,,% " ",Scanning folder
Loop, Files, % path "\*", F
{
	progress, % 100 * A_index/fcount
	flist .= A_LoopFileName "`n"
}


ExitApp
#Include xml.ahk
