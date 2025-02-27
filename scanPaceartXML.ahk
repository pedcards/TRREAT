/*	Scan paceart\done XML files
	Get Providers/PatientProvider/Provider and /ProviderType for each XML
	Output to CSV or XLSX
*/

#Requires AutoHotkey v1.1
#Include includes

path := "paceart\done"
csv := "MRN,NAME,DATE,FOLLOWING,REFERRING1,REFERRING2,OTHER`n"
fcount:=ComObjCreate("Scripting.FileSystemObject").GetFolder(path).Files.Count

progress,,% " ",Scanning folder
Loop, Files, % path "\*", F
{
	progress, % 100 * A_index/fcount
	k := A_LoopField
	flist0 .= A_LoopFileName "`n"
}
Sort flist0

progress,,% " ",Removing duplicates
Loop, Parse, flist0, `n, `r`n
{
	k := A_LoopField
	if (k="") {
		Break
	}
	fnam := RegExReplace(k,"WQ.xml$")
	progress, % 100 * A_Index/fcount, % fnam
	e := StrSplit(fnam, "_")
	if !(e.1=e0.1) {																	; novel MRN
		flist .= k "`n"
		Continue
	} 
	; same MRNxte
	if (e.1=e0.1) && !(e.3 > e0.3) {													; Same MRN, but new date not greater than last
		e0 := e
		Continue
	}
}

progress,,% " ",Scanning files
Loop, Parse, flist, `n, `r`n
{
	k := A_LoopField
	if (k="") {
		Break
	}
	fnam := RegExReplace(k,"WQ.xml$")
	progress, % 100 * A_Index/fcount, % fnam
	e := StrSplit(fnam, "_")
	if (e.1=e0.1) && !(e.3 > e0.3) {													; Same MRN, but new date not greater than last
		e0 := e
		Continue
	}
	e0 := e
	y := new XML("paceart\done\" k)

	res := provs(y.selectSingleNode("//Providers"))

	csv .= e.1 ",""" e.2 """," e.3 ","
	for key,val in res
	{
		csv .= """" val ""","
	}
	csv .= "`n"
}

FileAppend, % csv, % A_now ".csv"

ExitApp

provs(providers) {
	res := {}
	roles := ["FOLLOWING","REFERRING1","REFERRING2"]
	for key,role in roles
	{
		provider := providers.selectSingleNode("PatientProvider[ProviderType='" role "']/Provider")
		nameL := provider.selectSingleNode("LastName").Text
		nameF := provider.selectSingleNode("FirstName").Text
		res[role] := (nameL) ? nameL ", " nameF : ""
	}

 	loop % (p := providers.selectNodes("PatientProvider/ProviderType")).Length()
	{
		k := p.item(A_Index-1).Text
		if ObjHasValue(roles,k) {
			Continue 
		} else {
			res.Push(k)
		}
	}

	return res
}

ObjHasValue(aObj, aValue, rx:="") {
; modified from http://www.autohotkey.com/board/topic/84006-ahk-l-containshasvalue-method/	
	if (aValue="") {
		return, false, errorlevel := 1
	}
	for key, val in aObj
		if (rx) {
			if (val ~= aValue) {														; aObj contains set of regex strings
				return, key, Errorlevel := 0
			}
			if (aValue ~= val) {
				return, key, ErrorLevel := 0											; aValue contains a regex string
			}
		} else {
			if (val = aValue) {
				return, key, ErrorLevel := 0
			}
		}
	return, false, errorlevel := 1
}

#Include xml.ahk
