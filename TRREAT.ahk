/*	TRREAT - The Rhythm Recording Electronic Analysis Transmogrifier - PM
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%
FileInstall, pdftotext.exe, pdftotext.exe
#Include includes
IfInString, A_WorkingDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	isDevt := true
	binDir := ".\bin\"
	reportDir := ".\pending\"													; generated files pending signature
	complDir := ".\completed\"													; archive signed files and original PDF files
	chipDir := "..\CHIPOTLE\"													; CHIPOTLE files
	pdfDir := ".\USB\"
} else {
	isDevt := false
	binDir := "\\childrens\files\HCCardiologyFiles\EP\TRREAT_files\bin\"
	reportDir := "\\childrens\files\HCCardiologyFiles\EP\TRREAT_files\pending\"
	complDir := "\\childrens\files\HCCardiologyFiles\EP\TRREAT_files\completed\"
	chipDir := "\\childrens\files\HCChipotle\"
	pdfDir := ".\"
}
if !FileExist(reportDir) {
	MsgBox % "Requires pending dir`n""" reportDir """"
	ExitApp
}
if !FileExist(complDir) {
	MsgBox % "Requires completed dir`n""" complDir """"
	ExitApp
}
if !FileExist(chipDir) {
	MsgBox % "Requires CHIPOTLE dir`n""" chipDir """"
	ExitApp
}

user := A_UserName
newTxt := Object()
blk := Object()
blk2 := Object()
;~ docs := Object()
docs := {"TC":"783118","JS":"343079","SS":"358945"}

if ObjHasKey(docs,substr(user,1,2)) {											; User is in docs[]
	role := "Sign"																; set role to "Sign"
}

if (%0%) {																		; For each parameter: 
  fileArg := true																; Gets parameter passed to script/exe. 
  role := "Parse"
} else if (isDevt) {															; But if isDevt
	role := cMsgBox("Administrator"												; offer to either PARSE or SIGN
		, "Enter ROLE:"
		, "*&Parse PDFs|&Sign reports"
		, "Q","")
}

if instr(role,"Sign") {
	xl := new XML(binDir "worklist.xml")									; otherwise load existing worklist
	gosub signScan
}

if instr(role,"Parse") {
	Gui, Parse:Destroy
	Gui, Parse:Add, Listview, w600 -Multi Grid r12 gparsePat hwndHLV, Date|Name|Device|Serial|Status|PaceArt|FileName|MetaData|Report
	Gui, Parse:Default
	Gui, Show,, TRREAT Reports scan
	LV_ModifyCol(1, "Autohdr")													; when done, reformat the col widths
	LV_ModifyCol(2, "Autohdr")
	LV_ModifyCol(3, "Autohdr")
	LV_ModifyCol(4, "Autohdr")
	LV_ModifyCol(5, "Autohdr")
	LV_ModifyCol(6, "Autohdr")
	LV_ModifyCol(7, "0")														; hide the filename col
	LV_ModifyCol(8, "0")														; hide the metadata col
	LV_ModifyCol(9, "0")														; hide the report col
	
	gosub readList																; read the worklist
	
	if (fileArg) {																; process files dropped on TRREAT icon
		Loop, %0%
		{
			SplitPath, %A_Index%, fileIn 
			gosub fileLoop
		}
	} else {																	; otherwise scan the folders
		gosub readFiles
	}
}

WinWaitClose, TRREAT Reports
ExitApp

readList:
{
	tmp := []
	LV_Delete()
	fileNum := 0																; Start the worklist at fileNum 0
	if !FileExist(binDir "worklist.xml") {
		xl := new XML("<root/>")												; Create new XML if doesn't exist
		xl.addElement("work", "root", {ed: A_Now})
		xl.addElement("done", "root", {ed: A_Now})
		xl.save(binDir "worklist.xml")
	} else {
		xl := new XML(binDir "worklist.xml")									; otherwise load existing worklist
	}
	Loop, % (w_id := xl.selectNodes("/root/work/id")).length					; scan through each <id>
	{
		k := w_id.item(A_Index-1)												; put it into k
		if !IsObject(k) {														; skip if empty
			continue
		}
		tmp["date"] := k.getAttribute("date")
		tmp["name"] := k.selectSingleNode("name").text
		tmp["dev"]  := k.selectSingleNode("dev").text
		tmp["ser"]  := k.getAttribute("ser")
		tmp["status"] := k.selectSingleNode("status").text
		tmp["paceart"] := k.selectSingleNode("paceart").text
		tmp["file"] := k.selectSingleNode("file").text
		tmp["meta"] := k.selectSingleNode("meta").text
		tmp["report"] := k.selectSingleNode("report").text
		if ((tmp.report) && (tmp.status="Signed") && (tmp.paceart)) {
			MsgBox, 68, Tasks completed
				, % "Patient " tmp.name "`n"
				. 	"- Report done`n"
				. 	"- Report signed`n"
				. 	"- PaceArt filled`n`n"
				. 	"Archive this task?"
			IfMsgBox, Yes
			{
				archNode("/root/work/id[@date='" tmp.date "'][@ser='" tmp.ser "']")				; copy ID node to DONE
				xl.save(binDir "worklist.xml")
				continue
			}
		}
		
		fileNum += 1															; Add a row to the LV
		LV_Add("", tmp.date)								; col1 is date
		LV_Modify(fileNum,"col2", tmp.name)
		LV_Modify(fileNum,"col3", tmp.dev)
		LV_Modify(fileNum,"col4", tmp.ser)
		LV_Modify(fileNum,"col5", tmp.status)
		LV_Modify(fileNum,"col6", tmp.paceart)
		LV_Modify(fileNum,"col7", tmp.file)
		LV_Modify(fileNum,"col8", tmp.meta)
		LV_Modify(fileNum,"col9", tmp.report)
	}
	LV_ModifyCol(1, "Autohdr")													; when done, reformat the col widths
	LV_ModifyCol(2, "Autohdr")
	LV_ModifyCol(3, "Autohdr")
	LV_ModifyCol(4, "Autohdr")
	LV_ModifyCol(5, "Autohdr")
	LV_ModifyCol(6, "Autohdr")
return
}

readFiles:
{
/*	Read root - usually MEDT files
*/
	kmaxstr :=																	; clear string of files to skip
	Loop, files, % pdfDir "*.pdf"												; read all PDFs in root
	{
		patPDF := A_LoopFileName												; next file in PDFdir
		if instr(kmaxstr,patPDF) {												; in skiplist?
			continue
		}
		kmax := 1																; reset max k counter
		Loop, files, % pdfDir strX(patPDF,"",1,0,"_",0,1) "*.pdf"				; loop through all files with this "prefix"
		{
			i := A_LoopFileName													; i is filename in this inner loop
			n := substr(i,instr(i,"_",,-1))										; n is string up to final _#
			k := strX(i,"_",n,1,".",1)											; k is # between _ and .pdf
			if (k > kmax) {														; greater than previous kmax?
				j := substr(i,1,instr(i,"_",,-1)) (kmax) ".pdf"					; j is filename of previous kmax
				FileMove, % pdfDir j, % pdfDir j ".old"							; rename it to j.pdf.old
				kmax := k														; new kmax
				patPDF := i														; set patPDF as this new max (for when exits)
				kmaxstr .= i "`n"												; add to string of files to subsequently ignore
			}
		}
		tmp_name := strX(patPDF,"",1,0,"_",1,1,n)
		tmp_ser := strX(patPDF,"_",n-1,1,"_",1,1,n)
		tmp := parseDate(strX(patPDF,"_",n+1,1,".pdf",1,6))
		tmp_date := tmp.YYYY tmp.MM tmp.DD
		patPDF := pdfDir patPDF
		if ((xl.selectSingleNode("//id[@date='" tmp_date "']"))					; skip reprocessing if in either work or done lists
			&& (xl.selectSingleNode("//id[@ser='" tmp_ser "']"))) {
			continue
		}
		fileNum += 1															; Add a row to the LV
		LV_Add("", tmp_date)													; col1 is date
		LV_Modify(fileNum,"col2", tmp_name)
		LV_Modify(fileNum,"col3", "Medtronic")
		LV_Modify(fileNum,"col4", tmp_ser)
		LV_Modify(fileNum,"col5", "")
		LV_Modify(fileNum,"col6", "")
		LV_Modify(fileNum,"col7", patPDF)
		LV_Modify(fileNum,"col8", "")
	}

/* Read SJM "PDFs" folder
*/
	sjmDir := pdfDir "PDFs\Live.combined"
	Loop, Files, % sjmDir "\*", D
	{
		DateDir := A_LoopFileName
		Loop, Files, % sjmDir "\" DateDir "\*", D
		{
			patDir := A_LoopFileName
			tmp_date := RegExReplace(DateDir,"-")
			tmp_name := stregx(patDir,"",1,0,"_\d{2,}",1,n)
			tmp_dev  := stregx(patDir,"_",n,1,"_",1,n)
			tmp_ser  := strx(patDir,"_",n,1,"",0)
			if ((xl.selectSingleNode("//id[@date='" tmp_date "']"))				; skip reprocessing if in either work or done lists
				&& (xl.selectSingleNode("//id[@ser='" tmp_ser "']"))) {
				continue
			}
			Loop, Files, % sjmDir "\" DateDir "\" patDir "\*.pdf", F
			{
				patPDF := A_LoopFileName
				tmp_dev := strx(patPDF,"",1,0,"_",1,1)
				tmp_meta := (FileExist(pdfDir tmp_ser ".log")) ? pdfDir tmp_ser ".log" : ""
				fileNum += 1													; Add a row to the LV
				LV_Add("", tmp_date)											; col1 is date
				LV_Modify(fileNum,"col2", tmp_name)
				LV_Modify(fileNum,"col3", "SJM " tmp_dev)
				LV_Modify(fileNum,"col4", tmp_ser)
				LV_Modify(fileNum,"col5", "")
				LV_Modify(fileNum,"col6", "")
				LV_Modify(fileNum,"col7", patPDF)
				LV_Modify(fileNum,"col8", tmp_meta)
			}
		}
	}

/* Read BSCI "bsc" folder
*/
	bscDir := pdfDir "bsc\patientData\"
	loop, Files, % bscDir "*", D												; Loop through subdirs of patientData
	{
		patDir := bscDir A_LoopFileName
		loop, files, % patDir "\*.bnk"											; Find the current nnnnnn.bnk file (inactive files are .bn_ files)
		{
			patBnk := patDir "\" A_LoopFileName
		}
		FileRead, bscBnk, % patBnk												; read the patBnk file
		tmp := parseDate(trim(stregX(bscBnk,"Save Date:",1,1,"[\r\n]",1)))		; get the DATE array
		tmp_name := readBnk("PatientLastName") ", " readBnk("PatientFirstName")
		tmp_dev := "BSCI " readBnk("SystemName") " " strX(readBnk("SystemModelNumber"),"",1,0,"-",1)
		tmp_ser := readBnk("SystemSerialNumber")
		if ((xl.selectSingleNode("//id[@date='" tmp.YYYY tmp.MM tmp.DD "']"))	; skip reprocessing if in either work or done lists
			&& (xl.selectSingleNode("//id[@ser='" tmp_ser "']"))) {
			continue
		}
		
		Loop, files, % patDir "\report\Combined*" tmp.MMM "-" tmp.DD "-" tmp.YYYY "*.pdf"
		{
			patPDF := patDir "\report\" A_LoopFileName							; find the appropriate PDF matching this .bnk file
		}
		
		fileNum += 1															; Add a row to the LV
		LV_Add("", tmp.YYYY tmp.MM tmp.DD)										; col1 is date
		LV_Modify(fileNum,"col2", tmp_name)
		LV_Modify(fileNum,"col3", tmp_dev)
		LV_Modify(fileNum,"col4", tmp_ser)
		LV_Modify(fileNum,"col5", "")
		LV_Modify(fileNum,"col6", "")
		LV_Modify(fileNum,"col7", patPDF)
		LV_Modify(fileNum,"col8", patBnk)
	}
	
	LV_ModifyCol(1, "Autohdr")
	LV_ModifyCol(2, "Autohdr")
	LV_ModifyCol(3, "Autohdr")
	LV_ModifyCol(4, "Autohdr")
	LV_ModifyCol(5, "Autohdr")
	LV_ModifyCol(6, "Autohdr")
	WinActivate, TRREAT Reports
return
}

parsePat:
{
	if !(fileNum := LV_GetNext()) {
		return
	}
	LV_GetText(pat_date,fileNum,1)
	LV_GetText(pat_name,fileNum,2)
	LV_GetText(pat_dev,fileNum,3)
	LV_GetText(pat_ser,fileNum,4)
	LV_GetText(pat_status,fileNum,5)
	LV_GetText(pat_paceart,fileNum,6)
	LV_GetText(fileIn,fileNum,7)
	LV_GetText(pat_meta,fileNum,8)
	LV_GetText(pat_report,fileNum,9)
	
	if (pat_report) {
		opt := (pat_status="Pending")
			? "Modify report|Regenerate report|Mark entered in PaceArt"		; Not signed yet
			: "Mark entered in PaceArt"										; Report signed
		tmp := cMsgBox(pat_name " report","Do what?",opt,"Q","")
		if (tmp="Close") {
			return
		}
		if instr(tmp,"Modify") {
			RunWait, % "WordPad.exe """ pat_report """"						; launch fileNam in WordPad
			return
		}
		if instr(tmp,"Regenerate") {
			removeNode("/root/work/id[@date='" pat_date "'][@ser='" pat_ser "']")
			xl.save(binDir "worklist.xml")
			FileDelete, % pat_report
			gosub fileLoop
			return
		}
		if instr(tmp,"PaceArt") {
			xl.setText("/root/work/id[@date='" pat_date "'][@ser='" pat_ser "']/paceart","True")
			xl.save(binDir "worklist.xml")
			gosub readList
			gosub readFiles
			return
		}
	}
	
	gosub fileLoop
	
Return
}

fileLoop:
{
	blocks := Object()
	fields := Object()
	labels := Object()
	fldval := Object()
	leads := Object()
	summBl := summ := sjmLog := ""
	FileDelete, temp.txt
	SplitPath, fileIn,,,,fileOut
	RunWait, pdftotext.exe -table "%fileIn%" temp.txt , , hide
	FileRead, maintxt, temp.txt
	cleanlines(maintxt)
	
	if (maintxt~="Medtronic,\s+Inc") {											; PM and ICD reports use common subs
		gosub Medtronic
	}
	else if (maintxt~="Boston Scientific Corporation") {
		gosub BSCI
	}
	else if instr(pat_dev,"SJM") {												; SJM device clicked from LV
		gosub SJM
	} 
	else {
		MsgBox No match!`nAttempt OCR on PDF?
	}
	
	return
}

SignScan:
{
	l_users := {}
	l_numusers :=
	l_tabs := 
	Loop, %reportDir%*.rtf
	{
		fileNam := RegExReplace(A_LoopFileName,"i)\.rtf")						; fileNam is name only without extension, no path
		fileIn := A_LoopFileFullPath											; fileIn has complete path \\childrens\files\HCCardiologyFiles\EP\TRREAT reports\pending\steve.rtf
				
		l_user := strX(fileNam,"",1,0,"-")										; Get assigned EP from filename
		l_mrn  := strX(fileNam,"-",1,1," ",1,1)
		l_name := stregX(fileNam,"-\d+ ",1,1," #",1)		
		l_ser  := stregX(fileNam," #",1,1," \d{6,8}",1)
		l_date := strX(fileNam," ",0,1,"",0)
		
		if !Object(l_users[l_user]) {											; this user not present yet in l_users[]
			l_tabs .= l_user . "|"												; add user to tab titles string
		}
		l_users[l_user,A_index] := {filename:fileNam							; creates l_users[l_user, x], where x is just a number
			, name:l_name
			, date:l_date
			, ser:l_ser}
	}
	gosub signGUI
	
Return
}

SignGUI:
{
	Gui, sign:Destroy
	Gui, sign:Add, Tab3, w600 vRepLV hwndRepH, % l_tabs								; Create a tab control (hwnd=RepH) with titles l_tabs
	Gui, sign:Default
	for k in l_users															; loop through l_users
	{
		tmpHwnd := "HW" . k														; unique Hwnd (HWTC, etc)
		Gui, Tab, % k															; go to tab for the user
		Gui, Add, ListView, % "-Multi Grid NoSortHdr x10 y30 w600 h200 gSignRep vUsr" k " hwnd" tmpHwnd, file|serial|Date|Name
		for v in l_users[k]														; loop through users in l_users
		{
			i := l_users[k,v]													; i is the element for each V
			LV_Add(""
				, i.filename													; this is a hidden column 
				, i.ser															; this is a hidden column
				, i.date
				, i.name)
		}
		LV_ModifyCol()
		LV_ModifyCol(1, "0")
		LV_ModifyCol(2, "0")
		LV_ModifyCol(3, "Autohdr")
		LV_ModifyCol(4, "AutoHdr")
	}
	GuiControl, ChooseString, RepLV, % substr(user,1,2)							; make this user the active tab
	Gui, Show, AutoSize, TRREAT Reports Pile											; show GUI
	
	return
}

ParseGuiClose:
ExitApp

SignGUIClose:
ExitApp

SignRep:
{
	l_tab := A_GuiControl
	Gui, Sign:ListView, % l_tab													; Select the LV passed to A_GuiControl
	if !(l_row := LV_GetNext()) {												; will be 0 if selected row is an empty row
		return
	}
	Gui, Sign:Hide
	LV_GetText(fileNam,l_row,1)													; get hidden fileNam from LV(l_row,1)
	LV_GetText(l_ser,l_row,2)													; get hidden serial number
	LV_GetText(l_date,l_row,3)
	
	gosub SignActGUI
	Gui, Sign:Show
Return	
}

SignActGui:
{
	Gui, Act:Destroy
	Gui, Act:Default
	Gui, Add, Text,, % fileNam
	Gui, Add, Button, vS_PDF gActPDF, View PDF
	Gui, Add, Button, vS_rev gActSign Disabled, SIGN && SEND
	Gui, Color, EEAA99
	
	if !FileExist(complDir fileNam ".pdf") {
		GuiControl, Act:Disable, S_PDF
	}
	Gui, Act:+AlwaysOnTop -MinimizeBox -MaximizeBox
	Gui, Show
	
	RunWait, % "WordPad.exe """ reportDir fileNam ".rtf"""						; launch fileNam in WordPad
	GuiControl, Act:Enable, S_rev
Return
}

ActPDF:
{
	pdfNam := complDir fileNam ".pdf"
	run, % pdfNam
Return
}

ActSign:
{
	Gui, Act:Hide
	l_tab := substr(l_tab,-1)													; get last 2 chars of l_tab
	if (substr(user,1,2)=l_tab) {												; first 2 chars of Citrix login match l_tab?
		if !(S_num := SignVerify(user)) {										; but fails authentication
			return
		}
	} else {																	; Different than assigned MD
		MsgBox, 52, , % "Signing report for " l_tab "?"
		IfMsgBox Yes															; signing someone else's report
		{
			if !(S_num := SignVerify(user)) {									; err if typed login doesn't match Citrix or typed billing code doesn't match
				return
			}
			FileRead, tmp, % reportDir fileNam ".rtf"							; read the generated RTF file
			tmp := RegExReplace(tmp
				, "Dictating Phy #\\tab <8:(\d{6})>\\par"						; replace the original billing code
				, "Dictating Phy #\tab <8:" S_num ">\par")						; with yours
			FileDelete, % reportDir fileNam ".rtf"
			FileAppend, % tmp, % reportDir fileNam ".rtf"						; generate a new RTF file
		} else {
			return																; not signing this report, return
		}
	}
	
	MsgBox, 52, ARE YOU SURE?, Send report to ESIG?
	IfMsgBox, Yes
	{
		FileCopy, % reportDir fileNam ".rtf", % "\\PPWHIS01\Apps$\3mhisprd\Script\impunst\crd.imp\" . fileNam . ".rtf"
	}
	FileMove, % reportDir fileNam ".rtf", % complDir fileNam ".rtf", 1			; move copy to "completed" folder
	
	xl.setText("/root/work/id[@date='" l_date "'][@ser='" l_ser "']/status","Signed")
	xl.save(binDir "worklist.xml")
	
	Gosub signScan																; regenerate file list
Return
}

SignVerify(user)
{
	global docs
	InputBox, userIn, Sign, Enter CIS user name
	if !(userIn=user) {
		MsgBox, 16,, Wrong user name!
		return error
	}
	InputBox, numIn, Sign, Enter billing code
	if !(numIn=docs[substr(userIn,1,2)]) {
		MsgBox, 16,, Wrong billing code!
		return error
	}
	return numIn
}

Medtronic:
{
	if (maintxt~="Adapta") {													; Scan Adapta family of devices
		gosub mdtAdapta
	} else if (maintxt~="Quick Look II") {										; or scan more current QuickLook II reports
		gosub mdtQuickLookII
	} else {																	; or something else
		MsgBox NO MATCH
		return
	}
	
	if (fileArg) {
		fileNum += 1
		LV_Add("", fldval["dev-date"])
		LV_Modify(filenum,"col2",fldval["dev-Name"])
		LV_Modify(fileNum,"col3",fldval["dev-IPG"])
		LV_Modify(fileNum,"col4",fldval["dev-IPG_SN"])
		LV_ModifyCol(1,"AutoHdr")
		LV_ModifyCol(2,"AutoHdr")
		LV_ModifyCol(3,"AutoHdr")
		LV_ModifyCol(4,"AutoHdr")
		Gui, Show
	}
	
	gosub fetchDem
	
	if (fetchQuit) {
		return
	}
	
	gosub makeReport
	
return	
}

mdtQuickLookII:
{
	;~ iniRep := strX(columns(maintxt,"Therapy Summary","Medtronic, Inc",0,"Pacing\s+\("),"",1,0,"Pacing",1,0)
	inirep := columns(maintxt,"Clinical Status","Therapy Summary",0,"Cardiac Compass")
	
	fields[1] := ["VF","VT-NS","VT","^AT/AF"]
	labels[1] := ["VF","VTNS","VT","ATAF"]
	scanParams(stregX(inirep,"Monitored",1,0,"Therapy",1),1,"event",1)
	
	inirep := columns(maintxt,"Therapy Summary","Medtronic, Inc",0,"Pacing\s+\(")
	fields[1] := ["VT/VF-Pace-Terminated","VT/VF-Shock-Terminated","VT/VF-Total Shocks","VT/VF-Aborted Charges"
				, "AT/AF-Pace-Terminated","AT/AF-Shock-Terminated","AT/AF-Total Shocks","AT/AF-Aborted Charges"]
	labels[1] := ["V_Paced","V_Shocked","V_Total","V_Aborted"
				, "A_Paced","A_Shocked","A_Total","A_Aborted"]
	scanParams(parseTable(stregX(inirep,"Therapy Summary",1,0,"Observations|Pacing",1)),1,"event",1)
	
	iniRep := instr(iniRep,"Event Counters") ? oneCol(iniRep) : iniRep
	if instr(iniRep,"Sensed") {
		fields[2] := ["Sensed","Paced"]
		labels[2] := ["Sensed","Paced"]
	} else {
		fields[2] := ["AS.*VS","AS.*VP","AP.*VS","AP.*VP"]
		labels[2] := ["AsVs","AsVp","ApVs","ApVp"]
	}
	scanParams(iniRep,2,"dev",1)
	
	fintxt := stregX(maintxt,"Final: Session Summary",1,0,"Medtronic, Inc.",0)
	
	dev := stregX(fintxt,"Session Summary",1,1,"initial interrogation\)",0,n)
	fields[1] := ["Device","Serial Number","Date of Visit"
				, "Patient","ID","Physician","`n"
				, "Device Information","`n"
				, "Device", "Implanted","`n"
				, "Atrial", "Implanted","`n"
				, "RV", "Implanted","`n"
				, "LV", "Implanted","`n"
				, "Device Status", "Battery Voltage","Remaining Longevity","`n"]
	labels[1] := ["IPG","IPG_SN","Encounter"
				, "Name","MRN","Physician","null"
				, "null","null"
				, "IPG0", "IPG_impl","null"
				, "Alead", "Alead_impl","null"
				, "RVlead", "RVlead_impl","null"
				, "LVlead", "LVlead_impl","null"
				, "IPG_stat", "IPG_voltage","IPG_longevity","null"]
	fieldvals(dev,1,"dev")
	if !instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") {
		fldval["dev-Physician"] := "Dr. " . trim(tmp," `n")
	}
	fldfill("dev-IPG","Medtronic " RegExReplace(fldval["dev-IPG"],"Medtronic "))
	fldfill("dev-Alead", RegExReplace(fldval["dev-Alead"],"---"))
	fldfill("dev-RVlead", RegExReplace(fldval["dev-RVlead"],"---"))
	fldfill("dev-LVlead", RegExReplace(fldval["dev-LVlead"],"---"))
	
	fintbl := stregX(fintxt,"",n+1,0,"Parameter Summary",1,n)
	fields[2] := ["Atrial.*-Lead Impedance"
				, "Atrial.*-Pacing Impedance"
				, "Atrial.*-Capture Threshold"
				, "Atrial.*-Measured On"
				, "Atrial.*-In-Office Threshold"
				, "Atrial.*-Programmed Amplitude"
				, "Atrial.*-Measured .*Wave"
				, "Atrial.*-Programmed Sensitivity"
			, "RV.*-Lead Impedance"
				, "RV.*-Pacing Impedance"
				, "RV.*-Defibrillation Impedance"
				, "RV.*-Capture Threshold"
				, "RV.*-Measured On"
				, "RV.*-In-Office Threshold"
				, "RV.*-Programmed Amplitude"
				, "RV.*-Measured .*Wave"
				, "RV.*-Programmed Sensitivity"
			, "LV.*-Lead Impedance"
				, "LV.*-Pacing Impedance"
				, "LV.*-Capture Threshold"
				, "LV.*-Measured On"
				, "LV.*-In-Office Threshold"
				, "LV.*-Programmed Amplitude"
				, "LV.*-Measured .*Wave"
				, "LV.*-Programmed Sensitivity"]
	labels[2] := ["A_imp","A_imp","A_cap","A_date","A_Pthr","A_output","A_Sthr","A_sensitivity"
				, "RV_imp","RV_imp","RV_HVimp","RV_cap","RV_date","RV_Pthr","RV_output","RV_Sthr","RV_sensitivity"
				, "LV_imp","LV_imp","LV_cap","LV_date","LV_Pthr","LV_output","LV_Sthr","LV_sensitivity"]
	scanParams(parseTable(fintbl),2,"leads",1)
	
	fintbl := stregX(fintxt,"Detection",1,0,"(Changes)|(Enhancement)|(Clinical Status)",1)
	fields[2] := ["Rates-AT/AF","Rates-VF","Rates-FVT","Rates-VT"
				, "Therapies-AT/AF","Therapies-VF","Therapies-FVT","Therapies-VT"]
	labels[2] := ["AT/AF","VF","FVT","VT"
				, "Rx_AT/AF","Rx_VF","Rx_FVT","Rx_VT"]
	scanParams(parseTable(fintbl),2,"detect",1)
	
	fintxt := stregX(maintxt,"Final: Parameters",1,0,"Medtronic, Inc.",0)
	
	param := RegExReplace(stregx(fintxt,"Pacing Summary.",1,1,"Pacing Details",1),"Mode","----",,1)				; Replace the title "Mode" to prevent interference with param scan
	fields[1] := ["Mode Switch","Mode","Lower","Upper Track","Upper Sensor","V. Pacing","V-V Pace Delay","Paced AV","Sensed AV"]
	labels[1] := ["Mode Switch","Mode","LRL","URL","USR","CRT_VP","CRT_VV","PAV","SAV"]							; Scan for "Mode Switch" first, so can find plain "Mode" second
	scanParams(onecol(param),1,"par",1)
	
	par := parsetable(stregx(fintxt,"Pacing Details",1,0,"AV Therapies",1))
	fields[2] := ["Atrial.*-Capture Management","Atrial.*-Pace Polarity","Atrial.*-Sense Polarity"
				, "RV.*-Capture Management","RV.*-Pace Polarity","RV.*-Sense Polarity"
				, "LV.*-Capture Management","LV.*-Pace Polarity","LV.*-Sense Polarity"]
	labels[2] := ["A_Cap_Mgt","A_Pol_pace","A_Pol_Sens"
				, "RV_Cap_Mgt","RV_Pol_pace","RV_Pol_sens"
				, "LV_Cap_Mgt","LV_Pol_pace","LV_Pol_sens"]
	scanParams(par,2,"leads",1)
	
	if (fldval["leads-A_imp"]||fldval["leads-A_cap"]||fldval["leads-A_Sthr"]) {
		normLead("RA"
				,fldval["dev-Alead"],fldval["dev-Alead_impl"]
				,fldval["leads-A_imp"],fldval["leads-A_cap"],fldval["leads-A_output"],fldval["leads-A_Pol_pace"]
				,fldval["leads-A_Sthr"],fldval["leads-A_Sensitivity"],fldval["leads-A_Pol_sens"])
	}
	if (fldval["leads-RV_imp"]||fldval["leads-RV_cap"]||fldval["leads-RV_Sthr"]) {
		normLead("RV"
				,fldval["dev-RVlead"],fldval["dev-RVlead_impl"]
				,fldval["leads-RV_imp"],fldval["leads-RV_cap"],fldval["leads-RV_output"],fldval["leads-RV_Pol_pace"]
				,fldval["leads-RV_Sthr"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"])
	}
	if (fldval["leads-LV_imp"]||fldval["leads-LV_cap"]||fldval["leads-LV_Sthr"]) {
		normLead("LV"
				,fldval["dev-LVlead"],fldval["dev-LVlead_impl"]
				,fldval["leads-LV_imp"],fldval["leads-LV_cap"],fldval["leads-LV_output"],fldval["leads-LV_Pol_pace"]
				,fldval["leads-LV_Sthr"],fldval["leads-LV_Sensitivity"],fldval["leads-LV_Pol_sens"])
	}
return
}

mdtAdapta:
{
	iniRep := columns(maintxt,"Clinical Status","Medtronic, Inc",0,"Pacing \(")
	fields[1] := ["Atrial High Rate.*","Ventricular High Rate.*"]
	labels[1] := ["AHR","VHR"]
	scanParams(RegExReplace(inirep,"Episodes: ","Episodes:  "),1,"event",1)
	
	iniRep := instr(iniRep,"Event Counters") ? oneCol(iniRep) : iniRep
	if instr(iniRep,"Sensed") {
		fields[2] := ["Sensed","Paced"]
		labels[2] := ["Sensed","Paced"]
	} else {
		fields[2] := ["AS.*VS","AS.*VP","AP.*VS","AP.*VP"]
		labels[2] := ["AsVs","AsVp","ApVs","ApVp"]
	}
	scanParams(iniRep,2,"dev",1)
	
	splTxt := "Final Report"
	fin := StrSplit(StrReplace(maintxt,splTxt, "``" splTxt),"``")
	Loop, % fin.length()
	{
		fintxt := fin[A_index]
		if (fintxt~=splTxt ".*Pacemaker Status") {
			dev := strX(fintxt,"Final Report",1,0,"Lead Status:",1,0)
			fields[1] := ["Pacemaker Model","Serial Number","Date of Visit"
						, "Patient Name", "ID", "Physician","`n"
						, "Pacemaker Model", "Implanted"
						, "Atrial Lead", "Implanted"
						, "Ventricular Lead", "Implanted"
						, "Pacemaker Status", "Estimated remaining longevity"
						, "Battery Status", "Voltage", "Current", "Impedance", "Lead Status"]
			labels[1] := ["IPG","IPG_SN","Encounter"
						,"Name", "MRN", "Physician","null"
						, "IPG0", "IPG_impl"
						, "Alead", "Alead_impl"
						, "Vlead", "Vlead_impl"
						, "IPG_stat", "IPG_longevity"
						, "Battery_stat", "IPG_voltage", "Current", "Impedance", "null"]
			fieldvals(dev,1,"dev")
			if !instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") {
				fldval["dev-Physician"] := "Dr. " . trim(tmp)
			}
			fldfill("dev-IPG","Medtronic " RegExReplace(fldval["dev-IPG"],"Medtronic "))
			fldfill("dev-Name",RegExReplace(fldval["dev-Name"],"i)DOB.*"))
			fldfill("dev-Alead",RegExReplace(fldval["dev-Alead"],"---"))
			fldfill("dev-RVlead",RegExReplace(fldval["dev-RVlead"],"---"))
			
			finleads := stregX(fintxt,"Lead Status:",1,0,"Capture Management",1,21)
			finleads := columns(finleads,"Lead Status:","In-Office Threshold",0,"Sensing Assurance")
			fields[2] := ["Atrial lead-Output Energy","Atrial Lead-Measured Current"
						, "Atrial lead-Measured Impedance","Atrial Lead-Pace Polarity","endcolumn"
						, "Ventricular lead-Output Energy","Ventricular Lead-Measured Current"
						, "Ventricular lead-Measured Impedance","Ventricular Lead-Pace Polarity","endcolumn"]
			labels[2] := ["A_output","A_curr","A_imp","A_pol","null"
						, "V_output","V_curr","V_imp","V_pol","null"]
			fldfill("leads-date",strX(finleads,"Lead Status: ",1,13,"`n",1,0,n))
			tbl := stregX(substr(finleads,n),"",1,0,"In-Office Threshold",1)
			scanParams(parseTable(tbl),2,"leads")
			
			thresh := onecol(stregX(fintxt,"Threshold Test Results.",1,1,"Medtronic Software",1))
			fldfill("leads-AP_thr",parseStrDur(oneCol(stregx(thresh,"Atrial Pacing Threshold",1,1,"\n\n",0))))
			fldfill("leads-VP_thr",parseStrDur(oneCol(stregx(thresh,"Ventricular Pacing Threshold",1,1,"\n\n",0))))
			fldfill("leads-AS_thr",trim(stregx(thresh,"P-wave",1,1,"\n\n",0)," `r`n"))
			fldfill("leads-VS_thr",trim(stregx(thresh,"R-wave",1,1,"\n\n",0)," `r`n"))
		}
		if (fintxt~=splTxt ".*Permanent Parameters") {
			perm := oneCol(stregX(fintxt,"Permanent Parameters(.*?)`n",1,1,"Medtronic Software",1))
			param := strx(perm,"Permanent Parameters",1,0,"Refractory/Blanking",1,0)
			fields[1] := ["Mode","Lower Rate","Upper Tracking Rate","Upper Sensor Rate","ADL Rate","Paced AV","Sensed AV"]
			labels[1] := ["Mode","LRL","URL","USR","ADL","PAV","SAV"]
			scanParams(fintxt,1,"par")
			
			param_A := stregX(perm,"Atrial Lead",1,0,"(Ventricular Lead)|(Additional/Interventions)|(Additional Features)",1)
			fields[2] := ["Amplitude","Pulse Width","Sensitivity","Pace Polarity","Sense Polarity","Capture Management"]
			labels[2] := ["Amp","PW","Sens","Pol_pace","Pol_sens","Cap_Mgt"]
			scanParams(param_A,2,"Alead")
			
			param_V := stregX(perm,"Ventricular Lead",1,0,"(Additional/Interventions)|(Additional Features)|(>>>end)",1)
			fields[3] := ["Amplitude","Pulse Width","Sensitivity","Pace Polarity","Sense Polarity","Capture Management"]
			labels[3] := ["Amp","PW","Sens","Pol_pace","Pol_sens","Cap_Mgt"]
			scanParams(param_V,3,"Vlead")
		}
		
		if (fldval["dev-Alead_impl"]) {
			normLead("RA"
					,fldval["dev-Alead"],fldval["dev-Alead_impl"]
					,fldval["leads-A_imp"],fldval["leads-AP_thr"]
					,(fldval["Alead-Amp"]) ? fldval["Alead-Amp"] " at " fldval["Alead-PW"] : ""
					,fldval["Alead-Pol_pace"]
					,fldval["leads-AS_thr"],fldval["Alead-Sens"],fldval["Alead-Pol_sens"])
		}
		if (fldval["dev-Vlead_impl"]) {
			normLead("RV"
					,fldval["dev-Vlead"],fldval["dev-Vlead_impl"]
					,fldval["leads-V_imp"],fldval["leads-VP_thr"]
					,(fldval["Vlead-Amp"]) ? fldval["Vlead-Amp"] " at " fldval["Vlead-PW"] : ""
					,fldval["Vlead-Pol_pace"]
					,fldval["leads-VS_thr"],fldval["Vlead-Sens"],fldval["Vlead-Pol_sens"])
		}
	}
return
}

BSCI:
{
	gosub bsciZoomView
	if (fileArg) {
		fileNum += 1
		LV_Add("", fldval["dev-date"])
		LV_Modify(filenum,"col2",fldval["dev-Name"])
		LV_Modify(fileNum,"col3",fldval["dev-IPG"])
		LV_Modify(fileNum,"col4",fldval["dev-IPG_SN"])
		LV_ModifyCol(1,"AutoHdr")
		LV_ModifyCol(2,"AutoHdr")
		LV_ModifyCol(3,"AutoHdr")
		LV_ModifyCol(4,"AutoHdr")
		Gui, Show
	}
	
	gosub fetchDem
	
	if (fetchQuit) {
		return
	}
	
	gosub makeReport
	
return	
}

bsciZoomView:
{
	txt := onecol(stregX(maintxt,"",1,0,"My Alerts",1))
	fields[1] := ["Combined.*Report","Date of Birth","Device","/","Report Created","Last Office Interrogation","Implant Date",">>>end"]
	labels[1] := ["Name","DOB","IPG","IPG_SN","Encounter","Last_ck","IPG_impl"]
	fieldvals(txt,1,"dev")
	fldfill("dev-IPG_SN",RegExReplace(fldval["dev-IPG_SN"],"Tachy.*"))
	fldfill("dev-IPG","Boston Scientific " RegExReplace(fldval["dev-IPG"],"Boston Scientific "))
	fldfill("dev-Physician",readBnk("PatientPhysFirstName") " " readBnk("PatientPhysLastName"))
	
	txt := stregX(maintxt,"My Alerts",1,0,"Leads Data",1)
	fields[1] := ["Battery","Approximate.*Explant:","`n"]
	labels[1] := ["Batt_stat","IPG_voltage","null"]
	fieldvals(txt,1,"dev")
	fldfill("dev-Battery_stat",readBnk("BatteryStatus.BatteryPhase"))
	
	txt := stregX(maintxt,"(Ventricular )?Tachy Settings",1,0,"Brady Settings",1)
	if instr(txt,"Atrial Tachy") {
		txt := columns(txt "endcolumn","","endcolumn",0,"Atrial Tachy")
	}
	fields[1] := ["VF","VT","Detection Rate"]
	labels[1] := ["VF","VHR","VHR"]
	scanParams(txt,1,"tachy")
	
	txt := columns(maintxt,"Brady Settings","(.*?)Software Version",0,"Pacing Output")
	txt := RegExReplace(txt,"(?<=\d)\s+(ppm|ms|mV)"," $1")
	fields[1] := ["Mode","Lower Rate Limit","Maximum Tracking Rate","Maximum Sensor Rate"
				, "Paced AV Delay","Sensed AV Delay"]
	labels[1] := ["Mode","LRL","URL","USR","PAV","SAV"]
	scanParams(txt,1,"par")
	
	txt := strX(txt,"Pacing Output",1,0) "endcolumn"
	tmp := substr(fldval["par-Mode"],1,1)
	fields[1] := ["Pacing Output","Sensitivity","Leads Configuration \(Pace/Sense\)","(Rate Adaptive Pacing|Sensors)","endcolumn"]
	labels[1] := ["outp0","sens0","pol0","adaptive","null"]
	fieldvals(txt,1,"par")
	if (fldval["par-outp0"]~="(Atrial|Ventricular)") {
		fields[2] := ["Atrial","Ventricular"]
		labels[2] := ["AP_thr","VP_thr"]
		scanParams(RegExReplace(fldval["par-outp0"],"(Atrial|Ventricular)","$1:  "),2,"leads",1)
	} else {
		fldfill("leads-" tmp "P_thr",fldval["par-outp0"])
	}
	if (fldval["par-sens0"]~="(Atrial|Ventricular)") {
		fields[2] := ["Atrial","Ventricular"]
		labels[2] := ["AS_thr","VS_thr"]
		scanParams(RegExReplace(fldval["par-sens0"],"(Atrial|Ventricular)","$1:  "),2,"leads",1)
	} else {
		fldfill("leads-" tmp "S_thr",fldval["par-sens0"])
	}
	if (fldval["par-pol0"]~="(Atrial|Ventricular)") {
		fields[2] := ["Atrial","Ventricular"]
		labels[2] := ["A_Pol_pace","RV_Pol_pace"]
		scanParams(RegExReplace(fldval["par-pol0"],"(Atrial|Ventricular)","$1:  "),2,"leads",1)
	} else {
		fldfill("leads-R" tmp "_Pol_pace",fldval["par-pol0"])
	}
	
	txt := stregX(maintxt,"Leads Data",1,0,"Settings",1)
	hdr := strX(txt,"",1,0,"`n",1)
	fields[1] := ["Most Recent-Intrinsic Amplitude","Most Recent-Pace Impedance","Most Recent-Pace Threshold","Most Recent-Shock Impedance"]
	labels[1] := ["sensing","imp","cap","HVimp"]
	if instr(txt,"Atrial") {
		scanParams(parseTable(hdr . stregX(txt ">>>","Atrial",1,1,"Ventricular|>>>",1)),1,"Alead")
	}
	if instr(txt,"Ventricular") {
		scanParams(parseTable(hdr . stregX(txt ">>>","Ventricular",1,1,">>>",1)),1,"Vlead")
	} 
	if !(txt~="Atrial|Ventricular") {
		tmp := substr(fldval["par-Mode"],1,1)
		if !(tmp~="[AV]") {
			tmp := cMsgBox("Single Chamber","What type of lead?","A|V","Q","")
			if (tmp = "Close") {
				return
			}
		}
		scanParams(parseTable(hdr "`n" stregX(txt ">>>","Intrinsic Amplitude",1,0,">>>",1)),1,tmp "lead")
	}
	fldfill("leads-RV_HVimp",fldval["Vlead-HVimp"])
	
	fldfill("dev-Alead"
		, printQ(readBnk("PatientLeadAManufacturer"),"###") 
		. printQ(readBnk("PatientLeadAModelNum"), " ###") 
		. printQ(readBnk("PatientLeadASerialNum"), " (serial ###)"))
	fldfill("Alead-Pol_pace",readBnk("PatientLeadAPolarity"))
	
	fldfill("dev-RVlead"
		, printQ(readBnk("PatientLeadV1Manufacturer"),"###") 
		. printQ(readBnk("PatientLeadV1ModelNum"), " ###") 
		. printQ(readBnk("PatientLeadV1SerialNum"), " (serial ###)"))
	fldfill("RVlead-Pol_pace",readBnk("PatientLeadV1Polarity"))
	
	fldfill("dev-LVlead"
		, printQ(readBnk("PatientLeadV2Manufacturer"),"###") 
		. printQ(readBnk("PatientLeadV2ModelNum"), " ###") 
		. printQ(readBnk("PatientLeadV2SerialNum"), " (serial ###)"))
	fldfill("LVlead-Pol_pace",readBnk("PatientLeadV2Polarity"))
	
	ctr := stregX(maintxt,"(Ventricular )?Tachy Counters",1,0,"$",0)
	ctrT := stregX(ctr,"(Ventricular )?Episode Counters",1,0,"Brady Counters",1)
	fields[1] := ["Total Episodes","Nonsustained Episodes","ATP Delivered","Shocks Delivered","Shocks Diverted","SVT Episodes.*"]
	labels[1] := ["VHR","VTNS","V_Paced","V_Shocked","V_Aborted","AHR"]
	scanParams(ctrT,1,"event",1)

	ctrB := stregX(ctr,"Brady Counters",1,0,"$",0)
	if (ctr~="(A Paced)|(V Paced)") {
		fields[1] := ["% A Paced","% V Paced"]
		labels[1] := ["AP","VP"]
	} else {
		fields[1] := ["% Paced"]
		labels[1] := [substr(fldval["par-Mode"],1,1) "P"]
	}
	scanParams(ctrB,1,"dev",1)
	
	if (fldval["Alead-imp"]||fldval["Alead-cap"]||fldval["Alead-sensing"]) {
		normLead("RA"
				,fldval["dev-Alead"],fldval["dev-Alead_impl"]
				,fldval["Alead-imp"],fldval["Alead-cap"],fldval["leads-AP_thr"],fldval["Alead-Pol_pace"]
				,fldval["Alead-sensing"],fldval["leads-AS_thr"],fldval["leads-RA_Pol_sens"])
	}
	if (fldval["Vlead-imp"]||fldval["Vlead-cap"]||fldval["Vlead-sensing"]) {
		normLead("RV"
				,fldval["dev-RVlead"],fldval["dev-RVlead_impl"]
				,fldval["Vlead-imp"],fldval["Vlead-cap"],fldval["leads-VP_thr"],fldval["RVlead-Pol_pace"]
				,fldval["Vlead-sensing"],fldval["leads-VS_thr"],fldval["leads-RV_Pol_sens"])
	}
	if (fldval["leads-LV_imp"]||fldval["leads-LV_cap"]||fldval["leads-LV_Sthr"]) {
		normLead("LV"
				,fldval["dev-LVlead"],fldval["dev-LVlead_impl"]
				,fldval["leads-LV_imp"],fldval["leads-LV_cap"],fldval["leads-LV_output"],fldval["LVlead-Pol_pace"]
				,fldval["leads-LV_Sthr"],fldval["leads-LV_Sensitivity"],fldval["leads-LV_Pol_sens"])
	}

	return
}

SJM:
{
	if (pat_meta) {																; SJM device with metadata (ICD exported)
		FileRead, sjmLog, % pat_meta
		gosub SJM_meta															; 
	} else {
		MsgBox No metafile found!
		return
	}
	
	gosub fetchDem
	
	if (fetchQuit) {
		return
	}
	
	gosub makeReport
	
return
}

SJM_meta:
{
	fields[1] := ["Device Model Name","Device Model Number"
				,"Device Serial Number","Implant Date: Device"
				, "Patient ID","Patient Name","Device Last Interrogation"
				, "Manufacturer:.*Atrial Lead","Model Number:.*Atrial Lead","Implant Date:.*Atrial Lead","Atrial Lead Serial Number"
				, "Manufacturer:.*RV Lead","Model Number:.*RV Lead","Implant Date:.*RV Lead","RV Lead Serial Number"
				, "Manufacturer:.*LV Lead","Model Number:.*LV Lead","Implant Date:.*LV Lead","LV Lead Serial Number"
				, "Battery Voltage","Longevity Estimate","Percent Paced In Ventricle","Percent Paced in Atrium"]
	labels[1] := ["IPG","IPG_model","IPG_SN","IPG_impl"
				, "MRN","Name","Encounter"
				, "Alead_man","Alead_model","Alead_impl","Alead_SN"
				, "RVlead_man","RVlead_model","RVlead_impl","RVlead_SN"
				, "LVlead_man","LVlead_model","LVlead_impl","LVlead_SN"
				, "IPG_voltage","IPG_longevity","VP","AP"]
	sjmVals(1,"dev")
	fldfill("dev-AP",RegExReplace(fldval["dev-AP"]," %"))
	fldfill("dev-VP",RegExReplace(fldval["dev-VP"]," %"))
	fldfill("dev-Encounter",RegExReplace(fldval["dev-Encounter"]," \d\d:\d\d:\d\d"))
	fldfill("dev-IPG","SJM " fldval["dev-IPG"] printQ(fldval["dev-IPG_model"], " ###"))
	fldfill("dev-Alead",fldval["dev-Alead_man"] 
		. printQ(fldval["dev-Alead_model"], " ###") printQ(fldval["dev-Alead_SN"], ", serial ###"))
	fldfill("dev-RVlead",fldval["dev-RVlead_man"] 
		. printQ(fldval["dev-RVlead_model"], " ###") printQ(fldval["dev-RVlead_SN"], ", serial ###"))
	fldfill("dev-LVlead",fldval["dev-LVlead_man"] 
		. printQ(fldval["dev-LVlead_model"], " ###") printQ(fldval["dev-LVlead_SN"], ", serial ###"))
	
	fields[1] := ["Atrial Pulse Configuration","Atrial Pulse Width","Atrial Pulse Amplitude"
				, "Atrial Sense Configuration","Atrial Sensitivity","(?<!\s)Atrial Signal Amplitude"
				, "Atrial Pacing Lead Impedance","A. .* Capture Threshold","A. .* Test Pulse Width"
				, "RV Pulse Configuration","RV Pulse Width","RV Pulse Amplitude"
				, "Ventricular Sense Configuration","Ventricular Sensitivity","(?<!\s)Ventricular Signal Amplitude"
				, "RV Pacing Lead Impedance","V. .* Capture Threshold","V. .* Test Pulse Width"
				, "HV Lead Impedance"]
	labels[1] := ["RA_Pol_Pace","RA_Pace_PW","RA_Pace_Amp"
				, "RA_Pol_Sens","RA_Sensitivity","RA_Thr_Sens"
				, "RA_imp","RA_Thr_Amp","RA_Thr_PW"
				, "RV_Pol_Pace","RV_Pace_PW","RV_Pace_Amp"
				, "RV_Pol_Sens","RV_Sensitivity","RV_Thr_Sens"
				, "RV_imp","RV_Thr_Amp","RV_Thr_PW"
				, "RV_HVimp"]
	sjmVals(1,"leads")
	
	fields[1] := ["\x1CMode\x1c","Base Rate","Maximum Tracking Rate","Maximum Sensor Rate"
				, "Paced AV Delay","Sensed AV Delay"]
	labels[1] := ["Mode","LRL","URL","USR"
				, "PAV","SAV"]
	sjmVals(1,"par")
	
	fields[1] := ["AT/AF Episodes","VT/VF Episodes"]
	labels[1] := ["ATAF","VT"]
	sjmVals(1,"event")
	
	if (fldval["leads-RA_imp"]||fldval["leads-RA_Pace_Amp"]||fldval["leads-RA_Thr_Sens"]) {
		normLead("RA"
				,fldval["dev-Alead"],fldval["dev-Alead_impl"],fldval["leads-RA_imp"]
				,fldval["leads-RA_Thr_Amp"] " @ " fldval["leads-RA_Thr_PW"]
				,fldval["leads-RA_Pace_Amp"] " @ " fldval["leads-RA_Pace_PW"]
				,fldval["leads-RA_Pol_pace"]
				,fldval["leads-RA_Sensitivity"],fldval["leads-RA_Thr_Sens"],fldval["leads-RA_Pol_sens"])
	}
	if (fldval["leads-RV_imp"]||fldval["leads-RV_Pace_Amp"]||fldval["leads-RV_Thr_Sens"]) {
		normLead("RV"
				,fldval["dev-RVlead"],fldval["dev-RVlead_impl"],fldval["leads-RV_imp"]
				,fldval["leads-RV_Thr_Amp"] " @ " fldval["leads-RV_Thr_PW"]
				,fldval["leads-RV_Pace_Amp"] " @ " fldval["leads-RV_Pace_PW"]
				,fldval["leads-RV_Pol_pace"]
				,fldval["leads-RV_Sensitivity"],fldval["leads-RV_Thr_Sens"],fldval["leads-RV_Pol_sens"])
	}
	if (fldval["leads-LV_imp"]||fldval["leads-LV_Pace_Amp"]||fldval["leads-LV_Thr_Sens"]) {
		normLead("LV"
				,fldval["dev-LVlead"],fldval["dev-LVlead_impl"],fldval["leads-LV_imp"]
				,fldval["leads-LV_Thr_Amp"] " @ " fldval["leads-LV_Thr_PW"]
				,fldval["leads-LV_Pace_Amp"] " @ " fldval["leads-LV_Pace_PW"]
				,fldval["leads-LV_Pol_pace"]
				,fldval["leads-LV_Sensitivity"],fldval["leads-LV_Thr_Sens"],fldval["leads-LV_Pol_sens"])
	}

return
}


fldfill(var,val) {
/*	Nondestructively fill fields
	If val is empty, return
	Otherwise populate with new value
*/
	global fldval
	
	if (val=="") {																; val is null
		return																	; do nothing
	}
	
	fldval[var] := trim(val," `t`r`n")											; set var as val
	
return
}

parseStrDur(txt) {
/*	Parse a block of text for Strength Duration values
	and return as a formatted string
*/
	if !instr(txt,"Strength Duration") {										; must be a Strength Duration block
		return Error
	}
	n := 1
	txt := stregX(txt,"Strength Duration",1,1,">>>end",1)
	loop
	{
		RegExMatch(txt,"O)\d+[.]\d+ V(.*?)\d+[.]\d+ ms",val,n)					; find "0.50 V @ 0.4 ms"
		res := ((res) ? res " and " : "") . val.value()							; append to RES (if RES already exists, prepend "and")
		n+=val.Len()															; starting point for next instance
	} until (A_index > val.count())
	
return res
}

parseTable(txt) {
/*	2nd version
	First scans title row for header positions
	Then reads result of each column in each row into res arrays
	Consider flag for fuzzy start of columns?
*/
	col := {}																	; col[] = column position
	pre := {}																	; pre[] = header prefix
	res := {}																	; res[] = result of each column
	lastpos := 1																; necessary for first pos
	Loop, parse, txt, `n`r
	{
		i := A_LoopField
		if !(trim(i)) {															; completely blank line (no field, no values)
			break																; is end of table
		}
		
		if (A_index=1) {														; parse header row
			loop
			{
				pos := RegExMatch(i "  ","(?<=(\s{2}))[^\s]",,lastpos)			; get position of next column from lastpos
				
				col.Push(pos)													; add position to col[] array (0 when no more matches)
				pre.Push(strX(substr(i,pos),"",1,0,"  ",1,2))					; add header value
				
				if !(pos) {														; break out when no more headers
					break
				}
				
				lastpos := pos+1												; new starting pos for next search
			}
			continue															; move to next line in txt
		}
		
		fld := strX(i,"",1,0,"  ",1,2,n)										; field name is first column
		
		if !(trim(fld)) {														; null fld means no value
			continue
		}
			
		for k in col															; iterate each column
		{
			p1 := col[k]														; pos 1 is start of col
			while !(substr(i,p1-2,2)="  ") {									; check that there are no non-space chars before p1
				p1 := p1-1														; back p1 up a space
				if (p1<n) {														; will run into fld if results line is blank
					break
				}
			}
			p2 := (col[k+1]) ? col[k+1] : strlen(i)+1							; pos 2 is start of next col, or last pos in row
			
			res[k] .= pre[k] "-" trim(fld) ":  " 								; concat res[] for each column
					. cleanSpace(substr(i,p1,p2-p1)) "`n"
		}
	}																			; All cols done
	for k in col																; iterate each column
	{
		if !(col[k]) {															; quit if last col
			break
		}
		result .= res[k] . "endcolumn`n"										; concat result of each res[] column
	}
Return result
}

oneCol(txt) {
/*	Break text block into a single column 
	based on logical break points in title (first) row
*/
	lastpos := 1
	Loop																		; Iterate each column
	{
		Loop, parse, txt, `n,`r													; Read through text block
		{
			i := A_LoopField
			
			if (A_index=1) {
				pos := RegExMatch(i	"  "										; Add "  " to end of scan string
								,"O)(?<=(\s{2}))[^\s]"							; Search "  text" as each column 
								,col
								,lastpos+1)										; search position to find next "  "
				
				if !(pos) {														; no match beyond, have hit max column
					max := true
				}
			}
			
			len := (max) ? strlen(i) : pos-lastpos								; length of string to return (max gets to end of line)
			
			str := substr(i,lastpos,len)										; string to return
			
			result .= str "`n"													; add to result
			;~ MsgBox % result
		}
		if !(pos) {																; break out if at max column
			break
		}
		lastpos := pos															; set next start point
	}
	return result . ">>>end"
}

scanParams(txt,blk,pre:="par",rx:="") {
	global fields, labels, fldval
	colstr = (?<=(\s{2}))(\>\s*)?[^\s].*?(?=(\s{2}))
	Loop, parse, txt, `n,`r
	{
		i := A_LoopField "  "
		set := trim(strX(i,"",1,0,"  ",1,2)," :")								; Get leftmost column to first "  "
		val := objHasValue(fields[blk],set,rx)
		;~ if !(val) {
			;~ continue
		;~ }
		
		RegExMatch(i															; Add "  " to end of scan string
				,"O)" colstr													; Search "  text  " as each column 
				,col1)															; return result in var "col1"
		RegExMatch(i
				,"O)" colstr
				,col2
				,col1.pos()+1)
		
		res := col1.value()
		if (col2.value()~="^(\>\s*)(?=[^\s])") {
			res := RegExReplace(col2.value(),"^(\>\s*)(?=[^\s])") " (changed from " col1.value() ")"
		}
		if (col2.value()~="(Monitor.*)|(\d{2}J.*)") {
			res .= ", Rx " cleanSpace(col2.value())
		}
			
		;~ MsgBox % pre "-" labels[blk,val] ": " res
		fldfill(pre "-" labels[blk,val], res)
	}
	return
}

readBnk(lbl) {
	global bscBnk
	return stregX(bscBnk,lbl ",",1,1,"[\r\n]+",1)
}

readSjm(lbl) {
/*	SJM nnnnnn.log files output from Merlin programmer
	read like a HL7 stream:
	el1 | el2 | el3 | el4 | el5 \n
	el1 = entry numberk
	el2 = label
	el3 = value
	el4 = units
	el5 = ?
	|   = chr(28) = x1C = "section seperator"
*/
	global sjmLog
	Loop, parse, sjmLog, `n,`r													; Read sjmLog
	{
		line := A_LoopField
		if !(line~="i)" lbl) {													; lbl regexmatch in line?
			continue															; no, move along
		}
		StringSplit, el, line, % Chr(28), `n									; yes, split line on chr(28)
		break																	; and break out of loop
	}
	return el3 . printQ(el4," ###") . printQ(el5," ###")						; return: value ( units)( whatever el5 is)
}

pmPrint:
{
	if !(enc_MD) {
		return
	}
	rtfBody := "\fs22\b\ul DEVICE INFORMATION\ul0\b0\par`n\fs18 "
	. fldval["dev-IPG"] ", serial number " fldval["dev-IPG_SN"] 
	. printQ(fldval["dev-IPG_impl"],", implanted ###") . printQ(fldval["dev-Physician"]," by ###") ". `n"
	. printQ(fldval["dev-IPG_voltage"],"Generator cell voltage ###. ")
	. printQ(fldval["dev-Battery_stat"],"Battery status is ###. ") . printQ(fldval["dev-IPG_Longevity"],"Remaining longevity ###. ") "`n"
	. printQ(fldval["par-Mode"],"Brady programming mode is ### with lower rate " fldval["par-LRL"])
	. printQ(fldval["par-URL"],", upper tracking rate ###")
	. printQ(fldval["par-USR"],", upper sensor rate ###")
	. printQ(fldval["par-ADL"],", ADL rate ###") . ". `n"
	. printQ(fldval["par-Cap_Mgt"],"Adaptive mode is ###. `n")
	. printQ(fldval["par-PAV"],"Paced and sensed AV delays are " fldval["par-PAV"] " and " fldval["par-SAV"] ", respectively. `n")
	. printQ(fldval["dev-Sensed"],"Sensed ###. ") . printQ(fldval["dev-Paced"],"Paced ###. ")
	. printQ(fldval["dev-AsVs"],"AS-VS ###  ") . printQ(fldval["dev-AsVp"],"AS-VP ###  ")
	. printQ(fldval["dev-ApVs"],"AP-VS ###  ") . printQ(fldval["dev-ApVp"],"AP-VP ###  ")
	. printQ(fldval["dev-AP"],"A-paced ###%. ") . printQ(fldval["dev-VP"],"V-paced ###%. ")
	. printQ(fldval["detect-AT/AF"],"AT/AF detection ###, " fldval["detect-Rx_AT/AF"] ". ")
	. printQ(fldval["detect-VF"],"VF detection ###, " fldval["detect-Rx_VF"] ". ")
	. printQ(fldval["detect-FVT"],"FVT detection ###, " fldval["detect-Rx_FVT"] ". ")
	. printQ(fldval["detect-VT"],"VT detection ###, " fldval["detect-VT"] ". ") . "\par `n"
	. "\fs22\par `n"
	. "\b\ul LEAD INFORMATION\ul0\b0\par`n\fs18 "
	
	for k in leads
	{
		printLead(k)
	}
	
	printEvents()
	
	gosub PrintOut

Return
}

printQ(var1,txt) {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
*/
	if (var1="") {
		return error
	}
	return RegExReplace(txt,"###",var1)
}

normLead(lead				; RA, RV, LV
		,model				; Model name/ser
		,date				; Date implanted
		,P_imp				; Pacing impedance
		,P_thr				; Pacing capture threshold
		,P_out				; Pacing programmed output
		,P_pol				; Pacing polarity
		,S_thr				; Sensing threshold
		,S_sens				; Sensing programmed sensitivity
		,S_pol)				; Sensing polarity
{
	global leads, fldval
	leads[lead,"model"] 	:= model
	leads[lead,"date"]		:= date
	leads[lead,"imp"]  		:= P_imp 
							. ((fldval["leads-" lead "_HVimp"]) 
							? ". Defib impedance " fldval["leads-" lead "_HVimp"] : "")
	leads[lead,"cap"]  		:= P_thr
	leads[lead,"output"]	:= P_out
	leads[lead,"pace pol"] 	:= P_pol
	leads[lead,"sens"]		:= S_thr
	leads[lead,"sensitivity"] := S_sens
	leads[lead,"sens pol"] 	:= S_pol
return
}

printLead(lead) {
	global rtfBody, leads
	rtfBody .= "\b " lead " lead: \b0 " 
	. printQ(leads[lead,"model"],"###" printQ(leads[lead,"date"],", implanted ###") ". ")
	. printQ(leads[lead,"imp"],"Pacing impedance ###. ")
	. printQ(leads[lead,"cap"],"Capture threshold ###. ")
	. printQ(leads[lead,"output"],"Pacing output ###. ")
	. printQ(leads[lead,"pace pol"],"Pacing polarity ###. ")
	. printQ(leads[lead,"sens"],((lead="RA")?"P":"")((lead="RV")?"R":"") "-wave sensing ###. ")
	. printQ(leads[lead,"sensitivity"],"Sensitivity ###. ")
	. printQ(leads[lead,"sens pol"],"Sensing polarity ###. ")
	. "\par `n"
}

printEvents()
{
	global rtfBody, fldval
	txt := ""
	. printQ(fldval["event-AHR"]?fldval["event-AHR"]:"","There were ### Atrial High Rate episodes. ")
	. printQ(fldval["event-VHR"]?fldval["event-VHR"]:"","There were ### Ventricular High Rate episodes. ")
	. printQ(fldval["event-VF"]?fldval["event-VF"]:"","### VF episodes detected. ")
	. printQ(fldval["event-VT"]?fldval["event-VT"]:"","### VT episodes detected. ")
	. printQ(fldval["event-VTNS"]?fldval["event-VTNS"]:"","### NS-VT episodes detected. ")
	. printQ(fldval["event-ATAF"]?fldval["event-ATAF"]:"","### AT/AF episodes detected. ")
	. printQ(fldval["event-V_Paced"]?fldval["event-V_Paced"]:"","### VT episodes pace-terminated. ")
	. printQ(fldval["event-V_Shocked"]?fldval["event-V_Shocked"]:"","### VT/VF episodes shock-terminated. ")
	. printQ(fldval["event-V_Aborted"]?fldval["event-V_Aborted"]:"","### VT/VF episodes aborted. ")
	. printQ(fldval["event-A_Paced"]?fldval["event-A_Paced"]:"","### AT episodes pace-terminated. ")
	. printQ(fldval["event-A_Shocked"]?fldval["event-A_Shocked"]:"","### AT/AF episodes shock-terminated. ")
	. printQ(fldval["event-A_Aborted"]?fldval["event-A_Aborted"]:"","### AT/AF episodes aborted. ")
	
	rtfBody .= printQ(txt,"\fs22\par\b\ul EVENTS\ul0\b0\par\fs18 `n###\par `n") 
return	
}

PrintOut:
{
	FormatTime, enc_dictdate, A_now, yyyy MM dd hh mm t
	FormatTime, enc_date, A_now, MM/dd/yyyy
	dt := parseDate(fldval["dev-Encounter"])
	enc_dt := dt.MM "/" dt.DD "/" dt.YYYY 
	for k in leads
	{
		ctLeads := A_Index
	}
	if (ctLeads = 1) {
		enc_type := "Single"
	} else if (ctLeads > 2) {
		enc_type := "Multi"
	} else {
		enc_type := "Dual"
	}
	enc_type := (instr(leads["RV","imp"],"Defib"))
		? "ICD " enc_type
		: "PM " enc_type
	
	rtfHdr := "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil\fcharset0 Arial;}}`n"
			. "{\*\generator Riched20 10.0.14393}\viewkind4\uc1 `n"
			. "\pard\tx2160\fs22\lang9 `n"
			. "Transcriptionist\tab "	"<TrID:crd> \par `n"
			. "Document Type\tab "		"<7:Q8> \par `n"
			. "Clinic Title code\tab "	"<1035:PACE> \par `n"
			. "Medical Record #\tab "	"<1:" fldval["dev-MRN"] ">\par `n"
			. "Patient Name\tab "		"<2:" fldval["dev-Name"] ">\par `n"
			. "CIS Encounter #\tab "	"<3: " substr("0000" . fldval["dev-Enc"], -11) " >\par `n"
			. "Dictating Phy #\tab "	"<8:" docs[enc_MD] ">\par `n"
			. "Dictation Date\tab "		"<13:" enc_date ">\par `n"
			. "Job #\tab "				"<15:e> \par `n"
			. "Service Date\tab "		"<31:" enc_dt ">\par `n"
			. "Surgery Date\tab "		"<6:" enc_dt "> \par `n"
			. "Attending Phy #\tab "	"<9:" docs[enc_MD] "> \par `n"
			. "Transcription Date\tab "	"<TS:" enc_dictdate "> \par `n"
			. "<EndOfHeader>\par `n"
			. "\par `n"
	
	rtfFtr := "`n\fs22\par\par`n\{SEC XCOPY\} \par`n\{END\} \par`n}`r`n"
	
	rtfBody := "\tx1620\tx5220\tx7040" 
	. "\fs22\b\ul PROCEDURE DATE\ul0\b0\par\fs18 `n"
	. enc_dt "\par\par\fs22 `n"
	. "\b\ul ENCOUNTER TYPE\ul0\b0\par\fs18 `n"
	. "Device interrogation " enc_type "\par\par\fs22 `n"
	. rtfBody . "\fs22\par `n" 
	. "\b\ul ENCOUNTER SUMMARY\ul0\b0\par\fs18 `n"
	. summ . "\par `n"
	. "\par `n"
	. "\pard\tx2700\tx5220\tx7040 `n"
	
	rtfOut := rtfHdr . rtfBody . rtfFtr
	
	nm := fldval["dev-Name"]
	fileOut :=	enc_MD "-" encMRN " " 
			.	(instr(nm,",") ? strX(nm,"",1,0,",",1,1) : strX(nm," ",1,1,"",0)) " "
			.	"#" fldval["dev-IPG_SN"] " "
			.	dt.YYYY dt.MM dt.DD
	
	FileDelete, temp.rtf															; delete and generate RTF fileOut.rtf
	FileAppend, %rtfOut%, temp.rtf
	
	RunWait, % "WordPad.exe temp.rtf"												; launch fileNam in WordPad
	MsgBox, 36, , Report looks okay?
	IfMsgBox, Yes
	{
		FileMove, temp.rtf, % reportDir fileOut ".rtf", 1								; move RTF to the final directory
		FileCopy, % fileIn, % complDir fileOut ".pdf", 1								; copy PDF to complete directory
		if FileExist(pat_meta) {
			FileCopy, % pat_meta, % complDir fileOut ".bnk", 1							; copy BNK to complete directory
		}
		
		t_now := A_Now
		edID := "/root/work/id[@ed='" t_now "']"
		xl.addElement("id","/root/work",{date: dt.YYYY dt.MM dt.DD, ser:fldval["dev-IPG_SN"], ed:t_now, au:user})
			xl.addElement("name",edID,fldval["dev-Name"])
			xl.addElement("dev",edID,fldval["dev-IPG"])
			xl.addElement("status",edID,"Pending")
			xl.addElement("paceart",edID,"")
			xl.addElement("file",edID,complDir fileOut ".pdf")
			xl.addElement("meta",edID,complDir fileOut ".bnk")
			xl.addElement("report",edID,reportDir fileOut ".rtf")
		xl.save(binDir "worklist.xml")
	}
	gosub readList
	gosub readFiles
	
	return
}

columns(x,blk1,blk2,excl:="",col2:="",col3:="",col4:="") {
/*	Returns string as a single column.
	x 		= input string
	blk1	= leading regex string to start block
	blk2	= ending regex string to end block
	excl	= if null (default), leave blk1 string in result; if !null, remove blk1 string
	col2	= string demarcates start of COLUMN 2
	col3	= string demarcates start of COLUMN 3
	col4	= string demarcates start of COLUMN 4
*/
	blk1 := rxFix(blk1,"O",1)													; Adds "O)" to blk1, pad whitespace with "\s+"
	blk2 := rxFix(blk2,"O",1)
	RegExMatch(x,blk1,blo1)														; Creates blo1 object out of blk1 match in x
	RegExMatch(x,blk2,blo2)														; necessary to get final string result of regex match
	
	col2 := RegExReplace(col2,"\s+","\s+")										; pad whitespace of col regex strings
	col3 := RegExReplace(col3,"\s+","\s+")
	col4 := RegExReplace(col4,"\s+","\s+")
	
	txt := stRegX(x,blk1,1,(excl) ? blo1.len() : 0,blk2)						; get string between blk1 and blk2
	;~ MsgBox % txt
	
	loop, parse, txt, `n,`r														; find position of columns 2, 3, and 4
	{
		i:=A_LoopField
		if (!(pos2) && (t:=RegExMatch(i,col2)))									; get first occurence of pos2
			pos2:=t
		if (!(pos3) && (t:=RegExMatch(i,col3)))
			pos3:=t
		if (!(pos4) && (t:=RegExMatch(i,col4)))
			pos4:=t
	}
	
	loop, parse, txt, `n,`r														; Generate column text
	{
		i:=A_LoopField
		txt1 .= substr(i,1,pos2-1) . "`n"										; Add to txt1
		
		if (col4) {																; Handle 4 columns
			pos4ck := pos4
			while !(substr(i,pos4ck-1,1)=" ") {									; Can adjust leftward until finds true start of col4
				pos4ck := pos4ck-1
			}
			txt4 .= substr(i,pos4ck) . "`n"										; Add to txt4
			txt3 .= substr(i,pos3,pos4ck-pos3) . "`n"							; Add to txt3
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"								; Add to txt2
			continue
		} 
		if (col3) {																; Handle 3 columns
			txt3 .= substr(i,pos3) . "`n"										; Add to txt3
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"								; Add to txt2
			continue
		}
		txt2 .= substr(i,pos2) . "`n"											; Handle 2 columns
	}
	return txt1 . txt2 . txt3 . txt4
}

strVal(hay,n1,n2,BO:="",ByRef N:="") {
/*	hay = search haystack
	n1	= needle1 begin string
	n2	= needle2 end string
	BO	= trim offset, true or false
	N	= return end position
*/
	opt := "Oim)"
	RegExMatch(hay,opt . n1 "(.*?)" n2 ,res,(BO)?BO:1)
	N := res.pos()+res.len(1)

	return trim(res[1]," :`n")
}

rxFix(hay,req,spc:="") {
/*	Adds required options to regex string, pad whitespace with "\s+"
	hay = haystack baseline regex string
	req = required option codes to insert
	spc = if !null, pad whitespace with "\s+"; if null, leave space alone
*/
	opts:="^[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"									; all the regex opts I could think of
	
	out := (hay~=opts)															; prepend the required opt string 
		? req . hay 
		: req ")" hay
	
	out := (spc) 																; pad whitespace if needed
		? RegExReplace(out,"\s+","\s+") 
		: out
	
	return out
}

cellvals(x,blk1:="",blk2:="",type:="") {
/*	Parses block of text (x) for subtable. Data rows are delmited by ":", followed by columns of info.
	Separate sub-subroutines for pertinent tables (e.g. leads, detections, therapies). Otherwise assume columns are RA, RV, and LV.
	If result extends to left, move start position of column until finds whitespace.
	x		= input text
	blk1	= leading string to start block
	blk2	= ending string for block
	type	= sub-subroutine needed; if blank, assumes RA RV LV
*/
	cells := []
	txt := StrX(x,blk1,1,0,blk2,1,StrLen(blk2))
	if (type="leads") {
		if (strlen(tmp:=cleanspace(strX(txt,"Implant Date",0,13,"",1,0))) < 10) {
			txt := "Lead Manufacturer & Model  Serial Number  Implant Date`nRA  No model specified`nRV  No model specified`n"
		}
		Loop, parse, txt, `n,`r
		{
			i:=trim(A_LoopField)
			j = %i%
			if !(j)
				continue
			if (instr(i,"Serial Number")) {
				pos2:=instr(i,"Serial Number")
				pos3:=instr(i,"Implant Date")
				continue
			}
			data0 := data1 := data2 := data3 := ""
			data0 := trim(substr(i,1,3))
			data1 := trim(substr(i,3,pos2-3))
			data2 := trim(substr(i,pos2,pos3-pos2-1))
			data3 := trim(substr(i,pos3))
			if (data0="RV") {
				if (rv_set=true)
					data0 := "RV2"
				rv_set := true
			}
			cells.Insert(data0)
			cells[data0] := {model:data1, serial:data2, date:data3}
		}
	} 
	if (type="detect") {
		Loop, parse, txt, `n,`r
		{
			i:=A_LoopField
			j = %i%
			if !(j)
				continue
			if (instr(i,"Detection ")) {
				pos2:=instr(i,"Rate")
				pos3:=instr(i,"Interval")
				pos4:=instr(i,"Therapy")
			}
			data0 := data1 := data2 := data3 := ""
			if (instr(i,":")) {											; a data row
				data0 := "det " trim(strX(i,,1,1,":",1,1,nn))
				data1 := trim(substr(i,nn+1,pos3-nn))					; Rate
				cleanSpace(data1)
				pos4ck := pos4
				while !(substr(i,pos4ck,1)=" ") {
					pos4ck := pos4ck-1
				}
				data3 := trim(substr(i,pos4ck))							; Therapy
				cleanSpace(data3)
				data2 := trim(substr(i,pos3,pos4ck-pos3))				; Interval
				cleanSpace(data2)
				cells.Insert(data0)
				cells[data0] := {rate:data1, interval:data2, therapy:data3}
			}
		}
	}
	if (type="ther") {
		Loop, parse, txt, `n,`r
		{
			i:=A_LoopField
			j = %i%
			if !(j)
				continue
			if (RegExMatch(i,"RA  \s*RV")) {
				pos2:=instr(i,"RA")
				pos3:=instr(i,"RV")
			}
			data0 := data1 := data2 := data3 := ""
			if (instr(i,":")) {
				data0 := trim(strX(i,,1,1,":",1,1,nn))
				if (data0="Pacing Imp.")
					data0:="Pacing Impedance"
				if (data0="Capt. Amp.")
					data0:="Capture Amplitude"
				if (data0="Capt. Dur.")
					data0:="Capture Duration"
				if (data0="Sens. Amp.")
					data0:="Sensing Amplitude"
				pos3ck := pos3
				while !(substr(i,pos3ck,1)=" ") {
					pos3ck := pos3ck-1
				}
				data2 := trim(substr(i,pos3ck))							; Therapy
				if (RegExMatch(data2,"(V|ms|mV)"))
					data2 := trim(substr(data2,1,-2))
				data1 := trim(substr(i,nn+1,pos3ck-nn-1))				; Interval
				
				cells.Insert(data0)
				cells[data0] := {RA:data1, RV:data2}
			}
		}
	} else {
		Loop, parse, txt, `n,`r											; loop through lines
		{
			i:=A_LoopField
			j = %i%
			if !(j)														; skip blank lines
				continue
			if (RegExMatch(i,"RA  \s*RV")) {										; mark position of cols 2 and 3
				pos2:=instr(i,"RV")
				pos3:=instr(i,"LV")
				continue
			} 
			data0 := data1 := data2 := data3 := ""
			if (instr(i,":")) {											; a data row
				data0 := strX(i,,1,1,":",1,1,nn)
				data1 := trim(substr(i,nn+1,pos2-nn-1))					; RA
				pos3ck := pos3
				data4 := trim(substr(i,pos3ck))
				while !(substr(i,pos3ck,1)=" ") {
					pos3ck := pos3ck-1
				}
				data3 := trim(substr(i,pos3ck))							; LV
				if (RegExMatch(data3,"(V)|(ms)|(mV)"))
					data3 := trim(substr(data3,1,-3))
				data2 := trim(substr(i,pos2,pos3-pos2))				; RV
				;units := trim(substr(i,pos4))
;				MsgBox,,% data0, % "RA '" data1 "'`nRV '" data2 "'`nLV '" data3 "'`nunits " units
				cells.Insert(data0)
				cells[data0] := {RA:data1, RV:data2, LV:data3, units:units}
			}
		}
	}
	return cells
}

fieldvals(x,bl,pre:="") {
/*	Matches field values and results. Gets text between FIELDS[k] to FIELDS[k+1]. Excess whitespace removed. Returns results in array BLK[].
	x	= input text
	bl	= which FIELD number to use
	pre	= label prefix
*/
	global fields, labels, fldval
	
	for k, i in fields[bl]
	{
		j := fields[bl][k+1]
		m := (j) ?	strVal(x,i,j,n,n)			;trim(stRegX(x,i,n,1,j,1,n), " `n")
				:	trim(strX(SubStr(x,n),":",1,1,"",0)," `n")
		;~ MsgBox % i " ~ " j "`n" pre "-" lbl "`n" m
		lbl := labels[bl][A_index]
		
		cleanSpace(m)
		cleanColon(m)
		;~ fldval[pre "-" lbl] := m
		fldfill(pre "-" lbl, m)
		;~ MsgBox % i " ~ " j "`n" pre "-" lbl "`n" m
		;~ formatField(pre,lbl,m)
	}
}

sjmVals(bl,pre:="") {
	global fields, labels
	
	for k,i in fields[bl]
	{
		lbl := labels[bl][A_Index]
		val := readSJM(i)
		if (val="") {
			continue
		}
		if instr(i,"impedance") {
			val := round(val) " Ohms"
		}
		if instr(i,"voltage") {
			val := round(val,3) " V"
		}
		if instr(i,"implant date") {
			val := RegExReplace(val," 00:00:00")
		}
		fldfill(pre "-" lbl, val)
	}
}

cleanlines(ByRef txt) {
	Loop, Parse, txt, `n, `r
	{
		i := A_LoopField
		if !(i){
			continue
		}
		newtxt .= i "`n"
	}
	txt := newtxt
	return txt
}

cleancolon(txt) {
	if substr(txt,1,1)=":" {
		txt:=substr(txt,2)
		txt = %txt%
	}
	return txt
}

cleanspace(ByRef txt) {
	StringReplace txt,txt,`n`n,%A_Space%, All
	StringReplace txt,txt,%A_Space%.%A_Space%,.%A_Space%, All
	loop
	{
		StringReplace txt,txt,%A_Space%%A_Space%,%A_Space%, UseErrorLevel
		if ErrorLevel = 0	
			break
	}
	return txt
}

ObjHasValue(aObj, aValue, rx:="") {
; modified from http://www.autohotkey.com/board/topic/84006-ahk-l-containshasvalue-method/	
	if (rx="med") {
		med := true
	}
    for key, val in aObj
		if (rx) {
			if (med) {													; if a med regex, preface with "i)" to make case insensitive search
				val := "i)" val
			}
			if (aValue ~= val) {
				return, key, Errorlevel := 0
			}
		} else {
			if (val = aValue) {
				return, key, ErrorLevel := 0
			}
		}
    return, false, errorlevel := 1
}

FetchDem:
{
	if !(fldval["dev-MRN"]~="^\d{6,7}$") {				; Check MRN parsed from PDF
		fldval["dev-MRN"] := ""
	}
	y := new XML(chipDir "currlist.xml")
	yArch := new XML(chipDir "archlist.xml")
	SNstring := "/root/id/diagnoses/device[@SN='" fldval["dev-IPG_SN"] "']"
	if IsObject(k := y.selectSingleNode(SNstring)) {							; Device SN found
		dx := k.parentNode
		id := dx.parentNode
		fldval["dev-MRN"] := id.getAttribute("mrn")								; set dev-MRN based on device SN
		
	} else if IsObject(k := yArch.selectSingleNode(SNstring)) {					; Look in yArch if not in y
		dx := k.parentNode
		id := dx.parentNode
		fldval["dev-MRN"] := id.getAttribute("mrn")
	}
	
	/*	set test vals
	*/
		;~ fldval["dev-MRN"] := 746275
		;~ fldval["dev-Enc"] := 98224013
	/*
	*/
	
	getDem := true
	fetchQuit := false
	
	gosub fetchGUI
	
	while (getDem) {									; Repeat until we get tired of this
		clipboard :=
		ClipWait, 2
		if !ErrorLevel {								; clipboard has data
			clk := parseClip(clipboard)
			if !ErrorLevel {															; parseClip {field:value} matches valid data
				if (clk.field = "Account Number") {
					fldval["dev-Enc"] := clk.value
					eventlog("MouseGrab Account Number " clk.value ".")
				}
				if (clk.field = "MRN") {
					fldval["dev-MRN"] := clk.value
					eventlog("MouseGrab MRN " clk.value ".")
				}
			}
			gosub fetchGUI							; Update GUI with new info
		}
	}
	if (fetchQuit) {
		return
	}
	
	EncNum := fldval["dev-Enc"]
	EncMRN := fldval["dev-MRN"]
	
	gosub saveChip
	
	return
}

fetchGUI:
{
	fYd := 30,	fXd := 90														; fetchGUI delta Y, X
	fX1 := 12,	fX2 := fX1+fXd													; x pos for title and input fields
	fW1 := 80,	fW2 := 190														; width for title and input fields
	fH := 20																	; line heights
	fY := 10																	; y pos to start
	EncNum := fldval["dev-Enc"]													; we need these non-array variables for the Gui statements
	EncMRN := fldval["dev-MRN"]
	demBits := (EncNum && EncMRN)												; clear the error check
	Gui, fetch:Destroy
	Gui, fetch:+AlwaysOnTop
	Gui, fetch:Add, Text, % "x" fX1 " w" fW1 " h" fH " c" , MRN
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH " cDefault", % fldval["dev-MRN"]
	Gui, fetch:Add, Text, % "x" fX1 " y" (fY += fYd) " w" fW1 " h" fH " c" , Encounter
	Gui, fetch:Add, Edit, % "readonly x" fX2 " y" fY-4 " w" fW2 " h" fH , % fldval["dev-Enc"]
	Gui, fetch:Add, Button, % "x" fX1+10 " y" (fY += fYD) " h" fH+10 " w" fW1+fW2 " gfetchSubmit " ((demBits)?"":"Disabled"), Submit!
	Gui, fetch:Show, AutoSize, % fldval["dev-Name"]
	return
}

fetchGuiClose:
	Gui, fetch:destroy
	getDem := false																	; break out of fetchDem loop
	fetchQuit := true
	eventlog("Manual [x] out of fetchDem.")
Return

parseClip(clip) {
/*	If clip matches "val1:val2" format, and val1 in demVals[], return field:val
	If clip contains proper Encounter Type ("Outpatient", "Inpatient", "Observation", etc), return Type, Date, Time
*/
	;~ global demVals
	demVals := ["Account Number","MRN"]
	
	StringSplit, val, clip, :															; break field into val1:val2
	if (ObjHasValue(demVals, val1)) {													; field name in demVals, e.g. "MRN","Account Number","DOB","Sex","Loc","Provider"
		return {"field":trim(val1)
				, "value":trim(val2)}
	}
	
	return Error																		; Anything else returns Error
}

fetchSubmit:
{
/*	some error checking
	Check for required elements
demVals := ["MRN","Account Number","DOB","Sex","Loc","Provider"]
*/
	Gui, fetch:Submit
	Gui, fetch:Destroy
	
	getDem := false
	return
}

saveChip:
{
	MRNstring := "/root/id[@mrn='" EncMRN "']"
	if !IsObject(y.selectSingleNode(MRNstring)) {
		y.addElement("id", "root", {mrn: EncMRN})								; No MRN node exists, create it.
		FetchNode("demog")
		FetchNode("diagnoses")													; Check for existing node in Archlist,
		FetchNode("prov")														; retrieve old Dx, Prov. Otherwise, create placeholders.
	}
	yID := y.selectSingleNode(MRNstring)
	if IsObject(yDev := yID.selectSingleNode("diagnoses/device")) {				; Clear out any existing Device node
		yDev.parentNode.removeChild(yDev)
	}
	y.addElement("device"
		,MRNstring "/diagnoses"
		,{	au:A_UserName
		,	ed:A_Now
		,	model:fldval["dev-IPG"]
		,	SN:fldval["dev-IPG_SN"]} )
	pmNowString := MRNstring "/diagnoses/device"
		y.addElement("mode", pmNowString, fldval["par-Mode"])
		y.addElement("LRL", pmNowString, fldval["par-LRL"])
		y.addElement("URL", pmNowString, fldval["par-URL"])
		y.addElement("AVI", pmNowString, fldval["par-SAV"])
		y.addElement("PVARP", pmNowString, fldval["par-PVARP"])
		y.addElement("ApThr", pmNowString, leads["RA","cap"])
		y.addElement("AsThr", pmNowString, leads["RA","sens"])
		y.addElement("VpThr", pmNowString, leads["RV","cap"])
		y.addElement("VsThr", pmNowString, leads["RV","sens"])
		y.addElement("Ap", pmNowString, leads["RA","output"])
		y.addElement("As", pmNowString, leads["RA","sensitivity"])
		y.addElement("Vp", pmNowString, leads["RV","output"])
		y.addElement("Vs", pmNowString, leads["RV","sensitivity"])
	WriteOut(MRNstring "/diagnoses", "device")
	
	return
}

makeReport:
{
	ck := cMsgBox("Title","Choose a text","Normal device check|none","Q","")
	if instr(ck,"normal") {
		ck := "This represents a normal device check. The patient denies any pacemaker related symptoms. "
			. "The battery status is normal. Sensing and capture thresholds are good. The lead impedances are normal. "
			. "Routine follow up per implantable device protocol. "
	} else {
		ck := ""
	}
	InputBox, summ, Enter SUMMARY text..., % ck,,,,,,,,% ck
	;~ if ErrorLevel {
		;~ MsgBox clicked CANCEL
	;~ }
	
	gosub checkEP
	
	gosub pmPrint
	
	return
}

checkEP:
{
/*	Find responsible EP
	and/or assign to someone
*/
	yID := y.selectSingleNode(MRNstring)
	
	if !(yEP := yID.selectSingleNode("prov").getAttribute("EP")) {						; Assign a primary EP in prov if it does not exist
		yEP := cMsgBox("No associated EP found"
						,"Assign a primary EP`nClose [x] if none"
						,"T. Chun|J. Salerno|S. Seslar"
						,"Q","")
		if !(yEP=="Close") {
			yID.selectSingleNode("prov").setAttribute("EP", yEP)
			yID.selectSingleNode("prov").setAttribute("au", A_UserName)
			yID.selectSingleNode("prov").setAttribute("ed", A_Now)
			
			writeOut(MRNstring,"prov")
		} 
	}
	
	enc_MD := cMsgBox("Assign report"
					, "Send report to:`n`n(primary EP is " yEP ").`n`n"
					. "Close [x] window to skip this step."
					, ((yEP="T. Chun") ? "*" : "") . "&TC|"
					. ((yEP="J. Salerno") ? "*" : "") . "&JS|"
					. ((yEP="S. Seslar") ? "*" : "") . "&SS"
					, "Q","")
	
	if (enc_MD="Close") {
		enc_MD := ""
	}
	
	Return
}

FetchNode(node) {
	global
	local x, clone
	if IsObject(yArch.selectSingleNode(MRNstring "/" node)) {		; Node arch exists
		x := yArch.selectSingleNode(MRNstring "/" node)
		clone := x.cloneNode(true)
		y.selectSingleNode(MRNstring).appendChild(clone)			; using appendChild as no Child exists yet.
	} else {
		y.addElement(node, MRNstring)								; If no node arch exists, create placeholder
	}
}

archNode(node) {
	global
	local clone
	clone := xl.selectSingleNode(node).cloneNode(true)
	xl.selectSingleNode("/root/done").appendChild(clone)
	removeNode(node)
	return
}

RemoveNode(node) {
	global
	local q
	q := xl.selectSingleNode(node)
	q.parentNode.removeChild(q)
}

WriteOut(path,node) {
/* 
	Prevents concurrent writing of y.MRN data. If someone is saving data (.currlock exists), script will wait
	approx 6 secs and check every 50 msec whether the lock file is removed. When available it creates clones the y.MRN
	node, loads a fresh currlist into Z (latest update), replaces the z.MRN node with the cloned y.MRN node,
	saves it, then reloads this currlist into Y.
*/
	global y, chipDir
	filecheck()
	FileOpen(chipDir ".currlock", "W")													; Create lock file.
	
	locPath := y.selectSingleNode(path)
	locNode := locPath.selectSingleNode(node)
	clone := locNode.cloneNode(true)											; make copy of y.node
	
	z := y																		; temp Z will be most recent good currlist
	
	if !IsObject(z.selectSingleNode(path "/" node)) {
		If instr(node,"id[@mrn") {
			z.addElement("id","root",{mrn: strX(node,"='",1,2,"']",1,2)})
		} else {
			z.addElement(node,path)
		}
	}
	zPath := z.selectSingleNode(path)											; find same "node" in z
	zNode := zPath.selectSingleNode(node)
	zPath.replaceChild(clone,zNode)												; replace existing zNode with node clone
	
	z.save(chipDir "currlist.xml")												; write z into currlist
	y := z																		; make Y match Z, don't need a file op
	FileDelete, % chipDir ".currlock"											; release lock file.
	return
}

filecheck() {
	if FileExist(chipDir ".currlock") {
		err=0
		Progress, , Waiting to clear lock, File write queued...
		loop 50 {
			if (FileExist(chipDir ".currlock")) {
				progress, %p%
				Sleep 100
				p += 2
			} else {
				err=1
				break
			}
		}
		if !(err) {
			progress off
			return error
		}
	} 
	progress off
	return
}

eventlog(event) {
	global user
	comp := A_ComputerName
	FormatTime, sessdate, A_Now, yyyy.MM
	FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
	name := binDir "..\logs\" . sessdate . ".log"
	txt := now " [" user "/" comp "] " event "`n"
	filePrepend(txt,name)
;	FileAppend, % timenow " ["  user "/" comp "] " event "`n", % "logs/" . sessdate . ".log"
}

FilePrepend( Text, Filename ) { 
/*	from haichen http://www.autohotkey.com/board/topic/80342-fileprependa-insert-text-at-begin-of-file-ansi-text/?p=510640
*/
    file:= FileOpen(Filename, "rw")
    text .= File.Read()
    file.pos:=0
    File.Write(text)
    File.Close()
}

parseDate(x) {
; Disassembles "2/9/2015" or "2/9/2015 8:31" into Yr=2015 Mo=02 Da=09 Hr=08 Min=31
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	if (x~="i)(\d{1,2})[\-\s\.]|(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\-\s\.](\d{2,4})") {		; 03 Jan 2016
		StringSplit, DT, x, %A_Space%-.
		return {"DD":zDigit(DT1), "MM":zDigit(objHasValue(mo,DT2)), "MMM":DT2, "YYYY":year4dig(DT3)}
	}
	if (x~="\d{1,2}_\d{1,2}_\d{2,4}") {											; 03_06_17 or 03_06_2017
		StringSplit, DT, x, _
		return {"MM":zDigit(DT1), "DD":zDigit(DT2), "MMM":mo[DT2], "YYYY":year4dig(DT3)}
	}
	if (x~="\d{4}-\d{2}-\d{2}") {
		StringSplit, DT, x, -
		return {"YYYY":DT1, "MM":DT2, "DD":DT3}
	}
	StringSplit, DT, x, %A_Space%
	StringSplit, DY, DT1, /
	StringSplit, DHM, DT2, :
	return {"MM":zDigit(DY1), "DD":zDigit(DY2), "YYYY":year4dig(DY3), "hr":zDigit(DHM1), "min":zDigit(DHM2), "Date":DT1, "Time":DT2}
}

niceDate(x) {
	if !(x)
		return error
	FormatTime, x, %x%, MM/dd/yyyy
	return x
}

year4dig(x) {
	if (StrLen(x)=4) {
		return x
	}
	if (StrLen(x)=2) {
		return (x<50)?("20" x):("19" x)
	}
	return error
}

zDigit(x) {
; Add leading zero to a number
	return SubStr("0" . x, -1)
}

stRegX(h,BS="",BO=1,BT=0, ES="",ET=0, ByRef N="") {
/*	modified version: searches from BS to "   "
	h = Haystack
	BS = beginning string
	BO = beginning offset
	BT = beginning trim, TRUE or FALSE
	ES = ending string
	ET = ending trim, TRUE or FALSE
	N = variable for next offset
*/
	;~ BS .= "(.*?)\s{3}"
	rem:="^[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"										; All the possible regexmatch options
	
	pos0 := RegExMatch(h,((BS~=rem)?"Oim"BS:"Oim)"BS),bPat,((BO)?BO:1))
	/*	Ensure that BS begins with at least "Oim)" to return [O]utput, case [i]nsensitive, and [m]ultiline searching
		Return result in "bPat" (beginning pattern) object
		If (BO), start at position BO, else start at 1
	*/
	pos1 := RegExMatch(h,((ES~=rem)?"Oim"ES:"Oim)"ES),ePat,pos0+bPat.len())
	/*	Ensure that ES begins with at least "Oim)"
		Resturn result in "ePat" (ending pattern) object
		Begin search after bPat result (pos0+bPat.len())
	*/
	bmod := (BT) ? bPat.len() : 0
	emod := (ET) ? 0 : ePat.len()
	N := pos1+emod
	/*	Final position is start of ePat match + modifier
		If (ET), add nothing, else add ePat.len()
	*/
	return substr(h,pos0+bmod,(pos1+emod)-(pos0+bmod))
	/*	Start at pos0
		If (BT), add bPat.len(), else stay at pos0 (will include BS in result)
		substr length is position of N (either pos1 or include ePat) less starting pos0
	*/
}

#Include strx.ahk
#Include xml.ahk
#Include CMsgBox.ahk
