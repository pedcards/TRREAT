/*	TRREAT - The Rhythm Recording Electronic Analysis Transmogrifier - PM
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%
;~ FileInstall, pdftotext.exe, pdftotext.exe
#Include includes

Progress,100,Checking paths...,TRREAT
SplitPath, A_ScriptDir,,fileDir
user := A_UserName
if instr(user,"octe") {
	user:="TC"
}
IfInString, fileDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	isDevt := true
	readIni("adminpaths")
	eventlog(">>>>> Started in DEVT mode.")
} else {
	isDevt := false
	readIni("paths")
	eventlog(">>>>> Started in PROD mode. " A_ScriptName " ver " substr(tmp,1,12))
}
binDir:=trreatDir "bin\"
reportDir:=trreatDir "pending\"
complDir:=trreatDir "completed\"
paceartDir:=trreatDir "paceart\"
hl7inDir:=trreatDir "incoming\"

worklist := reportDir "worklist.xml"

initHL7()
hl7DirMap := {}

eventLog(">>>>> Session started...")
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

Progress, off

/*	Main part ==================================================================
*/
MainLoop:
{
	newTxt := Object()
	blk := Object()
	blk2 := Object()
	
	docs := {"TC":"783118","JS":"343079","SS":"358945"
			, "TCeml":"terrence.chun","JSeml":"jack.salerno","SSeml":"stephen.seslar"}
	if ObjHasKey(docs,substr(user,1,2)) {												; User is in docs[]
		role := "Sign"																	; set role to "Sign"
	} else {
		role := "Parse"
	}
	if instr(user,"TC") {																; But if is TC
		role := cMsgBox("Administrator"													; offer to either PARSE or SIGN
			, "Enter ROLE:"
			, "*&Parse PDFs|&Sign reports"
			, "Q","")
	}
	
	if instr(role,"Sign") {
		eventLog("SIGN module")
		xl := new XML(worklist)									; otherwise load existing worklist
		gosub signScan
	}

	if instr(role,"Parse") {
		eventLog("PARSE module")
		gosub parseGUI
	}
	
	WinWaitClose, TRREAT Reports
	eventLog("<<<<< Session ended.")
	ExitApp
	
/*	End Main part ==============================================================
*/
}

parseGUI:
{
	Gui, Parse:Destroy
	Gui, Parse:Default
	Gui, Add, Tab3, vWQtab +HwndWQtab,Interrogations|Paceart saves
	
	Gui, Tab, Interr
	Gui, Add, Listview, w800 -Multi Grid r12 gparsePat vWQlv hwndHLV, Date|Name|Device|Serial|Status|PaceArt|FileName|MetaData|Report
	Gui, Tab, Paceart
	Gui, Add, Listview, w800 -Multi Grid r12 gparsePat vWQlvP hwndHLVp, Date|Name|Device|Serial|Status|PaceArt|FileName|MetaData|Report
	
	gosub readList																		; read the worklist
	
	gosub readFiles																		; scan the folders
	
	fixWqlvCols("WQlv")
	
	progress, off
	
	Gui, Show,, TRREAT Reports and File Manager
	WinActivate, TRREAT Reports
	return
}

fixWQlvCols(lv) {
	Gui, ListView, % lv
	LV_ModifyCol(1, "Autohdr")													; when done, reformat the col widths
	LV_ModifyCol(2, "Autohdr")
	LV_ModifyCol(3, "Autohdr")
	LV_ModifyCol(4, "Autohdr")
	LV_ModifyCol(5, "Autohdr")
	LV_ModifyCol(6, "Autohdr")
	LV_ModifyCol(7, "0")														; hide the filename col
	LV_ModifyCol(8, "0")														; hide the metadata col
	LV_ModifyCol(9, "0")														; hide the report col
	return
}


readList:
{
	progress,20,Reading worklist,Scanning files
	
	Gui, ListView, WQlv
	
	LV_Delete()
	fileNum := 0																; Start the worklist at fileNum 0
	if !FileExist(worklist) {
		xl := new XML("<root/>")												; Create new XML if doesn't exist
		xl.addElement("work", "root", {ed: A_Now})
		xl.addElement("done", "root", {ed: A_Now})
		xl.save(worklist)
		eventlog("New worklist.xml created.")
	} else {
		xl := new XML(worklist)													; otherwise load existing worklist
		eventlog("Worklist.xml loaded.")
	}
	Loop, % (w_id := xl.selectNodes("/root/work/id")).length					; scan through each <id>
	{
		k := w_id.item(A_Index-1)												; put it into k
		if !IsObject(k) {														; skip if empty
			eventlog("Empty node skipped.")
			continue
		}
		tmp := []
		tmp["date"] := k.getAttribute("date")
		tmp["name"] := k.selectSingleNode("name").text
		tmp["dev"]  := k.selectSingleNode("dev").text
		tmp["ser"]  := k.getAttribute("ser")
		tmp["status"] := k.selectSingleNode("status").text
		tmp["paceart"] := k.selectSingleNode("paceart").text
		tmp["file"] := k.selectSingleNode("file").text
		tmp["meta"] := k.selectSingleNode("meta").text
		tmp["report"] := k.selectSingleNode("report").text
		if ((tmp.report) && (tmp.status="Signed") && (tmp.paceart)) {				; REPORT and SIGNED and PACEART all true
			fileNum += 1
			LV_Add("", tmp.date)
			LV_Modify(fileNum,"col2", tmp.name)										; add marker line if in DONE list
			LV_Modify(fileNum,"col3", "[DONE]")
			archNode("/root/work/id[@date='" tmp.date "'][@ser='" tmp.ser "']")		; copy ID node to DONE
			xl.save(worklist)
			eventlog("Node " tmp.date "/" tmp.ser "/" tmp.name " archived.")
			continue
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
	eventlog("Parse listview generated.")
return
}

readFiles:
{
	readFilesMDT()
	readFilesSJM()
	readFilesBSCI()
	readFilesPaceart()
	
return
}

readFilesMDT() {
/*	Read root - usually MEDT files
*/
	global pdfDir, xl, filenum, WQlvP, WQlv, HLVp, HLV
	
	progress, 40,Medtronic
	Loop, files, % pdfDir "*.pdf"												; read all PDFs in root
	{
		tmp := []
		tmp.file := A_LoopFileName												; next file in PDFdir
		if instr(tmp.maxstr,tmp.file) {											; in skiplist?
			continue
		}
		tmp.max := 1															; reset max k counter
		Loop, files, % pdfDir strX(tmp.file,"",1,0,"_",0,1) "*.pdf"				; loop through all files with this "prefix"
		{
			i := A_LoopFileName													; i is filename in this inner loop
			n := substr(i,instr(i,"_",,-1))										; n is string up to final _#
			k := strX(i,"_",n,1,".",1)											; k is # between _ and .pdf
			if (k > tmp.max) {													; greater than previous kmax?
				j := substr(i,1,instr(i,"_",,-1)) (tmp.max) ".pdf"				; j is filename of previous kmax
				FileMove, % pdfDir j, % pdfDir j ".old"							; rename it to j.pdf.old
				tmp.max := k													; new kmax
				tmp.file := i													; set patPDF as this new max (for when exits)
				tmp.maxstr .= i "`n"											; add to string of files to subsequently ignore
				eventlog("MDT: newer version of " j " _" k )
			}
		}
		fnam := StrSplit(tmp.file,"_")
		tmp.name := fnam.1			; strX(tmp.file,"",1,0,"_",1,1,n)
		tmp.ser := fnam.2			; strX(tmp.file,"_",n-1,1,"_",1,1,n)
		tmp.type := fnam.3
		tmp.date := parseDate(fnam.4 "-" fnam.5 "-" fnam.6).YMD
		tmp.file := pdfDir tmp.file
		tmp.node := "id[@date='" tmp.date "'][@ser='" tmp.ser "']"
		
		if IsObject(xl.selectSingleNode("/root/work/" tmp.node)) {
			eventlog("MDT: Skipping " tmp.file ", already in worklist.")
			continue															; skip reprocessing in WORK list
		}
		if IsObject(xl.selectSingleNode("/root/done/" tmp.node)) {
			fileNum += 1
			LV_Add("", tmp.date)
			LV_Modify(fileNum,"col2", tmp.name)									; add marker line if in DONE list
			LV_Modify(fileNum,"col3", "[DONE]")
			eventlog("MDT: File " tmp.file " already DONE.")
			continue
		}
		
		fileNum += 1															; Add a row to the LV
		LV_Add("", tmp.date)													; col1 is date
		LV_Modify(fileNum,"col2", tmp.name)
		LV_Modify(fileNum,"col3", "Medtronic")
		LV_Modify(fileNum,"col4", tmp.ser)
		LV_Modify(fileNum,"col5", "")
		LV_Modify(fileNum,"col6", "")
		LV_Modify(fileNum,"col7", tmp.file)
		LV_Modify(fileNum,"col8", "")
	}
	
	return
}

readFilesSJM() {
/* Read SJM "PDFs" folder
*/
	global pdfDir, xl, filenum, WQlvP, WQlv, HLVp, HLV
	
	progress, 60,St Jude/Abbott
	sjmDir := pdfDir "PDFs\Live.combined"
	Loop, Files, % sjmDir "\*", D
	{
		DateDir := A_LoopFileName
		Loop, Files, % sjmDir "\" DateDir "\*", D
		{
			tmp := []
			patDir := A_LoopFileName
			tmp.date := RegExReplace(DateDir,"-")
			tmp.name := stregx(patDir,"",1,0,"_\d{2,}",1,n)
			tmp.dev  := stregx(patDir,"_",n,1,"_",1,n)
			tmp.ser  := strx(patDir,"_",n,1,"",0)
			tmp.node := "id[@date='" tmp.date "'][@ser='" tmp.ser "']"
			
			if IsObject(xl.selectSingleNode("/root/work/" tmp.node)) {
				eventlog("SJM: Skipping " DateDir "\" tmp.ser ", already in worklist.")
				continue														; skip reprocessing in WORK list
			}
			if IsObject(xl.selectSingleNode("/root/done/" tmp.node)) {
				fileNum += 1
				LV_Add("", tmp.date)
				LV_Modify(fileNum,"col2", tmp.name)								; add marker line if in DONE list
				LV_Modify(fileNum,"col3", "[DONE]")
				eventlog("SJM: File " DateDir "\" tmp.ser " already DONE.")
				continue
			}
			
			Loop, Files, % sjmDir "\" DateDir "\" patDir "\*.pdf", F
			{
				tmp.file := A_LoopFileName
				tmp.full := A_LoopFileFullPath
				tmp.dev := strx(tmp.file,"",1,0,"_",1,1)
				Loop, Files, % pdfDir "*.log", F
				{
					k := RegExReplace(A_LoopFileName,".log")
					if InStr(tmp.ser,k) {
						tmp.meta := pdfDir k ".log"
						eventlog("SJM: " tmp.file " metafile " k ".log found")
					}
				}
				if !(tmp.meta) {
					eventlog("SJM: " tmp.file " has no metafile, skipping.")
					continue
				}
				fileNum += 1													; Add a row to the LV
				LV_Add("", tmp.date)											; col1 is date
				LV_Modify(fileNum,"col2", tmp.name)
				LV_Modify(fileNum,"col3", "SJM " tmp.dev)
				LV_Modify(fileNum,"col4", tmp.ser)
				LV_Modify(fileNum,"col5", "")
				LV_Modify(fileNum,"col6", "")
				LV_Modify(fileNum,"col7", tmp.full)
				LV_Modify(fileNum,"col8", tmp.meta)
			}
		}
	}
	
	return
}

readFilesBSCI() {
/* Read BSCI "bsc" folder
*/
	global xl, pdfDir, filenum, bscBnk, WQlvP, WQlv, HLVp, HLV
	
	progress, 80,Boston Scientific
	tmp := []
	bscDir := pdfDir "bsc\patientData\"
	loop, Files, % bscDir "*", D												; Loop through subdirs of patientData
	{
		patDir := bscDir A_LoopFileName
		loop, files, % patDir "\*.bnk"											; Find the current nnnnnn.bnk file (inactive files are .bn_ files)
		{
			tmp.bnk := patDir "\" A_LoopFileName
			eventlog("BSC: Metafile " A_LoopFileName " found.")
		}
		FileRead, bscBnk, % tmp.bnk												; need bscBnk for readBnk
		td := trim(stregX(bscBnk,"Save Date:",1,1,"\R",1))						; get the DATE array
		td := parseDate(RegExReplace(td," ","-"))
		tmp.date := td.YMD
		tmp.name := readBnk("PatientLastName") ", " readBnk("PatientFirstName")
		tmp.dev := "BSCI " readBnk("SystemName") " " strX(readBnk("SystemModelNumber"),"",1,0,"-",1)
		tmp.ser := readBnk("SystemSerialNumber")
		tmp.node := "id[@date='" tmp.date "'][@ser='" tmp.ser "']"
		
		if IsObject(xl.selectSingleNode("/root/work/" tmp.node)) {
			eventlog("BSC: Skipping " tmp.date "\" tmp.ser ", already in worklist.")
			continue															; skip reprocessing in WORK list
		}
		if IsObject(xl.selectSingleNode("/root/done/ " tmp.node)) {
			fileNum += 1
			LV_Add("", tmp.date)
			LV_Modify(fileNum,"col2", tmp.name)									; add marker line if in DONE list
			LV_Modify(fileNum,"col3", "[DONE]")
			eventlog("BSC: File " tmp.date "\" tmp.ser " already DONE.")
			continue
		}
		
		Loop, files, % patDir "\report\Combined*" td.MMM "-" td.DD "-" td.YYYY "*.pdf"
		{
			tmp.file := A_LoopFileFullPath										; find the appropriate PDF matching this .bnk file
			eventlog("BSC: " A_LoopFileName " found.")
		}
		
		fileNum += 1															; Add a row to the LV
		LV_Add("", tmp.date)										; col1 is date
		LV_Modify(fileNum,"col2", tmp.name)
		LV_Modify(fileNum,"col3", tmp.dev)
		LV_Modify(fileNum,"col4", tmp.ser)
		LV_Modify(fileNum,"col5", "")
		LV_Modify(fileNum,"col6", "")
		LV_Modify(fileNum,"col7", tmp.file)
		LV_Modify(fileNum,"col8", tmp.bnk)
	}
	
	return
}

readFilesPaceart() {
/*	read exported PDF reports from Paceart
	in .\paceart\ folder
*/
	global paceartDir, WQlvP, WQlv, HLVp, HLV
	
	progress, 100, Paceart imports
	
	Gui, Listview, WQLVp
	
	loop, files, % paceartDir "*.xml"
	{
		fileIn := paceartDir A_LoopFileName
		dem := []
		if (fileIn~="WQ.xml$") {
			fnam := StrSplit(RegExReplace(A_LoopFileName,"WQ.xml$"),"_")
			dem.mrn := fnam.1
			dem.nameL := fnam.2
			dem.encdate := fnam.3
		} 
		else {
			y := new XML(fileIn)
			dem.nameL := y.selectSingleNode("//PatientRecord/Demographics/LastName").text
			dem.mrn := y.selectSingleNode("//PatientRecord/IDs/ID[Type='MRN']/Value").text
			dem.encdate := parseDate(y.selectSingleNode("//PatientRecord/Encounters/Encounter/Date").text).YMD
			dem.devtype := y.selectSingleNode("//PatientRecord/ActiveDevices/PatientActiveDevice/Device/Type").text
			if !(dem.nameL && dem.mrn && dem.devtype) {									; probably not a Paceart report
				continue																; skip it
			}
			fileOut := paceartDir . dem.mrn "_" dem.nameL "_" dem.encdate "WQ.xml"
			FileMove, %fileIn%, %fileOut%, 1
			fileIn := fileOut
		}
		fileNum += 1																	; Add a row to the LV
		LV_Add("", dem.encdate)															; col1 is date
		LV_Modify(fileNum,"col2", dem.nameL)
		;~ LV_Modify(fileNum,"col3", "")
		;~ LV_Modify(fileNum,"col4", "")
		;~ LV_Modify(fileNum,"col5", "")
		;~ LV_Modify(fileNum,"col6", "")
		LV_Modify(fileNum,"col7", fileIn)
		;~ LV_Modify(fileNum,"col8", "")
	}
	fixWQlvCols("WQLVp")

	return	
}

ParseName(x) {
/*	Determine first and last name
*/
	if (x="") {
		return error
	}
	x := trim(x)																		; trim edges
	x := RegExReplace(x,"\'","^")														; replace ['] with [^] to avoid XPATH errors
	x := RegExReplace(x," \w "," ")														; remove middle initial: Troy A Johnson => Troy Johnson
	x := RegExReplace(x,"(,.*?)( \w)$","$1")											; remove trailing MI: Johnston, Troy A => Johnston, Troy
	x := RegExReplace(x,"i),?( JR| III| IV)$")											; Filter out name suffixes
	x := RegExReplace(x,"\s+"," ",ct)													; Count " "
	
	if instr(x,",") 																	; Last, First
	{
		last := trim(strX(x,"",1,0,",",1,1))
		first := trim(strX(x,",",1,1,"",0))
	}
	else if (ct=1)																		; First Last
	{
		first := strX(x,"",1,0," ",1)
		last := strX(x," ",1,1,"",0)
	}
	else if (ct>1)																		; James Jacob Jingleheimer Schmidt
	{
		x0 := x																			; make a copy to disassemble
		n := 1
		Loop
		{
			x0 := strX(x0," ",n,1,"",0)													; cut from first " " to end
			if (x0="") {
				q := trim(q,"|")
				break
			}
			q .= x0 "|"																	; add to button q
		}
		last := cmsgbox("Name check",x "`n" RegExReplace(x,".","--") "`nWhat is the patient's`nLAST NAME?",q)
		if (last~="close|xClose") {
			return {first:"",last:x}
		}
		first := RegExReplace(x," " last)
	}
	
	return {first:first
			,last:last
			,firstlast:first " " last
			,lastfirst:last ", " first
			,apostr:RegExReplace(x,"\^","'")}
}

parsePat:
{
	agc := A_GuiControl
	Gui, ListView, %agc%
	if !(fileNum := LV_GetNext()) {
		return
	}
	
	pat_date:=
	pat_name:=
	pat_dev:=
	pat_ser:=
	pat_status:=
	pat_paceart:=
	pat_meta:=
	pat_report:=
	is_remote:=
	LV_GetText(pat_date,fileNum,1)
	LV_GetText(pat_name,fileNum,2)
	LV_GetText(pat_dev,fileNum,3)
	LV_GetText(pat_ser,fileNum,4)
	LV_GetText(pat_status,fileNum,5)
	LV_GetText(pat_paceart,fileNum,6)
	LV_GetText(fileIn,fileNum,7)
	LV_GetText(pat_meta,fileNum,8)
	LV_GetText(pat_report,fileNum,9)
	eventlog(pat_date " " pat_name " selected.")
	
	if (pat_report) {
		pat_node := "/root/work/id[@date='" pat_date "'][@ser='" pat_ser "']"
		opt := (pat_status="Pending")
			? "Modify report|Regenerate report|Mark entered in PaceArt"		; Not signed yet
			: "Mark entered in PaceArt"										; Report signed
		tmp := cMsgBox(pat_name " report","Do what?",opt,"Q","")
		if (tmp="Close") {
			return
		}
		if instr(tmp,"Modify") {
			RunWait, % "WordPad.exe """ pat_report """"						; launch fileNam in WordPad
			eventlog(pat_report " modified.")
			return
		}
		if instr(tmp,"Regenerate") {
			removeNode(pat_node)
			xl.save(worklist)
			eventlog("Node " pat_node " removed from worklist.")
			FileDelete, % pat_report
			eventlog("File " pat_report " deleted.")
			gosub fileLoop
			return
		}
		if instr(tmp,"PaceArt") {
			xl.setText(pat_node "/paceart","True")
			xl.save(worklist)
			eventlog("PaceArt marked true.")
			gosub ParseGUI
			return
		}
	}
	
	gosub fileLoop
	
Return
}

fileLoop:
{
/*	Read PDF file from clicked LV entry
*/
	blocks := Object()
	fields := Object()
	labels := Object()
	fldval := Object()
	leads := Object()
	yp := maintxt := summBl := summ := sjmLog := ""
	
	if (fileIn~="i).pdf$") {
		Run, %fileIn%
		SplitPath, fileIn,,,,fileOut
		FileDelete, %binDir%%fileOut%.txt
		RunWait, %binDir%pdftotext.exe -table "%fileIn%" "%binDir%%fileOut%.txt" , , hide
		eventlog("pdftotext " fileIn " -> " binDir fileOut ".txt")
		FileRead, maintxt, %binDir%%fileOut%.txt
		cleanlines(maintxt)
	}
	
	if (maintxt~="Medtronic,\s+Inc") {											; PM and ICD reports use common subs
		eventlog("Medtronic identified.")
		gosub Medtronic
	}
	else if (maintxt~="Boston Scientific Corporation") {
		eventlog("Boston Scientific identified.")
		gosub BSCI
	}
	else if instr(pat_dev,"SJM") {												; SJM device clicked from LV
		eventlog("St Jude identified.")
		gosub SJM
	} 
	else if (fileIn~="i).xml$") {
		eventlog("Opened " fileIn)
		gosub PaceartXml
	} 
	else {
		eventlog("No file match.")
		MsgBox No match!														; Attempt OCR on PDF?
	}
	
	return
}

SignScan:
{
	if !FileExist(hisDir) {
		MsgBox % "Requires 3M HIS dir`n""" hisDir """"
		ExitApp
	}
	l_users := {}
	l_numusers :=
	l_tabs := 
	Loop, %reportDir%*.rtf
	{
		fileNam := RegExReplace(A_LoopFileName,"i)\.rtf")						; fileNam is name only without extension, no path
		fileIn := A_LoopFileFullPath											; fileIn has complete path \\childrens\files\HCCardiologyFiles\EP\TRREAT reports\pending\steve.rtf
		
		l_user := strX(fileNam,"",1,0,"-",1)										; Get assigned EP from filename
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
	eventlog("Report RTF dir scanned.")
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
eventlog("<<<<< Parse session closed.")
ExitApp

SignGUIClose:
eventlog("<<<<< Sign session closed.")
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
	
	eventlog("Selected '" fileNam "'")
	
	tmp_usr := substr(fileNam,1,2)
	l_usr := substr(user,1,2)
	if !(l_usr=stmp_usr) {														; first user doesn't match that on filename?
		MsgBox, 262196,
			, % "Did you mean to open this report?`n`n"
			. "Was originally assigned to " tmp_usr "."
		IfMsgBox, No
		{
			eventlog("Oops. Didn't mean to open that.")
			gosub SignScan
			return
		}
	}
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
	Gui, Add, Button, vS_rev gActSign Disabled, SEND TO ESIG
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
	eventlog("PDF opened.")
Return
}

ActSign:
{
	Gui, Act:Hide
	l_tab := substr(l_tab,-1)													; get last 2 chars of l_tab
	l_usr := substr(user,1,2)
	if !(l_usr=l_tab) {													; first 2 chars of Citrix login don't match l_tab?
		MsgBox, 52, 
			, % "Sign this report?`n`n"
			. "Was originally assigned to " l_tab "."
		IfMsgBox Yes															; signing someone else's report
		{
			FileRead, tmp, % reportDir fileNam ".rtf"							; read the generated RTF file
			tmp := RegExReplace(tmp
				, "Dictating Phy #\\tab <8:(\d{6})>\\par"						; replace the original billing code
				, "Dictating Phy #\tab <8:" docs[l_usr] ">\par")				; with yours
			tmp := RegExReplace(tmp
				, "Attending Phy #\\tab <9:(\d{6})>\\par"						; and replace the assigned Attg
				, "Attending Phy #\tab <9:" docs[l_usr] ">\par")
			FileDelete, % reportDir fileNam ".rtf"
			FileAppend, % tmp, % reportDir fileNam ".rtf"						; generate a new RTF file
			eventlog(l_tab " report signed by " l_usr ".") 
		} else {
			eventlog("Oops. Don't sign " l_tab "'s report.")
			return																; not signing this report, return
		}
	}
	if !(isDevt) {
		FileCopy, % reportDir fileNam ".rtf", % hisDir . fileNam . ".rtf"
		eventlog("Sent to HIS.")
	}
	FileMove, % reportDir fileNam ".rtf", % complDir fileNam ".rtf", 1			; move copy to "completed" folder
	
	xl.setText("/root/work/id[@date='" l_date "'][@ser='" l_ser "']/status","Signed")
	xl.save(worklist)
	
	eventlog("Worklist.xml updated.")
	
	Gosub signScan																; regenerate file list
Return
}

Medtronic:
{
	if (maintxt~="Adapta|Sensia") {												; Scan Adapta family of devices
		eventlog("Adapta report.")
		gosub mdtAdapta
	} else if (maintxt~="(Quick Look II)|(Final:\s+Session Summary)") {							; or scan more current QuickLook II reports
		eventlog("QuickLookII report.")
		gosub mdtQuickLookII
	} else {																	; or something else
		eventlog("No match.")
		MsgBox NO MATCH
		return
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
/*	INITIAL INTERROGATION: QUICK LOOK II
	- Arrhythmia counters
	- Therapy counters
	- Pacing counters
*/
	qltxt := stregX(maintxt,"Quick Look II",1,0,"Observations\s+\(",1)
	inirep := stregX(qltxt,"Quick Look II",1,1,"Device Status",1,n)
	fields[1] := ["Device","Serial Number","Date of Visit"
				, "Patient","ID","Physician","`n"]
	labels[1] := ["IPG","IPG_SN","Encounter"
				, "Name","MRN","Physician","null"]
	fieldvals(inirep,1,"dev")
	if !instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") {
		fldval["dev-Physician"] := "Dr. " . trim(tmp," `n")
	}
	fldfill("dev-IPG","Medtronic " RegExReplace(fldval["dev-IPG"],"Medtronic "))
	
	inirep := stregX(qltxt,"Device Status",1,0,"Parameter Summary",1)
	fields[1] := ["\(Implanted: ","\)"
				, "Battery Voltage","`n"
				, "Remaining Longevity","`n"]
	labels[1] := ["IPG_impl","null"
				, "IPG_voltage","null"
				, "IPG_longevity","null"]
	fieldvals(inirep,1,"dev")
	fldfill("IPG_longevity",cleanspace(strX(inirep,"Remaining Longevity",1,19,"`n",1)))
	
	qltbl := stregX(qltxt,"Remaining Longevity",1,0,"Parameter Summary",1,n)
	qltbl := RegExReplace(qltbl,"\s+RRT.*years")
	qltbl := RegExReplace(qltbl,"\s+\(based on initial interrogation\)")
	qltbl := stregX(qltbl "<<<", "[\r\n]+   ",1,0,"<<<",1)
	qltbl := stregX(qltbl "<<<", "   ",1,0,"<<<",1)
	fields[2] := ["Atrial.*-Lead Impedance"
				, "Atrial.*-Pacing Impedance"
				, "Atrial.*-Capture Threshold"
				, "Atrial.*-Measured On"
				, "Atrial.*-In-Office Threshold"
				, "Atrial.*-Programmed Amplitude"
				, "Atrial.*-Measured .*Wave"
				, "Atrial.*-In-Office .*Wave"
				, "Atrial.*-Programmed Sensitivity"
			, "RV.*-Lead Impedance"
				, "RV.*-Pacing Impedance"
				, "RV.*-Defibrillation Impedance"
				, "RV.*-Capture Threshold"
				, "RV.*-Measured On"
				, "RV.*-In-Office Threshold"
				, "RV.*-Programmed Amplitude"
				, "RV.*-Measured .*Wave"
				, "RV.*-In-Office .*Wave"
				, "RV.*-Programmed Sensitivity"
			, "LV.*-Lead Impedance"
				, "LV.*-Pacing Impedance"
				, "LV.*-Capture Threshold"
				, "LV.*-Measured On"
				, "LV.*-In-Office Threshold"
				, "LV.*-Programmed Amplitude"
				, "LV.*-Measured .*Wave"
				, "LV.*-Programmed Sensitivity"]
	labels[2] := ["A_imp","A_imp","A_cap","A_date","A_Pthr","A_output","A_Sthr","A_Sthr","A_sensitivity"
				, "RV_imp","RV_imp","RV_HVimp","RV_cap","RV_date","RV_Pthr","RV_output","RV_Sthr","RV_Sthr","RV_sensitivity"
				, "LV_imp","LV_imp","LV_cap","LV_date","LV_Pthr","LV_output","LV_Sthr","LV_sensitivity"]
	scanParams(parseTable(qltbl),2,"leads",1)
	
	normLead("RA"
			,fldval["dev-Alead"],fldval["dev-Alead_impl"]
			,fldval["leads-A_imp"],fldval["leads-A_cap"],fldval["leads-A_output"],fldval["leads-A_Pol_pace"]
			,fldval["leads-A_Sthr"],fldval["leads-A_Sensitivity"],fldval["leads-A_Pol_sens"])
	normLead("RV"
			,fldval["dev-RVlead"],fldval["dev-RVlead_impl"]
			,fldval["leads-RV_imp"],fldval["leads-RV_cap"],fldval["leads-RV_output"],fldval["leads-RV_Pol_pace"]
			,fldval["leads-RV_Sthr"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"])
	normLead("LV"
			,fldval["dev-LVlead"],fldval["dev-LVlead_impl"]
			,fldval["leads-LV_imp"],fldval["leads-LV_cap"],fldval["leads-LV_output"],fldval["leads-LV_Pol_pace"]
			,fldval["leads-LV_Sthr"],fldval["leads-LV_Sensitivity"],fldval["leads-LV_Pol_sens"])
	
	inirep := stregX(qltxt,"Parameter Summary",1,1,"Clinical Status",1)
	qltbl := stregX(inirep,"Mode",1,0,"Detection",0)
	qltbl := columns(qltbl,"Mode","Detection",0,"Lower\s+Rate",(qltbl~="Paced AV")?"Paced AV":"")
	qltbl := RegExReplace(qltbl,"Lower  Rate","Lower Rate ")
	qltbl := RegExReplace(qltbl,"Upper  Track","Upper Track ")
	qltbl := RegExReplace(qltbl,"Upper  Sensor","Upper Sensor ")
	fields[2] := ["Mode Switch","Mode","V. Pacing","AdaptivCRT"
				, "Lower\s+Rate","Upper\s+Track","Upper\s+Sensor"
				, "Paced AV","Sensed AV"]
	labels[2] := ["Mode Switch","Mode","CRT_VP","CRT_VV","LRL","URL","USR","PAV","SAV"]
	scanParams(qltbl,2,"par",1)
	
	qltbl := stregX(inirep "<<<","Detection",1,0,"<<<",1)
	fields[2] := ["Rates-AT/AF","Rates-VF","Rates-FVT","Rates-VT"
				, "Therapies-AT/AF","Therapies-VF","Therapies-FVT","Therapies-VT"]
	labels[2] := ["AT/AF","VF","FVT","VT"
				, "Rx_AT/AF","Rx_VF","Rx_FVT","Rx_VT"]
	scanParams(parseTable(qltbl),2,"detect",1)
	
	inirep := columns(qltxt,"Clinical Status","Therapy Summary|Pacing",0,"Cardiac Compass")
	
	fields[1] := ["VF","VT-NS","VT","^AT/AF"]
	labels[1] := ["VF","VTNS","VT","ATAF"]
	scanParams(stregX(inirep,"Monitored",1,0,"Therapy|Pacing",1),1,"event",1)
	
	inirep := columns(qltxt "<<<","Therapy Summary|(\s+)?Pacing","<<<",0,"Pacing\s+\(")
	fields[1] := ["VT/VF-Pace-Terminated","VT/VF-Shock-Terminated","VT/VF-Total Shocks","VT/VF-Aborted Charges"
				, "AT/AF-Pace-Terminated","AT/AF-Shock-Terminated","AT/AF-Total Shocks","AT/AF-Aborted Charges"]
	labels[1] := ["V_Paced","V_Shocked","V_Total","V_Aborted"
				, "A_Paced","A_Shocked","A_Total","A_Aborted"]
	scanParams(parseTable(stregX(inirep,"Therapy Summary",1,0,"Observations|Pacing",1)),1,"event",1)
	
	iniRep := instr(iniRep,"Event Counters") ? oneCol(iniRep) : iniRep
	if instr(iniRep,"Sensed") {															; No chamber specified
		fields[2] := ["Sensed","Paced"]
		labels[2] := ["Sensed","Paced"]
	} else {
		fields[2] := ["AS.*VS","AS.*VP","AP.*VS","AP.*VP","^AS","^AP","^VS","^VP"]
		labels[2] := ["AsVs","AsVp","ApVs","ApVp","As","Ap","Vs","Vp"]
	}
	scanParams(iniRep,2,"dev",1)
	
	qlObs := stregX(maintxt,"Observations\s+\(",1,0,"\d+ Software Version",1)
	fldfill("event-Obs",qlObs)
	
/*	Pacing Threshold Test Report
	- Atrial, RV, LV amplitude threshold test
	- to Capture Management
*/
	ptr := 1
	While (i := stregX(maintxt,"Pacing Threshold Test Report",ptr,1,"Medtronic, Inc",1,ptr)) {
		thrTest := stregX(i,"\w+\s+Amplitude Threshold Test",1,0,"Capture Management",0)
		thrLead := stregX(thrTest,"\w+",1,0,"\s+Amplitude",1)
		thrTbl := parseTable(stregX(thrTest "<<<","   ",1,0,"<<<",1))
		fields[1] := ["Threshold-.*Amplitude","Threshold-.*Pulse Width"]
		labels[1] := ["Amp","PW"]
		scanParams(thrTbl,1,"tmp",1)
		if (thrVal := fldval["tmp-Amp"] printQ(fldval["tmp-PW"]," / ###")) {
			leads[(thrLead="Atrial")?"RA":thrLead,"cap"] := thrVal
		}
	}
	
/*	FINAL: SESSION SUMMARY
	- Device info, implant info
	- Lead parameters and measurements
	- Detections
*/
	fintxt := stregX(maintxt,"Final: Session Summary",1,0,"Medtronic, Inc.",0)
	
	dev := stregX(fintxt,"Session Summary",1,1,"Parameter Summary",1,n)
	fields[1] := ["Device","Serial Number","Date of Visit"
				, "Patient","ID","Physician","`n"]
	labels[1] := ["IPG","IPG_SN","Encounter"
				, "Name","MRN","Physician","null"]
	fieldvals(dev,1,"dev")
	if !instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") {
		fldval["dev-Physician"] := "Dr. " . trim(tmp," `n")
	}
	
	dev := stregX(fintxt,"Device Status",1,1,"Parameter Summary",1)
	fields[1] := ["Device Status", "Battery Voltage","Remaining Longevity","`n"]
	labels[1] := ["IPG_stat", "IPG_voltage","IPG_longevity","null"]
	fieldvals(dev,1,"dev")
	fldfill("IPG_longevity",cleanspace(strX(dev,"Remaining Longevity",1,19,"`n",1)))
	
	dev := stregX(fintxt,"Device Information",1,1,"Device Status",1)
	scanDevInfo(dev)
	fldfill("dev-IPG","Medtronic " RegExReplace(fldval["dev-IPG"],"Medtronic "))
	fldfill("dev-Alead", RegExReplace(fldval["dev-Alead"],"---"))
	fldfill("dev-RVlead", RegExReplace(fldval["dev-RVlead"],"---"))
	fldfill("dev-LVlead", RegExReplace(fldval["dev-LVlead"],"---"))
	
	fintbl := stregX(fintxt,"Remaining Longevity",1,0,"Parameter Summary",1,n)
	fintbl := RegExReplace(fintbl,"\s+RRT.*years")
	fintbl := RegExReplace(fintbl,"\s+\(based on initial interrogation\)")
	fintbl := stregX(fintbl "<<<", "[\r\n]+   ",1,0,"<<<",1)
	fintbl := stregX(fintbl "<<<", "   ",1,0,"<<<",1)
	fields[2] := ["Atrial.*-Lead Impedance"
				, "Atrial.*-Pacing Impedance"
				, "Atrial.*-Capture Threshold"
				, "Atrial.*-Measured On"
				, "Atrial.*-In-Office Threshold"
				, "Atrial.*-Programmed Amplitude"
				, "Atrial.*-Measured .*Wave"
				, "Atrial.*-In-Office .*Wave"
				, "Atrial.*-Programmed Sensitivity"
			, "RV.*-Lead Impedance"
				, "RV.*-Pacing Impedance"
				, "RV.*-Defibrillation Impedance"
				, "RV.*-Capture Threshold"
				, "RV.*-Measured On"
				, "RV.*-In-Office Threshold"
				, "RV.*-Programmed Amplitude"
				, "RV.*-Measured .*Wave"
				, "RV.*-In-Office .*Wave"
				, "RV.*-Programmed Sensitivity"
			, "LV.*-Lead Impedance"
				, "LV.*-Pacing Impedance"
				, "LV.*-Capture Threshold"
				, "LV.*-Measured On"
				, "LV.*-In-Office Threshold"
				, "LV.*-Programmed Amplitude"
				, "LV.*-Measured .*Wave"
				, "LV.*-Programmed Sensitivity"]
	labels[2] := ["A_imp","A_imp","A_cap","A_date","A_Pthr","A_output","A_Sthr","A_Sthr","A_sensitivity"
				, "RV_imp","RV_imp","RV_HVimp","RV_cap","RV_date","RV_Pthr","RV_output","RV_Sthr","RV_Sthr","RV_sensitivity"
				, "LV_imp","LV_imp","LV_cap","LV_date","LV_Pthr","LV_output","LV_Sthr","LV_sensitivity"]
	scanParams(parseTable(fintbl),2,"leads",1)
	
	fintbl := stregX(fintxt,"Detection",1,0,"(Changes)|(Enhancement)|(Clinical Status)",1)
	fields[2] := ["Rates-AT/AF","Rates-VF","Rates-FVT","Rates-VT"
				, "Therapies-AT/AF","Therapies-VF","Therapies-FVT","Therapies-VT"]
	labels[2] := ["AT/AF","VF","FVT","VT"
				, "Rx_AT/AF","Rx_VF","Rx_FVT","Rx_VT"]
	scanParams(parseTable(fintbl),2,"detect",1)
	
/*	FINAL: PARAMETERS
	- Modes, timing values
	- Programmed thresholds and outputs
*/
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
	
	normLead("RA"
			,fldval["dev-Alead"],fldval["dev-Alead_impl"]
			,fldval["leads-A_imp"],fldval["leads-A_cap"],fldval["leads-A_output"],fldval["leads-A_Pol_pace"]
			,fldval["leads-A_Sthr"],fldval["leads-A_Sensitivity"],fldval["leads-A_Pol_sens"])
	normLead("RV"
			,fldval["dev-RVlead"],fldval["dev-RVlead_impl"]
			,fldval["leads-RV_imp"],fldval["leads-RV_cap"],fldval["leads-RV_output"],fldval["leads-RV_Pol_pace"]
			,fldval["leads-RV_Sthr"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"])
	normLead("LV"
			,fldval["dev-LVlead"],fldval["dev-LVlead_impl"]
			,fldval["leads-LV_imp"],fldval["leads-LV_cap"],fldval["leads-LV_output"],fldval["leads-LV_Pol_pace"]
			,fldval["leads-LV_Sthr"],fldval["leads-LV_Sensitivity"],fldval["leads-LV_Pol_sens"])
	
return
}

mdtAdapta:
{
	isAdapta := true
	ptr := 1
	While (iniRep := stregX(maintxt,"Initial Interrogation Report",ptr,0,"Medtronic Software",1,ptr)) {
		if instr(iniRep,"Pacemaker Status") {
			fields[1] := ["Pacemaker\s+Model","Serial\s+Number","Date\s+of\s+Visit","`n"
						, "Patient\s+Name","ID","Physician","`n"
						, "History","`n"
						, "Implanted","\)"]
			labels[1] := ["IPG","IPG_SN","Encounter","null"
						, "Name","MRN","Physician","null","History","null","IPG_impl","null"]
			fieldvals(inirep,1,"dev")
			
			iniBlk := stregX(inirep,"Pacemaker Status",1,0,"Parameter Summary",1)
			
			iniTbl := columns(iniBlk "<<<","Pacemaker Status","<<<",0,"Battery Status")
			iniFld := stregX(iniTbl,"Battery Status",1,0,"Lead Summary",1)
			fields[1] := ["Estimated.*longevity","Voltage.Impedance"]
			labels[1] := ["IPG_longevity","IPG_voltage"]
			scanParams(iniFld,1,"dev",1)
			
			iniTbl := parseTable(stregX(iniTbl "<<<","Lead Summary",1,0,"<<<",1))
			fields[1] := ["Atrial-Measured Threshold"
						, "Atrial-Date Measured"
						, "Atrial-Programmed Output"
						, "Atrial-Capture"
						, "Atrial-Measured.*Wave"
						, "Atrial-Programmed Sensitivity"
						, "Atrial-Measured Impedance"
						, "Atrial-Lead Status"
						, "Atrial-Lead Model"
						, "Atrial-Implanted"
						, "Ventricular-Measured Threshold"
						, "Ventricular-Date Measured"
						, "Ventricular-Programmed Output"
						, "Ventricular-Capture"
						, "Ventricular-Measured.*Wave"
						, "Ventricular-Programmed Sensitivity"
						, "Ventricular-Measured Impedance"
						, "Ventricular-Lead Status"
						, "Ventricular-Lead Model"
						, "Ventricular-Implanted"]
			labels[1] := ["A_cap","A_date","A_output","A_mgt","A_Sthr","A_sensitivity","A_imp","A_stat","A_model","A_impl"
						, "V_cap","V_date","V_output","V_mgt","V_Sthr","V_sensitivity","V_imp","V_stat","V_model","V_impl"]
			scanParams(iniTbl,1,"leads",1)
			
			iniBlk := stregX(inirep,"Parameter Summary",1,1,"Clinical Status",1)
			iniTbl := stregX(iniBlk "<<<","Mode",1,0,"<<<",1)
			iniTbl := columns(iniTbl "<<<","Mode","<<<",0,"Lower Rate",instr(iniTbl,"Paced AV")?"Paced AV":"")
			fields[1] := ["Mode","Mode Switch","Detection Rate"
						, "Lower Rate","Upper Tracking Rate","Upper Sensor Rate"
						, "Search AV+","Paced AV","Sensed AV"]
			labels[1] := ["Mode","ModeSwitch","ModeSwitchRate"
						, "LRL","URL","USR"
						, "SearchAV","PAV","SAV"]
			scanParams(iniTbl,1,"par")
			
			iniBlk := stregX(inirep "<<<","Clinical Status",1,0,"<<<",0)
			iniBlk := columns(iniBlk,"Clinical Status","<<<",0,"Pacing\s+\(")
			fields[1] := ["Atrial High Rate Episodes","Ventricular High Rate Episodes"]
			labels[1] := ["AHR","VHR"]
			scanParams(RegExReplace(iniBlk,"Episodes: ","Episodes:  "),1,"event",1)
			iniTbl := stregX(iniBlk "<<<","Pacing\s+\(",1,0,"<<<",1)
			iniTbl := instr(iniTbl,"Event Counters") ? oneCol(iniTbl) : iniTbl
			fields[2] := ["AS.*VS","AS.*VP","AP.*VS","AP.*VP","Sensed","Paced"]
			labels[2] := ["AsVs","AsVp","ApVs","ApVp","Sensed","Paced"]
			scanParams(iniTbl,2,"dev",1)
			
			iniTbl := stregX(iniTbl "<<<","Event Counters",1,0,"<<<",1)
			fields[3] := ["PVC singles","PVC runs","PAC runs"]
			labels[3] := ["PVC","PVCruns","PACruns"]
			scanParams(iniTbl,3,"event",1)
		}
	}
	
	ptr := 1
	while (finRep := stregX(maintxt,"Final Report",ptr,0,"Medtronic Software",1,ptr)) {
		if instr(finRep,"Pacemaker Status") {
			fields[1] := ["Pacemaker Model","Serial Number","Date of Visit","`n"
						, "Patient Name","ID","Physician","`n"]
			labels[1] := ["IPG","IPG_SN","Encounter","null"
						, "Name","MRN","Physician","null"]
			fieldvals(finRep,1,"dev")
			
			finBlk := stregX(finRep,"Patient Name",1,0,"Pacemaker Status",0)
			finBlk := stregX(finBlk,"Pacemaker Model",1,0,"Pacemaker Status",1)
			scanDevInfo(finBlk)
			fldval["dev-IPG"] := RegExReplace(fldval["dev-IPG"],fldval["dev-IPG_SN"])
			
			finBlk := stregX(finRep,"Pacemaker Status",1,0,"Lead Status",1)
			fields[1] := ["Battery Status","Voltage"]
			labels[1] := ["IPG_stat","IPG_voltage"]
			scanParams(finBlk,1,"dev",1)
			
			finBlk := stregX(finRep,"Lead Status",1,0,".. Capture Management|Sensing Assurance",1)
			finBlk := stregX(finBlk "<<<","[\r\n]+   ",1,0,"<<<",1)
			finBlk := stregX(finBlk "<<<","   ",1,0,"<<<",1)
			finBlk := parseTable(finBlk)
			fields[1] := ["Atrial.*-Measured Impedance"
						, "Atrial.*-Pace Polarity"
						, "Ventricular.*-Measured Impedance"
						, "Ventricular.*-Pace Polarity"]
			labels[1] := ["A_imp","A_pol","V_imp","V_pol"]
			scanParams(finBlk,1,"leads",1)
			
			finBlk := stregX(finRep "<<<","In-Office Threshold",1,0,"<<<",0)
			finBlk := columns(finBlk,"In-Office Threshold","<<<",0,"\w+ Sensing Threshold") "<<<"
			finFld := stregX(finBlk,"Atrial Pacing Threshold",1,1,"(<<<)|(Ventricular Pacing Threshold)",1)
			fldfill("leads-A_cap",parseStrDur(finFld))
			finFld := stregX(finBlk,"Ventricular Pacing Threshold",1,1,"<<<",1)
			fldfill("leads-V_cap",parseStrDur(finFld))
			fields[1] := ["P-wave","R-wave"]
			labels[1] := ["AS_thr","VS_thr"]
			scanParams(finBlk,1,"leads",1)
		}
		if instr(finRep,"Permanent Parameters") {
			perm := oneCol(stregX(finRep,"Permanent Parameters(.*?)`n",1,1,"Medtronic Software",1))
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
	}
	normLead("RA"
			,fldval["dev-Alead"],fldval["dev-Alead_impl"]
			,fldval["leads-A_imp"],fldval["leads-A_cap"]
			,(fldval["Alead-Amp"]) ? fldval["Alead-Amp"] " at " fldval["Alead-PW"] : ""
			,fldval["Alead-Pol_pace"]
			,fldval["leads-AS_thr"],fldval["Alead-Sens"],fldval["Alead-Pol_sens"])
	normLead("RV"
			,fldval["dev-Vlead"],fldval["dev-Vlead_impl"]
			,fldval["leads-V_imp"],fldval["leads-V_cap"]
			,(fldval["Vlead-Amp"]) ? fldval["Vlead-Amp"] " at " fldval["Vlead-PW"] : ""
			,fldval["Vlead-Pol_pace"]
			,fldval["leads-VS_thr"],fldval["Vlead-Sens"],fldval["Vlead-Pol_sens"])
	isAdapta := 
return
}

BSCI:
{
	if (pat_meta) {
		FileRead, bscbnk, % pat_meta
	}
	gosub bsciZoomView
	
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
	fldfill("dev-DOB",parseDate(RegExReplace(fldval["dev-DOB"]," ","-")).MDY)
	fldfill("dev-Encounter",parseDate(RegExReplace(fldval["dev-Encounter"]," ","-")).MDY)
	fldfill("dev-Last_ck",parseDate(RegExReplace(fldval["dev-Last_ck"]," ","-")).MDY)
	fldfill("dev-IPG_impl",parseDate(RegExReplace(fldval["dev-IPG_impl"]," ","-")).MDY)
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
	
	normLead("RA"
			,fldval["dev-Alead"],fldval["dev-Alead_impl"]
			,fldval["Alead-imp"],fldval["Alead-cap"],fldval["leads-AP_thr"],fldval["Alead-Pol_pace"]
			,fldval["Alead-sensing"],fldval["leads-AS_thr"],fldval["leads-RA_Pol_sens"])
	normLead("RV"
			,fldval["dev-RVlead"],fldval["dev-RVlead_impl"]
			,fldval["Vlead-imp"],fldval["Vlead-cap"],fldval["leads-VP_thr"],fldval["RVlead-Pol_pace"]
			,fldval["Vlead-sensing"],fldval["leads-VS_thr"],fldval["leads-RV_Pol_sens"])
	normLead("LV"
			,fldval["dev-LVlead"],fldval["dev-LVlead_impl"]
			,fldval["leads-LV_imp"],fldval["leads-LV_cap"],fldval["leads-LV_output"],fldval["LVlead-Pol_pace"]
			,fldval["leads-LV_Sensitivity"],fldval["leads-LV_Sthr"],fldval["leads-LV_Pol_sens"])

	return
}

SJM:
{
	if !(pat_meta) {																; SJM device with metadata (ICD exported)
		MsgBox No metafile found!
		return
	} 
	FileRead, sjmLog, % pat_meta
	if (sjmLog~="Microny|Zephyr") {
		gosub SJM_old
	} else {
		gosub SJM_meta															; 
	}
	
	gosub fetchDem
	
	if (fetchQuit) {
		return
	}
	
	gosub makeReport
	
return
}

SJM_old:
{
	fields[1] := ["Device Name","Model:", "Serial:","Implant Date:"
				, "Voltage","Current","Impedance"
				, "Last Interrogated On:"]
	labels[1] := ["IPG","IPG_model","IPG_SN","IPG_impl"
				, "IPG_voltage","IPG_current","IPG_imped"
				, "Encounter"]
	sjmVals(1,"dev")
	fldfill("dev-Name",pat_name)
	fldfill("dev-IPG","SJM " fldval["dev-IPG"] printQ(fldval["dev-IPG_model"], " ###"))
	fldfill("dev-Encounter", parseDate(fldval["dev-Encounter"]).MDY)
	fldfill("dev-IPG_impl",niceDate(fldval["dev-IPG_impl"]))
	
	fields[1] := ["Lead Chamber","Lead Type"
				, ".. Pulse Amplitude",".. Pulse Width","Lead Impedance","P/R Sensitivity",
				, "Vario Capture Threshold","Test Pulse Width","E/R Sensitivity"]
	labels[1] := ["Chamber","Type"
				, "Pace_Amp","Pace_PW","Imped","Sensitivity"
				, "Thr_Amp","Thr_PW","Thr_Sens"]
	sjmVals(1,"leads")
	
	fields[1] := ["(\x1C)Mode(\x1C)","Base Rate","Max Sensor Rate"]
	labels[1] := ["Mode","LRL","USR"]
	sjmVals(1,"par")
	
	normLead("R" (InStr(fldval["leads-Chamber"],"V")?"V":"A")
		,fldval["dev-RVlead"],fldval["dev-RVlead_impl"],fldval["leads-Imped"]
		,printQ(fldval["leads-Thr_Amp"],"###" printQ(fldval["leads-Thr_PW"]," @ ###"))
		,printQ(fldval["leads-Pace_Amp"],"###" printQ(fldval["leads-Pace_PW"]," @ ###"))
		,fldval["leads-RV_Pol_pace"]
		,fldval["leads-Thr_Sens"],fldval["leads-Sensitivity"],fldval["leads-RV_Pol_sens"])
	
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
	fldfill("dev-Encounter", parseDate(fldval["dev-Encounter"]).MDY)
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
	
	fields[1] := ["(\x1C)Mode(\x1c)","Base Rate","Maximum Tracking Rate","Maximum Sensor Rate"
				, "Paced AV Delay","Sensed AV Delay"]
	labels[1] := ["Mode","LRL","URL","USR"
				, "PAV","SAV"]
	sjmVals(1,"par")
	
	fields[1] := ["(\x1C)VF Detection Interval","(\x1C)VT-1 Detection Interval"
				, "VT-1 Therapy 1 Type","VF Therapy 1 Type","VF Voltage 1"]
	labels[1] := ["VF","VT"
				, "Rx_VT","VF0","VF1"]
	sjmVals(1,"detect")
	fldfill("detect-VF",fldval["detect-VF"]?round(60000/RegExReplace(fldval["detect-VF"],"\D")):"")
	fldfill("detect-VT",fldval["detect-VT"]?round(60000/RegExReplace(fldval["detect-VT"],"\D")):"")
	fldfill("detect-Rx_VF",fldval["detect-VF0"] printQ(fldval["detect-VF1"],", ###"))
	
	fields[1] := ["AT/AF Episodes","VT/VF Episodes"]
	labels[1] := ["ATAF","VT"]
	sjmVals(1,"event")
	
	normLead("RA"
			,fldval["dev-Alead"],fldval["dev-Alead_impl"],fldval["leads-RA_imp"]
			,printQ(fldval["leads-RA_Thr_Amp"],"###" printQ(fldval["leads-RA_Thr_PW"]," @ ###"))
			,printQ(fldval["leads-RA_Pace_Amp"],"###" printQ(fldval["leads-RA_Pace_PW"]," @ ###"))
			,fldval["leads-RA_Pol_pace"]
			,fldval["leads-RA_Thr_Sens"],fldval["leads-RA_Sensitivity"],fldval["leads-RA_Pol_sens"])
	normLead("RV"
			,fldval["dev-RVlead"],fldval["dev-RVlead_impl"],fldval["leads-RV_imp"]
			,printQ(fldval["leads-RV_Thr_Amp"],"###" printQ(fldval["leads-RV_Thr_PW"]," @ ###"))
			,printQ(fldval["leads-RV_Pace_Amp"],"###" printQ(fldval["leads-RV_Pace_PW"]," @ ###"))
			,fldval["leads-RV_Pol_pace"]
			,fldval["leads-RV_Thr_Sens"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"])
	normLead("LV"
			,fldval["dev-LVlead"],fldval["dev-LVlead_impl"],fldval["leads-LV_imp"]
			,printQ(fldval["leads-LV_Thr_Amp"],"###" printQ(fldval["leads-LV_Thr_PW"]," @ ###"))
			,printQ(fldval["leads-LV_Pace_Amp"],"###" printQ(fldval["leads-LV_Pace_PW"]," @ ###"))
			,fldval["leads-LV_Pol_pace"]
			,fldval["leads-LV_Thr_Sens"],fldval["leads-LV_Sensitivity"],fldval["leads-LV_Pol_sens"])

return
}

PaceartXml:
{
	progress,,,Scanning...
	yp := new XML(fileIn)
	fldval["dev-type"] := yp.selectSingleNode("//ActiveDevices/PatientActiveDevice/Device/Type").text
	
	if (fldval["dev-type"]) {
		eventlog("Paceart " fldval["dev-type"]" report.")
		gosub PaceartReadXml
	}
	else {
		progress,off
		eventlog("Paceart no match.")
		MsgBox NO MATCH
		return
	}
	
	progress,off
	gosub fetchDem
	
	if (fetchQuit) {
		return
	}
	
	gosub makeReport
	
return	
}

PaceartReadXml:
{
	fields[1] := ["IDs/ID[Type='MRN']/Value:MRN"
				, "Demographics/FirstName:nameF"
				, "Demographics/LastName:nameL"
				, "Diagnoses/PatientDiagnosis/Diagnosis/Code:dx_code"
				, "Diagnoses/PatientDiagnosis/Diagnosis/Description:dx_desc"
				, "/Encounter/Evaluation/MiscellaneousComment:summary"
				. ""]
	xmlFld("//PatientRecord",1,"dev")
	fldfill("dev-name",fldval["dev-nameL"] ", " fldval["dev-nameF"])
	fldfill("indication",printQ(fldval["dev-dx_code"],"### - ") fldval["dev-dx_desc"])
	
	fields[1] := ["Device/Manufacturer:manufacturer"
				, "Device/Model:model"
				, "SerialNumber:IPG_SN"
				, "ImplantDate:IPG_impl"
				, "FirstImplantingProvider/LastName:Physician"
				. ""]
	xmlFld("//ActiveDevices/PatientActiveDevice[Status='ACTIVE']",1,"dev")
	fldfill("dev-IPG_impl",parseDate(fldval["dev-IPG_impl"]).MDY)
	
	fields[1] := ["Date:Encounter"
				, "Type:EncType"
				, "/Battery/Status[@nonconformingData]:Battery_stat"
				, "/Battery/RemainingLongevity:IPG_Longevity[months]"
				, "/Battery/Voltage:IPG_voltage[V]"
				, "/Battery/Impedance:IPG_impedance[ohms]"
				. ""]
	xmlFld("//Encounter",1,"dev")
	fldfill("dev-IPG",printQ(fldval["dev-manufacturer"],"###") printQ(fldval["dev-model"]," ###"))
	fldfill("dev-Encounter",parseDate(fldval["dev-Encounter"]).MDY)
	
	fields[1] := ["PacingMode:Mode"
				, "LowerRate:LRL"
				, "TrackingRate:URL"
				, "MaxSensorRate:USR"
				, "RateModulation/ADLRate:ADL"
				, "PacingData[Chamber='RIGHT_VENTRICLE']/AdaptationMode:Cap_Mgt"
				, "AVDelay/Sensed:SAV"
				, "AVDelay/Paced:PAV"
				, "/SensingData[Chamber='RIGHT_ATRIUM']"
					. "//RefractoryPeriod[PreviousEventChamber='VENTRICLE']"
					. "/Interval:PVARP"
				, "AutomaticModeSwitch/Status:ModeSwitch"
				, "AutomaticModeSwitch/Detection/Criteria/Rate:AMSRate"
				. ""]
	xmlFld("//Programming/Bradycardia",1,"par")
	fldval["par-AMS"] := fldval["par-ModeSwitch"]="ENABLED" ? fldval["par-AMSRate"] : "Off"
	
	fields[1] := ["/VentricularFirstChamberPaced:CRT_VP"
				, "/VVDelay:CRT_VV[ms]"
				. ""]
	xmlFld("//Programming/HeartFailure",1,"par")
	fldval["par-CRT_VP"] := fldval["par-CRT_VP"]~="LEFT" ? "LV>RV" : "RV<LV"
	
	fields[1] := ["APVPPercent:ApVp[%]"
				, "ASVPPercent:AsVp[%]"
				, "APVSPercent:ApVs[%]"
				, "ASVSPercent:AsVs[%]"
				, "/PercentPaced[Chamber='RIGHT_ATRIUM']/Percent:AP"
				, "/PercentPaced[Chamber='RIGHT_VENTRICLE']/Percent:VP"
				, "/PercentPaced[Chamber='LEFT_VENTRICLE']/Percent:LVP"
				. ""]
	xmlFld("//BradycardiaCollection/Bradycardia",1,"dev")
	
	fields[1] := ["/Zone[Type='VENTRICULAR_FIBRILLATION']//Summary:VF"
				, "/Zone[Type='VENTRICULAR_TACHYCARDIA']//Summary:VT"
				, "/Zone[Type='VENTRICULAR_TACHYCARDIA_1']//Summary:VT1"
				, "/Zone[Type='VENTRICULAR_TACHYCARDIA_2']//Summary:VT2"
				, "/Zone[Type='ATRIAL_TACHYCARDIA']//Summary:AT"
				, "/Zone[Type='ATRIAL_FIBRILLATION']//Summary:AF"
				. ""]
	xmlFld("//Programming/Tachycardia",1,"detect")
	
	fields[1] := ["/Episode[Type='AF_AT']/Count:AT/AF"
				, "/Episode[Type='VF_VT']/Count:VT/VF"
				, "/Episode[Type='SVT']/Count:SVT"
				, "/Episode[Type='V_NST']/Count:VNST"
				, "/Episode[Type='VT']/Count:VT"
				, "/Episode[Type='FVT']/Count:FVT"
				, "/Therapy[Chamber='RIGHT_ATRIUM']/ATP/Delivered:Rx_AT/AF"
				, "/Therapy[Chamber='RIGHT_ATRIUM']/Shocks/Delivered:A_Shocked"
				, "/Therapy[Chamber='RIGHT_ATRIUM']/Shocks/Aborted:A_Aborted"
				, "/Therapy[Chamber='RIGHT_VENTRICLE']/ATP/Delivered:Rx_VATP"
				, "/Therapy[Chamber='RIGHT_VENTRICLE']/Shocks/Delivered:V_Shocked"
				, "/Therapy[Chamber='RIGHT_VENTRICLE']/Shocks/Aborted:V_Aborted"
				. ""]
	xmlFld("//Statistics/Detections_Therapies",1,"detect")
	
	loop, % (i:=yp.selectNodes("//PatientPassiveDevice[Status='ACTIVE']")).length
	{
		k := readXmlLead(i.item(A_Index-1))
		normLead(k.ch
			, printQ(k.manu, "### ") printQ(k.model, "###") printQ(k.ser,", serial ###"), k.impl
			, k.pacing_imped
			, k.cap_amp printQ(k.cap_pw," / ###")
			, k.pacing_amp printQ(k.pacing_pw," / ###")
			, k.pacing_pol
			, k.sensing_thr
			, k.sensitivity_amp
			, k.sensitivity_pol)
	}
	
	return
}

readXmlLead(k) {
	global fldval, leads
	
	res := []
	res.ser := k.selectSingleNode("SerialNumber").text
	res.manu := k.selectSingleNode("Device/Manufacturer").text
	res.model := k.selectSingleNode("Device/Model").text
	res.impl := parseDate(k.selectSingleNode("ImplantDate").text).MDY
	res.chamb := k.selectSingleNode("Chamber").text
	res.ch := RegExReplace(res.chamb,"(L|R).*?_(A|V).*?$","$1$2")
	if IsObject(leads[res.ch]) {
		for key in leads
		{
			if (res.ch ~= key) {
				num ++
			}
		}
		res.ch .= num+1
	}
	if (k.selectSingleNode("Device/Comments").text~="HV") {
		fldval["leads-" res.ch "_HVimp"] := printQ(readNodeVal("//Statistics//HighPowerChannel//Impedance//Value"),"### ohms")
	}
	if (res.model ~= "6937") {
		return res
	}
	
	base := "//Programming//PacingData[Chamber='" res.chamb "']"
	res.pacing_pol := printQ(readNodeVal(base "//Polarity"),"###")
	res.pacing_amp := printQ(readNodeVal(base "/Amplitude"),"### V")
	res.pacing_pw := printQ(readNodeVal(base "/PulseWidth"),"### ms")
	res.pacing_adaptive := printQ(readNodeVal(base "/AdaptationMode"),"###")
	
	base := "//Programming//SensingData[Chamber='" res.chamb "']"
	res.sensitivity_pol := printQ(readNodeVal(base "//Polarity"),"###")
	res.sensitivity_amp := printQ(readNodeVal(base "//Amplitude"),"### mV")
	
	base := "//Statistics//Lead[Chamber='" res.chamb "']"
	res.cap_amp := printQ(readNodeVal(base "/LowPowerChannel//Capture//Amplitude"),"### V") 
	res.cap_pw := printQ(readNodeVal(base "/LowPowerChannel//Capture//Duration"),"### ms") 
	res.sensing_thr := printQ(readNodeVal(base "/LowPowerChannel//Sensitivity//Amplitude"),"### mV") 
	res.pacing_imped := printQ(readNodeVal(base "/LowPowerChannel//Impedance//Value"),"### ohms")
	;~ fldval["leads-" res.ch "_HVimp"] := printQ(readNodeVal("//Statistics//HighPowerChannel//Impedance//Value"),"### ohms")
	
	return res
}

xmlFld(base,blk,pre="") {
/*	Reads xxxxxx:yyyy from array blk
		xxxxxx = xpath appended to base, if xxxxxx[@aaa] will getAttribute @aaa
		yyyy = fldval[label], if yyyy[bbb] will append bbb units to result from xxxxxx 
*/
	global fldval, fields
	
	loop,
	{
		i := A_Index
		k := fields[blk][i]
		fld := strX(k,"",1,0,":",1,1)
		lbl := strX(k,":",1,1,"",0)
		if (fld="") {
			break
		}
		
		res := readNodeVal(base "/" fld)
		unit := strX(lbl,"[",1,1,"]",0,1)
		lbl := strX(lbl,"",1,0,"[",1,1)
		
		fldval[pre "-" lbl] := printQ(res, "###" . printQ(unit," ###"))
	}
	return
}

readNodeVal(fld) {
/*	Reads a result from Xpath node 'fld'
	xxxxx returns text from node
	xxxxx[@yyy] returns value from attribute yyy
*/
	global yp
	
	if (fld="") {
		return error
	}
	if RegExMatch(fld,"\[@(.*)?\]$",d) {
		fld := strX(fld,"",1,0,"[@",1,2)
		res := yp.selectSingleNode(fld).getAttribute(d1)
	} else {
		res := yp.selectSingleNode(fld).text
	}
	
	return res
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
	n := 1
	While (pos:=RegExMatch(txt,"O)\d+[.]\d+ V(.*?)\d+[.]\d+ ms",val,n)) {		; find "0.50 V @ 0.4 ms"
		res := ((res) ? res " and " : "") . val.value()							; append to RES (if RES already exists, prepend "and")
		n+=pos+val.Len()														; starting point for next instance
	}
	
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
				
				if !(pos) {														; break out when no more headers
					break
				}
				
				col.Push(pos)													; add position to col[] array (0 when no more matches)
				pre.Push(strX(substr(i,pos),"",1,0,"  ",1,2))					; add header value
				
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

scanDevInfo(txt) {
	global fldval,isAdapta
	fields := ["Device","Atrial","RA","RV","LV"]
	labels := ["IPG","Alead","Alead","RVlead","LVlead"]
	if (isAdapta) {
		fields := ["Pacemaker Model:","Atrial Lead:","Ventricular Lead:"]
		labels := ["IPG","Alead","Vlead"]
	}
	Loop, parse, txt, `n,`r
	{
		i := trim(A_LoopField)
		set := strX(i,"",1,0,"   ",1,3,n)
		val := objHasValue(fields,set)
		
		if !(val) {
			continue
		}
		
		res := substr(i,n)
		model := cleanspace(strX(res,"",1,0,"Implanted:",1,10))
		date := trim(strx(i,"Implanted:",1,10,"",0))
		
		fldfill("dev-" labels[val], model)
		fldfill("dev-" labels[val] "_impl", date)
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
	return RegExReplace(el3,"[^[:ascii:]]") 
		. printQ(RegExReplace(el4,"[^[:ascii:]]")," ###") 
		. printQ(RegExReplace(el5,"[^[:ascii:]]")," ###")						; return: value ( units)( whatever el5 is)
}

pmPrint:
{
	if !(enc_MD) {
		return
	}
	rtfBody := "\b\ul DEVICE INFORMATION AND INITIAL SETTINGS\ul0\b0\par`n"
	. fldval["dev-IPG"] ", serial number " fldval["dev-IPG_SN"] 
	. printQ(fldval["dev-IPG_impl"],", implanted ###") . printQ(fldval["dev-Physician"]," by ###") ". `n"
	. printQ(fldval["dev-IPG_voltage"],"Generator cell voltage ###. ")
	. printQ(fldval["dev-Battery_stat"],"Battery status is ###. ") . printQ(fldval["dev-IPG_Longevity"],"Remaining longevity ###. ") "`n"
	. printQ(fldval["par-Mode"],"Brady programming mode is ### with lower rate " fldval["par-LRL"])
	. printQ(fldval["par-URL"],", upper tracking rate ###")
	. printQ((substr(fldval["par-Mode"],0,1)="R"),printQ(fldval["par-USR"],", upper sensor rate ###"))
	. printQ(fldval["par-ADL"],", ADL rate ###") . ". `n"
	. printQ(fldval["par-Cap_Mgt"],"Adaptive mode is ###. `n")
	. printQ(fldval["par-PAV"],"Paced and sensed AV delays are " fldval["par-PAV"] " and " fldval["par-SAV"] ", respectively. `n")
	. printQ(fldval["dev-Sensed"],"Sensed ###. ") . printQ(fldval["dev-Paced"],"Paced ###. ")
	. printQ(fldval["dev-AsVs"],"AS-VS ###  ") . printQ(fldval["dev-AsVp"],"AS-VP ###  ")
	. printQ(fldval["dev-ApVs"],"AP-VS ###  ") . printQ(fldval["dev-ApVp"],"AP-VP ###  ")
	. printQ(fldval["dev-AP"],"A-paced ###%. ") . printQ(fldval["dev-VP"],"V-paced ###%. ")
	. printQ(fldval["detect-AT/AF"],"AT/AF detection ###" printQ(fldval["detect-Rx_AT/AF"],", Rx ###") ". ")
	. printQ(fldval["detect-VF"],"VF detection ###" printQ(fldval["detect-Rx_VF"],", Rx ###") ". ")
	. printQ(fldval["detect-FVT"],"FVT detection ###" printQ(fldval["detect-Rx_FVT"],", Rx ###") ". ")
	. printQ(fldval["detect-VT"],"VT detection ###" printQ(fldval["detect-Rx_VT"],", Rx ###") ". ") 
	. "\par `n"
	. "\par `n"
	. "\b\ul LEAD INFORMATION\ul0\b0\par`n"
	
	for k in leads
	{
		printLead(k)
	}
	
	printEvents()
	
	gosub PrintOut

Return
}

printQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	return (var1="") ? null : RegExReplace(txt,"###",var1)
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
	if (!P_imp && !P_thr && !P_out && !P_pol && !S_thr && !S_sens && !S_pol) {			; ALL parameters in pre or post are NULL
		eventlog("Lead " lead " all null values!")
		;~ return error																	; Do not populate leads[]
	}
	global leads, fldval
	leads[lead,"model"] 	:= model
	leads[lead,"date"]		:= date
	leads[lead,"imp"]  		:= printQ(P_imp,"Pacing impedance ###") 
							. printQ(fldval["leads-" lead "_HVimp"]
							, printQ(P_imp,". ") " Defib impedance ###")
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
	. printQ(leads[lead,"imp"],"###. ")
	. printQ(leads[lead,"cap"],"Capture threshold ###. ")
	. printQ(leads[lead,"output"],"Pacing output ###. ")
	. printQ(leads[lead,"pace pol"],"Pacing polarity ###. ")
	. printQ(leads[lead,"sens"],((lead="RA")?"P":"")((lead="RV")?"R":"") "-wave sensing " 
		. ((leads[lead,"sens"]~="N/R")?"not measured/detected":"###") ". ")
	. printQ(leads[lead,"sensitivity"],"Sensitivity ###. ")
	. printQ(leads[lead,"sens pol"],"Sensing polarity ###. ")
	. "\par `n"
}

printEvents()
{
	global rtfBody, fldval
	if (fldval["leads-RV_HVimp"]) {
		txt := ""
		. printQ(fldval["event-AHR"]?fldval["event-AHR"]:"0","There were ### Atrial High Rate episodes. ")
		. printQ(fldval["event-VHR"]?fldval["event-VHR"]:"0","There were ### Ventricular High Rate episodes. ")
		. printQ(fldval["event-VF"]?fldval["event-VF"]:"0","### VF episodes detected. ")
		. printQ(fldval["event-VT"]?fldval["event-VT"]:"0","### VT episodes detected. ")
	}
	txt .= ""
	. printQ(fldval["event-VTNS"]?fldval["event-VTNS"]:"","### NS-VT episodes detected. ")
	. printQ(fldval["event-ATAF"]?fldval["event-ATAF"]:"","### AT/AF episodes detected. ")
	. printQ(fldval["event-V_Paced"]?fldval["event-V_Paced"]:"","### VT episodes pace-terminated. ")
	. printQ(fldval["event-V_Shocked"]?fldval["event-V_Shocked"]:"","### VT/VF episodes shock-terminated. ")
	. printQ(fldval["event-V_Aborted"]?fldval["event-V_Aborted"]:"","### VT/VF episodes aborted. ")
	. printQ(fldval["event-A_Paced"]?fldval["event-A_Paced"]:"","### AT episodes pace-terminated. ")
	. printQ(fldval["event-A_Shocked"]?fldval["event-A_Shocked"]:"","### AT/AF episodes shock-terminated. ")
	. printQ(fldval["event-A_Aborted"]?fldval["event-A_Aborted"]:"","### AT/AF episodes aborted. ")
	. printQ(fldval["event-Obs"],"\par ### ")
	
	rtfBody .= printQ(txt,"\par\b\ul EVENTS\ul0\b0\par`n###\par `n") 
return	
}

PrintOut:
{
	FormatTime, enc_dictdate, A_now, yyyy MM dd hh mm t
	if (is_remote) {
		enc_type := "REMOTE "
		enc_dt := parseDate(substr(A_now,1,8))											; report date is date run (today)
		enc_trans := parseDate(fldval["dev-Encounter"])									; transmission date is date sent
	} else {
		enc_type := "IN-OFFICE "
		enc_dt := parseDate(fldval["dev-Encounter"])									; report date is day of encounter
		enc_trans :=																	; transmission date is null
	}
	
	for k in leads
	{
		ctLeads := A_Index
	}
	enc_type .= (instr(leads["RV","imp"],"Defib"))
		? "ICD "
		: "PM "
	if (ctLeads = 1) {
		enc_type .= "Single"
	} else if (ctLeads = 2) {
		enc_type .= "Dual"
	} else if (ctLeads > 2) {
		enc_type .= "Multi"
	}
	
	rtfHdr := "{\rtf1`n{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f2\fnil\fcharset2 Wingdings;}}\pard`n"
	
	rtfFtr := "`n}`n"
	
	rtfBody := "\b\ul ANALYSIS DATE:\ul0\b0  " enc_dt.MDY "\par\par`n"
			. strQ(is_remote
				, "\b\ul TRANSMISSION DATE:\ul0\b0 " enc_trans.MDY "\par\par`n")
			. "\b\ul ENCOUNTER TYPE\ul0\b0\par`n"
			. "Device interrogation "enc_type "\par`n"
			. "Performed by " tech ".\par\par`n"
			. strQ(fldval["indication"]
				, "\b\ul INDICATION FOR DEVICE\ul0\b0\par`n"
				. "###\par\par`n")
			. strQ(fldval["dependent"]
				, "\b\ul PACEMAKER DEPENDENT\ul0\b0\par`n"
				. "###\par\par`n")
			. rtfBody "\par`n"
			. "\b\ul ENCOUNTER SUMMARY\ul0\b0\par`n"
			. summ "\par\par`n"
	
	rtfOut := rtfHdr . rtfBody . rtfFtr
	
	nm := fldval["dev-Name"]
	RegExMatch(fileIn,"\....$",ext)
	fileOut :=	enc_MD "-" encMRN " " 
			.	(instr(nm,",") ? strX(nm,"",1,0,",",1,1) : strX(nm," ",1,1,"",0)) " "
			.	"#" fldval["dev-IPG_SN"] " "
			.	enc_dt.YMD
	
	FileDelete, %binDir%%fileOut%.rtf													; delete and generate RTF fileOut.rtf
	FileAppend, %rtfOut%, %binDir%%fileOut%.rtf
	
	eventlog("Print output generated in " binDir)
	
	RunWait, WordPad.exe "%binDir%%fileOut%.rtf"										; launch fileNam in WordPad
	MsgBox, 262180, , Report looks okay?
	IfMsgBox, Yes
	{
		eventlog("RTF, " ext " copied to " complDir)
		if (pat_meta) {
			FileMove, %pat_meta%, %complDir%%fileOut%.meta, 1							; copy BNK to complete directory
			eventlog("META copied to " complDir)
		}
		if (ext=".xml") {
			nBytes := Base64Dec( yp.selectSingleNode("//Encounter//Attachment//FileData").text, Bin )
			ed_File := FileOpen( complDir . fileOut ".pdf", "w")
			ed_File.RawWrite(Bin, nBytes)
			ed_File.Close
			
			fileWQ := enc_dt.MDY "," 			 										; date processed and MA user
					. """" nm """" ","													; CIS name
					. """" encMRN """" ","												; CIS MRN
					. """" fldval["dev-Enc"] """"										; Acct Num
					. "`n"
			FileAppend, %fileWQ%, %trreatDir%logs\trreatWQ.csv									; Add to logs\fileWQ list
			FileCopy, %trreatDir%logs\trreatWQ.csv, %chipDir%trreatWQ-copy.csv, 1
		}
		FileMove, %binDir%%fileOut%.rtf, %reportDir%%fileOut%.rtf, 1					; move RTF to the final directory
		FileCopy, %fileIn%, %complDir%%fileOut%%ext%, 1									; copy PDF to complete directory
		fileDelete, %fileIn%
		
		t_now := A_Now
		edID := "/root/work/id[@ed='" t_now "']"
		xl.addElement("id","/root/work",{date: enc_dt.YMD, ser:fldval["dev-IPG_SN"], ed:t_now, au:user})
			xl.addElement("order",edID,{id:fldval["dev-wqid"]})
			xl.addElement("name",edID,fldval["dev-Name"])
			xl.addElement("dev",edID,fldval["dev-IPG"])
			xl.addElement("status",edID,"Pending")
			xl.addElement("paceart",edID,printQ(is_remote,"True"))
			xl.addElement("file",edID,complDir fileOut ext)
			xl.addElement("meta",edID,(pat_meta) ? complDir fileOut ".meta" : "")
			xl.addElement("report",edID,reportDir fileOut ".rtf")
		xl.save(worklist)
		eventlog("Record added to worklist.xml")
		
		if !(isDevt) {
			whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")							; initialize http request in object whr
			whr.Open("GET"																; set the http verb to GET file "change"
				, "https://depts.washington.edu/pedcards/change/direct.php?" 
					. "do=trreat" 
					. "&to=" enc_MD
				, true)
			whr.Send()																	; SEND the command to the address
			eventlog("Notification email sent to " enc_MD)
			MsgBox, 64,, % "Email sent to " enc_MD
			;~ whr.WaitForResponse()	
			;~ err := whr.ResponseText													; the http response
		}
	}
	gosub parseGUI
	
	return
}

Base64Dec( ByRef B64, ByRef Bin ) {  ; By SKAN / 18-Aug-2017
; from https://autohotkey.com/boards/viewtopic.php?t=35964
Local Rqd := 0, BLen := StrLen(B64)                 ; CRYPT_STRING_BASE64 := 0x1
  DllCall( "Crypt32.dll\CryptStringToBinary", "Str",B64, "UInt",BLen, "UInt",0x1
         , "UInt",0, "UIntP",Rqd, "Int",0, "Int",0 )
  VarSetCapacity( Bin, 128 ), VarSetCapacity( Bin, 0 ),  VarSetCapacity( Bin, Rqd, 0 )
  DllCall( "Crypt32.dll\CryptStringToBinary", "Str",B64, "UInt",BLen, "UInt",0x1
         , "Ptr",&Bin, "UIntP",Rqd, "Int",0, "Int",0 )
Return Rqd
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

parseORM() {
/*	parse fldval values to values
	including aliases for both WQlist and readWQorder
*/
	global fldval, sitesLong
	
	monType:=(tmp:=fldval.OBR_TestName)~="i)14 DAY" ? "BGM"
		: tmp~="i)24 HOUR" ? "HOL"
		: tmp~="i)RECORDER" ? "BGH"
		: ""
	encType:=(tmp:=fldval.PV1_PtClass)="O" ? "Outpatient" 
		: tmp="I" ? "Inpatient"
		: "Other"
	switch encType
	{
		case "Outpatient":
			location := sitesLong[fldval.PV1_Location]
		case "Inpatient":
			location := fldval.PV1_Location
		case "SurgCntr":
			location := "SurgCntr"
		case "Emergency":
			location := "Emergency"
		default:
			location := encType
	}
	;~ location := (encType="Outpatient") ? sitesLong[fldval.PV1_Location]
		;~ : encType
	
	return {date:parseDate(fldval.PV1_DateTime).YMD
		, encDate:parseDate(fldval.PV1_DateTime).YMD
		, nameL:fldval.PID_NameL
		, nameF:fldval.PID_NameF
		, name:fldval.PID_NameL strQ(fldval.PID_NameF,", ###")
		, mrn:fldval.PID_PatMRN
		, sex:(fldval_.PID_sex~="F") ? "Female" : "Male"
		, DOB:parseDate(fldval.PID_DOB).MDY
		, monitor:monType
		, mon:monType
		, provider:fldval.ORC_ProvNameL strQ(fldval.ORC_ProvNameF,", ###")
		, prov:fldval.ORC_ProvNameL strQ(fldval.ORC_ProvNameF,", ###")
		, type:encType
		, loc:location
		, Account:fldval.ORC_ReqNum
		, order:fldval.ORC_ReqNum
		, accession:fldval.ORC_FillerNum
		, acct:location strQ(fldval.ORC_ReqNum,"_###") strQ(fldval.ORC_FillerNum,"-###")
		, UID:tobase(fldval.ORC_ReqNum fldval.ORC_FillerNum,36)
		, ind:strQ(fldval.OBR_ReasonCode,"###") strQ(fldval.OBR_ReasonText,"^###")
		, indication:strQ(fldval.OBR_ReasonCode,"###") strQ(fldval.OBR_ReasonText,"^###")
		, indicationCode:fldval.OBR_ReasonCode
		, orderCtrl:fldval.ORC_OrderCtrl
		, ctrlID:fldval.MSH_CtrlID}
}

FetchDem:
{
	if !(fldval["dev-MRN"]~="^\d{6,7}$") {				; Check MRN parsed from PDF
		fldval["dev-MRN"] := ""
	}
	y := new XML(chipDir "currlist.xml")
	yArch := new XML(chipDir "archlist.xml")
	SNstring := "/root/id[data/device[@SN='" fldval["dev-IPG_SN"] "']]"
	if IsObject(k := y.selectSingleNode(SNstring)) {							; Device SN found
		fldval["dev-MRN"] := k.getAttribute("mrn")								; set dev-MRN based on device SN
		fldval["dev-NameL"] := k.selectSingleNode("demog/name_last").text
		fldval["dev-NameF"] := k.selectSingleNode("demog/name_first").text
		fldval["dev-Name"] := fldval["dev-NameL"] strQ(fldval["dev-NameF"],", ###")
		eventlog("Device " fldval["dev-IPG_SN"] " found in currlist (" fldval["dev-MRN"] ").")
	} else if IsObject(k := yArch.selectSingleNode(SNstring)) {					; Look in yArch if not in y
		fldval["dev-MRN"] := k.getAttribute("mrn")
		fldval["dev-NameL"] := k.selectSingleNode("demog/name_last").text
		fldval["dev-NameF"] := k.selectSingleNode("demog/name_first").text
		fldval["dev-Name"] := fldval["dev-NameL"] strQ(fldval["dev-NameF"],", ###")
		eventlog("Device " fldval["dev-IPG_SN"] " found in archlist (" fldval["dev-MRN"] ").")
	}
	
	fetchQuit := false
	scanOrders()
	matchOrder()
	
	if (fetchQuit) {
		return
	}
	
	return
}

scanOrders() {
	global xl, hl7inDir, fldval, worklist
	
	xl := new XML(worklist)																; refresh worklist
	if !IsObject(xl.selectSingleNode("/root/orders")) {
		xl.addElement("orders","/root")
	}
	
	Loop, files, % hl7InDir "*"															; Scan incoming folder for new orders and add to Orders node
	{
		e0 := {}
		fileIn := A_LoopFileName
		if RegExMatch(fileIn,"_([a-zA-Z0-9]{4,})Z.hl7",i) {								; skip old files
			continue
		}
		processhl7(A_LoopFileFullPath)
		e0:=parseORM()
		e0.orderNode := "/root/orders/order[ordernum='" e0.order "']"
		if IsObject(k:=xl.selectSingleNode(e0.orderNode)) {								; ordernum node exists
			e0.nodeCtrlID := k.selectSingleNode("ctrlID").text
			if (e0.CtrlID < e0.nodeCtrlID) {											; order CtrlID is older than existing, somehow
				FileDelete, % hl7InDir fileIn
				eventlog("Order msg " fileIn " is outdated.")
				continue
			}
			if (e0.orderCtrl="CA") {													; CAncel an order
				FileDelete, % hl7InDir fileIn											; delete this order message
				FileDelete, % hl7InDir "*_" e0.UID "Z.hl7"								; and the previously processed hl7 file
				removeNode(e0.orderNode)												; and the accompanying node
				eventlog("Cancelled order " e0.order ".")
				continue
			}
			FileDelete, % hl7InDir "*_" e0.UID "Z.hl7"									; delete previously processed hl7 file
			removeNode(e0.orderNode)													; and the accompanying node
			eventlog("Cleared order " e0.order " node.")
		}
		if (e0.orderCtrl="XO") {														; change an order
			e0.orderNode := "/root/orders/order[accession='" e0.accession "']"
			k := xl.selectSingleNode(e0.orderNode)
			e0.nodeUID := k.getAttribute("id")
			FileDelete, % hl7InDir "*_" e0.nodeUID "Z.hl7"
			removeNode(e0.orderNode)
			eventlog("Removed node id " e0.nodeUID " for replacement.")
		}
		
		newID := "/root/orders/order[@id='" e0.UID "']"								; otherwise create a new node
		xl.addElement("order","/root/orders",{id:e0.UID})
		xl.addElement("ordernum",newID,e0.order)
		xl.addElement("accession",newID,e0.accession)
		xl.addElement("ctrlID",newID,e0.CtrlID)
		xl.addElement("date",newID,e0.date)
		xl.addElement("name",newID,e0.name)
		xl.addElement("mrn",newID,e0.mrn)
		xl.addElement("sex",newID,e0.sex)
		xl.addElement("dob",newID,e0.dob)
		xl.addElement("mon",newID,e0.mon)
		xl.addElement("prov",newID,e0.prov)
		xl.addElement("site",newID,e0.loc)
		xl.addElement("acct",newID,e0.acct)
		xl.addElement("ind",newID,e0.ind)
		eventlog("Added order ID " e0.UID ".")
		
		fileOut := e0.MRN "_" 
			. fldval["PID_nameL"] "^" fldval["PID_nameF"] "_"
			. e0.date "_"
			. e0.uid "Z.hl7"
			
		FileMove, %A_LoopFileFullPath%													; and rename ORM file
			, % hl7InDir . fileOut
		
	}
	xl.save(worklist)
		
	return
}

matchOrder() {
	global fldval, xl, fetchQuit
	static selbox, selbut
	key := {}
	
	fldName := Format("{:U}",fldval["dev-Name"])
	Loop, % (k:=xl.selectNodes("/root/orders/order")).length							; generate list of orders with fuzz levels
	{
		node := k.item(A_index-1)
		nodeName := node.selectSingleNode("name").text
		nodeMRN := node.selectSingleNode("mrn").text
		nodeID := node.getAttribute("id")
		nodeOrdernum := node.selectSingleNode("ordernum").text
		nodeAccession := node.selectSingleNode("accession").text
		fuzz := fuzzysearch(nodename,fldName)
		list .= fuzz "|" nodeID "|" nodeName "|" nodeMRN "|" nodeOrdernum "|" nodeAccession "`n"
	}
	Sort, list																			; sort by fuzz level
	Loop, parse, list, `n
	{
		k := A_LoopField
		if (k="") {
			break
		}
		vals:=strsplit(k,"|")
		key[A_Index] := {name:vals[3]													; build array of key{name,id,etc}
						,id:vals[2]
						,mrn:vals[4]
						,ordernum:vals[5]
						,accession:vals[6]}
		keylist .= vals[3] "|"															; keylist just contains the name
	}
	
	Gui, dev:Destroy
	Gui, dev:Default
	Gui, -MinimizeBox
	Gui, Add, Text, w180 +Wrap
		, % "Select the order that matches this patient:"
	Gui, Font, s12
	Gui, Add, ListBox, h100 vSelBox -vScroll AltSubmit gMatchOrderSelect, % keylist		; listbox and button
	Gui, Add, Button, h30 vSelBut gMatchOrderSubmit Disabled, Submit					; disabled by default
	Gui, Show, AutoSize, Select order
	Gui, +AlwaysOnTop
	
	winwaitclose, Select order
	
	if !(selbox) {																		; no selection
		fetchQuit := true
		return
	}
	
	res := key[selbox]
	if (fuzzysearch(res.name , fldval["dev-name"]) > 0.20) {						; possible bad match
		MsgBox, 262196
			, Possible name mismatch
			, % "Order name: " res.Name "`n"
			. "Report name: " fldval["dev-name"] "`n`n"
			. "Keep " res.name "?"
		IfMsgBox, No 
		{
			fetchQuit:=true
			return
		}
	}
	
	fldval["dev-name"] := res.name
		fldval["dev-nameL"] := parseName(res.name).last
		fldval["dev-nameF"] := parseName(res.name).first
	fldval["dev-MRN"] := res.mrn
	fldval["dev-wqid"] := res.id
	fldval["dev-ordernum"] := res.ordernum
	fldval["dev-accession"] := res.accession
	
	return 
	
	matchOrderSelect:
	{
		GuiControl, dev:Enable, Submit
		return
	}
	
	matchOrderSubmit:
	{
		Gui, dev:Submit
		return
	}
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
	EncName := (fldval["dev-Name"]~="[A-Z \-]+, [A-Z\-](?!=\s)")
	demBits := ((EncNum~="\d{8}") && (EncMRN~="\d{6,7}") && EncName)			; clear the error check
	/*	set this as true to skip demographics validation
	*/
		;~ dembits := true
	/*
	*/
	
	Gui, fetch:Destroy
	Gui, fetch:+AlwaysOnTop
	
	Gui, fetch:Add, Text, % "x" fX1 " w" fW1 " h" fH " c" ((encName)?"Default":"Red") , Name
	Gui, fetch:Add, Edit, % "x" fX2 " yP-4" " w" fW2 " h" fH 
		. " readonly c" ((encName)?"Default":"Red") , % fldval["dev-Name"]
	
	Gui, fetch:Add, Text, % "x" fX1 " w" fW1 " h" fH " c" ((encMRN~="\d{6,7}")?"Default":"Red") , MRN
	Gui, fetch:Add, Edit, % "x" fX2 " yP-4" " w" fW2 " h" fH 
		. " readonly c" ((encMRN~="\d{6,7}")?"Default":"Red"), % fldval["dev-MRN"]
	
	Gui, fetch:Add, Text, % "x" fX1 " w" fW1 " h" fH " c" ((encNum~="\d{8}")?"Default":"Red") , Encounter
	Gui, fetch:Add, Edit, % "x" fX2 " yP-4" " w" fW2 " h" fH 
		. " readonly c" ((encNum~="\d{8}")?"Default":"Red"), % fldval["dev-Enc"]
	
	Gui, fetch:Add, Button, % "x" fX1 " yP+" fYD " h" fH+10 " w" fW1+fW2+10 " gfetchSubmit " ((demBits)?"":"Disabled"), Submit!
	Gui, fetch:Show, AutoSize, % fldval["dev-Name"]
	return
}

fetchGuiClose:
{
	Gui, fetch:destroy
	getDem := false																	; break out of fetchDem loop
	fetchQuit := true
	eventlog("Manual [x] out of fetchDem.")
Return
}

parseClip(clip) {
/*	If clip matches "val1:val2" format, and val1 in demVals[], return field:val
	If clip contains proper Encounter Type ("Outpatient", "Inpatient", "Observation", etc), return Type, Date, Time
*/
	if (clip~="[A-Z \-]+, [A-Z \-]+") {													; matches name format "SMITH, WILLIAM JAMES"
		nameL := trim(strX(clip,"",1,0,",",1,1))
		nameF := trim(strX(clip,",",1,1," ",1,1))
		return {field:"Name", value:nameL ", " nameF}
	}
	
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
	yID := y.selectSingleNode(MRNstring)
	
	if IsObject(q := yID.selectSingleNode("diagnoses/epdevice")) {				; Clear prior <epdevice>
		q.parentNode.removeChild(q)
	}
	y.addElement("epdevice", MRNstring "/diagnoses")
	y.addElement("dependent", MRNstring "/diagnoses/epdevice", fldval["dependent"])
	y.addElement("indication", MRNstring "/diagnoses/epdevice", fldval["indication"])
	WriteOut(MRNstring "/diagnoses", "epdevice")
	
	if IsObject(yDev := yID.selectSingleNode("data/device")) 	{				; Clear out any existing Device node
		yDev.parentNode.removeChild(yDev)
		eventlog("Removed existing <device> node.","C")							; chipotle\logs
		eventlog("Removed existing <device> node from currlist.")				; trreat\logs
	}
	y.addElement("device"
		,MRNstring "/data"
		,{	au:A_UserName
		,	ed:A_Now
		,	model:fldval["dev-IPG"]
		,	SN:fldval["dev-IPG_SN"]} )
	pmNowString := MRNstring "/data/device"
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
	WriteOut(MRNstring "/data", "device")
	eventlog("Add new <device> node.","C")
	eventlog("Add new <device> node to currlist.")
	
	return
}

makeReport:
{
	is_remote := (fldval["dev-EncType"]="REMOTE") ? true : ""
	EncMRN := fldval["dev-MRN"]
	MRNstring := "/root/id[@mrn='" EncMRN "']"
	if !IsObject(y.selectSingleNode(MRNstring)) {
		y.addElement("id", "root", {mrn: EncMRN})								; No MRN node exists, create it.
		FetchNode("demog")
		FetchNode("diagnoses")													; Check for existing node in Archlist,
		FetchNode("prov")														; retrieve old Dx, Prov. Otherwise, create placeholders.
		FetchNode("data")
	}
	if !IsObject(y.selectSingleNode(MRNstring "/data")) {						; Make sure <data> exists
		y.addElement("data",MRNstring)
	}
	
	fldval["dependent"] := y.selectSingleNode(MRNstring "/diagnoses/epdevice/dependent").text
	fldval["indication"] := y.selectSingleNode(MRNstring "/diagnoses/epdevice/indication").text
	ciedGUI()
	if (fetchQuit) {
		return
	}
	
	tech := cMsgBox("Technician","Device check performed by:","Jenny Keylon, RN|Device rep","Q","")
	if (tech="Close") {
		fetchQuit := true
		return
	}
	
	summ := fldval["dev-summary"]
	if (summ="") {
		summ := cMsgBox("Title","Choose a text","Normal device check|none","Q","")
		if (summ="Close") {
			fetchQuit := true
			return
		}
		if instr(summ,"normal") {
			summ := "This represents a normal " format("{:L}",fldval["dev-EncType"]) " device check. The patient denies any device related symptoms. "
				. "The battery status is normal. Sensing and capture thresholds are good. The lead impedances are normal. "
				. "Routine follow up per implantable device protocol. "
			eventlog("Normal summary template selected.")
		} else {
			summ := ""
			eventlog("Blank report summary.")
		}
	}
	
	gosub saveChip
	
	gosub checkEP
	
	gosub pmPrint
	
	return
}

ciedGUI() {
	global fldval, tmpBtn, fetchQuit
	static DepY, DepN, DepX, Ind
	tmpBtn := ""
	
	gui, cied:Destroy
	gui, cied:Add, Text, , Pacemaker dependent?
	gui, cied:Add, Radio, % "vDepY Checked" (fldval["dependent"]="Yes"), Yes
	gui, cied:Add, Radio, % "vDepN Checked" (fldval["dependent"]="No") , No
	gui, cied:Add, Radio, vDepX, Clear
	gui, cied:Add, Text
	gui, cied:Add, Text, , Indication for device
	gui, cied:Add, Edit, r3 w200 vInd, % fldval["indication"]
	gui, cied:Add, Text
	gui, cied:Add, Button, w100 h30, OK
	
	gui, cied:Show, AutoSize
	
	loop
	{
		if (tmpBtn) {
			break
		}
	}
	gui, cied:Submit, NoHide
	gui, cied:Destroy
	
	if (tmpBtn="x") {
		fetchQuit := true
		return
	}
	
	fldval["dependent"] := (depY) 
		? "Yes"
			: (depN)
		? "No"
			: ""
	fldval["indication"] := Ind
	
	return
}

ciedGuiEscape:
ciedGuiClose:
{
	tmpBtn := "x"
	return
}

ciedButtonOK:
{
	tmpBtn := "ok"
	return
}

checkEP:
{
/*	Find responsible EP
	and/or assign to someone
*/
	yID := y.selectSingleNode(MRNstring)
	
	if !(yEP := yID.selectSingleNode("prov").getAttribute("EP")) {						; Assign a primary EP in prov if it does not exist
		eventlog("No primary EP found.")
		yEP := cMsgBox("No associated EP found"
						,"Assign a primary EP`nClose [x] if none"
						,"T. Chun|J. Salerno|S. Seslar"
						,"Q","")
		if !(yEP=="Close") {
			yID.selectSingleNode("prov").setAttribute("EP", yEP)
			yID.selectSingleNode("prov").setAttribute("au", A_UserName)
			yID.selectSingleNode("prov").setAttribute("ed", A_Now)
			eventlog(yEP " set as primary EP.")
			eventlog(yEP " set as primary EP.","C")
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
	eventlog("Report assigned to " enc_MD ".")
	
	Return
}

readWQ(idx) {
	global xl
	
	res := []
	k := xl.selectSingleNode("//order[@id='" idx "']")
	Loop, % (ch:=k.selectNodes("*")).Length
	{
		i := ch.item(A_index-1)
		node := i.nodeName
		val := i.text
		res[node]:=val
	}
	res.node := k.parentNode.nodeName 
	
	return res
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
	eventlog(path "/" node " saved.","C")
	eventlog("CHIPOTLE currlist updated.")
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

eventlog(event,ch:="") {
	global user, trreatDir, chipDir
	dir := (ch="C") ? chipDir "logs\" : trreatDir "logs\"
	comp := A_ComputerName
	FormatTime, sessdate, A_Now, yyyyMM
	FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
	name := dir . sessdate . ".log"
	txt := now " [" user "/" comp "] " event "`n"
	filePrepend(txt,name)
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

ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-_/]"
	date := []
	time := []
	x := RegExReplace(x,"[,\(\)]")
	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",d) {			; 03-Jan-2015
		date.dd := zdigit(d1)
		date.mmm := d2
		date.mm := zdigit(objhasvalue(mo,d2))
		date.yyyy := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"i)(" moStr "|\d{1,2})" dSep "(\d{1,2})" dSep "(\d{4}|\d{2})",d) {	; Jan-03-2015, 01-03-2015
		date.dd := zdigit(d2)
		date.mmm := objhasvalue(mo,d1) 
			? d1
			: mo[d1]
		date.mm := objhasvalue(mo,d1)
			? zdigit(objhasvalue(mo,d1))
			: zdigit(d1)
		date.yyyy := (d3~="\d{4}")
			? d3
			: (d3>50)
				? "19" d3
				: "20" d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"\b(\d{4})-?(\d{2})-?(\d{2})\b",d) {								; 20150103 or 2015-01-03
		date.yyyy := d1
		date.mm := d2
		date.mmm := mo[d2]
		date.dd := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"\b(\d{4})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",d)  {			; 20150103174307
		date.yyyy := d1
		date.mm := d2
		date.mmm := mo[d2]
		date.dd := d3
		date.date := d1 "-" d2 "-" d3
		
		time.hr := d5
		time.min := d6
		time.sec := d7
		time.time := d5 ":" d6 . strQ(d7,":###")
	}
	
	if RegExMatch(x,"iO)(\d{1,2}):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",t) {				; 17:42 PM
		hasDays := (t.value[4]) ? true : false 												; 4 nums has days
		time.days := (hasDays) ? t.value[1] : ""
		time.hr := zdigit(t.value[1+hasDays])
		time.min := trim(t.value[2+hasDays]," :")
		time.sec := trim(t.value[3+hasDays]," :")
		time.ampm := trim(t.value[5])
		time.time := trim(t.value)
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, days:time.days, hr:time.hr, min:time.min, sec:time.sec, ampm:time.ampm, time:time.time}
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

strQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	return (var1="") ? null : RegExReplace(txt,"###",var1)
}

ToBase(n,b) {
/*	from https://autohotkey.com/board/topic/15951-base-10-to-base-36-conversion/
	n >= 0, 1 < b <= 36
*/
   Return (n < b ? "" : ToBase(n//b,b)) . ((d:=mod(n,b)) < 10 ? d : Chr(d+55))
}

readIni(section) {
/*	Reads a set of variables
	[section]					==	 		var1 := res1, var2 := res2
	var1=res1
	var2=res2
	
	[array]						==			array := ["ccc","bbb","aaa"]
	=ccc
	bbb
	=aaa
	
	[objet]						==	 		objet := {aaa:10,bbb:27,ccc:31}
	aaa:10
	bbb:27
	ccc:31
*/
	global
	local x, i, key, val
		, i_res := object()
		, i_type := []
		, i_lines := []
	i_type.var := i_type.obj := i_type.arr := false
	IniRead,x,includes\trreat.ini,%section%
	Loop, parse, x, `n,`r																; analyze section struction
	{
		i := A_LoopField
		if (i~="(?<!"")[=]")															; find = not preceded by "
		{
			if (i ~= "^=") {															; starts with "=" is an array list
				i_type.arr := true
			} else {																	; "aaa=123" is a var declaration
				i_type.var := true
			}
		} else																			; does not contain a quoted =
		{
			if (i~="(?<!"")[:]") {														; find : not preceded by " is an object
				i_type.obj := true
			} else {																	; contains neither = nor : can be an array list
				i_type.arr := true
			}
		}
	}
	if ((i_type.obj) + (i_type.arr) + (i_type.var)) > 1 {								; too many types, return error
		return error
	}
	Loop, parse, x, `n,`r																; now loop through lines
	{
		i := A_LoopField
		if (i_type.var) {
			key := strX(i,"",1,0,"=",1,1)
			val := strX(i,"=",1,1,"",0)
			%key% := trim(val,"""")
		}
		if (i_type.obj) {
			key := strX(i,"",1,0,":",1,1)
			val := strX(i,":",1,1,"",0)
			i_res[key] := trim(val,"""")
		}
		if (i_type.arr) {
			i := RegExReplace(i,"^=")													; remove preceding =
			i_res.push(trim(i,""""))
		}
	}
	return i_res
}

#Include strx.ahk
#Include xml.ahk
#Include CMsgBox.ahk
#Include sift3.ahk
#Include hl7.ahk
