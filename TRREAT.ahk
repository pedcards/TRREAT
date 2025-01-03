/*	TRREAT - The Rhythm Recording Electronic Analysis Transmogrifier - PM
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
;~ SetWorkingDir %A_ScriptDir%							; Don't set workingdir so can read files from run dir
;~ FileInstall, pdftotext.exe, pdftotext.exe
#Include %A_ScriptDir%\includes
#Requires AutoHotkey v1.1

Progress,100,Checking paths...,TRREAT
IfInString, A_ScriptDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	isDevt := true
	trreatDir := ".\"
	path:=readIni("devpaths")
	eventlog(">>>>> Started in DEVT mode.")
} else {
	isDevt := false
	trreatDir := A_ScriptDir "\"														; need to define this before readIni
	path:=readIni("paths")
	eventlog(">>>>> Started in PROD mode. " A_ScriptName " ver " substr(tmp,1,12))
}

;~ chipDir:=path.chip																	; CHIPOTLE root
;~ pdfDir:=path.pdf																		; USB root
path.trreat		:= trreatDir															; TRREAT root
path.files		:= path.trreat "files\"													; ini and xml files
path.report		:= path.trreat "pending\"												; parsed reports and rtf pending
path.compl		:= path.trreat "completed\"												; signed rtf with PDF and meta files
path.paceart	:= path.trreat "paceart\"												; PaceArt import xml
path.hl7in		:= path.trreat "epic\Orders\"											; inbound Epic ORM
path.outbound	:= path.trreat "epic\OutboundHL7\"										; outbound ORU for Ensemble
path.onbase		:= path.trreat "onbase\import\"											; onbase DRIP folder for PDFs

worklist := path.files "worklist.xml"

utcDiff := setUTC()

user := instr(A_UserName,"octe") ? "tchun1" : A_UserName
docs := readIni("docs")
parsedocs(docs)

initHL7()
hl7DirMap := {}

eventLog(">>>>> Session started...")
if !FileExist(path.report) {
	MsgBox % "Requires pending dir`n""" path.report """"
	ExitApp
}
if !FileExist(path.compl) {
	MsgBox % "Requires completed dir`n""" path.compl """"
	ExitApp
}
if !FileExist(path.chip) {
	MsgBox % "Requires CHIPOTLE dir`n""" path.chip """"
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
	
	gosub parseGUI
	
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
	Gui, Add, Tab3, vWQtab +HwndWQtab,Paceart saves
	
	Gui, Tab, Paceart
	Gui, Add, Listview, w800 -Multi Grid r12 gparsePat vWQlvP hwndHLVp, Date|Name|Device|Serial|Status|PaceArt|FileName|MetaData|Report
	
	gosub readFiles																		; scan the folders
	
	fixWqlvCols("WQlv")
	
	progress, off
	
	Gui, Show,, TRREAT Reports and File Manager
	WinActivate, TRREAT Reports
	return
}

ParseGuiClose:
{
	Loop, files, % path.files "tmp\*"
	{
		dt := A_now
		dt -= A_LoopFileTimeModified, Days
		if (dt > 30) {
			FileDelete, %A_LoopFileLongPath%
		}
	}
	eventlog("<<<<< Parse session closed.")
	ExitApp
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
		xl.transformXML()
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
		if (tmp.status="Sent") && (tmp.paceart="True" || tmp.file~="\.xml")			; SENT and in PACEART
		{
			fileNum += 1
			LV_Add("", tmp.date)
			LV_Modify(fileNum,"col2", tmp.name)										; add marker line if in DONE list
			LV_Modify(fileNum,"col3", "[DONE]")
			archNode("/root/work/id[@date='" tmp.date "'][@ser='" tmp.ser "']")		; copy ID node to DONE
			xl.transformXML()
			xl.save(worklist)
			eventlog("Node " tmp.date "/" tmp.ser "/" tmp.name " archived.")
			continue
		}
		
		fileNum += 1																; Add a row to the LV
		LV_Add("", tmp.date)														; col1 is date
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
	; readFilesRootMDT()
	; readFilesSJM()
	; readFilesBSCI()
	readFilesPaceart()
	
return
}

readFilesRootMDT() {
/*	Read root - usually MEDT files
*/
	global path, xl, filenum, WQlvP, WQlv, HLVp, HLV
		, fields, labels, fldval
	
	progress, 40,Medtronic
	Loop, files, % path.pdf "SmartSyncPDF*.pdf"									; first pass: Look for SmartSyncPDF files, rename
	{
		fields := []
		labels := []
		fldval := []
		fullpath := A_LoopFileFullPath
		filename := A_LoopFileName
		
		txt := readPDF(fullpath)
		fields[1] := ["Device","Serial Number","Date of Visit"
					, "Patient","ID","Physician","History","`n"]
		labels[1] := ["IPG","IPG_SN","Encounter"
					, "Name","MRN","Physician","Indication","null"]
		fieldvals(txt,1)
		dt := parseDate(fldval.Encounter)
		newfnam := fldval.Name "_" fldval.IPG_SN "_SmartSync_" dt.mm "_" dt.dd "_" dt.YYYY ".pdf"
		FileMove, % path.pdf filename, % path.pdf newfnam
	}
	
	Loop, files, % path.pdf "*.pdf"												; read all PDFs in root
	{
		tmp := []
		tmp.file := A_LoopFileName												; next file in PDFdir
		if instr(tmp.maxstr,tmp.file) {											; in skiplist?
			continue
		}
		
		tmp.max := 1															; reset max k counter
		Loop, files, % path.pdf strX(tmp.file,"",1,0,"_",0,1) "*.pdf"			; loop through all files with this "prefix"
		{
			i := A_LoopFileName													; i is filename in this inner loop
			n := substr(i,instr(i,"_",,-1))										; n is string up to final _#
			k := strX(i,"_",n,1,".",1)											; k is # between _ and .pdf
			if (k > tmp.max) {													; greater than previous kmax?
				j := substr(i,1,instr(i,"_",,-1)) (tmp.max) ".pdf"				; j is filename of previous kmax
				FileMove, % path.pdf j, % path.pdf j ".old"							; rename it to j.pdf.old
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
			tmp.file := path.pdf tmp.file
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
	global path, xl, filenum, WQlvP, WQlv, HLVp, HLV
	
	progress, 60,St Jude/Abbott
	sjmDir := path.pdf "PDFs\Live.combined"
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
				Loop, Files, % path.pdf "*.log", F
				{
					k := RegExReplace(A_LoopFileName,".log")
					if InStr(tmp.ser,k) {
						tmp.meta := path.pdf k ".log"
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
	global xl, path, filenum, bscBnk, WQlvP, WQlv, HLVp, HLV
	
	progress, 80,Boston Scientific
	tmp := []
	bscDir := path.pdf "bsc\patientData\"
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
		tmp.name := readBnk("PatientLastName") ", " readBnk("PatientFirstName")
		tmp.dev := "BSCI " readBnk("SystemName") " " strX(readBnk("SystemModelNumber"),"",1,0,"-",1)
		tmp.ser := readBnk("SystemSerialNumber")
		tmp.node := "id[@date='" td.YMD "'][@ser='" tmp.ser "']"
		
		if IsObject(xl.selectSingleNode("/root/work/" tmp.node)) {
			eventlog("BSC: Skipping " td.YMD "\" tmp.ser ", already in worklist.")
			continue															; skip reprocessing in WORK list
		}
		if IsObject(xl.selectSingleNode("/root/done/ " tmp.node)) {
			fileNum += 1
			LV_Add("", td.YMD)
			LV_Modify(fileNum,"col2", tmp.name)									; add marker line if in DONE list
			LV_Modify(fileNum,"col3", "[DONE]")
			eventlog("BSC: File " td.YMD "\" tmp.ser " already DONE.")
			continue
		}
		
		Loop, files, % patDir "\report\Combined*" td.MMM "-" td.DD "-" td.YYYY "*.pdf"
		{
			tmp.file := A_LoopFileFullPath										; find the appropriate PDF matching this .bnk file
			eventlog("BSC: " A_LoopFileName " found.")
		}
		
		fileNum += 1															; Add a row to the LV
		LV_Add("", td.YMD)														; col1 is date
		LV_Modify(fileNum,"col2", tmp.name)
		LV_Modify(fileNum,"col3", tmp.dev)
		LV_Modify(fileNum,"col4", tmp.ser)
		LV_Modify(fileNum,"col5", "")
		LV_Modify(fileNum,"col6", "")
		LV_Modify(fileNum,"col7", tmp.file)
		LV_Modify(fileNum,"col8", tmp.bnk)
	}
	
	progress, 100
	bscDir := path.pdf "DataFiles\"
	loop, Files, % bscDir "*", D												; Loop through subdirs of Emblem datafiles
	{
		tmp := A_LoopFileName
		MMDD := substr(tmp,1,4)
		YYYY := SubStr(tmp,5,4)
		dirdate := YYYY MMDD 													; correct their weird date format
		dt := A_Now
		dt -= dirdate , Days													; calculate days since check
		dirlist .= Format("{:04}",dt) "|" dirdate "|" tmp "`n"
	}
	sort dirlist																; sort from most recent
	dirlist := strX(dirlist,"",1,0,"`n",1,1)
	Loop , parse, dirlist, `n, `n
	{
		dirName := StrSplit(A_LoopField,"|").3 "\Sessions\"
		name := {}
		loop, Files, % bscDir dirName "*.hl7", F
		{
			snam := RegExReplace(A_LoopFileName,"__(.*)","__")
			maxdate :=
			loop, Files, % bscDir dirName snam "*.hl7"
			{
				fnam := A_LoopFileName
				RegExMatch(fnam,"(.*)-(.*)-(.*)-(.*)__(.*)\.hl7",x)
				dt := parseDate(x4).YMD RegExReplace(x5,"[.:]")
				if (dt>maxdate) {
					maxdate:=dt
					maxfnam:=fnam
				}
			}
			tmp := []
			tmp.fnam := maxfnam
			RegExMatch(tmp.fnam,"(.*)-(.*)-(.*)-(.*)__(.*)\.hl7",x)
			tmp.name := x1
			tmp.model := "BSCI " x2
			tmp.ser := x3
			tmp.date := parseDate(x4).YMD 
			tmp.time := RegExReplace(x5,"[.:]")
			dt := tmp.date tmp.time 
			tmp.node := "id[@date='" tmp.date "'][@ser='" tmp.ser "']"
			if instr(name[tmp.name],dt) {
				continue
			}
			if IsObject(xl.selectSingleNode("/root/work/" tmp.node)) {
				eventlog("BSC: Skipping " tmp.date "\" tmp.ser ", already in worklist.")
				continue																; skip reprocessing in WORK list
			}
			if IsObject(xl.selectSingleNode("/root/done/ " tmp.node)) {
				fileNum += 1
				LV_Add("", tmp.date)
				LV_Modify(fileNum,"col2", tmp.name)										; add marker line if in DONE list
				LV_Modify(fileNum,"col3", "[DONE]")
				eventlog("BSC: File " tmp.date "\" tmp.ser " already DONE.")
				continue
			}
			name[tmp.name] .= dt "`n"
			tmp.fnam := bscDir dirName tmp.fnam
			tmp.file := bscDir dirName RegExReplace(fnam,".hl7",".pdf")
			
			fileNum += 1																; Add a row to the LV
			LV_Add("", tmp.date)														; col1 is date
			LV_Modify(fileNum,"col2", tmp.name)
			LV_Modify(fileNum,"col3", tmp.model)
			LV_Modify(fileNum,"col4", tmp.ser)
			LV_Modify(fileNum,"col5", tmp.dev)
			LV_Modify(fileNum,"col6", "")
			LV_Modify(fileNum,"col7", tmp.file)
			LV_Modify(fileNum,"col8", tmp.fnam)
		}
	}
	
	return
}

readFilesPaceart() {
/*	read exported PDF reports from Paceart
	in .\paceart\ folder
*/
	global path, WQlvP, WQlv, HLVp, HLV
	
	progress, 100, Paceart imports
	
	Gui, Listview, WQLVp
	
	loop, files, % path.paceart "*.xml"
	{
		fileIn := path.paceart A_LoopFileName
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
			fileOut := path.paceart . dem.mrn "_" dem.nameL "_" dem.encdate "WQ.xml"
			FileMove, % fileIn, % fileOut, 1
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
	; x := RegExReplace(x,"\'","^")														; replace ['] with [^] to avoid XPATH errors
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
	Gui, ListView, % agc
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
	is_remoteAlert:=
	is_postop:=
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
			xl.transformXML()
			xl.save(worklist)
			eventlog("Node " pat_node " removed from worklist.")
			FileDelete, % pat_report
			eventlog("File " pat_report " deleted.")
			gosub fileLoop
			return
		}
		if instr(tmp,"PaceArt") {
			xl.setText(pat_node "/paceart","True")
			xl.transformXML()
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
		maintxt:=readPDF(fileIn)
		Run, % fileIn
	}
	
	if (maintxt~="Medtronic,\s+Inc|Medtronic Software") {						; PM and ICD reports use common subs
		eventlog("Medtronic identified.")
		gosub Medtronic
	}
	else if (maintxt~="(Boston Scientific Corporation|800\.CARDIAC)") {
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
	
	gosub parseGUI
	return
}

readPDF(fileIn, args="-table") {
/*	Convert PDF to text using pdftotext.exe with optional args
	output file generated in .\tmp\xxxxx.txt
	text returned in result
*/
	global path
	SplitPath, fileIn,,,,fileOut
	FileDelete, % path.files "tmp\" fileOut ".txt"
	RunWait, % path.files "pdftotext.exe " args " """ fileIn """ """ path.files "tmp\" fileOut ".txt""" , , hide
	eventlog("pdftotext " fileIn " -> " path.files "tmp\" fileOut ".txt")
	FileRead, txt, % path.files "tmp\" fileOut ".txt"
	
	return cleanlines(txt)
}

; Builds Epic ORU using values stored in <order\> node.
makeORU(wqid) {
	global xl, fldval, hl7out, docs, path, filenam, isRemote, user
	dict:=readIni("EpicResult")
	
	order := readWQ(wqid)
	eventlog("makeORU: ordertype='" order.ordertype "'")
	
	hl7time := A_Now
	hl7out := Object()
	
	buildHL7("MSH"
		,{1:"^~\&"
		, 2:"CVTRREAT"
		, 3:"CVTRREAT"
		, 4:"HS"
		, 6:hl7time
		, 8:"ORU^R01"
		, 9:wqid
		, 10:"T"
		, 11:"2.5.1"})
	
	buildHL7("PID"
		,{2:order.mrn
		, 3:order.mrn "^^^^CHRMC"
		, 5:parseName(order.name).last "^" parseName(order.name).first
		, 7:parseDate(order.dob).YMD
		, 8:substr(order.sex,1,1)
		, 18:order.accountnum})
	
	buildHL7("PV1"
		,{19:order.encnum
		, 50:wqid})
	
	tmpDoc := docs[order.reading]
	buildHL7("OBR"
		,{2:order.ordernum
		, 3:order.accession
		, 4:order.ordertype "^IMGEAP"
		, 7:order.date
		, 16:order.prov "^^^^^^MSOW_ORG_ID"
		, 25:"P"
		, 32:tmpDoc.NPI "^" tmpDoc.nameL "^" tmpDoc.nameF })
	
	File := path.report fileNam ".rtf"
	FileRead, rtfStr, %File%
	rtfStr := RegExReplace(rtfStr,"\\par\R","\par ")									; correct "\par" from WordPad
	rtfStr := RegExReplace(rtfStr,"\R","")												; remove CRLF 
	rtfStr := StrReplace(rtfStr,"\","\E\")												; replace "\" chars (HL7 "\E\" esc)
	rtfStr := StrReplace(rtfStr,"|","\F\")												; replace "|" chars (HL7 field separator)
	rtfStr := StrReplace(rtfStr,"~","\R\")												; replace "~" chars (HL7 repetition seperator)
	rtfStr := StrReplace(rtfStr,"^","\S\")												; replace "^" chars (HL7 component seperator)
	rtfStr := StrReplace(rtfStr,"&","\T\")												; replace "&" chars (HL7 subcomponent seperator)
	buildHL7("OBX"
		,{2:"FT"
		, 3:"&GDT^PACEMAKER/ICD INTERROGATION"
		, 5:rtfStr
		, 11:"P"
		, 14:hl7time})
	
	for key,val in dict																	; Loop through all values in Dict (from ini)
	{
		str:=StrSplit(val,"^")
		buildHL7("OBX"																	; generate OBX for each value
			,{2:"TX"
			, 3:key "^" str[1] "^IMGLRR"
			, 5:order[str[2]] 
			, 11:"F"
			, 14:hl7time})
	}
	
	return
}

matchEAP(txt) {
	
	EAP := readIni("EpicOrderEAP")
	top := 1.00
	for key,val in EAP
	{
		str := RegExReplace(val,"^.*?\^")
		fuzz := fuzzysearch(txt,str)
		if (fuzz<top) {
			top := fuzz
			best := key
		}
		if instr(str,txt) {
			match := str
		}
	}
	
	eventlog("matchEAP: '" txt "' => '" EAP[best] "' (" best ")")
	eventlog("matchEAP: match='" match "'")
	return EAP[best]
}

Medtronic:
{
	if (maintxt~="Adapta|Sensia") {												; Scan Adapta family of devices
		eventlog("Adapta report.")
		gosub mdtAdapta
	} else if (maintxt~="Medtronic\s+Application ID") {							; or new iPad report
		eventlog("Medtronic Application report.")
		gosub mdtApplication
	} else if (maintxt~="(Quick Look II)|(Final:\s+Session Summary)") {			; or scan more current QuickLook II reports
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

mdtApplication:
{
/*	INITIAL: QUICK LOOK
	- Demographics
	- Device info
	- Check info
*/
	qltxt := maintxt
	inirep := stregX(qltxt,"",1,1,"Device Status",1)
	fields[1] := ["Device","Serial Number","Date of Visit"
				, "Patient","ID","Physician","History","`n"]
	labels[1] := ["IPG","IPG_SN","Encounter"
				, "Name","MRN","Physician","Indication","null"]
	fieldvals(inirep,1,"dev")
	fldval["dev-Encounter"] := parsedate(fldval["dev-Encounter"]).MDY
	fldval["dev-Physician"] := instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") 
		? tmp 
		: "Dr. " trim(tmp," `n")
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
	qltbl := RegExReplace(qltbl,"\s+RRT.*?[\r\n]+")
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
			,fldval["leads-RV_Sthr"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"],fldval["leads-RV_HVimp"])
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
	fields[2] := ["Mode  ","V. Pacing","AdaptivCRT","V-V Pace Delay"
				, "Lower\s+Rate","Upper\s+Track","Upper\s+Sensor"
				, "Paced AV","Sensed AV","Mode Switch"]
	labels[2] := ["Mode","CRT_VP","CRT_VV","CRT_VV","LRL","URL","USR","PAV","SAV","Mode Switch"]
	scanParams(qltbl,2,"par",1)
	
	qltbl := stregX(inirep "<<<","Detection",1,0,"<<<",1)
	fields[2] := ["Rates-AT/AF","Rates-VF","Rates-FVT","Rates-VT"
				, "Therapies-AT/AF","Therapies-VF","Therapies-FVT","Therapies-VT"]
	labels[2] := ["ATAF","VF","FVT","VT"
				, "Rx_ATAF","Rx_VF","Rx_FVT","Rx_VT"]
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
	
	qlObs := stregX(maintxt,"Observations\s+\(",1,0,"Medtronic Software",1)
	fldfill("event-Obs",qlObs)
	
	
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
	fldval["dev-Physician"] := instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") 
		? tmp 
		: "Dr. " trim(tmp," `n")
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
			,fldval["leads-RV_Sthr"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"],fldval["leads-RV_HVimp"])
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
	fields[2] := ["Mode  ","V. Pacing","AdaptivCRT","V-V Pace Delay"
				, "Lower\s+Rate","Upper\s+Track","Upper\s+Sensor"
				, "Paced AV","Sensed AV","Mode Switch"]
	labels[2] := ["Mode","CRT_VP","CRT_VV","CRT_VV","LRL","URL","USR","PAV","SAV","Mode Switch"]
	scanParams(qltbl,2,"par",1)
	
	qltbl := stregX(inirep "<<<","Detection",1,0,"<<<",1)
	fields[2] := ["Rates-AT/AF","Rates-VF","Rates-FVT","Rates-VT"
				, "Therapies-AT/AF","Therapies-VF","Therapies-FVT","Therapies-VT"]
	labels[2] := ["ATAF","VF","FVT","VT"
				, "Rx_ATAF","Rx_VF","Rx_FVT","Rx_VT"]
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
		if (thrVal := fldval["tmp-Amp"] strQ(fldval["tmp-PW"]," / ###")) {
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
	fldval["dev-Physician"] := instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") 
		? tmp 
		: "Dr. " trim(tmp," `n")
	
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
	fintbl := RegExReplace(fintbl,"\s+\(.*?based on.*?\)")
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
	labels[2] := ["ATAF","VF","FVT","VT"
				, "Rx_ATAF","Rx_VF","Rx_FVT","Rx_VT"]
	scanParams(parseTable(fintbl),2,"detect",1)
	
/*	FINAL: PARAMETERS
	- Modes, timing values
	- Programmed thresholds and outputs
*/
	fintxt := stregX(maintxt,"Final: Parameters",1,0,"Medtronic, Inc.",0)
	
	param := RegExReplace(stregx(fintxt,"Pacing Summary.",1,1,"Pacing Details",1),"Mode","----",,1)				; Replace the title "Mode" to prevent interference with param scan
	fields[1] := ["Mode  ","Lower","Upper Track","Upper Sensor"
		,"V. Pacing","AdaptivCRT","V-V Pace Delay"
		,"Paced AV","Sensed AV","Mode Switch"]
	labels[1] := ["Mode","LRL","URL","USR","CRT_VP","CRT_VV","CRT_VV","PAV","SAV","Mode Switch"]				; Scan for "Mode Switch" first, so can find plain "Mode" second
	tmp := onecol(param)
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
			,fldval["leads-RV_Sthr"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"],fldval["leads-RV_HVimp"])
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
			fldval["dev-Physician"] := instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") 
				? tmp 
				: "Dr. " trim(tmp," `n")
			
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
			iniTbl := RegExReplace(iniTbl,"(\w)  (\w)","$1 $2")
			fields[1] := ["Mode   ","Mode Switch","Detection Rate"
						, "Lower Rate","Upper Tracking Rate","Upper Sensor Rate"
						, "Search AV\+","Paced AV","Sensed AV"]
			labels[1] := ["Mode","ModeSwitch","ModeSwitchRate"
						, "LRL","URL","USR"
						, "SearchAV","PAV","SAV"]
			scanParams(iniTbl,1,"par",1)
			
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
			fldval["dev-Physician"] := instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") 
				? tmp 
				: "Dr. " trim(tmp," `n")
			
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
			fields[1] := ["Mode   ","Lower Rate","Upper Tracking Rate","Upper Sensor Rate","ADL Rate","Paced AV","Sensed AV"]
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
	if instr(maintxt,"800.CARDIAC") {
		eventlog("Boston Scientific Emblem identified.")
		gosub SICD
	} else {
		gosub bsciZoomView
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
		, strQ(readBnk("PatientLeadAManufacturer"),"###") 
		. strQ(readBnk("PatientLeadAModelNum"), " ###") 
		. strQ(readBnk("PatientLeadASerialNum"), " (serial ###)"))
	fldfill("Alead-Pol_pace",readBnk("PatientLeadAPolarity"))
	
	fldfill("dev-RVlead"
		, strQ(readBnk("PatientLeadV1Manufacturer"),"###") 
		. strQ(readBnk("PatientLeadV1ModelNum"), " ###") 
		. strQ(readBnk("PatientLeadV1SerialNum"), " (serial ###)"))
	fldfill("RVlead-Pol_pace",readBnk("PatientLeadV1Polarity"))
	
	fldfill("dev-LVlead"
		, strQ(readBnk("PatientLeadV2Manufacturer"),"###") 
		. strQ(readBnk("PatientLeadV2ModelNum"), " ###") 
		. strQ(readBnk("PatientLeadV2SerialNum"), " (serial ###)"))
	fldfill("LVlead-Pol_pace",readBnk("PatientLeadV2Polarity"))
	
	ctr := stregX(maintxt,"(Ventricular )?Tachy Counters",1,0,"$",0)
	ctrT := stregX(ctr,"(Ventricular )?Episode Counters",1,0,"Brady Counters",1)
	fields[1] := ["Total Episodes","Nonsustained Episodes","ATP Delivered","Shocks Delivered","Shocks Diverted","SVT Episodes.*"]
	labels[1] := ["VHR","VTNS","V_Paced","V_Shocked","V_Aborted","AHR"]
	scanParams(ctrT,1,"event",1)

	ctrB := stregX(ctr,"Brady Counters",1,0,"Page \d+ of",0)
	if (ctr~="(A Paced)|(V Paced)") {
		fields[1] := ["Since Last Reset-% A Paced","Since Last Reset-% V Paced"]
		labels[1] := ["AP","VP"]
	} else {
		fields[1] := ["% Paced"]
		labels[1] := [substr(fldval["par-Mode"],1,1) "P"]
	}
	scanParams(parseTable(ctrB),1,"dev",1)
	
	normLead("RA"
			,fldval["dev-Alead"],fldval["dev-Alead_impl"]
			,fldval["Alead-imp"],fldval["Alead-cap"],fldval["leads-AP_thr"],fldval["Alead-Pol_pace"]
			,fldval["Alead-sensing"],fldval["leads-AS_thr"],fldval["leads-RA_Pol_sens"])
	normLead("RV"
			,fldval["dev-RVlead"],fldval["dev-RVlead_impl"]
			,fldval["Vlead-imp"],fldval["Vlead-cap"],fldval["leads-VP_thr"],fldval["RVlead-Pol_pace"]
			,fldval["Vlead-sensing"],fldval["leads-VS_thr"],fldval["leads-RV_Pol_sens"],fldval["leads-RV_HVimp"])
	normLead("LV"
			,fldval["dev-LVlead"],fldval["dev-LVlead_impl"]
			,fldval["leads-LV_imp"],fldval["leads-LV_cap"],fldval["leads-LV_output"],fldval["LVlead-Pol_pace"]
			,fldval["leads-LV_Sensitivity"],fldval["leads-LV_Sthr"],fldval["leads-LV_Pol_sens"])

	return
}

SICD:
{
	txt := onecol(stregX(maintxt,"",1,0,"Programmable\s+Parameters",1))
	txt := RegExReplace(txt,": ",":  ")
	fields[1] := ["Patient Name","^Follow-up Date","^Last Follow-up Date","Implant Date"
				, "Device Model#","Device Serial#","Electrode Model#","Electrode Serial#"]
	labels[1] := ["Name","Encounter","Last_ck","IPG_impl"
				, "IPG","IPG_SN","HV","HV_SN"]
	scanparams(txt,1,"dev",1)
	fldfill("dev-IPG","Boston Scientific " RegExReplace(fldval["dev-IPG"],"EMBLEMTM","Emblem(TM)"))
	fldfill("dev-HVlead"
		, strQ(fldval["dev-HV"],"Boston Scientific ###")
		. strQ(fldval["dev-HV_SN"]," (serial ###)"))
	
	txt := onecol(stregX(maintxt,"Programmable\s+Parameters.*?\R",1,1,"Shock Polarity.*?\R",0))
	txt := RegExReplace(txt,": ",":  ")
	txt1 := stregX(txt,"Current Device Settings",1,0,"Initial Device Settings",1)
	txt2 := stregX(txt,"Initial Device Settings",1,0,">>>end",1)
	fields[1] := ["^Shock Zone","^Conditional Shock Zone"]
	labels[1] := ["VF","VT"]
	scanparams(txt1,1,"tachy",1)
	fields[1] := ["^Post Shock Pacing","^Gain Setting","^Sensing Configuration"]
	labels[1] := ["Mode","Gain","Pol_Sens"]
	scanparams(txt1,1,"par",1)
	fldval["par-Mode"] := strQ(fldval["par-Mode"],"Post-shock pacing ###")
	
	txt := onecol(stregx(maintxt,"Episode\s+Summary.*?\R",1,1,"1.800.CARDIAC",1))
	txt := stregx(txt,"",1,0,"Americas",0)
	txt := RegExReplace(txt,": ",":  ")
	fields[1] := ["Remaining Battery Life to ERI"]
	labels[1] := ["Battery_stat"]
	scanparams(txt,1,"dev",1)
	fields[1] := ["^Untreated Episodes","^Treated Episodes"]
	labels[1] := ["V_Aborted","V_Shocked"]
	scanparams(txt,1,"event",1)
	
	normLead("HV"
			,fldval["dev-HVlead"],fldval["dev-HVlead_impl"] 
			,"","","",""
			,"",fldval["par-Gain"],fldval["par-Pol_Sens"])
	
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
		eventlog("Old SJM report.")
		gosub SJM_old
	} else {
		eventlog("Newer SJM report.")
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
	fldfill("dev-IPG","SJM " fldval["dev-IPG"] strQ(fldval["dev-IPG_model"], " ###"))
	fldfill("dev-Encounter", parseDate(fldval["dev-Encounter"]).MDY)
	fldfill("dev-IPG_impl",parseDate(fldval["dev-IPG_impl"]).MDY)
	
	fields[1] := ["Lead Chamber","Lead Type"
				, ".. Pulse Amplitude",".. Pulse Width","Lead Impedance","P/R Sensitivity",
				, "Vario Capture Threshold","Test Pulse Width","P/R Signal"]
	labels[1] := ["Chamber","Type"
				, "Pace_Amp","Pace_PW","Imped","Sensitivity"
				, "Thr_Amp","Thr_PW","Thr_Sens"]
	sjmVals(1,"leads")
	
	fields[1] := ["(\x1C)Mode(\x1C)","Base Rate","Max Sensor Rate"]
	labels[1] := ["Mode","LRL","USR"]
	sjmVals(1,"par")
	
	tmp := xl.selectSingleNode("//id[@ser='" pat_ser "']/lead")
	fldval["dev-lead"] := fldval["dev-RVlead"] := tmp.text
	fldval["dev-leadimpl"] := fldval["dev-RVlead_impl"] := tmp.getAttribute("date")
	
	normLead("R" (InStr(fldval["leads-Chamber"],"V")?"V":"A")
		,fldval["dev-RVlead"],fldval["dev-RVlead_impl"],fldval["leads-Imped"]
		,strQ(fldval["leads-Thr_Amp"],"###" strQ(fldval["leads-Thr_PW"]," @ ###"))
		,strQ(fldval["leads-Pace_Amp"],"###" strQ(fldval["leads-Pace_PW"]," @ ###"))
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
	fldfill("dev-IPG","SJM " fldval["dev-IPG"] strQ(fldval["dev-IPG_model"], " ###"))
	fldfill("dev-Alead",fldval["dev-Alead_man"] 
		. strQ(fldval["dev-Alead_model"], " ###") strQ(fldval["dev-Alead_SN"], ", serial ###"))
	fldfill("dev-RVlead",fldval["dev-RVlead_man"] 
		. strQ(fldval["dev-RVlead_model"], " ###") strQ(fldval["dev-RVlead_SN"], ", serial ###"))
	fldfill("dev-LVlead",fldval["dev-LVlead_man"] 
		. strQ(fldval["dev-LVlead_model"], " ###") strQ(fldval["dev-LVlead_SN"], ", serial ###"))
	
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
				, "Paced AV Delay","Sensed AV Delay","Ventricular Pacing Chamber","Interventricular Pace Delay"]
	labels[1] := ["Mode","LRL","URL","USR"
				, "PAV","SAV","CRT_VP","CRT_VV"]
	sjmVals(1,"par")
	
	fields[1] := ["(\x1C)VF Detection Interval","(\x1C)VT-1 Detection Interval"
				, "VT-1 Therapy 1 Type","VF Therapy 1 Type","VF Voltage 1"]
	labels[1] := ["VF","VT"
				, "Rx_VT","VF0","VF1"]
	sjmVals(1,"detect")
	fldfill("detect-VF",fldval["detect-VF"]?round(60000/RegExReplace(fldval["detect-VF"],"\D")):"")
	fldfill("detect-VT",fldval["detect-VT"]?round(60000/RegExReplace(fldval["detect-VT"],"\D")):"")
	fldfill("detect-Rx_VF",fldval["detect-VF0"] strQ(fldval["detect-VF1"],", ###"))
	
	fields[1] := ["AT/AF Episodes","VT/VF Episodes"]
	labels[1] := ["ATAF","VT"]
	sjmVals(1,"event")
	
	normLead("RA"
			,fldval["dev-Alead"],fldval["dev-Alead_impl"],fldval["leads-RA_imp"]
			,strQ(fldval["leads-RA_Thr_Amp"],"###" strQ(fldval["leads-RA_Thr_PW"]," @ ###"))
			,strQ(fldval["leads-RA_Pace_Amp"],"###" strQ(fldval["leads-RA_Pace_PW"]," @ ###"))
			,fldval["leads-RA_Pol_pace"]
			,fldval["leads-RA_Thr_Sens"],fldval["leads-RA_Sensitivity"],fldval["leads-RA_Pol_sens"])
	normLead("RV"
			,fldval["dev-RVlead"],fldval["dev-RVlead_impl"],fldval["leads-RV_imp"]
			,strQ(fldval["leads-RV_Thr_Amp"],"###" strQ(fldval["leads-RV_Thr_PW"]," @ ###"))
			,strQ(fldval["leads-RV_Pace_Amp"],"###" strQ(fldval["leads-RV_Pace_PW"]," @ ###"))
			,fldval["leads-RV_Pol_pace"]
			,fldval["leads-RV_Thr_Sens"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"],fldval["leads-RV_HVimp"])
	normLead("LV"
			,fldval["dev-LVlead"],fldval["dev-LVlead_impl"],fldval["leads-LV_imp"]
			,strQ(fldval["leads-LV_Thr_Amp"],"###" strQ(fldval["leads-LV_Thr_PW"]," @ ###"))
			,strQ(fldval["leads-LV_Pace_Amp"],"###" strQ(fldval["leads-LV_Pace_PW"]," @ ###"))
			,fldval["leads-LV_Pol_pace"]
			,fldval["leads-LV_Thr_Sens"],fldval["leads-LV_Sensitivity"],fldval["leads-LV_Pol_sens"])

return
}

PaceartXml:
{
	progress,,,Scanning...
	yp := new XML(fileIn)
	fldval["dev-type"] := yp.selectSingleNode("//ActiveDevices/PatientActiveDevice[Status='ACTIVE']/Device/Type").text
	
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
	fldfill("indication",strQ(fldval["dev-dx_code"],"### - ") fldval["dev-dx_desc"])
	
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
				, "/Battery/RemainingPercentage:Battery_percent"
				. ""]
	xmlFld("//Encounter",1,"dev")
	fldfill("dev-IPG",strQ(fldval["dev-manufacturer"],"###") strQ(fldval["dev-model"]," ###"))
	fldfill("dev-Encounter",parseDate(utcTime(fldval["dev-Encounter"])).MDY)
	
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
	fldfill("par-CRT_VP",strQ(fldval["par-CRT_VP"],fldval["par-CRT_VP"]~="LEFT" ? "LV>RV" : "RV>LV"))
	
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
	if (fldval["dev-model"]~="i)Emblem") {
		fields[1] := ["/Zone[Type='VENTRICULAR_FIBRILLATION']//Detection//Interval:VF_ms"
					, "/Zone[Type='VENTRICULAR_TACHYCARDIA']//Detection//Interval:VT_ms" ]
	}
	xmlFld("//Programming/Tachycardia",1,"detect")
	fldfill("detect-VF",strQ(fldval["detect-VF_ms"],round(60000/fldval["detect-VF_ms"])))
	fldfill("detect-VT",strQ(fldval["detect-VT_ms"],round(60000/fldval["detect-VT_ms"])))
	
	fields[1] := ["/Episode[Type='AF_AT']/Count:ATAF"
				, "/Episode[Type='VF_VT']/Count:VTVF"
				, "/Episode[Type='SVT']/Count:SVT"
				, "/Episode[Type='V_NST']/Count:VNST"
				, "/Episode[Type='VT']/Count:VT"
				, "/Episode[Type='FVT']/Count:FVT"
				, "/Therapy[Chamber='RIGHT_ATRIUM']/ATP/Delivered:Rx_ATAF"
				, "/Therapy[Chamber='RIGHT_ATRIUM']/Shocks/Delivered:A_Shocked"
				, "/Therapy[Chamber='RIGHT_ATRIUM']/Shocks/Aborted:A_Aborted"
				, "/Therapy[Chamber='RIGHT_VENTRICLE']/ATP/Delivered:Rx_VATP"
				, "/Therapy[Chamber='RIGHT_VENTRICLE']/Shocks/Delivered:V_Shocked"
				, "/Therapy[Chamber='RIGHT_VENTRICLE']/Shocks/Aborted:V_Aborted"
				. ""]
	xmlFld("//Statistics/Detections_Therapies",1,"event")
	
	loop, % (i:=yp.selectNodes("//PatientPassiveDevice[Status='ACTIVE']")).length
	{
		k := readXmlLead(i.item(A_Index-1))
		normLead(k.ch
			, strQ(k.manu, "### ") strQ(k.model, "###") strQ(k.ser,", serial ###"), k.impl
			, k.pacing_imped
			, k.cap_amp strQ(k.cap_pw," / ###")
			, k.pacing_amp strQ(k.pacing_pw," / ###")
			, k.pacing_pol
			, k.sensing_thr
			, k.sensitivity_amp
			, k.sensitivity_pol
			, k.HV_imped)
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
	if (fldval["dev-model"]~="i)Emblem") {
		res.chamb := "HV"
		res.ch := "HV"
		fldval["leads-" res.ch "_HVimp"] := strQ(readNodeVal("//Statistics//HighPowerChannel//Impedance//Value"),"### ohms")
	}
	if (k.selectSingleNode("Device/Comments").text~="HV") {
		if !(res.chamb) {
			res.chamb := "HV"
			res.ch := "HV"
		}
		fldval["leads-" res.ch "_HVimp"] := strQ(readNodeVal("//Statistics//HighPowerChannel//Impedance//Value"),"### ohms")
	}
	
	base := "//Programming//PacingData[Chamber='" res.chamb "']"
	res.pacing_pol := strQ(readNodeVal(base "//Polarity"),"###")
	res.pacing_vector := strQ(readNodeVal(base "//PathwaysSummary"),"###")
	res.pacing_amp := strQ(readNodeVal(base "/Amplitude"),"### V")
	res.pacing_pw := strQ(readNodeVal(base "/PulseWidth"),"### ms")
	res.pacing_adaptive := strQ(readNodeVal(base "/AdaptationMode"),"###")
	
	base := "//Programming//SensingData[Chamber='" res.chamb "']"
	res.sensitivity_pol := strQ(readNodeVal(base "//Polarity"),"###")
	res.sensitivity_amp := strQ(readNodeVal(base "//Amplitude"),"### mV")
	
	base := "//Statistics//Lead[Chamber='" res.chamb "']"
	pathway := "[PolarityConfiguration/PathwaysSummary='" res.pacing_vector "']"
	res.cap_amp := strQ(readNodeVal(base "//CaptureCollection" pathway "//Capture//Amplitude"),"### V") 
	res.cap_pw := strQ(readNodeVal(base "//CaptureCollection " pathway "//Capture//Duration"),"### ms") 
	res.sensing_thr := strQ(readNodeVal(base "//SensitivityCollection" pathway "//Sensitivity//Amplitude"),"### mV") 
	res.pacing_imped := strQ(readNodeVal(base "//ImpedanceCollection" pathway "//Impedance//Value"),"### ohms")
	res.HV_imped := strQ(readNodeVal("//Statistics//HighPowerChannel//Impedance//Value"),"### ohms")
	
	return res
}

xmlFld(base,blk,pre="") {
/*	Reads xxxxxx:yyyy from array fields[blk]
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
		
		fldval[pre "-" lbl] := strQ(res, "###" . strQ(unit," ###"))
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
	if (val=="N/R") {
		val := ""
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
		if (val<1) {
			continue
		}
		
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
		. strQ(RegExReplace(el4,"[^[:ascii:]]")," ###") 
		. strQ(RegExReplace(el5,"[^[:ascii:]]")," ###")						; return: value ( units)( whatever el5 is)
}

pmPrint:
{
	if !(enc_MD) {
		return
	}
	rtfBody := "\b\ul DEVICE INFORMATION AND INITIAL SETTINGS\ul0\b0\par "
	. fldval["dev-IPG"] ", serial number " fldval["dev-IPG_SN"] 
	. strQ(fldval["dev-IPG_impl"],", implanted ###") . strQ(fldval["dev-Physician"]," by ###") ". "
	. strQ(fldval["dev-IPG_voltage"],"Generator cell voltage ###. ")
	. strQ(fldval["dev-Battery_stat"],"Battery status is ###. ") . strQ(fldval["dev-IPG_Longevity"],"Remaining longevity ###. ")
	. strQ(fldval["dev-Battery_percent"],"Battery percentage remaining ###%. ")
	. strQ(fldval["par-Mode"],"Brady programming mode is ###")
	. strQ(fldval["par-LRL"],", lower rate ###")
	. strQ(fldval["par-URL"],", upper tracking rate ###")
	. strQ((substr(fldval["par-Mode"],0,1)="R"),strQ(fldval["par-USR"],", upper sensor rate ###"))
	. strQ(fldval["par-ADL"],", ADL rate ###") . ". "
	. strQ(fldval["par-Cap_Mgt"],"Adaptive mode is ###. ")
	. strQ(fldval["par-PAV"],"Paced and sensed AV delays are " fldval["par-PAV"] " and " fldval["par-SAV"] ", respectively. ")
	. strQ(fldval["par-CRT_VP"],"Ventricular pace sequence ###" . strQ(fldval["par-CRT_VV"]," (###)") ". ")
	. strQ(fldval["dev-Sensed"],"Sensed ###. ") . strQ(fldval["dev-Paced"],"Paced ###. ")
	. strQ(fldval["dev-AsVs"],"AS-VS ###  ") . strQ(fldval["dev-AsVp"],"AS-VP ###  ")
	. strQ(fldval["dev-ApVs"],"AP-VS ###  ") . strQ(fldval["dev-ApVp"],"AP-VP ###  ")
	. strQ(fldval["dev-AP"],"A-paced ###%. ") . strQ(fldval["dev-VP"],"V-paced ###%. ")
	. strQ(fldval["detect-ATAF"],"AT/AF detection ###" strQ(fldval["detect-Rx_ATAF"],", Rx ###") ". ")
	. strQ(fldval["detect-VF"],"VF detection ###" strQ(fldval["detect-Rx_VF"],", Rx ###") ". ")
	. strQ(fldval["detect-FVT"],"FVT detection ###" strQ(fldval["detect-Rx_FVT"],", Rx ###") ". ")
	. strQ(fldval["detect-VT"],"VT detection ###" strQ(fldval["detect-Rx_VT"],", Rx ###") ". ") 
	. "\par\par "
	. "\b\ul LEAD INFORMATION\ul0\b0\par "
	
	for k,v in ["RA","RV","RV2","RV3","LV","LV2","HV"]
	{
		if !isobject(leads[v]) {
			continue
		}
		printLead(v)
	}
	
	printEvents()
	
	gosub PrintOut

Return
}

strQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	return (var1="") ? null : RegExReplace(txt,"###",var1)
}

normLead(lead				; RA, RV, LV 												) {
		,model				; Model name/ser
		,date				; Date implanted
		,P_imp				; Pacing impedance
		,P_thr				; Pacing capture threshold
		,P_out				; Pacing programmed output
		,P_pol				; Pacing polarity
		,S_thr				; Sensing threshold
		,S_sens				; Sensing programmed sensitivity
		,S_pol				; Sensing polarity
		,HV_imp="")			; HV impedance
{
	if (!P_imp && !P_thr && !P_out && !P_pol && !S_thr && !S_sens && !S_pol && !HV_imp) {	; ALL parameters in pre or post are NULL
		eventlog("Lead " lead " all null values!")
		return error																	; Do not populate leads[]
	}
	global leads, fldval
	leads[lead,"model"] 	:= model
	leads[lead,"date"]		:= date
	leads[lead,"imp"]  		:= strQ(P_imp,"Pacing impedance ###") 
							. strQ(HV_imp,strQ(P_imp,". ") "Defib impedance ###")
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
	rtfBody .= "\b " lead " lead:\b0  " 
	. strQ(leads[lead,"model"],"###" strQ(leads[lead,"date"],", implanted ###") ". ")
	. strQ(leads[lead,"imp"],"###. ")
	. strQ(leads[lead,"cap"],"Capture threshold ###. ")
	. strQ(leads[lead,"output"],"Pacing output ###. ")
	. strQ(leads[lead,"pace pol"],"Pacing polarity ###. ")
	. strQ(leads[lead,"sens"],((lead="RA")?"P":"")((lead="RV")?"R":"") "-wave sensing " 
		. ((leads[lead,"sens"]~="N/R")?"not measured/detected":"###") ". ")
	. strQ(leads[lead,"sensitivity"],"Sensitivity ###. ")
	. strQ(leads[lead,"sens pol"],"Sensing polarity ###. ")
	. "\par "
}

printEvents()
{
	global rtfBody, fldval
	if (fldval["leads-RV_HVimp"]) {
		txt := ""
		. strQ(fldval["event-AHR"]?fldval["event-AHR"]:"0","There were ### Atrial High Rate episodes. ")
		. strQ(fldval["event-VHR"]?fldval["event-VHR"]:"0","There were ### Ventricular High Rate episodes. ")
		. strQ(fldval["event-VF"]?fldval["event-VF"]:"0","### VF episodes detected. ")
		. strQ(fldval["event-VT"]?fldval["event-VT"]:"0","### VT episodes detected. ")
	}
	txt .= ""
	. strQ(fldval["event-VTNS"]?fldval["event-VTNS"]:"","### NS-VT episodes detected. ")
	. strQ(fldval["event-ATAF"]?fldval["event-ATAF"]:"","### AT/AF episodes detected. ")
	. strQ(fldval["event-V_Paced"]?fldval["event-V_Paced"]:"","### VT episodes pace-terminated. ")
	. strQ(fldval["event-V_Shocked"]?fldval["event-V_Shocked"]:"","### VT/VF episodes shock-terminated. ")
	. strQ(fldval["event-V_Aborted"]?fldval["event-V_Aborted"]:"","### VT/VF episodes aborted. ")
	. strQ(fldval["event-A_Paced"]?fldval["event-A_Paced"]:"","### AT episodes pace-terminated. ")
	. strQ(fldval["event-A_Shocked"]?fldval["event-A_Shocked"]:"","### AT/AF episodes shock-terminated. ")
	. strQ(fldval["event-A_Aborted"]?fldval["event-A_Aborted"]:"","### AT/AF episodes aborted. ")
	. strQ(fldval["event-Obs"],"\par ")
	
	rtfBody .= strQ(txt,"\par\b\ul EVENTS\ul0\b0\par ###\par ") 

	if (fldval["dev-type"]="MONITOR") {
		printMonitor()
	}
return	
}

printMonitor()
{
/*	ILR reports do not have all of the same elements as other devices
	Replaces rtfBody with barebones report
*/
	global rtfBody, fldval, yp

	rtfBody := "\b\ul DEVICE INFORMATION\ul0\b0\par "
	. fldval["dev-IPG"] ", serial number " fldval["dev-IPG_SN"] 
	. strQ(fldval["dev-IPG_impl"],", implanted ###") . strQ(fldval["dev-Physician"]," by ###") ". \par\par"
	. "\b\ul EVENTS\ul0\b0\par "

	epstr := "//InterrogatedDeviceData/Episodes//Episode"
	loop % (eps:=yp.selectNodes(epstr)).Length
	{
		ep := eps.item(A_Index-1)
		epId := ep.selectSingleNode("Id").Text
		epType := ep.selectSingleNode("Type").getAttribute("nonconformingData") 
				. ep.selectSingleNode("Type").Text
		epDate := utcTime(ep.selectSingleNode("Start").Text)
		epDur := ep.selectSingleNode("Duration").Text
		epAvgRate := ep.selectSingleNode("AverageVentricularRate").Text
		epMaxRate := ep.selectSingleNode("MaximumVentricularRate").Text

		rtfBody .= "Episode #" epId ": " ParseDate(epDate).DT ", Type """ epType """, "
			. "Duration (HH:MM:SS) " calcDuration(epDur).HMS " \par "
			. "*** \par\par "
	}	
	Return
}

PrintOut:
{
	summ := strQ(fldval["dev-summary"],"###"
			, "This represents a normal " strQ(format("{:L}",fldval["dev-EncType"]),"### ") "device check. The patient denies any device related symptoms. "
			. "The battery status is normal. Sensing and capture thresholds are good. The lead impedances are normal. "
			. "Routine follow up per implantable device protocol. ")
	if (fldval["dev-type"]="MONITOR") {
		summ := strQ(fldval["dev-summary"],"###","***")
	}
	
	; rtfHdr := "{\rtf1{\fonttbl{\f0\fnil Segoe UI;}}"
	rtfHdr := "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil Segoe UI;}}\viewkind4\uc1"
	
	rtfFtr := "}"
	
	rtfBody := "\pard\f0\fs22"
			. "\b\ul ANALYSIS DATE:\ul0\b0  " enc_dt.MDY "\par\par "
			. strQ(is_remote
				, "\b\ul TRANSMISSION DATE:\ul0\b0  " enc_trans.MDY "\par\par ")
			. "\b\ul ENCOUNTER TYPE\ul0\b0\par "
			. "Device interrogation " enc_type "\par\par "
			. strQ(fldval["indication"]
				, "\b\ul INDICATION FOR DEVICE\ul0\b0\par "
				. "###\par\par ")
			. strQ(fldval["dependent"]
				, "\b\ul PACEMAKER DEPENDENT\ul0\b0\par "
				. "###\par\par ")
			. rtfBody "\par "
			. "\b\ul ENCOUNTER SUMMARY\ul0\b0\par "
			. summ "\par\par "
	
	rtfOut := rtfHdr . rtfBody . rtfFtr
	
	nm := fldval["dev-Name"]
	RegExMatch(fileIn,"\....$",ext)
	fileOut :=	enc_MD "-" encMRN " " 
			.	(instr(nm,",") ? strX(nm,"",1,0,",",1,1) : strX(nm," ",1,1,"",0)) " "
			.	"#" fldval["dev-IPG_SN"] " "
			.	enc_dt.YMD
	onbaseFile := "TRREAT_"
			. fldval["dev-ordernum"] "_" 
			. enc_dt.YMD "_" 
			. fldval["dev-NameL"] "_" 
			. fldval["dev-MRN"]
	
	FileDelete, % path.files "tmp\" fileOut ".rtf"										; delete and generate RTF fileOut.rtf
	FileAppend, % rtfOut, % path.files "tmp\" fileOut ".rtf"
	
	eventlog("Print output generated in " path.files "tmp")
	
	RunWait, % "WordPad.exe """ path.files "tmp\" fileOut ".rtf"""						; launch fileNam in WordPad
	MsgBox, 262180, , Report looks okay?
	IfMsgBox, Yes
	{
		eventlog("RTF, " ext " copied to " path.compl)
		if (pat_meta) {
			FileMove, % pat_meta, % path.compl fileOut ".meta", 1						; copy BNK to complete directory
			eventlog("META copied to " path.compl)
		}
		if (ext=".xml") {
			extractXmlPdfs(fileOut,onbaseFile)
			
			fileWQ := enc_dt.MDY "," 			 										; date processed and MA user
					. """" nm """" ","													; CIS name
					. """" encMRN """" ","												; CIS MRN
					. """" fldval["dev-Enc"] """"										; Acct Num
					. "`n"
			FileAppend, % fileWQ, % path.trreat "logs\trreatWQ.csv"						; Add to logs\fileWQ list
			FileCopy, % path.trreat "logs\trreatWQ.csv", % path.chip "trreatWQ-copy.csv", 1
			
			FileCopy, % fileIn, % path.paceart "done\"
		}
		FileRead, rtfOut, % path.files "tmp\" fileOut ".rtf"							; reload edited RTF
		FileMove, % path.files "tmp\" fileOut ".rtf", % path.report fileOut ".rtf", 1	; move RTF to the final directory
		FileCopy, % fileIn, % path.compl fileOut ext, 1									; copy PDF to complete directory
		fileDelete, % fileIn
		
		t_now := A_Now
		edID := "/root/done/id[@ed='" t_now "']"
		xl.addElement("id","/root/done",{date: enc_dt.YMD, ser:fldval["dev-IPG_SN"], ed:t_now, au:user})
			xl.addElement("wqid",edID,fldval["dev-wqid"])
			xl.addElement("name",edID,fldval["dev-Name"])
			xl.addElement("dev",edID,fldval["dev-IPG"])
			; xl.addElement("lead",edID,{date:fldval["dev-leadimpl"]},fldval["dev-lead"])
			; xl.addElement("status",edID,"Sent")
			; xl.addElement("paceart",edID,strQ(is_remote,"True"))
			; xl.addElement("ordernum",edID,fldval["dev-ordernum"])
			; xl.addElement("accession",edID,fldval["dev-accession"])
			; xl.addElement("file",edID,path.compl fileOut ext)
			; xl.addElement("meta",edID,(pat_meta) ? path.compl fileOut ".meta" : "")
			; xl.addElement("report",edID,path.compl fileOut ".rtf")
		eventlog("Record added to worklist.xml")
		
		l_wqid := fldval["dev-wqid"]
		fileNam := fileOut
		makeORU(l_wqid)
		hl7out.file := "TRREAT_ORU_" A_now
		FileAppend, % hl7out.msg, % path.report hl7out.file								; create ORU file in pending
		FileCopy, % path.report hl7out.file, % path.outbound							; copy ORU to outbound folder for Ensemble
		FileMove, % path.report hl7out.file, % path.compl fileNam ".hl7", 1				; move ORU to completed folder, renamed fileNam.hl7
		FileMove, % path.report fileNam ".rtf", % path.compl fileNam ".rtf", 1			; move RTF from pending to completed folder
		eventlog("ORU sent to outbound.")
		
		FileMove, % path.hl7in "*_" l_wqid "Z.hl7", % path.hl7in "done\*.*", 1
		removeNode("/root/orders/order[@id='" l_wqid "']")
		xl.transformXML()
		xl.save(worklist)
		
		eventlog("Worklist.xml updated.")
		
		if !(isDevt) {
			whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")							; initialize http request in object whr
			whr.Open("GET"																; set the http verb to GET file "change"
				, "https://depts.washington.edu/pedcards/change/direct.php?" 
					. "do=sign" 
					. "&to=" enc_MD
				, true)
			whr.Send()																	; SEND the command to the address
			eventlog("Notification email sent to " enc_MD)
			MsgBox, 64,, % "Email sent to " enc_MD
			;~ whr.WaitForResponse()	
			;~ err := whr.ResponseText													; the http response
		}
	}
	
	return
}

extractXmlPdfs(fileOut,onbaseFile) {
	global yp, path

	loop, % (att := yp.selectNodes("//Encounter//Attachment//FileData")).Length
	{
		suffix := "_" A_index ".pdf"
		fName := path.compl fileOut suffix
		nBytes := Base64Dec( att.item(A_Index-1).text, Bin )
		ed_File := FileOpen( fName, "w")
		ed_File.RawWrite(Bin, nBytes)
		ed_File.Close

		FileCopy, % fName, % path.onbase onbaseFile suffix, 1							; copy PDF from complete dir to OnBase dir
	}
	Return
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
	
	if (pre!="") {
		pre := pre "-"
	}
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
		fldfill(pre . lbl, m)
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

parseORM() {
/*	parse fldval values to values
	including aliases for both WQlist and readWQorder
*/
	global fldval
	
	switch fldval.PV1_PtClass
	{
		case "O": encType:="Outpatient"
		case "I": encType:="Inpatient"
		case "DS": encType:="Inpatient"
		default: encType:="Outpatient"
	}
	location := encType
	
	return {date:parseDate(fldval.PV1_DateTime).YMD
		, encDate:parseDate(fldval.PV1_DateTime).YMD
		, nameL:fldval.PID_NameL
		, nameF:fldval.PID_NameF
		, name:fldval.PID_NameL strQ(fldval.PID_NameF,", ###")
		, mrn:fldval.PID_PatMRN
		, sex:fldval.PID_sex
		, DOB:parseDate(fldval.PID_DOB).MDY
		, prov:strQ(fldval.ORC_ProvCode
			, fldval.ORC_ProvCode "^" fldval.ORC_ProvNameL "^" fldval.ORC_ProvNameF
			, fldval.OBR_ProviderCode "^" fldval.OBR_ProviderNameL "^" fldval.OBR_ProviderNameF)
		, type:encType
		, loc:location
		, accountnum:fldval.PID_AcctNum
		, encnum:fldval.PV1_VisitNum
		, order:fldval.ORC_ReqNum
		, accession:fldval.ORC_FillerNum
		, acct:location strQ(fldval.ORC_ReqNum,"_###") strQ(fldval.ORC_FillerNum,"-###")
		, UID:tobase(fldval.ORC_ReqNum RegExReplace(fldval.ORC_FillerNum,"[^0-9]"),36)
		, ind:strQ(fldval.OBR_ReasonCode,"###") strQ(fldval.OBR_ReasonText,"^###")
		, indication:strQ(fldval.OBR_ReasonCode,"###") strQ(fldval.OBR_ReasonText,"^###")
		, indicationCode:fldval.OBR_ReasonCode
		, ordertype:fldval.OBR_TestCode "^" fldval.OBR_TestName
		, orderCtrl:fldval.ORC_OrderCtrl
		, ctrlID:fldval.MSH_CtrlID}
}

FetchDem:
{
	if !(fldval["dev-MRN"]~="^\d{6,7}$") {				; Check MRN parsed from PDF
		fldval["dev-MRN"] := ""
	}
	y := new XML(path.chip "currlist.xml")
	yArch := new XML(path.chip "archlist.xml")
	SNstring := "/root/id[data/device[@SN='" fldval["dev-IPG_SN"] "']]"
	if IsObject(k := y.selectSingleNode(SNstring)) {							; Device SN found
		fldval["dev-MRN"] := k.getAttribute("mrn")								; set dev-MRN based on device SN
		fldfill("dev-NameL",k.selectSingleNode("demog/name_last").text)
		fldfill("dev-NameF",k.selectSingleNode("demog/name_first").text)
		fldfill("dev-Name",fldval["dev-NameL"] strQ(fldval["dev-NameF"],", ###"))
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
	global xl, path, fldval, worklist
	
	xl := new XML(worklist)																; refresh worklist
	if !IsObject(xl.selectSingleNode("/root/orders")) {
		xl.addElement("orders","/root")
	}

	progress,,Reading worklist,Scanning orders
	fcount:=ComObjCreate("Scripting.FileSystemObject").GetFolder(path.hl7in).Files.Count

	Loop, files, % path.hl7in "*"														; Scan incoming folder for new orders and add to Orders node
	{
		e0 := {}
		fileIn := A_LoopFileName
		Progress, % 100*A_Index/fcount, % fileIn
		if RegExMatch(fileIn,"_([a-zA-Z0-9]{4,})Z.hl7",i) {								; skip old files
			continue
		}
		processhl7(A_LoopFileFullPath)
		e0:=parseORM()
		e0.orderNode := "/root/orders/order[ordernum='" e0.order "']"
		if IsObject(k:=xl.selectSingleNode(e0.orderNode)) {								; ordernum node exists
			e0.nodeCtrlID := k.selectSingleNode("ctrlID").text
			if (e0.CtrlID < e0.nodeCtrlID) {											; order CtrlID is older than existing, somehow
				FileDelete, % path.hl7in fileIn
				eventlog("Order msg " fileIn " is outdated.")
				continue
			}
			if (e0.orderCtrl="CA") {													; CAncel an order
				FileDelete, % path.hl7in fileIn											; delete this order message
				FileDelete, % path.hl7in "*_" e0.UID "Z.hl7"							; and the previously processed hl7 file
				removeNode(e0.orderNode)												; and the accompanying node
				eventlog("Cancelled order " e0.order ".")
				continue
			}
			FileDelete, % path.hl7in "*_" e0.UID "Z.hl7"								; delete previously processed hl7 file
			removeNode(e0.orderNode)													; and the accompanying node
			eventlog("Cleared order " e0.order " node.")
		}
		if (e0.orderCtrl="XO") {														; change an order
			e0.orderNode := "/root/orders/order[accession='" e0.accession "']"
			k := xl.selectSingleNode(e0.orderNode)
			e0.nodeUID := k.getAttribute("id")
			FileDelete, % path.hl7in "*_" e0.nodeUID "Z.hl7"
			removeNode(e0.orderNode)
			eventlog("Removed node id " e0.nodeUID " for replacement.")
		}
		
		newID := "/root/orders/order[@id='" e0.UID "']"								; otherwise create a new node
		xl.addElement("order","/root/orders",{id:e0.UID})
		xl.addElement("ordernum",newID,e0.order)
		xl.addElement("accession",newID,e0.accession)
		xl.addElement("ctrlID",newID,e0.CtrlID)
		xl.addElement("accountnum",newID,e0.accountnum)
		xl.addElement("ordertype",newID,e0.ordertype)
		xl.addElement("encnum",newID,e0.encnum)
		xl.addElement("loctype",newID,e0.type)
		xl.addElement("loc",newID,e0.loc)
		xl.addElement("date",newID,e0.date)
		xl.addElement("name",newID,e0.name)
		xl.addElement("mrn",newID,e0.mrn)
		xl.addElement("dob",newID,parsedate(e0.dob).YMD)
		xl.addElement("sex",newID,e0.sex)
		xl.addElement("prov",newID,e0.prov)
		eventlog("Added order ID " e0.UID ".")
		
		fileOut := e0.MRN "_" 
			. fldval["PID_nameL"] "^" fldval["PID_nameF"] "_"
			. e0.date "_"
			. e0.uid "Z.hl7"
			
		FileMove, %A_LoopFileFullPath%													; rename ORM file
			, % path.hl7in fileOut
	}
	xl.transformXML()
	xl.save(worklist)
	progress, off

	return
}

matchOrder(full:="") {
	global fldval, xl, fetchQuit
	static selbox, selbut
	key := {}
	thresh := (full) ? 1.0 : 0.15														; % fuzz tolerated, "all" = everything
	
	fldName := Format("{:U}",fldval["dev-Name"])
	Loop, % (k:=xl.selectNodes("/root/orders/order")).length							; generate list of orders with fuzz levels
	{
		node := k.item(A_index-1)
		nodeName := node.selectSingleNode("name").text
		nodeMRN := node.selectSingleNode("mrn").text
		nodeID := node.getAttribute("id")
		nodeOrdernum := node.selectSingleNode("ordernum").text
		nodeAccession := node.selectSingleNode("accession").text
		nodeOrdertype := node.selectSingleNode("ordertype").text
		nodeLocation := node.selectSingleNode("loctype").text
		nodeDate := node.selectSingleNode("date").text
		fuzz := (full) ? "" : fuzzysearch(nodename,fldName)
		if (fuzz<thresh) {
			list .= fuzz "|" nodeName "|" nodeID "|" nodeMRN "|" nodeDate "|" nodeLocation "|" nodeOrdernum "|" nodeAccession "|" nodeOrdertype "`n"
		}
	}
	Sort, list, R																		; sort by fuzz level
	Loop, parse, list, `n
	{
		k := A_LoopField
		if (k="") {
			break
		}
		vals:=strsplit(k,"|")
		key[A_Index] := {name:vals[2]													; build array of key{name,id,etc}
						,id:vals[3]
						,mrn:vals[4]
						,date:vals[5]
						,location:vals[6]
						,ordernum:vals[7]
						,accession:vals[8]
						,ordertype:vals[9]}
		keylist .= key[A_index].name " [" parseDate(key[A_Index].date).mdy "] " 
				; . key[A_Index].ordernum 
				. regexreplace(key[A_index].ordertype,".*?- ")
				. "|"
	}
	
	if (keylist="") {
		MsgBox, 262160, Empty orders, No active orders available!`n`nBe sure to "check-in" the order to send to TRREAT.
		fetchQuit := true
		eventlog("No orders released to incoming.")
		return
	}
	
	Gui, dev:Destroy
	Gui, dev:Default
	Gui, -MinimizeBox
	Gui, Add, Text, +Wrap
		, % "Select the order that matches this patient:`n"
	Gui, Font, s12
	Gui, Add, ListBox																	; listbox and button
		, h100 w640 r6 vSelBox VScroll AltSubmit gMatchOrderSelect
		, % keylist
	Gui, Add, Button, h30 vSelBut gMatchOrderSubmit Disabled, Select order				; disabled by default
	Gui, Add, Button, h30 yp xp+120 gLoadAllOrders Disabled, View all orders
	Gui, Show, AutoSize, Active orders
	Gui, +AlwaysOnTop
	
	if (full) {
		GuiControl, dev:Disable, View all orders
	}
	
	winwaitclose, Active orders
	
	if !(selbox) {																		; no selection
		fetchQuit := true
		return
	}
	
	res := key[selbox]
	fuzz := fuzzysearch(res.name , fldval["dev-name"]) 
	if (fuzz > 0.20) {																	; possible bad match
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
	fldval["dev-location"] := res.location
	
	return 
	
	matchOrderSelect:
	{
		GuiControl, dev:Enable, Select order
		return
	}
	
	matchOrderSubmit:
	{
		Gui, dev:Submit
		return
	}

	loadAllOrders:
	{
		matchOrder("all")
		Return
	}
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
		,{	au:user
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
	
	orderString := "//orders/order[@id='" fldval["dev-wqid"] "']"
		xl.addElement("ordertype", orderString, matchEAP(enc_type))
		xl.addElement("reading", orderString, enc_MD)
		xl.addElement("dependent", orderString, fldval["dependent"])
		xl.addElement("model",	orderString, fldval["dev-IPG"])
		xl.addElement("ser",	orderString, fldval["dev-IPG_SN"])
		xl.addElement("mode",	orderString, fldval["par-Mode"])
		xl.addElement("LRL",	orderString, fldval["par-LRL"])
		xl.addElement("URL",	orderString, fldval["par-URL"])
		xl.addElement("SAV",	orderString, fldval["par-SAV"])
		xl.addElement("PAV",	orderString, fldval["par-PAV"])
		xl.addElement("PVARP",	orderString, fldval["par-PVARP"])
		xl.addElement("ApThr",	orderString, leads["RA","cap"])
		xl.addElement("AsThr",	orderString, leads["RA","sens"])
		xl.addElement("VpThr",	orderString, leads["RV","cap"])
		xl.addElement("VsThr",	orderString, leads["RV","sens"])
		xl.addElement("Ap",   	orderString, leads["RA","output"])
		xl.addElement("As",   	orderString, leads["RA","sensitivity"])
		xl.addElement("Vp",   	orderString, leads["RV","output"])
		xl.addElement("Vs",   	orderString, leads["RV","sensitivity"])
	xl.transformXML()
	xl.save(worklist)
	
	return
}

makeReport:
{
/*	Generate the elements of the report
	- Pull Chipotle data if it exists (dependent, indication, primary EP), make necessary data node
	- Validate data values
	- Generate OBR_4 string, store in <order> for makeORU
	- Device check performed by
	- Normal text insert
	- Save values to Chipotle and Orders
	- Final print output and file routing
*/
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
	fldval["primaryEP"] := y.selectSingleNode(MRNstring "/prov").getAttribute("EP")
	
	ciedQuery()
	if (fetchQuit) {
		eventlog("fetchQuit ciedQuery.")
		return
	}
	
	ciedCheck()
	if (fetchQuit) {
		eventlog("fetchQuit ciedCheck.")
		return
	}

	buildEncType()
	
	gosub saveChip
	
	gosub pmPrint
	
	return
}

ciedQuery() {
/*	For setting values related to this patient/device
	Values are saved in Chipotle currlist.xml
*/
	global fldval, leads, fetchQuit, docs
		, tmpLead, tmpLDate
	static DepY, DepN, DepX, Ind
		, DocGroup, tmpEP
	tmpLead := ""
	tmpLDate := ""
	
	tmpEP := []
	
	gui, cied:Destroy
	
	if (fldval["dev-IPG"]~="Microny") {
		tmpLead := fldval["dev-lead"]
		tmpLDate := fldval["dev-leadimpl"]
		gui, cied:Add, Text, , Pacing lead
		gui, cied:Add, Edit, w200 vtmpLead, % tmpLead
		gui, cied:Add, Text, , Lead implant date
		gui, cied:Add, Edit, w200 vtmpLDate, % tmpLDate
		gui, cied:Add, Text
	}
	if !(fldval["dev-type"]="MONITOR") {
		gui, cied:Add, Text, , Pacemaker dependent?
		gui, cied:Add, Radio, % "vDepY Checked" (fldval["dependent"]="Yes"), Yes
		gui, cied:Add, Radio, % "vDepN Checked" (fldval["dependent"]="No") , No
		gui, cied:Add, Radio, vDepX, Clear
		gui, cied:Add, Text
	}
	gui, cied:Add, Text, , Indication for device
	gui, cied:Add, Edit, r3 w200 vInd, % fldval["indication"]
	gui, cied:Add, Text
	gui, cied:Add, Text, , Primary EP
	gui, cied:Add, Radio, vDocGroup, Outside/None
	for key,val in docs
	{
		tmpName := val.abbrev
		tmpEP[A_Index+1]:=tmpName
		gui, cied:Add, Radio, % "Checked" (tmpName=fldval.primaryEP), % tmpName
	}
	gui, cied:Add, Text
	gui, cied:Add, Button, w100 h30 , OK
	
	gui, cied:Show, AutoSize, CIED Query
	
	WinWaitClose, CIED Query
	
	gui, cied:Destroy
	return

	ciedGuiEscape:
	ciedGuiClose:
	{
		fetchQuit := true
		gui, cied:Cancel
		return
	}
	
	ciedButtonOK:
	{
		gui, cied:Submit
		
		fldval["dependent"] := (depY) 
			? "Yes"
				: (depN)
			? "No"
				: ""
		
		fldval["indication"] := Ind
		
		fldval["primaryEP"] := checkEP(tmpEP[docGroup])
		
		if (tmpLead) {
			tmp:=instr(fldval["leads-chamber"],"V") ? "RV" : "RA"
			leads[tmp].model := fldval["dev-lead"] := fldval["dev-RVlead"] := tmpLead
			leads[tmp].date := fldval["dev-leadimpl"] := fldval["dev-RVlead_impl"] := tmpLDate
		}
		
		return
	}
}

ciedCheck() {
/*	For setting values related to the performance of this check
	To aid in determining procedure performed
*/
	global fldval, fetchQuit, enc_type, is_postop, is_remote, is_remoteAlert
	static PeriOp_Y, PeriOp_N, chk_peri
		, RemoteAlert_Y, RemoteAlert_N
		, ChkInt, ChkPrg
	tmpEP := []
	
	gui, cied2:Destroy
	
	if (fldval["dev-location"]="Inpatient") {											; Inpatient encounter, ask if peri-op check
		chk_peri := true
		gui, cied2:Font, w Bold Underline
		gui, cied2:Add, Text, , Is this a peri-procedural check?
		gui, cied2:Font, w Norm
		gui, cied2:Add, Radio, vPeriOp_Y gcied2click, Yes								; is_postop := true
		gui, cied2:Add, Radio, vPeriOp_N gcied2click, No								; is_postop := ""
		gui, cied2:Add, Text
	}
	if (fldval["dev-EncType"]="REMOTE") {												; Inpatient encounter, ask if peri-op check
		is_remote := true
		gui, cied2:Font, w Bold Underline
		gui, cied2:Add, Text, , Was this a scheduled check or remote alert?
		gui, cied2:Font, w Norm
		gui, cied2:Add, Radio, vRemoteAlert_N gcied2click, Scheduled					; is_remoteAlert := ""
		gui, cied2:Add, Radio, vRemoteAlert_Y gcied2click, Acute						; is_remoteAlert := true
		gui, cied2:Add, Text
	} else {
		gui, cied2:Font, w Bold Underline
		gui, cied2:Add, Text, w400, Did this check involve checking thresholds, changing settings, or any other device programming?
		gui, cied2:Font, w Norm
		gui, cied2:Add, Radio, vChkPrg gcied2click, Yes									; fldval["dev-CheckType"] := "Programming"
		gui, cied2:Add, Radio, vChkInt gcied2click, No									; fldval["dev-CheckType"] := "Interrogation"
		gui, cied2:Add, Text
	}
	
	gui, cied2:Add, Button, w100 h30 Disabled, OK
	
	gui, cied2:Show, AutoSize, Device Check Info
	
	WinWaitClose, Device Check Info
	
	gui, cied2:Destroy
	
	return
	
	cied2click:
	{
		gui, cied2:Submit, NoHide
		prg := (ChkPrg or ChkInt)
		peri := (PeriOp_Y or PeriOp_N)
		alert := (RemoteAlert_Y or RemoteAlert_N)
		
		if (chk_peri) {																	; PeriOp
			valid := (prg && peri)
		} else {
			valid := (prg)																; Inpt not PeriOp, all outpt
		}
		if (is_remote) {																; Remote
			valid := (alert)
		}
			
		if valid {
			GuiControl, cied2:Enable, OK
		}
		return
	}
	
	cied2GuiEscape:
	cied2GuiClose:
	{
		fetchQuit := true
		gui, cied2:Cancel
		return
	}
	
	cied2ButtonOK:
	{
		gui, cied2:Submit
		
		is_postop := (PeriOp_Y)
		is_remoteAlert := (RemoteAlert_Y)
		fldval["dev-CheckType"] := ChkPrg ? "PROGRAMMING" : "INTERROGATION"
		
		return
	}
}

checkEP(name) {
/*	Find responsible EP
	and/or assign to someone
*/
	global y, fldval, mrnString, enc_MD, docs
	yID := y.selectSingleNode(MRNstring)
	
	if (name!=fldval.PrimaryEP) {
		MsgBox, 262180, Change, % "Change primary EP `n"
			. "from '" fldval.PrimaryEP "'`n"
			. "to '" name "'?"
		IfMsgBox, Yes
		{
			yID.selectSingleNode("prov").setAttribute("EP", name)
			yID.selectSingleNode("prov").setAttribute("au", user)
			yID.selectSingleNode("prov").setAttribute("ed", A_Now)
			eventlog(name " set as primary EP.")
			eventlog(name " set as primary EP.","C")
			writeOut(MRNstring,"prov")
		} else {
			name := fldval.PrimaryEP
		}
	}
	
	for key,val in docs
	{
		
		doclist .= (name=val.abbrev ? "*" : "") format("{:U}",key) " - " val.abbrev "|"
	}
	tmp := cMsgBox("Assign report"
					, "Send report to:`n`n(primary EP is " name ").`n`n"
					. "Close [x] window to skip this step."
					, trim(doclist," |")
					, "Q","")
	
	if (tmp="Close") {
		enc_MD := ""
	} else {
		enc_MD := substr(tmp,1,2)
	}
	eventlog("Report assigned to " enc_MD ".")
	
	Return name
}

; Create enc_type based on office/remote/periop, device type 
buildEncType() {
/*	Determine certain statuses
	- is_remote = remote check
	- is_postop = postop check
	
	- enc_dt = date object of device check date
	- enc_trans = date object of transmit date from Paceart XML
	
	- enc_type = text of check type: OUTPATIENT, INPATIENT, POSTOP, REMOTE
	- enc_type = append device type: PM, ICD, ILR, Leadless(?)
	- enc_type = append lead type: Single, Dual, Multi
*/
	global fldval, leads
		, is_remote, is_remoteAlert, is_postop
		, enc_dt, enc_trans, enc_type
	dict := readIni("EpicOrderEAP")
	
	if (fldval["dev-location"]="Outpatient") {
		enc_type := "OUTPT "
		enc_dt := parseDate(fldval["dev-Encounter"])									; report date is day of encounter
		enc_trans :=																	; transmission date is null
	}
	if (is_remote) {
		enc_type := "REMOTE "
		enc_dt := parseDate(substr(A_now,1,8))											; report date is date run (today)
		enc_trans := parseDate(fldval["dev-Encounter"])									; transmission date is date sent
		enc_type .= (is_remoteAlert) ? "ALERT " : ""									; include if remote alert
		fldval["dev-CheckType"] := ""
	} 
	if (fldval["dev-location"]="Inpatient") {
		enc_type := "INPT "
		enc_dt := parseDate(fldval["dev-Encounter"])									; report date is day of encounter
		enc_trans :=																	; transmission date is null
	}
	if (is_postop) {
		enc_type := "PERI-PROCEDURE "
		enc_dt := parseDate(fldval["dev-Encounter"])									; report date is day of encounter
		enc_trans :=																	; transmission date is null
	}
	
	enc_type .= (instr(leads["RV","imp"],"Defib") || IsObject(leads["HV"]))
		? "ICD "
		: "PM "
/*	Need to add in other types here for leadless, ILR, and SICD
	Might need to insert these for Epic testing
*/
	if !(is_remote || is_postop) {
		if (IsObject(leads["RA"])) {
			leads.A := true
		}
		if (IsObject(leads["RV"]) || IsObject(leads["LV"])) {
			leads.V := true
		}
		if (IsObject(leads["RV"]) && IsObject(leads["LV"])) {
			leads.M := true
		}
		if (leads.M) {
			enc_type .= "BIV "
		} else
		if (leads.A && leads.V) {
			enc_type .= "DUAL "
		} else
		{
			enc_type .= "SINGLE "
		}
	} 
	
	if (fldval["dev-type"]="MONITOR") {
		enc_type := "REMOTE MONITOR "
		enc_dt := parseDate(substr(A_now,1,8))											; report date is date run (today)
		enc_trans := parseDate(fldval["dev-Encounter"])									; transmission date is date sent
	}

	enc_type .= fldval["dev-CheckType"]
	eventlog("enc_type builder: " enc_type)
	
	return
}

readWQ(idx) {
	global xl
	
	res := []
	k := xl.selectSingleNode("//orders/order[@id='" idx "']")
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

WriteOut(parentpath,node) {
/* 
	Prevents concurrent writing of y.MRN data. If someone is saving data (.currlock exists), script will wait
	approx 6 secs and check every 50 msec whether the lock file is removed. When available it creates clones the y.MRN
	node, loads a fresh currlist into Z (latest update), replaces the z.MRN node with the cloned y.MRN node,
	saves it, then reloads this currlist into Y.
*/
	global y, path
	filecheck()
	FileOpen(path.chip ".currlock", "W")										; Create lock file.
	
	locPath := y.selectSingleNode(parentpath)
	locNode := locPath.selectSingleNode(node)
	clone := locNode.cloneNode(true)											; make copy of y.node
	
	z := y																		; temp Z will be most recent good currlist
	
	if !IsObject(z.selectSingleNode(parentpath "/" node)) {
		If instr(node,"id[@mrn") {
			z.addElement("id","root",{mrn: strX(node,"='",1,2,"']",1,2)})
		} else {
			z.addElement(node,parentpath)
		}
	}
	zPath := z.selectSingleNode(parentpath)										; find same "node" in z
	zNode := zPath.selectSingleNode(node)
	zPath.replaceChild(clone,zNode)												; replace existing zNode with node clone
	
	z.save(path.chip "currlist.xml")											; write z into currlist
	eventlog(parentpath "/" node " saved.","C")
	eventlog("CHIPOTLE currlist updated.")
	y := z																		; make Y match Z, don't need a file op
	FileDelete, % path.chip ".currlock"											; release lock file.
	return
}

filecheck() {
	global path
	if FileExist(path.chip ".currlock") {
		err=0
		Progress, , Waiting to clear lock, File write queued...
		loop 50 {
			if (FileExist(path.chip ".currlock")) {
				progress, % p
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
	global user, path
	logdir := (ch="C") ? path.chip "logs\" : path.trreat "logs\"
	comp := A_ComputerName
	FormatTime, sessdate, A_Now, yyyyMM
	FormatTime, now, A_Now, yyyy.MM.dd||HH:mm:ss
	name := logdir . sessdate . ".log"
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

; Parse date strings in a variety of ways
ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-\._/]"
	date := []
	time := []
	x := RegExReplace(x,"[,\(\)]")
	
	if (x~="\d{4}.\d{2}.\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z") {
		x := RegExReplace(x,"[TZ]","|")
	}
	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",d) {			; 03-Jan-2015
		date.dd := zdigit(d1)
		date.mmm := d2
		date.mm := zdigit(objhasvalue(mo,d2))
		date.yyyy := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"i)\b(" moStr "|\d{1,2})" dSep "(\d{1,2})" dSep "(\d{4}|\d{2})",d) {	; Jan-03-2015, 01-03-2015
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
	else if RegExMatch(x,"i)(" moStr ")\s+(\d{1,2}),?\s+(\d{4})",d) {					; Dec 21, 2018
		date.mmm := d1
		date.mm := zdigit(objhasvalue(mo,d1))
		date.dd := zdigit(d2)
		date.yyyy := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"\b(\d{4})[\-\.](\d{2})[\-\.](\d{2})\b",d) {					; 2015-01-03
		date.yyyy := d1
		date.mm := d2
		date.mmm := mo[d2]
		date.dd := d3
		date.date := trim(d)
	}
	else if RegExMatch(x,"\b(19|20\d{2})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",d)  {	; 20150103174307 or 20150103
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
	
	if RegExMatch(x,"iO)(\d+):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",t) {				; 17:42 PM
		hasDays := (t.value[4]) ? true : false 											; 4 nums has days
		time.days := (hasDays) ? t.value[1] : ""
		time.hr := trim(t.value[1+hasDays])
		if (time.hr>23) {
			time.days := floor(time.hr/24)
			time.hr := mod(time.hr,24)
			DHM:=true
		}
		time.min := trim(t.value[2+hasDays]," :")
		time.sec := trim(t.value[3+hasDays]," :")
		time.ampm := trim(t.value[5])
		time.time := trim(t.value)
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, days:zdigit(time.days)
			, hr:zdigit(time.hr), min:zdigit(time.min), sec:zdigit(time.sec)
			, ampm:time.ampm, time:time.time
			, DHM:zdigit(time.days) ":" zdigit(time.hr) ":" zdigit(time.min) " (DD:HH:MM)" 
 			, DT:date.mm "/" date.dd "/" date.yyyy " at " zdigit(time.hr) ":" zdigit(time.min) ":" zdigit(time.sec) }
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

; Convert duration secs to DDHHMMSS
calcDuration(sec) {
	DD := divTime(sec,"D")
	HH := divTime(DD.rem,"H")
	MM := divTime(HH.rem,"M")
	SS := MM.rem

	return { DHM: zDigit(DD.val) ":" zDigit(HH.val) ":" zDigit(MM.val)
			, DHMS: zDigit(DD.val) ":" zDigit(HH.val) ":" zDigit(MM.val) ":" zDigit(SS) 
			, HMS: zDigit(HH.val) ":" zDigit(MM.val) ":" zDigit(SS) }
}

divTime(sec,div) {
	static T:={D:86400,H:3600,M:60,S:1}
	xx := Floor(sec/T[div])
	rem := sec-xx*T[div]
	Return {val:xx,rem:rem}
}

utcTime(x) {
/*	convert UTC time to local time

*/
	global utcDiff

	k := ParseDate(x)
	dt := k.ymd k.hr k.min k.sec
	dt += utcDiff, Hours

	Return dt
}

setUTC() {
/*	Get UTC offset
	Should compensate for DST setting on local machine
*/
	tdif := A_Now
	tdif -= A_NowUTC, Hours

	Return tdif
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
	IniRead,x,% trreatDir "files\trreat.ini",% section
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

parsedocs(ByRef list) {
/*	List = {XX:aaa^bbb^ccc^ddd, BB:eee^fff^ggg^hhh}
	map = [last,first,npi,cumg]
	
	returns list = {
					XX={last:aaa,first:bbb,npi:ccc,cumg:ddd}
					BB={last:eee,first:fff,npi:ggg,cumg:hhh}
					}
*/	
	map := ["nameL","nameF","NPI","CUMG"]
	
	for key,val in list
	{
		init := substr(key,1,2)
		ele:=StrSplit(val,"^")
		node := list[init] := {}
		for key2,val2 in map
		{
			node[val2] := ele[key2]
		}
		node.hl7 := node.NPI "^" node.nameL "^" node.nameF
		node.eml := node.nameF "." node.nameL
		node.abbrev := substr(node.nameF,1,1) ". " node.nameL
		node.user := key
		list.Delete(key)
	}
	return list
}

#Include strx.ahk
#Include xml.ahk
#Include CMsgBox.ahk
#Include sift3.ahk
#Include hl7.ahk
