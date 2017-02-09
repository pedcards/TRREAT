/*	TRREAT - The Rhythm Recording Electronic Analysis Transmogrifier - PM
	Converts file
		Drag-and-drop onto window
		Monitor folder for changes
	Inputs a text file
		Probably converted from PDF using XPDF's "PDFtoTEXT"
		Use the -layout or -table option
		Only need the first 1-2 pages
	Identifies type of report:
		PaceArt device check
		ZioPatch Holter
		LifeWatch (or other) Holter
		LifeWatch (or other) Event Recorder
	Extracts salient data
	Generates report using mail merge or template in Word
	Sends report to HIM
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%
IfInString, A_WorkingDir, AhkProjects					; Change enviroment if run from development vs production directory
{
	isAdmin := true
	holterDir := ".\Holter PDFs\"
	importFld := ".\Import\"
	chipDir := ".\Chipotle\"
} else {
	isAdmin := false
	holterDir := "\\chmc16\Cardio\EP\HoltER Database\Holter PDFs\"
	importFld := "\\chmc16\Cardio\EP\HoltER Database\Import\"
	chipDir := "\\childrens\files\HCChipotle\"
}
user := A_UserName

Gui, Add, Listview, w600 -Multi NoSortHdr Grid r12 hwndHLV, Filename|Name|Device|Report|Fix
Gui, Add, Button, Disabled w600 h50 , Reload
Gui, Show
newTxt := Object()
blk := Object()
blk2 := Object()
docs := Object()
docs := {"Chun, Terrence":"783118","Salerno, Jack":"343079","Seslar, Stephen":"358945"}

Loop, *.pdf
{ 
	blocks := Object()
	fields := Object()
	labels := Object()
	fldval := Object()
	leads := Object()
	summBl := summ := ""
	fileIn := A_LoopFileName
	SplitPath, fileIn,,,,fileOut
	RunWait, pdftotext.exe -table "%fileIn%" temp.txt , , hide
	FileRead, maintxt, temp.txt
	;~ RunWait, pdftotext.exe -raw -nopgbrk "%fileIn%" tempraw.txt , , hide
	;~ FileRead, mainraw, tempraw.txt
	cleanlines(maintxt)
	if (maintxt~="Medtronic,\s+Inc") {
		if (instr(maintxt,"Defibrillation")) {									; All ICD reports will have this text
			;~ MsgBox MDT icd
			gosub MDTpm
		}
		;~ if (instr(maintxt,"Pacemaker Model")) {
		else {																	; Can't find PM specific text, other than not being an ICD
			gosub MDTpm															; maybe use an array of Brady devices?
		}
	}
	if (maintxt~="Boston Scientific Corporation") {
		if (instr(maintxt,"Shock")) {
			MsgBox BSCI icd
		}
		else {
			MsgBox BSCI pm
		}
		ExitApp
	}
}

MsgBox Directory scan complete.
GuiControl, Enable, Reload

ExitApp


ButtonReload:
	Reload
Return

GuiClose:
ExitApp

MDTpm:
{
	fileNum += 1
	LV_Add("", fileIN)
	LV_Modify(fileNum,"col3","PM")
	Gui, Show
	
	if (maintxt~="Adapta") {
		gosub mdtAdapta
	}
	if (maintxt~="Quick Look II") {
		gosub mdtQuickLookII
	}
	
	gosub pmPrint
	;~ clipboard := rtfBody
	;~ MsgBox % rtfBody
	
return	
}

mdtQuickLookII:
{
	iniRep := strX(columns(maintxt,"Clinical Status","Medtronic, Inc",0,"Pacing \("),"Pacing",1,0)
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
				, "Name","ID","Physician","null"
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
	fldval["dev-Alead"] := RegExReplace(fldval["dev-Alead"],"---")
	fldval["dev-RVlead"] := RegExReplace(fldval["dev-RVlead"],"---")
	fldval["dev-LVlead"] := RegExReplace(fldval["dev-LVlead"],"---")
	
	fintbl := stregX(fintxt,"",n+1,0,"Parameter Summary",1)
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
	
	if (fldval["dev-Alead"]) {
		normLead("RA"
				,fldval["dev-Alead"],fldval["dev-Alead_impl"]
				,fldval["leads-A_imp"],fldval["leads-A_cap"],fldval["leads-A_output"],fldval["leads-A_Pol_pace"]
				,fldval["leads-A_Sthr"],fldval["leads-A_Sensitivity"],fldval["leads-A_Pol_sens"])
	}
	if (fldval["dev-RVlead"]) {
		normLead("RV"
				,fldval["dev-RVlead"],fldval["dev-RVlead_impl"]
				,fldval["leads-RV_imp"],fldval["leads-RV_cap"],fldval["leads-RV_output"],fldval["leads-RV_Pol_pace"]
				,fldval["leads-RV_Sthr"],fldval["leads-RV_Sensitivity"],fldval["leads-RV_Pol_sens"])
	}
	if (fldval["dev-LVlead"]) {
		normLead("LV"
				,fldval["dev-LVlead"],fldval["dev-LVlead_impl"]
				,fldval["leads-LV_imp"],fldval["leads-LV_cap"],fldval["leads-LV_output"],fldval["leads-LV_Pol_pace"]
				,fldval["leads-LV_Sthr"],fldval["leads-LV_Sensitivity"],fldval["leads-LV_Pol_sens"])
	}
return
}

mdtAdapta:
{
	iniRep := strX(columns(maintxt,"Clinical Status","Medtronic, Inc",0,"Pacing \("),"Pacing",1,0)
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
						, "Patient Name", "DOB", "ID", "Physician","`n"
						, "Pacemaker Model", "Implanted"
						, "Atrial Lead", "Implanted"
						, "Ventricular Lead", "Implanted"
						, "Pacemaker Status", "Estimated remaining longevity"
						, "Battery Status", "Voltage", "Current", "Impedance", "Lead Status"]
			labels[1] := ["IPG","IPG_SN","Encounter"
						,"Name", "DOB", "MRN", "Physician","null"
						, "IPG0", "IPG_impl"
						, "Alead", "Alead_impl"
						, "Vlead", "Vlead_impl"
						, "IPG_stat", "IPG_longevity"
						, "Battery_stat", "IPG_voltage", "Current", "Impedance", "null"]
			fieldvals(dev,1,"dev")
			if !instr(tmp := RegExReplace(fldval["dev-Physician"],"\s(-+)|(\d{3}.\d{3}.\d{4})"),"Dr.") {
				fldval["dev-Physician"] := "Dr. " . trim(tmp)
			}
			fldval["dev-Alead"] := RegExReplace(fldval["dev-Alead"],"---")
			fldval["dev-RVlead"] := RegExReplace(fldval["dev-RVlead"],"---")
			
			finleads := strX(fintxt,"Lead Status:",1,0,"Capture Management",1,21)
			fields[2] := ["Atrial lead-Output Energy","Atrial Lead-Measured Current"
						, "Atrial lead-Measured Impedance","Atrial Lead-Pace Polarity","endcolumn"
						, "Ventricular lead-Output Energy","Ventricular Lead-Measured Current"
						, "Ventricular lead-Measured Impedance","Ventricular Lead-Pace Polarity","endcolumn"]
			labels[2] := ["A_output","A_curr","A_imp","A_pol","null"
						, "V_output","V_curr","V_imp","V_pol","null"]
			fldval["leads-date"] := strX(finleads,"Lead Status: ",1,13,"`n",1,0,n)
			tbl := substr(finleads,n)
			scanParams(parseTable(tbl),2,"leads")
			
			thresh := onecol(stregX(fintxt,"Threshold Test Results.",1,1,"Medtronic Software",1))
			fldval["leads-AP_thr"] := parseStrDur(oneCol(stregx(thresh,"Atrial Pacing Threshold",1,1,"\n\n",0)))
			fldval["leads-VP_thr"] := parseStrDur(oneCol(stregx(thresh,"Ventricular Pacing Threshold",1,1,"\n\n",0)))
			fldval["leads-AS_thr"] := trim(stregx(thresh,"P-wave",1,1,"\n\n",0)," `r`n")
			fldval["leads-VS_thr"] := trim(stregx(thresh,"R-wave",1,1,"\n\n",0)," `r`n")
			
		}
		if (fintxt~=splTxt ".*Permanent Parameters") {
			perm := oneCol(strX(fintxt,"Permanent Parameters",1,0,"Medtronic Software",1,0))
			param := strx(perm,"Permanent Parameters",1,0,"Refractory/Blanking",1,0)
			fields[1] := ["Mode","Lower Rate","Upper Tracking Rate","Upper Sensor Rate","ADL Rate","Paced AV","Sensed AV"]
			labels[1] := ["Mode","LRL","URL","USR","ADL","PAV","SAV"]
			scanParams(fintxt,1,"par")
			
			param_A := stregX(perm,"Atrial Lead",1,0,"Ventricular Lead",1)
			fields[2] := ["Amplitude","Pulse Width","Sensitivity","Pace Polarity","Sense Polarity","Capture Management"]
			labels[2] := ["Amp","PW","Sens","Pol_pace","Pol_sens","Cap_Mgt"]
			scanParams(param_A,2,"Alead")
			
			param_V := stregX(perm,"Ventricular Lead",1,0,">>>end",1)
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
/*	Analyze text block for vertical table format
	Top row must be title row
	First column beginning row2 are parameters
*/
	nextpos := 1
	Loop																		; Iterate for each column found
	{
		Loop, parse, txt, `n,`r													; Read through text block
		{
			i := A_LoopField
			if !(i) {															; OK to skip entirely blank lines
				continue
			}
			
			if (A_Index=1) {
				pos := RegExMatch(i "  "										; Add "  " to end of scan string
								,"O)(?<=(\s{2}))[^\s](.*?)(?=(\s{2}))"			; Search "  text  " as each column 
								,col											; return result in var "col"
								,nextpos)										; search position at next column
								
				pre := col.value()												; result is column name
				
				if !(pos) {														; break if no more matches
					break
				}
				continue														; pre header is set, move to next row
			} 
			
			fld := strX(i,"",1,0,"  ",1,2)										; field name is first column
			
			str := strX(substr(i,pos),"",1,0,"  ",1)							; result is substr from pos of header column to "  "
			
			result .= pre "-" fld ":  " str "`n"								; concat results into a single column
		}
		
		if !(pos) {																; break when no more hits
			break
		}
		
		result .= "endcolumn`n"													; not sure if need "endcolumn" delimiter any more
		nextpos := pos+1														; start next search 1 space over from last result
	}
	;~ MsgBox % result
return result
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
		if !(val) {
			continue
		}
		
		RegExMatch(i															; Add "  " to end of scan string
				,"O)" colstr													; Search "  text  " as each column 
				,col1)															; return result in var "col1"
		RegExMatch(i
				,"O)" colstr
				,col2
				,col1.pos()+1)
		
		res :=	(col2.value()~="^(\>\s*)(?=[^\s])")
			?	RegExReplace(col2.value(),"^(\>\s*)(?=[^\s])") " (changed from " col1.value() ")"
			:	col1.value()
			
		fldval[pre "-" labels[blk,val]] := res
	}
	return
}

pmPrint:
{
	rtfBody := "\fs22\b\ul DEVICE INFORMATION\ul0\b0\par`n\fs18"
	. fldval["dev-IPG"] ", serial number " fldval["dev-IPG_SN"] 
	. printQ(fldval["dev-IPG_impl"],", implanted ###") . printQ(fldval["dev-Physician"]," by ###") ". `n"
	. printQ(fldval["dev-IPG_voltage"],"Generator cell voltage ###. ")
	. printQ(fldval["dev-Battery_stat"],"Battery status is ###. ") . printQ(fldval["dev-IPG_Longevity"],"Remaining longevity ###. ") "`n"
	. printQ(fldval["par-Mode"],"Brady programming mode is ### with lower rate " fldval["par-LRL"])
	. printQ(fldval["par-URL"],", upper tracking rate ###")
	. printQ(fldval["par-USR"],", upper sensor rate ###")
	. printQ(fldval["par-ADL"],", ADL rate ###") . ". `n"
	. printQ(fldval["par-Cap_Mgt"],"Adaptive mode is ###. `n")
	. printQ(fldval["par-PAV"]&&fldval["par-SAV"],"Paced and sensed AV delays are " fldval["par-PAV"] " and " fldval["par-SAV"] ", respectively. `n")
	. printQ(fldval["dev-Sensed"],"Sensed ###. ") . printQ(fldval["dev-Paced"],"Paced ###. ")
	. printQ(fldval["dev-AsVs"],"AS-VS ###  ") . printQ(fldval["dev-AsVp"],"AS-VP ###  ")
	. printQ(fldval["dev-ApVs"],"AP-VS ###  ") . printQ(fldval["dev-ApVp"],"AP-VP ###  ") . "\par`n"
	. "\fs22\par`n"
	. "\b\ul LEAD INFORMATION\ul0\b0\par`n\fs18 "
	
	for k in leads
	{
		printLead(k)
	}
	
	gosub PrintOut

Return
}

printQ(var1,txt) {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
*/
	if !(var1) {
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
	rtfBody .= "\b " lead " lead: \b0 " leads[lead,"model"] ", implanted " leads[lead,"date"] ". "
	. printQ(leads[lead,"imp"],"Pacing impedance ###. ")
	. printQ(leads[lead,"cap"],"Capture threshold ###. ")
	. printQ(leads[lead,"output"],"Pacing output ###. ")
	. printQ(leads[lead,"pace pol"],"Pacing polarity ###. ")
	. printQ(leads[lead,"sens"],((lead="RA")?"P":"")((lead="RV")?"R":"") "-wave sensing ###. ")
	. printQ(leads[lead,"sensitivity"],"Sensitivity ###. ")
	. printQ(leads[lead,"sens pol"],"Sensing polarity ###. ")
	. "\par`n"
}

PrintOut:
{
	;~ if (RegExMatch(summ,"\<\d*\>")) {
		;~ enc_FIN:=strX(summ,"<",1,1,">",1,1,nn)
		;~ summ := trim(substr(summ,nn+1))
	;~ } else {
		;~ InputBox , enc_FIN, % blk["Patient Name"] " - " enc_date, REQUIRED:`n`nEncounter number`n(8 digits)
	;~ }

	FormatTime, enc_dictdate, A_now, yyyy MM dd hh mm t
	
	rtfHdr := "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Arial;}}`n"
	. "{\*\generator Msftedit 5.41.21.2510;}\viewkind4\uc1\lang9\margl1440\margr1440\margt1440\margb1440`n"
	. "{\pard\f0\fs22`n"
	. "{\tx2160`n"
	. "Transcriptionist\tab "	"<TrID:crd> \par`n"
	. "Document Type\tab "		"<7:Q8> \par`n"
	. "Clinic Title code\tab "	"<1035:PACE> \par`n"
	. "Medical Record #\tab "	"<1:" blk["Patient ID"] ">\par`n"
	. "Patient Name\tab "		"<2:" blk["Patient Name"] ">\par`n"
	. "CIS Encounter #\tab "	"<3: " substr("0000" . enc_FIN, -11) " >\par`n"
	. "Dictating Phy #\tab "	"<8:" enc_MD ">\par`n"
	. "Dictation Date\tab "		"<13:" enc_signed ">\par`n"
	. "Job #\tab "				"<15:e> \par`n"
	. "Service Date\tab "		"<31:" enc_date ">\par`n"
	. "Surgery Date\tab "		"<6:" enc_date "> \par`n"
	. "Attending Phy #\tab "	"<9:" enc_MD "> \par`n"
	. "Transcription Date\tab "	"<TS:" enc_dictdate "> \par`n"
	. "<EndOfHeader>\par}`n"
	. "\par`n"

	rtfFtr := "`n\fs22\par\par`n\{SEC XCOPY\} \par`n\{END\} \par`n}`r`n"

	rtfBody := "\tx1620\tx5220\tx7040" 
	. "\fs22\b\ul PROCEDURE DATE\ul0\b0\par\fs18`n"
	. fldval["dev-Encounter"] "\par\par\fs22`n"
	. rtfBody . "\fs22\par`n" 
	. "\b\ul ENCOUNTER SUMMARY\ul0\b0\par\fs18`n"
	. summ . "\par\par{\tx2700\tx5220\tx7040`n"

	rtfOut := rtfHdr . rtfBody . rtfFtr

	LV_Modify(filenum,"col4","YES")
	Gui, Show
	FileDelete, %fileOut%.rtf
	FileAppend, %rtfOut%, %fileOut%.rtf
	outDir := (isAdmin) 
		? ".\completed\"
		: ".\test\"
;		: "\\PPWHIS01\Apps$\3mhisprd\Script\impunst\crd.imp\" . fileOut . ".rtf"

	FileCopy, %fileOut%.rtf, %outDir%%fileOut%.rtf, 1			; copy to the final directory
	FileMove, %fileOut%.rtf, completed\%fileout%.rtf ,1			; store in Completed, is this necessary?
	;FileMove, %fileIn%, archive\%fileout%-done.pdf, 1			; archive the PDF. Comment out if don't want to keep moving test PDF.
	
	
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
		fldval[pre "-" lbl] := m
		;~ MsgBox % i " ~ " j "`n" pre "-" lbl "`n" m
		;~ formatField(pre,lbl,m)
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
