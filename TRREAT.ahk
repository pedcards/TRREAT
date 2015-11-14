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
IfInString, A_WorkingDir, Dropbox					; Change enviroment if run from development vs production directory
{
	isAdmin := true
	holterDir := ".\Holter PDFs\"
	importFld := ".\Import\"
	chipDir := ".\Chipotle\"
} else {
	isAdmin := false
	holterDir := "\\chmc16\Cardio\EP\HoltER Database\Holter PDFs\"
	importFld := "\\chmc16\Cardio\EP\HoltER Database\Import\"
	chipDir := "\\chmc16\Cardio\Inpatient List\chipotle\"
}
user := A_UserName

Gui, Add, Listview, w600 -Multi NoSortHdr Grid r12 hwndHLV, Filename|Name|Device|Report|Fix
Gui, Add, Button, Disabled w600 h50 , Reload
Gui, Show
blocks := Object()
fields := Object()
newTxt := Object()
blk := Object()
blk2 := Object()
docs := Object()
docs := {"Chun, Terrence":"783118","Salerno, Jack":"343079","Seslar, Stephen":"358945"}

Loop, *.pdf
{ 
	summBl := summ := ""
	fileIn := A_LoopFileName
	SplitPath, fileIn,,,,fileOut
	RunWait, pdftotext.exe -l 2 -table -fixed 3 "%fileIn%" temp.txt , , hide
	FileRead, maintxt, temp.txt

	IfInString, maintxt, � Medtronic
		gosub PaceArt
	;~ if RegExMatch(maintxt,"i)�.*Medtronic")
		;~ gosub PaceArt
}

MsgBox Directory scan complete.
GuiControl, Enable, Reload

Exit


ButtonReload:
	Reload
Return

GuiClose:
ExitApp

PaceArt:
{
	fileNum += 1
	LV_Add("", fileIN)
	newtxtL:="", newTxtR:="", pos2:=200
	Loop, parse, maintxt, `n,`r									; first pass, determine graph column positions
	{
		i:=A_LoopField
		pos2:=((p:=RegExMatch(i,"(RA|RV|LV)[ ]+(Capture|Pacing|Sensing)[ ]+(Duration|Amplitude|Impedance)")) ? ((p<pos2) ? p : pos2) : pos2)
	}
	Loop, parse, maintxt, `n,`r									; second pass clean up
	{
		i:=A_LoopField
		if !(i)													; skip entirely blank lines
			continue
		j := substr(i,1,pos2-1)
		if (j ~= "Patient[ ]+Information")				; fix the most common spacing errors
			j := "Patient Information"
		if (j ~= "Detections[ ]+and[ ]+Therapies")
			j := "Detections and Therapies"
		newtxtL .= j . "`n"										; strip left from right columns
	}
	devtype := trim(strX(maintxt,":",RegExMatch(maintxt,"Device\s*Type:"),1,"`n",1,2))		; PM or ICD?
	
	FileDelete tempfile.txt
	FileAppend %newtxtL%, tempfile.txt

	if (devtype="Pacemaker") {
		gosub PaceArtPM
		if (reportErr) {
			LV_Modify(fileNum,"col4","no")
			LV_Modify(fileNum,"col5",reportErr)
			Gui, Show
			reportErr := ""
			return
		}
		; No errors, now generate report
		rtfBody := "\fs22\b\ul`nPATIENT INFORMATION\ul0\b0\par`n"
		. "\fs18\b Rhythm:\b0\tab " blk["Rhythm"] " \tab\b Referring MD:\b0\tab " blk["Referring"] "\par`n"
		. "\b Dependency:\b0\tab " blk["Dependency"] " \tab\b Following:\b0\tab " blk["Following"] "\par`n"
		. "\b Diagnosis:\b0\tab " StrX(blk["Diagnosis"]," - ",1,3,,1,1) "\par`n"
		. "}\fs22\par`n"
		. "\b\ul DEVICE INFORMATION\ul0\b0\par`n"
		. "\fs18 " blk["Manufacturer and Model"] ", serial number " blk["Serial Number"] 
		. ((pm_imp:=blk["Implant Date"]) ? ", implanted " pm_imp ((pm_impMD:=blk["Implant Provider"]) ? " by " pm_impMD : "") : "") ". `n"
		. "Generator cell voltage " (instr(tmp:=blk["Battery Voltage"],"(ERT = V )") ? tmp : substr(tmp,1,instr(tmp,"(ERT")-2)) ". "
		. ((pm_bat:=blk["Battery Status"]) ? "Battery status is " pm_bat ", with r" : "R") "emaining longevity " blk["Remaining Longevity"] ". `n"
		. "Brady programming mode is " blk["Mode"] " with lower rate " blk["Lower Rate"]
		. ((pm_URL:=blk["Upper Rate"])="bpm" ? "" : ", upper tracking rate " pm_URL)
		. ((pm_USR:=blk["Upper Sensor"])="bpm" ? "" : ", upper sensor rate " pm_USR) . ". `n"
		. ((pm_ADL:=blk["ADL Rate"])="bpm" ? "" : "ADL rate is " pm_ADL ". ")
		. ((pm_adap:=blk["Adaptive"]) ? "Adaptive mode is " pm_adap ". " : "")
		. (((pm_PAV:=blk["Paced"])="ms" or (pm_SAV:=blk["Sensed"])="ms") ? "" : "Paced and sensed AV delays are " pm_PAV " and " pm_SAV ", respectively. `n")
		. ((pm_RA:=blk["RA"])="%" ? "" : "RA pacing " pm_RA ". ") . ((pm_RV:=blk["RV"])="%" ? "" : "RV Pacing " pm_RV ". ")
		. ((pm_ASVS:=blk["AS-VS"])="%" ? "" : "AS-VS " pm_ASVS " ") . ((pm_ASVP:=blk["AS-VP"])="%" ? "" : "AS-VP " pm_ASVP " ")
		. ((pm_APVS:=blk["AP-VS"])="%" ? "" : "AP-VS " pm_APVS " ") . ((pm_APVP:=blk["AP-VP"])="%" ? "" : "AP-VP " pm_APVP) . "\par`n"
		. "\fs22\par`n"
		. "\b\ul LEAD INFORMATION\ul0\b0\par`n\fs18 "
		for k in leads
		{
			if (leads[pmlead:=leads[k],"model"]) {
				gosub PaceArtLeads
			}
		}
	}
	if (devtype="ICD") {
		gosub PaceArtICD
		if (reportErr) {
			LV_Modify(fileNum,"col4","no")
			LV_Modify(fileNum,"col5",reportErr)
			Gui, Show
			reportErr := ""
			return
		}
		; No errors, now generate report
		rtfBody := "\fs22\b\ul`nPATIENT INFORMATION\ul0\b0\par`n"
		. "\fs18\b Rhythm:\b0\tab " blk["Rhythm"] " \tab\b Referring MD:\b0\tab " blk["Referring"] "\par`n"
		. "\b Dependency:\b0\tab " blk["Dependency"] " \tab\b Following:\b0\tab " blk["Following"] "\par`n"
		. "\b Diagnosis:\b0\tab " blk["Diagnosis"] "\par`n"
		. "}\fs22\par`n"
		. "\b\ul DEVICE INFORMATION\ul0\b0\par`n"
		. "\fs18 " blk["Manufacturer and Model"] ", serial number " blk["Serial Number"] 
		. ((pm_imp:=blk["Implant Date"]) ? ", implanted " pm_imp ((pm_impMD:=blk["Implant Provider"]) ? " by " pm_impMD : "") : "") ". `n"
		. ((substr(pm_cell:=blk["battery voltage"],1,1)="V") ? "" : "Generator cell voltage " blk["Battery Voltage"] ". `n" )
		. ((pm_bat:=blk["Battery Status"]) ? "Battery status is " pm_bat : "")
		. (((pm_long:=blk["Remaining Longevity"])="months") ? ". `n" : ", with remaining longevity " pm_long ". `n")
		. ((blk["Mode"]) ? "Brady programming mode is " blk["Mode"] " with lower rate " blk["Lower Rate"] : "")
		. ((pm_URL:=blk["Upper Rate"])="bpm" ? "" : ", upper tracking rate " pm_URL)
		. ((pm_USR:=blk["Upper Sensor"])="bpm" ? "" : ", upper sensor rate " pm_USR) . ". `n"
		. ((pm_ADL:=blk["ADL Rate"])="bpm" ? "" : "ADL rate is " pm_ADL ". ")
		. ((pm_adap:=blk["Adaptive"]) ? "Adaptive mode is " pm_adap ". " : "")
		. ((blk["Paced"]="ms") or (blk["Sensed"]="ms") ? "" : "Paced and sensed AV delays are " blk["paced"] " and " blk["sensed"] ", respectively. `n")
		. ((pm_RA:=blk["RA"])="%" ? "" : "RA pacing " pm_RA ". ") . ((pm_RV:=blk["RV"])="%" ? "" : "RV Pacing " pm_RV ". ")
		. ((pm_ASVS:=blk["AS-VS"])="%" ? "" : "AS-VS " pm_ASVS " ") . ((pm_ASVP:=blk["AS-VP"])="%" ? "" : "AS-VP " pm_ASVP " ")
		. ((pm_APVS:=blk["AP-VS"])="%" ? "" : "AP-VS " pm_APVS " ") . ((pm_APVP:=blk["AP-VP"])="%" ? "" : "AP-VP " pm_APVP) . "\par`n"
		. "\fs22\par`n"
		. "\b\ul LEAD INFORMATION\ul0\b0\par`n\fs18 "
		for k in leads
		{
			if (leads[pmlead:=leads[k],"model"]) {
				gosub PaceArtLeads
			}
		}
		rtfBody .= "\fs22\par\b\ul DETECTIONS AND THERAPIES\ul0\b0\par`n\fs18 "
		. (((tmp:=ther["det VF (VHR)","rate"]) and !(tmp="bpm") and !(tmp="blank")) 
			? "VF zone " tmp ", with " ((((tmp:=ther["det VF (VHR)","therapy"])="DISABLED") or (tmp="ms"))
			? "monitor only. " : tmp ". ") : "")
		rtfBody .= (((tmp:=ther["det Fast VT","rate"]) and !(tmp="bpm") and !(tmp="blank")) 
			? "Fast VT zone " tmp ", with " ((((tmp:=ther["det Fast VT","therapy"])="DISABLED") or (tmp="ms"))
			? "monitor only. " : tmp ". ") : "")
		rtfBody .= (((tmp:=ther["det Slow VT","rate"]) and !(tmp="bpm") and !(tmp="blank")) 
			? "Slow VT zone " tmp ", with " ((((tmp:=ther["det Slow VT","therapy"])="DISABLED") or (tmp="ms"))
			? "monitor only. " : tmp ". ") : "")
		rtfBody .= (((tmp:=ther["det Vslow VT","rate"]) and !(tmp="bpm") and !(tmp="blank")) 
			? "Very slow VT zone " tmp ", with " ((((tmp:=ther["det Vslow VT","therapy"])="DISABLED") or (tmp="ms"))
			? "monitor only. " : tmp ". ") : "")
		rtfBody .= (((tmp:=ther["det VT-NS","rate"]) and !(tmp="bpm") and !(tmp="blank")) 
			? "NS-VT zone " tmp ", with " ((((tmp:=ther["det VT-NS","therapy"])="DISABLED") or (tmp="ms"))
			? "monitor only. " : tmp ". ") : "")
		. "There have been "
		. ((det_vv:=((tmp:=blk["VF (VHR)"]) ? tmp : 0)) ? tmp " VF, " : "")
		. ((det_vv+=((tmp:=blk["Fast VT"]) ? tmp : 0)) ? tmp " Fast VT, " : "")
		. ((det_vv+=((tmp:=blk["Slow VT"]) ? tmp : 0)) ? tmp " Slow VT, " : "")
		. ((det_vv+=((tmp:=blk["Vslow VT"]) ? tmp : 0)) ? tmp " Very slow VT, and " : "")
		. ((det_vv+=((tmp:=blk["VT-NS"]) ? tmp : 0)) ? tmp " NS-VT" : "")
		. ((det_vv) ? "" : "no ventricular arrhythmia")
		. " episodes detected. `n"
		. "There have been "
		. ((det_sv:=((tmp:=blk["AF (AHR)"]) ? tmp : 0)) ? tmp " AF, " : "")
		. ((det_sv+=((tmp:=blk["AT"]) ? tmp : 0)) ? tmp " atrial tach, " : "")
		. ((det_sv+=((tmp:=blk["SVT"]) ? tmp : 0)) ? tmp " SVT, " : "")
		. ((det_sv+=((tmp:=blk["AT-NS"]) ? tmp : 0)) ? tmp " NS-AT, and " : "")
		. ((det_sv+=((tmp:=blk["Brady"]) ? tmp : 0)) ? tmp " brady" : "")
		. ((det_sv) ? "" : "no atrial arrhythmia")
		. " episodes detected.\par`n"
	}
	if (devtype~="ILR") {
		gosub PaceArtLINQ
		if (reportErr) {
			LV_Modify(fileNum,"col4","no")
			LV_Modify(fileNum,"col5",reportErr)
			Gui, Show
			reportErr := ""
			return
		}
		; No errors, now generate report
		rtfBody := "\fs22\b\ul`nPATIENT INFORMATION\ul0\b0\par`n"
		. "\fs18"
		. "\b Diagnosis:\b0\tab " StrX(blk["Diagnosis"]," - ",1,3,,1,1) "\par`n"
		. "\b Referring MD:\b0\tab " blk["Referring"] "\par`n"
		. "\b Following:\b0\tab " blk["Following"] "\par`n"
		. "}\fs22\par`n"
		. "\b\ul DEVICE INFORMATION\ul0\b0\par`n"
		. "\fs18 " blk["Manufacturer and Model"] ", serial number " blk["Serial Number"] 
		. ((pm_imp:=blk["Implant Date"]) ? ", implanted " pm_imp : "") ". "
		. ((pm_bat:=blk["Battery Status"]) ? "Battery status is " pm_bat "." : "") "\par`n"
		. "\fs22\par`n"
		. "\b\ul DETECTION CRITERIA\ul0\b0\par`n"
		. "\fs18 "
		. (RegExMatch((d_VF:=blk["det_VF (VHR)"]),"i)(Enabled|Disabled)") ? "VF: " RegExReplace(d_VF,"\d\sbpm\sms\ss") ".\par`n" : "")
		. (RegExMatch((d_FVT:=blk["det_Fast VT"]),"i)(Enabled|Disabled)") ? "Fast VT: " d_FVT ".\par`n" : "")
		. (RegExMatch((d_SlowVT:=blk["det_Slow VT"]),"i)(Enabled|Disabled)") ? "Slow VT: " d_SlowVT ".\par`n" : "")
		. (RegExMatch((d_VSlow:=blk["det_V-Slow VT"]),"i)(Enabled|Disabled)") ? "V-Slow VT: " d_VSlow ".\par`n" : "")
		. (RegExMatch((d_AF:=blk["det_AF (AHR)"]),"i)(Enabled|Disabled)") ? "AF: " RegExReplace(d_AF,"\d\sbpm\sms\ss") ".\par`n" : "")
		. (RegExMatch((d_AT:=blk["det_AT"]),"i)(Enabled|Disabled)") ? "AT: " d_AT ".\par`n" : "")
		. (RegExMatch((d_Asys:=blk["det_Asystole"]),"i)(Enabled|Disabled)") ? "Asystole: " d_Asys ".\par`n" : "")
		. (RegExMatch((d_Brady:=blk["det_Brady"]),"i)(Enabled|Disabled)") ? "Brady: " d_Brady ".\par`n" : "")
		. "\fs22\par`n"
		. "\b\ul LEAD INFORMATION\ul0\b0\par`n\fs18 "
	}
	gosub PaceArtPrint
	Gui, Show
Return
}

PaceArtPM:
{
	LV_Modify(fileNum,"col3","PM")
	Gui, Show
	blocks := ["Patient Information"
		,"Device and Lead Information"
		,"Lead Manufacturer"
		,"Brady Programming"
		,"Measurements"
		,"Encounter Summary","� Medtronic"]
	fields[1] := ["Patient Name:","Patient ID:","Date of Birth:","Gender:","Rhythm:","Dependency:","Next In-Clinic:","Next Remote:"
		,"Referring:","Following:","Blood Press.:","Diagnosis:"]
	fields[2] := ["Serial Number:","Implant Date:","Implant Provider:","Battery Voltage:","Battery Status:","Remaining Longevity:"]
	fields[3] := ["Mode:","Lower Rate:","Upper Rate:","Upper Sensor:","ADL Rate:","Hysteresis:","Sleep Rate:","Detection:","Fallback Rate:","Fallback Mode:"
		,"Amplitude:","Pulse Width:","Pace Polarity:","Sensitivity:","Blanking:","Refractory:","Sense Polarity:","LV Pace Path:","VV Delay:"
		,"Adaptive:","Paced:","Sensed:","Paced Min:","Sensed Min:","PMT Int.:","PVC Resp.:","Notes"]
	fields[4] := ["Presenting","Rate:","AV Delay:","Magnet Mode","Rate:","Interval:","AV Delay:","Duration:","Capture:","Sensing:"
		,"Pacing Impedance:","Capture Amplitude:","Capture Duration:","Sensing Amplitude:","Lead Information"
		,"Lead Status:","Integrity Count:","Polarization:","Evoked Response:"
		,"Percent Pacing","RA:","RV:","LV:","CRT:","AS-VS:","AS-VP:","AP-VS:","AP-VP:"]
	fields[5] := ["Electronically Signed By:","Last Modified By:","Signed Date:","Encounter Date:","Encounter Type:"]

	; Get the PATIENT INFORMATION block
	ptInfo := columns(newtxtL,blocks[1],"Comments:",,"Referring:")
	fieldvals(ptInfo,1)

	; Get the DEVICE INFORMATION block
	devInfo := columns(newtxtL,blocks[2],blocks[3],,"Battery Voltage:")
	fieldvals(devInfo,2)
	tmp := trim(strX(newtxtL,"Manufacturer and Model:",1,23,"Device",1,6), " `n")
	blk["Manufacturer and Model"] := tmp								; Has different column width

	; Get the LEAD INFORMATION block
	leadInfo := columns(newtxtL,blocks[3],blocks[4],1)						; Also different table widths
	leads := cellvals(leadInfo,,,"leads")

	; BRADY PROGRAMMING parameters
	bradyParam := columns(newtxtL,blocks[4],blocks[5],"leads","Amplitude:","Adaptive:")
	fieldvals(bradyParam,3)

	; PACING AND SENSING subtable
	outputs := cellvals(bradyParam,"Pacing and Sensing","Heart Failure")
	val := "Sensitivity", chamber := "LV"

	; MEASUREMENTS table
	meas := columns(newtxtL,blocks[5],blocks[6],,"Pacing Impedance:","RA:")
	fieldvals(meas,4)

	; LEAD IMPEDANCE AND THRESHOLDS subtable
	thr := cellvals(meas,"Lead Impedance and Thresholds","Lead Information")
	val := "Capture Duration", chamber := "rv"

	; ENCOUNTER SUMMARY block
	summBl := trim(columns(maintxt,blocks[6],blocks[7])," `n")
	cleanSpace(summBl)
	if (instr(summBl,"(Since Last Reset)",1)) {
		reportErr .= "Save file in 'Encounter Brady (no strips)' format. "
	} 
	if !(instr(summBl,"Electronically Signed By:")) {
		reportErr .= "Report not signed. "
	}
	if !(summ:=trim(SubStr(summBl,1,RegExMatch(summBl,"(Electronically Signed By)|(Last Modified By)|(Encounter Date)")-1))) {
		reportErr .= "No summary. "
	}
	fieldvals(summBl,5)
	enc_MD := docs[strX(blk["Electronically Signed By"],,1,1," MD",1,3)]
	enc_signed := strX(blk["Signed Date"],,1,1," ",1,1)
	enc_date := strX(blk["Encounter Date"],,1,1," ",1,1)
	if !(enc_MD) {
		reportErr .= "Not MD signed. "
	}
Return
}

PaceArtICD:
{
	LV_Modify(fileNum,"col3","ICD")
	Gui, Show
	blocks := ["Patient Information"
		,"Device and Lead Information"
		,"Lead Manufacturer"
		,"Detections and Therapies"
		,"Counters (Since Last Reset)"
		,"Brady Programming"
		,"Lead Data"
		,"Lead Status:"
		,"Encounter Summary"
		,"� Medtronic"]
	fields[1] := ["Patient Name:","Patient ID:","Date of Birth:","Gender:","Rhythm:","Dependency:","Next In-Clinic:","Next Remote:","Comments:"
		,"Referring:","Following:","Blood Press.:","Diagnosis:"]
	fields[2] := ["Serial Number:","Implant Date:","Implant Provider:","Battery Voltage:","Battery Status:","Remaining Longevity:"]
	fields[3] := ["VF (VHR):","AF (AHR):","Fast VT:","AT:","Slow VT:","SVT:","VSlow VT:","AT-NS:","VT-NS:","Brady:"]
	fields[4] := ["RA:","AS-VS:","RV:","AS-VP:","LV:","AP-VS:","CRT:","AP-VP:"]
	fields[5] := ["Mode:","Lower Rate:","Upper Rate:","Upper Sensor:","ADL Rate:","Hysteresis:","Sleep Rate:","Detection:","Fallback Rate:","Fallback Mode:"
		,"Amplitude:","Pulse Width:","Pace Polarity:","Sensitivity:","Blanking:","Refractory:","Sense Polarity:","LV Pace Path:","VV Delay:"
		,"Adaptive:","Paced:","Sensed:","Paced Min:","Sensed Min:","PMT Int.:","PVC Resp.:","Notes"]
	fields[7] := ["Electronically Signed By:","Last Modified By:","Signed Date:","Encounter Date:","Encounter Type:"]

	; Get the PATIENT INFORMATION block
	ptInfo := columns(newtxtL,blocks[1],blocks[2],,"Referring:")
	fieldvals(ptInfo,1)

	; Get the DEVICE INFORMATION block
	devInfo := columns(newtxtL,blocks[2],blocks[3],,"Battery Voltage:")
	fieldvals(devInfo,2)
	tmp := trim(strX(newtxtL,"Manufacturer and Model:",1,23,"Device",1,6), " `n")
	blk["Manufacturer and Model"] := tmp								; Has different column width

	; Get the LEAD INFORMATION block
	leadInfo := columns(newtxtL,blocks[3],blocks[4],1)						; Also different table widths
	leads := cellvals(leadInfo,,,"leads")

	; Get DETECTIONS AND THERAPIES
	detTher := columns(newtxtL,blocks[4],blocks[5],,"Configuration Comments")
	Ther := cellvals(detTher,,,"detect")
	
	; Get Detection Counters
	ctrs := columns(newtxtL,blocks[5],blocks[6],,"Shocks Delivered:","RA:")
	ctrs_D := strX(ctrs,"Detections",1,0,"Brady:",1,7)
	fieldvals(ctrs_D,3)
	
	; Get Therapy counters
	therDel := cellvals(ctrs,"Therapies","Mode Switch Detections","ther")
	
	; Get Pacing counters
	paceCtr := columns(ctrs,"Pacing","Burden","AS-VS")
	fieldvals(paceCtr,4)

	; BRADY PROGRAMMING parameters
	bradyParam := columns(newtxtL,blocks[6],blocks[7],"leads","Amplitude:","Adaptive:")
	fieldvals(bradyParam,5)

	; PACING AND SENSING subtable
	outputs := cellvals(bradyParam,"Pacing and Sensing","Heart Failure")
	val := "Sensitivity", chamber := "RV"

	; Get Lead values
	meas := columns(newtxtL,blocks[7],blocks[8])
	thr := cellvals(meas,"Lead Impedance / Thresholds","Lead Information","ther")
	
	; ENCOUNTER SUMMARY block
	summBl := trim(columns(maintxt,blocks[9],blocks[10])," `n")
	cleanSpace(summBl)
	if (instr(summBl,"(Since Last Reset)",1)) {
		reportErr .= "Save file in 'Encounter Tachy (detailed)' format. "
	}
	if !(instr(summBl,"Electronically Signed By:")) {
		reportErr .= "Report not signed. "
	}
	if !(summ:=trim(SubStr(summBl,1,RegExMatch(summBl,"(Electronically Signed By)|(Last Modified By)|(Encounter Date)")-1))) {
		reportErr .= "No summary. "
	}
	fieldvals(summBl,7)
	enc_MD := docs[strX(blk["Electronically Signed By"],,1,1," MD",1,3)]
	enc_signed := strX(blk["Signed Date"],,1,1," ",1,1)
	enc_date := strX(blk["Encounter Date"],,1,1," ",1,1)
	if !(enc_MD) {
		reportErr .= "Not MD signed. "
	}
Return
}

PaceArtLeads:
{
	rtfBody .= "\b " pmlead " lead: \b0 " leads[pmlead,"model"] 
	. ((tmp:=leads[pmlead,"serial"]) ? ", serial number " tmp : "")
	. ((tmp:=leads[pmlead,"date"]) ? ", implanted " tmp : "") ". `n"
	. ((tmp:=thr["pacing impedance",pmlead]) ? "Lead impedance " tmp " ohms. " : "")
	. ((tmp:=thr["capture amplitude",pmlead]) ? "Capture threshold " tmp " V at " thr["capture duration",pmlead] " ms. " : "")
	. ((tmp:=thr["sensing amplitude",pmlead]) ? ((pmlead="RA") ? "P wave " : "R wave ") "sensing " tmp " mV. " : "")
	. ((tmp:=outputs["amplitude",pmlead]) ? "Pacing output " tmp " V at " outputs["pulse width",pmlead] " ms" ((tmp:=outputs["pace polarity",pmlead]) ? " (" tmp "). " : ". ") : "")
	. ((tmp:=outputs["sensitivity",pmlead]) ? "Sensitivity " tmp " mV" ((tmp:=outputs["sense polarity",pmlead]) ? " (" tmp "). " : ". ") : "")
	. "\par`n"

Return	
}

PaceArtLINQ:
{
	LV_Modify(fileNum,"col3","ILR")
	Gui, Show
	blocks := ["Patient Information"
		,"Device Information"
		,"Past Encounters"
		,"Detections"
		,"Encounter Summary","� Medtronic"]
	fields[1] := ["Patient Name:","Patient ID:","Date of Birth:","Gender:","Blood Pressure:"
		,"Referring:","Following:","Rhythm:"
		,"Next In-clinic:","Next Remote:","Diagnosis:","Dependency:"]
	fields[2] := ["Implant Date:","Serial Number:","Battery Status:"]
	fields[3] := ["VF (VHR):","VT:","SVT:","VT-NS:","AF (AHR):","AT:","AT-NS:"
		,"Switch:","Activated:","Asystole:","Brady:","Other:"]
	fields[4] := ["VF (VHR):","Fast VT:","Slow VT:","V-Slow VT:","AF (AHR):","AT:","Asystole:","Brady:"]
	fields[5] := ["Electronically Signed By:","Last Modified By:","Signed Date:","Encounter Date:","Encounter Type:"]

	; Get the PATIENT INFORMATION block
	ptInfo := columns(newtxtL,blocks[1],"Comments:",,"Referring:","Next In-Clinic:")
	fieldvals(ptInfo,1)

	; Get the DEVICE INFORMATION block
	devInfo := columns(newtxtL,blocks[2],blocks[3],,"Device Type:","Serial Number:")
	fieldvals(devInfo,2)
	tmp := trim(strX(newtxtL,"Manufacturer and Model:",1,23,"`n",1,1), " `n")
	blk["Manufacturer and Model"] := tmp								; Has different column width

	; Get the EPISODES and DETECTIONS block
	epdet := columns(newtxtL,blocks[4],blocks[5],,"Detection")
	epBlk := columns(epdet,"","Detection",,"Asystole:")
	fieldvals(epBlk,3,,"ep")
	detBlk := strX(epdet,"Detection",1,0)
	fieldvals(detBlk,4,,"det")

	; ENCOUNTER SUMMARY block
	summBl := trim(columns(maintxt,blocks[5],blocks[6])," `n")
	cleanSpace(summBl)
	;~ if !(instr(summBl,"Electronically Signed By:")) {
		;~ reportErr .= "Report not signed. "
	;~ }
	;~ if !(summ:=trim(SubStr(summBl,1,RegExMatch(summBl,"(Electronically Signed By)|(Last Modified By)|(Encounter Date)")-1))) {
		;~ reportErr .= "No summary. "
	;~ }
	fieldvals(summBl,5)
	enc_MD := docs[strX(blk["Electronically Signed By"],,1,1," MD",1,3)]
	enc_signed := strX(blk["Signed Date"],,1,1," ",1,1)
	enc_date := strX(blk["Encounter Date"],,1,1," ",1,1)
	;~ if !(enc_MD) {
		;~ reportErr .= "Not MD signed. "
	;~ }
Return
}

PaceArtPrint:
{
	if (RegExMatch(summ,"\<\d*\>")) {
		enc_FIN:=strX(summ,"<",1,1,">",1,1,nn)
		summ := trim(substr(summ,nn+1))
	} else {
		InputBox , enc_FIN, % blk["Patient Name"] " - " enc_date, REQUIRED:`n`nEncounter number`n(8 digits)
	}

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

	rtfBody := "\tx1620\tx5220\tx7040" . rtfBody . "\fs22\par`n" 
	. "\b\ul ENCOUNTER SUMMARY\ul0\b0\par\fs18`n"
	. summ . "\par\par{\tx2700\tx5220\tx7040`n"
	. "\b Electronically Signed By:\b0\tab " blk["Electronically Signed By"] "\tab\b Encounter Type:\b0\tab " blk["Encounter Type"] "\par`n"
	. "\b Signed Date:\b0\tab " blk["Signed Date"] "\tab\b Encounter Date:\b0\tab " blk["Encounter Date"] "\par}`n"

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
	
Return	
}

columns(x,blk1,blk2,incl:="",col2:="",col3:="",col4:="") {
/*	Returns string as a single column.
	x 		= input string
	blk1	= leading string to start block
	blk2	= ending string to end block
	incl	= if null, exclude blk1 string; if !null, remove blk1 string
	col2	= string demarcates start of COLUMN 2
	col3	= string demarcates start of COLUMN 3
	col4	= string demarcates start of COLUMN 4
*/
	txt := strX(x,blk1,1,(incl ? 0 : StrLen(blk1)),blk2,1,StrLen(blk2))
	StringReplace, col2, col2, %A_space%, [ ]+, All
	StringReplace, col3, col3, %A_space%, [ ]+, All
	StringReplace, col4, col4, %A_space%, [ ]+, All
	
	loop, parse, txt, `n,`r										; find position of columns 2, 3, and 4
	{
		i:=A_LoopField
		if (t:=RegExMatch(i,"i)" col2))
			pos2:=t
		if (t:=RegExMatch(i,"i)" col3))
			pos3:=t
		if (t:=RegExMatch(i,"i)" col4))
			pos4:=t
	}
	loop, parse, txt, `n,`r
	{
		i:=A_LoopField
		if (instr(i,"Manufacturer and Model:")) {				; Skip the M and M line, screws up the table formatting
			continue
		}
		txt1 .= substr(i,1,pos2-1) . "`n"
		if (col4) {
			pos4ck := pos4
			while !(substr(i,pos4ck-1,1)=" ") {
				pos4ck := pos4ck-1
			}
			txt4 .= substr(i,pos4ck) . "`n"
			txt3 .= substr(i,pos3,pos4ck-pos3) . "`n"
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"
			continue
		} 
		if (col3) {
			txt2 .= substr(i,pos2,pos3-pos2) . "`n"
			txt3 .= substr(i,pos3) . "`n"
			continue
		}
		txt2 .= substr(i,pos2) . "`n"
	}
	return txt1 . txt2 . txt3 . txt4
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

fieldvals(x,bl,bl2:="",pre:="") {
/*	Matches field values and results. Gets text between FIELDS[k] to FIELDS[k+1]. Excess whitespace removed. Returns results in array BLK[].
	x	= input text
	bl	= which FIELD number to use
	bl2	= if present, use blk2
	pre	= if present, prefix name
*/
	global blocks, fields, blk, blk2
	blk2 := ""
	blk2 := Object()
	for k, i in fields[bl]
	{
		j := fields[bl][k+1]
		m := trim(strX(x,i,n,StrLen(i),j,1,StrLen(j)+1,n), " `n")
		cleanSpace(m)
		if (pre="det") {
			if !(m~="i)(Enabled|Disabled)") {
				m := ""
			} else {
				m := RegExReplace(m,"d\sbpm\sms\ss","d")
			}
		}
		if (substr(i,0)=":") {
			StringTrimRight i,i,1
		}
		if (pre) {
			i := pre "_" i
		}
		if (bl2) {
			blk2[i] := cleancolon(m)
		} else {
			blk[i] := m
			;MsgBox,,% i, % m
		}
	}
	if (bl2) {
		blk[bl2] := blk2
	}
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
	;BS .= "(.*?)\s{3}"
	MsgBox % ES
	;rem:="[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"
	pos0 := RegExMatch(h,BS,bPat,((BO<1)?1:BO))
	pos1 := RegExMatch(h,ES,ePat,pos0+bPat.len)
	MsgBox % pos0 " - " pos1
	N := pos1+((ET)?0:(ePat.len))
	return substr(h,pos0+((BT)?(bPat.len):0),N-pos0-bPat.len)
}

#Include strx.ahk
