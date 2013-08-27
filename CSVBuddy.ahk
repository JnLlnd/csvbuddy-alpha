;===============================================
/* CSV Buddy v0.1
Written using AutoHotkey_L v1.1.09.03 (http://l.autohotkey.net/)
By jlalonde on AHK forum
2013-08-16
*/

#NoEnv
#SingleInstance force
#Include <JLDev>
#Include %A_ScriptDir%\..\ObjCSV\lib\ObjCSV.ahk

EM_GETLINECOUNT = 0xBA

global strApplicationName := "CSV Buddy"
global strApplicationVersion := "0.1"


Gui, 1:New, +Resize, %strApplicationName%

Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Text, x10, %strApplicationName%

Gui, 1:Font, s10 w700, Verdana
Gui, 1:Add, Tab2, w950 r4 vtabCSVBuddy gChangedTabCSVBuddy, % " 1) Load CSV File     ||     2) Edit Columns     |     3) Save CSV File     |     About     "
Gui, 1:Font

Gui, 1:Tab, 1
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToLoad w85 right, CSV &file to load:
Gui, 1:Add, Edit,		yp		x100	vstrFileToLoad disabled gChangedFileToLoad
Gui, 1:Add, Button,	yp		x+5		vbtnHelpFileToLoad gButtonHelpFileToLoad, ?
Gui, 1:Add, Button,	yp		x+5		vbtnSelectFileToLoad gButtonSelectFileToLoad default, &Select
GuiControl, 1:Focus, btnSelectFileToLoad
Gui, 1:Add, Text,		y+10	x10 	vlblHeader w85 right, CSV file &Header:
Gui, 1:Add, Edit,		yp		x100	vstrFileHeaderEscaped disabled
Gui, 1:Add, Button,	yp		x+5		vbtnHelpHeader gButtonHelpHeader, ?
Gui, 1:Add, Button,	yp		x+5		vbtnPreviewFile gButtonPreviewFile hidden, &Preview
Gui, 1:Add, Radio,	y+10	x100	vradGetHeader gClickRadGetHeader checked, &Get header from CSV file
Gui, 1:Add, Radio,	yp		x+5		vradSetHeader gClickRadSetHeader, Set &CSV header
Gui, 1:Add, Button,	yp		x+0		vbtnHelpSetHeader gButtonHelpSetHeader, ?
Gui, 1:Add, Text,		xp		x+45	vlblFieldDelimiter1, Field &delimiter:
Gui, 1:Add, Edit,		yp		x+5		vstrFieldDelimiter1 gChangedFieldDelimiter1 w20 limit1 center, `, 
Gui, 1:Add, Button,	yp		x+5		vbtnHelpFieldDelimiter1 gButtonHelpFieldDelimiter1, ?
Gui, 1:Add, Text,		yp		x+45	vlblFieldEncapsulator1, Field e&ncapsulator:
Gui, 1:Add, Edit,		yp		x+5		vstrFieldEncapsulator1 gChangedFieldEncapsulator1 w20 limit1 center, `"
Gui, 1:Add, Button,	yp		x+5		vbtnHelpEncapsulator1 gButtonHelpEncapsulator1, ?
Gui, 1:Add, Checkbox,	yp		x+45	vblnMultiline1 checked, &Multi-line fields
Gui, 1:Add, Button,	yp		x+0		vbtnHelpMultiline1 gButtonHelpMultiline1, ?
Gui, 1:Add, Button,	yp		x+5		vbtnLoadFile gButtonLoadFile, &Load

Gui, 1:Tab, 2
Gui, 1:Add, Text,		y+10	x10		vlblRenameFields w85 right, Ren&ame fields:
Gui, 1:Add, Edit,		yp		x100	vstrRenameEscaped
Gui, 1:Add, Button,	yp		x+0		vbtnSetRename gButtonSetRename, &Rename
Gui, 1:Add, Button,	yp		x+5		vbtnHelpRename gButtonHelpRename, ?
Gui, 1:Add, Text,		y+10	x10		vlblSelectFields w85 right, Selec&t fields:
Gui, 1:Add, Edit,		yp		x100	vstrSelectEscaped
Gui, 1:Add, Button,	yp		x+0		vbtnSetSelect gButtonSetSelect, S&elect
Gui, 1:Add, Button,	yp		x+5		vbtnHelpSelect gButtonHelpSelect, ?
Gui, 1:Add, Text,		y+10	x10		vlblOrderFields w85 right, Order f&ields:
Gui, 1:Add, Edit,		yp		x100	vstrOrderEscaped
Gui, 1:Add, Button,	yp		x+0		vbtnSetOrder gButtonSetOrder, &Order
Gui, 1:Add, Button,	yp		x+5		vbtnHelpOrder gButtonHelpOrder, ?

Gui, 1:Tab, 3
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToSave w85 right, CS&V file to save:
Gui, 1:Add, Edit,		yp		x100	vstrFileToSave gChangedFileToSave
Gui, 1:Add, Button,	yp		x+5		vbtnHelpFileToSave gButtonHelpFileToSave, ?
Gui, 1:Add, Button,	yp		x+5		vbtnSelectFileToSave gButtonSelectFileToSave default, &Select
GuiControl, 1:Focus, btnSelectFileToSave
Gui, 1:Add, Text,		y+10	x100	vlblFieldDelimiter3, Field delimiter:
Gui, 1:Add, Edit,		yp		x200	vstrFieldDelimiter3 w20 limit1 center, `, 
Gui, 1:Add, Button,	yp		x+5		vbtnHelpFieldDelimiter3 gButtonHelpFieldDelimiter3, ?
Gui, 1:Add, Text,		y+10	x100	vlblFieldEncapsulator3, Field encaps&ulator:
Gui, 1:Add, Edit,		yp		x200	vstrFieldEncapsulator3 w20 limit1 center, `"
Gui, 1:Add, Button,	yp		x+5		vbtnHelpEncapsulator3 gButtonHelpEncapsulator3, ?
Gui, 1:Add, Radio,	y100	x300	vradSaveWithHeader checked, Save &with CSV header
Gui, 1:Add, Radio,	y+10	x300	vradSaveNoHeader, Save without CSV header
Gui, 1:Add, Button,	y100	x450	vbtnHelpSaveHeader gButtonHelpSaveHeader, ?
Gui, 1:Add, Radio,	y100	x500	vradSaveMultiline gClickRadSaveMultiline checked, Save multi-line
Gui, 1:Add, Radio,	y+10	x500	vradSaveSingleline gClickRadSaveSingleline, Save single-line
Gui, 1:Add, Button,	y100	x620	vbtnHelpMultiline gButtonHelpSaveMultiline, ?
Gui, 1:Add, Text,		y+25	x500	vlblEndoflineReplacement hidden, End-of-line replacement:
Gui, 1:Add, Edit,		yp		x620	vstrEndoflineReplacement hidden w50 center, % chr(182)
Gui, 1:Add, Button,	y105	x+5		vbtnSaveFile gButtonSaveFile, Save
Gui, 1:Add, Button,	y137	x+5		vbtnCheckFile hidden gButtonCheckFile, Check

Gui, 1:Tab, 4
Gui, 1:Add, Text,		y+10	x10		vlblAboutText, CSV Buddy v0.1 ALPHA`nby Jean Lalonde (JnLlnd on AHK forum)`nAll rights reserved (c)2013 - DO NOT DISTRIBUTE WITHOUT AUTHOR AUTORIZATION`n`nUsing ObjCSV AHK_L library: www.github.com/JnLlnd/ObjCSV`nIcon: Visual Pharm - http://www.visualpharm.com

Gui, 1:Tab

Gui, 1:Add, ListView, x10 r24 w200 vlvData -ReadOnly gListViewEvents

Gui, 1:Show, Autosize

; ###
; GuiControl, Choose, tabCSVBuddy, 3

return



ChangedTabCSVBuddy:
Gui, 1:Submit, NoHide
;  " 1) Load CSV File     ||     2) Edit Columns     |     3) Save CSV File     "
if InStr(tabCSVBuddy, "Load")
	GuiControl, 1:+Default, btnSelectFileToLoad
else if InStr(tabCSVBuddy, "Edit")
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnReady
	else
	{
		MsgBox, 48, %strApplicationName%, First load a CSV file in the first tab.
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
else if InStr(tabCSVBuddy, "Save")
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnSelectFileToSave
	else
	{
		MsgBox, 48, %strApplicationName%, First load a CSV file in the first tab.
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
else if InStr(tabCSVBuddy, "About")
{
	; do nothing
	; GuiControl, 1:+Default, ###
	; ###(tabCSVBuddy)
}
else
	###(tabCSVBuddy . " !?!")
return




; --------------------- TAB 1 --------------------------


ButtonHelpFileToLoad:
Help("CSV File To Load", "Hit ""Select"" to choose the CSV file to load. When other options are OK, hit ""Load"" to import the file in the list below.`n`nNote that a maximum of 200 fields can be loaded.")
return



ButtonSelectFileToLoad:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
FileSelectFile, strInputFile, 3, %A_ScriptDir%, Select CSV File to load
if !(StrLen(strInputFile))
	return
GuiControl, 1:, strFileToLoad, %strInputFile%
if (radGetHeader) or !(StrLen(strFileHeaderEscaped))
{
	FileReadLine, strCurrentHeader, %strInputFile%, 1
	GuiControl, 1:, strFileHeaderEscaped, % StrEscape(strCurrentHeader)
}
GuiControl, 1:+Default, btnLoadFile
GuiControl, 1:Focus, btnLoadFile
return


ChangedFileToLoad:
Gui, 1:Submit, NoHide
loop
{
	SplitPath, strFileToLoad, , strOutDir, strOutExtension, strOutNameNoExt
	strNewName := strOutDir . "\" . strOutNameNoExt . " (" . A_Index  . ")." . strOutExtension
	GuiControl, 1:, strFileToSave, %strNewName%
	if !FileExist(strNewName)
		break
}
if FileExist(strFileToLoad)
	GuiControl, 1:Show, btnPreviewFile
else
	GuiControl, 1:Hide, btnPreviewFile
return



ButtonHelpHeader:
Help("CSV Header", "Most of the time, the first line of a CSV file contains the CSV header, a list of field names, separated by a field delimiter. If your file contains a CSV Header, select the radio button ""Get CSV Header"". When you select a file (using the ""Select"" button), the ""CSV Header"" zone displays the content of the first line of the file.`n`nNote that invisible characters used as delimiters (for example Tab) are displayed with an escape character. For example, Tabs are shown as ""``t"".`n`nIf the file does not contain a CSV header, select the radio button ""Set CSV Header"" and enter in the ""CSV Header"" zone the field names for each column of data in the file, seperated by the field delimiter.")
return



ButtonPreviewFile:
Gui, 1:Submit, NoHide
if !StrLen(strFileToLoad)
{
	MsgBox, 48, %strApplicationName%, First use the "Select" button to choose the CSV file you want to load.
	return
}
run notepad.exe %strFileToLoad%
return



ClickRadGetHeader:
Gui, 1:Submit, NoHide
FileReadLine, strCurrentHeader, %strFileToLoad%, 1
GuiControl, 1:, strFileHeaderEscaped, % StrEscape(strCurrentHeader)
GuiControl, 1:Disable, strFileHeaderEscaped
GuiControl, 1:, lblHeader, CSV Header:
return



ClickRadSetHeader:
GuiControl, 1:Enable, strFileHeaderEscaped
GuiControl, 1:, lblHeader, CSV &Header:
GuiControl, 1:Focus, strFileHeaderEscaped
return



ButtonHelpSetHeader:
Gui, 1:Submit, NoHide
Help("CSV Get/Set CSV Header", "If the first line of the CSV file contains the list of data columns field names, click ""Get header from CSV file"". If not, click ""Set CSV header"" and enter the list of field names separated by the Field delimiter.")
return



ChangedFieldDelimiter1:
Gui, 1:Submit, NoHide
GuiControl, 1:, strFieldDelimiter3, %strFieldDelimiter1%
;  effacer la ligne suivante quand je serai certain que ça ne servait à rien
; GuiControl, 1:, strFieldDelimiter1, %strFieldDelimiter1%
return



ButtonHelpFieldDelimiter1:
Help("Field Delimiter", "Each field in the CSV header or in data rows of the file must be separated by a field delimiter. This is often comma ( , ), semicolon ( `; ) or Tab.`n`nEnter any single character such as comma, semicolon or one of these letters for special characters:`n`nt`tTab (HT)`nn`tLinefeed (LF)`nr`tCarriage return (CR)`nf`tFormfeed (FF)`n`nUse the ""Preview"" button to find what is the field delimiter in this file.")
return



ChangedFieldEncapsulator1:
Gui, 1:Submit, NoHide
GuiControl, 1:, strFieldEncapsulator3, %strFieldEncapsulator1%
;  effacer la ligne suivante quand je serai certain que ça ne servait à rien
; GuiControl, 1:, strFieldEncapsulator1, %strFieldEncapsulator1%
return



ButtonHelpEncapsulator1:
Help("Field Encapsulator", "When data fields in a CSV file contain characters used as delimiter or end-of-line, they must be enclosed in a field encapsulator. This encapsulator is often double-quotes ( ""..."" ) or single quotes ( '...' ). For example, if comma is used as field delimiter in a CSV file, the data field ""Smith, John"" must be encapsulated because it contains a comma.`n`nIf a field contains a character used as encapsulator, this character must be doubled. For example, the data ""John ""Junior"" Smith"" must be stated as ""John """"Junior"""" Smith"").`n`nUse the ""Preview"" button to find what is the field encapsulator in this file.")
return



ButtonHelpMultiline1:
Help("Multi-line Fields", "Most CSV files do not contain line breaks inside text field. But some do. For example, you can find multi-lines ""Notes"" fields in Google or Outlook contacts exported files.`n`nFor safety, this option is already checked (ON). If you are sure text fields in your CSV file do NOT contain line breaks, unselect this checkbox to turn this option OFF. This will improve loading performance.")
return



ButtonLoadFile:
Gui, 1:+OwnDialogs
Gui, 1:Submit, NoHide
if !StrLen(strFileToLoad)
{
	MsgBox, 48, %strApplicationName%, First use the "Select" button to choose the CSV file you want to load.
	return
}
if !StrLen(strFileHeaderEscaped) and (radSetHeader)
{
	MsgBox, 52, %strApplicationName%, CSV Header is not specified. Numbers will be used as field names. Do you want to continue?
	IfMsgBox, No
		return
}
if LV_GetCount("Column")
{
	MsgBox, 36, %strApplicationName%, Replace the current content of the list?
	IfMsgBox, Yes
	{
		LV_Delete() ; delete all rows - better performance on large files when we delete rows before columns
		loop, % LV_GetCount("Column")
			LV_DeleteCol(1) ; delete all columns
	}
	IfMsgBox, No
	{
		MsgBox, 36, %strApplicationName%, If the CSV file you want to load have the same fields, in the same order, you can add file data to the current list.`n`nDo you want to add to the content of this file to the list?
		IfMsgBox, No
			return
	}
}
strCurrentHeader := StrUnEscape(strFileHeaderEscaped)
; ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames [, blnHeader = true, blnMultiline = 1, blnProgress = 0, strFieldDelimiter = ",", strFieldEncapsulator = """", strRecordDelimiter = "`n", strOmitChars = "`r"])
obj := ObjCSV_CSV2Collection(strFileToLoad, strCurrentHeader, radGetHeader, blnMultiline1, 1, StrConvertFieldDelimiter(strFieldDelimiter1), strFieldEncapsulator1)
; ObjCSV_Collection2ListView(objCollection [, strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", strSortFields = "", strSortOptions = "", blnProgress = "0"])
ObjCSV_Collection2ListView(obj, "1", "lvData", strCurrentHeader, StrConvertFieldDelimiter(strFieldDelimiter1), strFieldEncapsulator1, , , 1)
if !LV_GetCount()
{
	MsgBox, 16, %strApplicationName%, CSV file not loaded.`n`nNote that %strApplicationName% support files with a maximum of 200 fields.
	return
}
else
{
	gosub UpdateCurrentHeader
	Help("Ready to edit","Your CSV file is loaded. You can review its content and sort rows by clicking on column headers (note that numeric sorting is not available yet).`n`nYou can use the ""2) Edit Columns"" tab to edit field names, select fields to keep or change fields order.`n`nWhen you will be ready, change to the ""3) Save CSV File"" tab to save all or selected rows in a new CSV file.")
}
obj := ; release object
return



; --------------------- TAB 2 --------------------------


ButtonSetRename:
Gui, 1:Submit, NoHide
if !LV_GetCount()
{
	MsgBox, 48, %strApplicationName%, First load a CSV file in the first tab.
	return
}
if !StrLen(strRenameEscaped)
{
	MsgBox, 52, %strApplicationName%, In "Rename fields:", enter the list of field names separated by the field delimiter ( %strFieldDelimiter1% ). Field names are automatically filled when you load a CSV file in the first tab.`n`nIf no field names are provided, numbers will be used as field names. Do you want to use numbers as field names? 
	IfMsgBox, No
		return
}
objNewHeader := ReturnDSVObjectArray(StrUnEscape(strRenameEscaped), StrConvertFieldDelimiter(strFieldDelimiter1), strFieldEncapsulator1)
Loop, % LV_GetCount("Column")
{
	if StrLen(objNewHeader[A_Index])
		LV_ModifyCol(A_Index, "", objNewHeader[A_Index])
	else
		LV_ModifyCol(A_Index, "", A_Index)
	LV_ModifyCol(A_Index, "AutoHdr")
}
gosub UpdateCurrentHeader
objNewHeader := ; release object
return



ButtonHelpRename:
Gui, 1:Submit, NoHide
Help("Rename Fields", "To change field names (column headers), enter a new name for each fields, in the order they actually appear in the list, separated by the field delimiter ( " . strFieldDelimiter1 . " ) and click ""Rename"".`n`nIf you enter less names than the number of fields (or no field name at all), numbers will be used as field names for remaining columns.`n`nField names including the separator character ( " . strFieldDelimiter1 . " ) must be enclosed by the encapsulator character ( " . strFieldEncapsulator1 . " ).`n`nTo save the file, click on the last tab ""3) Save CSV File"".")
return



ButtonSetSelect:
Gui, 1:Submit, NoHide
if !StrLen(strSelectEscaped)
{
	MsgBox, 48, %strApplicationName%, First enter the names of the fields you want to keep in the list, separated by the field delimiter ( %strFieldDelimiter1% ), keeping their current order.`n`nField names are automatically filled when you load a CSV file in the fisrt tab.
	return
}
if !LV_GetCount()
{
	MsgBox, 48, %strApplicationName%, First load a CSV file in the fisrt tab.
	return
}
strDelimiter := StrConvertFieldDelimiter(strFieldDelimiter1)
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strDelimiter, strFieldEncapsulator1)
objNewHeader := ReturnDSVObjectArray(StrUnEscape(strSelectEscaped), strDelimiter, strFieldEncapsulator1)
intMaxCurrent := objCurrentHeader.MaxIndex()
intMaxNew := objNewHeader.MaxIndex()
intIndexCurrent := 1
intIndexNew := 1
intDeleted := 0
Loop
{
	if (objCurrentHeader[intIndexCurrent] = objNewHeader[intIndexNew])
	{
		; ##(strCurrentHeader . "`n" . strNewHeader . "`n`n" . intIndexCurrent . " / " . intIndexNew . "`n`nobjCurrentHeader[intIndexCurrent] " . objCurrentHeader[intIndexCurrent] . " = objNewHeader[intIndexNew] " . objNewHeader[intIndexNew] . "`n`nConserver col")
		intIndexCurrent := intIndexCurrent + 1
		intIndexNew := intIndexNew + 1
	}
	else
	{
		; ##(strCurrentHeader . "`n" . strNewHeader . "`n`n" . intIndexCurrent . " / " . intIndexNew . "`n`nobjCurrentHeader[intIndexCurrent] " . objCurrentHeader[intIndexCurrent] . " <> objNewHeader[intIndexNew] " . objNewHeader[intIndexNew] . "`n`nSUPPRIMER col")
		LV_DeleteCol(intIndexCurrent - intDeleted)
		intDeleted := intDeleted + 1
		intIndexCurrent := intIndexCurrent + 1
	}
	if (intIndexCurrent > intMaxCurrent)
		break
}
gosub UpdateCurrentHeader
objCurrentHeader := ; release object
objNewHeader := ; release object
return



ButtonHelpSelect:
Gui, 1:Submit, NoHide
Help("Select Fields", "To remove fields (columns) from the list, enter the name of fields you want to keep, in the order they actually appear in the list, separated by the field delimiter ( " . strFieldDelimiter1 . " ) and click ""Select"".`n`nField names including the separator character ( " . strFieldDelimiter1 . " ) must be enclosed by the encapsulator character ( " . strFieldEncapsulator1 . " ).`n`nTo save the file, click on the last tab ""3) Save CSV File"".")
return



ButtonSetOrder:
/*
ObjCSV_Collection2ListView(objCollection, strGuiID := "", strListViewID := "", strFieldOrder := "", strFieldDelimiter := ",", strEncapsulator := """", strSortFields := "", strSortOptions := "", blnProgress := "0")
ObjCSV_ListView2Collection(strGuiID := "", strListViewID := "", strFieldOrder := "", strFieldDelimiter := ",", strFieldEncapsulator := """", blnProgress := "0")
*/
Gui, 1:Submit, NoHide
if !StrLen(strSelectEscaped)
{
	MsgBox, 48, %strApplicationName%, First enter the names of the fields you want to keep in the list, in the desired order, separated by the field delimiter ( %strFieldDelimiter1% ).`n`nField names are automatically filled when you load a CSV file in the fisrt tab.
	return
}
if !LV_GetCount()
{
	MsgBox, 48, %strApplicationName%, First load a CSV file in the fisrt tab.
	return
}
objNewCollection := ObjCSV_ListView2Collection("1", "lvData", StrUnEscape(strOrderEscaped), StrConvertFieldDelimiter(strFieldDelimiter1), strFieldEncapsulator1, 1)
LV_Delete() ;  better performance on large files when we delete rows before columns
loop, % LV_GetCount("Column")
	LV_DeleteCol(1) ; delete all rows
ObjCSV_Collection2ListView(objNewCollection, "1", "lvData", StrUnEscape(strOrderEscaped), StrConvertFieldDelimiter(strFieldDelimiter1), strFieldEncapsulator1, , , 1)
gosub UpdateCurrentHeader
objNewCollection := ; release object
return



ButtonHelpOrder:
Gui, 1:Submit, NoHide
Help("Order Fields", "To change the order of fields (columns) in the list, enter the name of fields in the new order you want to apply, separated by the field delimiter ( " . strFieldDelimiter1 . " ) and click ""Order"".`n`nField names including the separator character ( " . strFieldDelimiter1 . " ) must be enclosed by the encapsulator character ( " . strFieldEncapsulator1 . " ).`n`nIf you enter less fields than in the original header, fields not included in the new order will be removed from the list. However, if you only want to remove fields from the list (without changing the order), the ""Select"" button gives better performance on large files.`n`nTo save the file, click on the last tab ""3) Save CSV File"" and select the destination file.")
return



; --------------------- TAB 3 --------------------------


ButtonHelpFileToSave:
Help("CSV File To Save", "Enter the name of the CSV file destination file (the current program's directory will be used if an absolute path isn't specified) or hit ""Select"" to choose the CSV destination file. When other options are OK, hit ""Save"" to save all or selected rows to the CSV file.`n`nNote that all rows are saved by default. You can select one row (using Click), a series of adjacent rows (using Shift-Click) or non contiguous rows (using Ctrl-Click or Shift-Ctrl-Click).`n`nNote that fields will be saved in the order they appear in the list.")
return



ButtonSelectFileToSave:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
FileSelectFile, strOutputFile, 2, %A_ScriptDir%, Select CSV File to save
if !(StrLen(strOutputFile))
	return
GuiControl, 1:, strFileToSave, %strOutputFile%
GuiControl, 1:+Default, btnSaveFile
GuiControl, 1:Focus, btnSaveFile
return



ChangedFileToSave:
Gui, 1:Submit, NoHide
if FileExist(strFileToSave)
	GuiControl, 1:Show, btnCheckFile
else
	GuiControl, 1:Hide, btnCheckFile
return



ButtonHelpFieldDelimiter3:
Help("Field Delimiter", "Each field in the CSV header or in data rows of the file must be separated by a field delimiter. Enter the field delimiter character to use in the saved file.`n`nIt can be comma ( , ), semicolon ( `; ), Tab or any single character.`n`nFor the following special characters, use the letters on the left:`n`nt`tTab (HT)`nn`tLinefeed (LF)`nr`tCarriage return (CR)`nf`tFormfeed (FF)")
return



ButtonHelpEncapsulator3:
Help("Field Encapsulator", "When data fields in a CSV file contain characters used as delimiter or end-of-line, they must be enclosed in a field encapsulator. Enter the field encapsulator character to use in the saved file.`n`nThe encapsulator is often double-quotes ( ""..."" ) or single quotes ( '...' ). For example, if comma is used as field delimiter in the saved CSV file, the data field ""Smith, John"" will be encapsulated because it contains a comma.`n`nIf a field contains the character used as encapsulator, this character will be doubled. For example, the data ""John ""Junior"" Smith"" will be entered as ""John """"Junior"""" Smith"").")
return



ButtonHelpSaveHeader:
Gui, 1:Submit, NoHide
Help("CSV Get/Set CSV Header", "To save the field names as the first line of the CSV file, select ""Save with CSV header"".`n`nIf you select ""Save without CSV header"", the first line of the file will contain the data of the first row to save.")
return



ClickRadSaveMultiline:
Gui, 1:Submit, NoHide
GuiControl, 1:Hide, lblEndoflineReplacement
GuiControl, 1:Hide, strEndoflineReplacement
return



ClickRadSaveSingleline:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, lblEndoflineReplacement
GuiControl, 1:Show, strEndoflineReplacement
return



ButtonHelpSaveMultiline:
Gui, 1:Submit, NoHide
Help("Saving multi-line fields", "If text fields contain line breaks, you can decide if line breaks will be saved as is or replaced with a character (or a sequence of characters) in order to keep these fields on a single line.`n`nIf you select ""Save multi-line"", line breaks are saved unchanged.`n`nIf you select ""Save single-line"", enter the replacement sequence for line breaks in the ""End-of-line replacement:"" zone. By default, the replacement character is """ . chr(182) . """ (ASCII code 182).")
return



ButtonSaveFile:
Gui, 1:Submit, NoHide
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = "0"])
if FileExist(strFileToSave)
{
	MsgBox, 35, %strApplicationName% - File exists, File exists:`n%strFileToSave%`n`nDo you want to overwrite this file?`n`nYes: The file will be overwritten.`nNo: Data will be added to the existing file.
	IfMsgBox, Yes
		blnOverwrite := True
	IfMsgBox, No
		blnOverwrite := False
	IfMsgBox, Cancel
		return
}
if (LV_GetCount("Selected") = 1)
{
	MsgBox, 35, %strApplicationName% - One record selected, Only one record is selected. Do you want to save only this record?`n`nYes: Only one record will be saved.`nNo: All records will be saved.
	IfMsgBox, No
		LV_Modify(0, "Select") ; select all records
	IfMsgBox, Cancel
		return
}
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
if (radSaveMultiline)
	strEol := ""
else
	strEol := strEndoflineReplacement
; ObjCSV_Collection2CSV(objCollection, strFilePath [, blnHeader = 0, strFieldOrder = "", blnProgress = 0, blnOverwrite = 0, strFieldDelimiter = ",", strEncapsulator = """", strEndOfLine = "`n", strEolReplacement = ""])
ObjCSV_Collection2CSV(obj, strFileToSave, radSaveWithHeader, GetHeader(strFieldDelimiter3, strFieldEncapsulator3), 1, blnOverwrite, StrConvertFieldDelimiter(strFieldDelimiter3), strFieldEncapsulator3, , strEol)
if FileExist(strFileToSave)
{
	GuiControl, 1:Show, btnCheckFile
	GuiControl, 1:+Default, btnCheckFile
	GuiControl, 1:Focus, btnCheckFile
}
obj := ; release object
return



ButtonCheckFile:
Gui, 1:Submit, NoHide
if !StrLen(strFileToSave)
{
	MsgBox, 48, %strApplicationName%, First use the "Select" button to choose the CSV file you want to save to.
	return
}
if !FileExist(strFileToSave)
{
	MsgBox, 48, %strApplicationName%, File does not exist yet.
	return
}
run notepad.exe %strFileToSave%
return






; --------------------- GUI2 LISTVIEW EVENTS --------------------------



ListViewEvents:
if (A_GuiEvent = "DoubleClick") or (A_GuiEvent = "R")
{
	intRowNumber := A_EventInfo
	Gui, 1:Submit, NoHide
	intGui1WinID := WinExist("A")
	Gui, 2:New, +Resize , %strApplicationName% - Edit row
	Gui, 2:+Owner1
	Gui, 1:Default
	SysGet, intMonWork, MonitorWorkArea 
	; ###(intMonWorkRight)
	intColWidth := 380
	intEditWidth := intColWidth - 20
	intMaxNbCol := Floor(intMonWorkRight / intColWidth)
	; MsgBox, WorkArea: %intMonWorkRight% x %intMonWorkBottom%`nColonne: %intNbCol% de %intColWidth%
	intX := 10
	intY := 5
	intCol := 1
	loop, % LV_GetCount("Column")
	{
		if ((intY + 100) > intMonWorkBottom)
		{
			if (intCol = 1)
			{
				Gui, 2:Add, Button, y%intY% x10 vbtnSaveRecord gButtonSaveRecord, Save
				Gui, 2:Add, Button, yp x+5 vbtnCancel gButtonCancel, Cancel
			}
			if (intCol = intMaxNbCol)
			{
				intYLabel := intY
				Gui, 2:Add, Text, y%intYLabel% x%intX% vstrLabelMissing, *** FIELDS MISSING ***
				break
			}
			intCol := intCol + 1
			intX := intX + intColWidth
			intY := 5
			; ###("intCol: " . intCol . " / intMaxNbCol: " . intMaxNbCol)
		}
		intYLabel := intY
		; ###(intYLabel)
		intYEdit := intY + 15
		; ###(A_Index . " " . intYLabel . " " . intYEdit . " " . intMonWorkBottom)
		LV_GetText(strColHeader, 0, A_Index)
		LV_GetText(strColData, intRowNumber, A_Index)
		Gui, 2:Add, Text, y%intYLabel% x%intX% vstrLabel%A_Index%, %strColHeader%
		Gui, 2:Add, Edit, y%intYEdit% x%intX% w%intEditWidth% vstrEdit%A_Index% +HwndstrEditHandle, %strColData%
		ShrinkEditControl(strEditHandle, 2, "2")
		GuiControlGet, intPosEdit, 2:Pos, %strEditHandle%
		; ###("!" . intPosEditH)
		intY := intY + intPosEditH + 19
		intNbFieldsOnScreen := A_Index ; incremented at each occurence of the loop
	}
	if (intCol = 1)
	{
		Gui, 2:Add, Button, y%intY% x10 vbtnSaveRecord gButtonSaveRecord, Save
		Gui, 2:Add, Button, yp x+5 vbtnCancel gButtonCancel, Cancel
	}
	Gui, 2:Show, AutoSize Center
	Gui, 1:+Disabled
}
return



ButtonSaveRecord:
if (intRowNumber < 1)
	###("Pas normal! intRowNumber: " . intRowNumber)
Gui, 2:Submit
Gui, 1:Default
loop, % LV_GetCount("Column")
	LV_Modify(intRowNumber, "Col" . A_Index, strEdit%A_Index%)
Goto, 2GuiClose
return


ButtonCancel:
Goto, 2GuiClose
return


2GuiClose:
2GuiEscape:
Gui, 1:-Disabled
Gui, 2:Destroy
WinActivate, ahk_id %intGui1WinID%
return



GuiSize: ; Expand or shrink the ListView in response to the user's resizing of the window.
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
; Otherwise, the window has been resized or maximized. Resize the controls to match.
GuiControl, 1:Move, tabCSVBuddy, % "W" . (A_GuiWidth - 20)

GuiControl, 1:Move, strFileToLoad, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpFileToLoad, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnSelectFileToLoad, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, strFileHeaderEscaped, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpHeader, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnPreviewFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnLoadFile, % "X" . (A_GuiWidth - 65)

GuiControl, 1:Move, strRenameEscaped, % "W" . (A_GuiWidth - 205)
GuiControl, 1:Move, btnSetRename, % "X" . (A_GuiWidth - 95)
GuiControl, 1:Move, btnHelpRename, % "X" . (A_GuiWidth - 40)
GuiControl, 1:Move, strSelectEscaped, % "W" . (A_GuiWidth - 205)
GuiControl, 1:Move, btnSetSelect, % "X" . (A_GuiWidth - 95)
GuiControl, 1:Move, btnHelpSelect, % "X" . (A_GuiWidth - 40)
GuiControl, 1:Move, strOrderEscaped, % "W" . (A_GuiWidth - 205)
GuiControl, 1:Move, btnSetOrder, % "X" . (A_GuiWidth - 95)
GuiControl, 1:Move, btnHelpOrder, % "X" . (A_GuiWidth - 40)

GuiControl, 1:Move, strFileToSave, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpFileToSave, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnSelectFileToSave, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnSaveFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnCheckFile, % "X" . (A_GuiWidth - 65)

GuiControl, 1:Move, lvData, % "W" . (A_GuiWidth - 20) . " H" . (A_GuiHeight - 190)

return





2GuiSize:  ; Expand or shrink the ListView in response to the user's resizing of the window.
; ###("intNbFieldsOnScreen: " . intNbFieldsOnScreen . " / intWidthSize: " . intWidthSize)
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
; MsgBox, A_GuiWidth: %A_GuiWidth% / intCol: %intCol%
GuiControl, 2:Move, btnSaveRecord, % "X" . (A_GuiWidth - 100)
GuiControl, 2:Move, btnCancel, % "X" . (A_GuiWidth - 50)
if intCol > 1  ; The window has been minimized.  No action needed.
    return
intWidthSize := A_GuiWidth - 20
Loop, %intNbFieldsOnScreen%
{
 	GuiControl, 2:Move, strEdit%A_Index%, % "W" . intWidthSize
}
return



; --------------------- OTHER COMMANDS --------------------------



1GuiQuit:
ExitApp



UpdateCurrentHeader:
Gui, 1:Submit, NoHide
strCurrentHeader := GetHeader(strFieldDelimiter1, strFieldEncapsulator1)
strEscapedCurrentHeader := StrEscape(strCurrentHeader)
GuiControl, 1:, strFileHeaderEscaped, %strEscapedCurrentHeader%
GuiControl, 1:, strRenameEscaped, %strEscapedCurrentHeader%
GuiControl, 1:, strSelectEscaped, %strEscapedCurrentHeader%
GuiControl, 1:, strOrderEscaped, %strEscapedCurrentHeader%
return



GetHeader(strFieldDelimiter, strFieldEncapsulator)
{
	strDelimiter := StrConvertFieldDelimiter(strFieldDelimiter)
	strHeader := ""
	Loop, % LV_GetCount("Column")
	{
		LV_GetText(strColumnHeader, 0, A_Index)
		strHeader := strHeader . Format4CSV(strColumnHeader, strDelimiter, strFieldEncapsulator) . strDelimiter
	}
	StringTrimRight, strHeader, strHeader, 1 ; remove extra delimiter
	return strHeader
}



/*
SetSelectColumn(strCurrentHeader, strNewHeader, strDelimiter, strEncapsulator)
{
	objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strDelimiter, strEncapsulator)
	objNewHeader := ReturnDSVObjectArray(strNewHeader, strDelimiter, strEncapsulator)
	intMaxCurrent := objCurrentHeader.MaxIndex()
	intMaxNew := objNewHeader.MaxIndex()
	intIndexCurrent := 1
	intIndexNew := 1
	intDeleted := 0
	Loop
	{
		if (objCurrentHeader[intIndexCurrent] = objNewHeader[intIndexNew])
		{
			; ###(strCurrentHeader . "`n" . strNewHeader . "`n`n" . intIndexCurrent . " / " . intIndexNew . "`n`nobjCurrentHeader[intIndexCurrent] " . objCurrentHeader[intIndexCurrent] . " = objNewHeader[intIndexNew] " . objNewHeader[intIndexNew] . "`n`nConserver col")
			intIndexCurrent := intIndexCurrent + 1
			intIndexNew := intIndexNew + 1
		}
		else
		{
			; ###(strCurrentHeader . "`n" . strNewHeader . "`n`n" . intIndexCurrent . " / " . intIndexNew . "`n`nobjCurrentHeader[intIndexCurrent] " . objCurrentHeader[intIndexCurrent] . " <> objNewHeader[intIndexNew] " . objNewHeader[intIndexNew] . "`n`nSUPPRIMER col")
			LV_DeleteCol(intIndexCurrent - intDeleted)
			intDeleted := intDeleted + 1
			intIndexCurrent := intIndexCurrent + 1
		}
		if (intIndexCurrent > intMaxCurrent)
			break
	}
}
*/



StrEscape(strEscaped)
{
	StringReplace, strEscaped, strEscaped, ``, £¡ƒ†, All ; temporary string replacement for escape char
	StringReplace, strEscaped, strEscaped, `t, ``t, All ; Tab (HT)
	StringReplace, strEscaped, strEscaped, `n, ``n, All ; Linefeed (LF)
	StringReplace, strEscaped, strEscaped, `r, ``r, All ; Carriage return (CR)
	StringReplace, strEscaped, strEscaped, `f, ``f, All ; Form feed (FF)
	StringReplace, strEscaped, strEscaped, £¡ƒ†, ````, All ; from temporary string to escape char
	return strEscaped
}



StrUnEscape(strUnEscaped)
{
	StringReplace, strUnEscaped, strUnEscaped, ````, £¡ƒ†, All ; temporary string replacement for escape char
	StringReplace, strUnEscaped, strUnEscaped, ``t, `t, All ; Tab (HT)
	StringReplace, strUnEscaped, strUnEscaped, ``n, `n, All ; Linefeed (LF)
	StringReplace, strUnEscaped, strUnEscaped, ``r, `r, All ; Carriage return (CR)
	StringReplace, strUnEscaped, strUnEscaped, ``f, `f, All ; Form feed (FF)
	StringReplace, strUnEscaped, strUnEscaped, £¡ƒ†, ``, All ; from temporary string to escape char
	return strUnEscaped
}



StrConvertFieldDelimiter(strConverted)
{
	StringReplace, strConverted, strConverted, t, `t, All ; Tab (HT)
	StringReplace, strConverted, strConverted, n, `n, All ; Linefeed (LF)
	StringReplace, strConverted, strConverted, r, `r, All ; Carriage return (CR)
	StringReplace, strConverted, strConverted, f, `f, All ; Form feed (FF)
	return strConverted
}



Help(strTitle, strMessage)
{
	Gui, 1:+OwnDialogs 
	MsgBox, 0, %strApplicationName% (v%strApplicationVersion%) - %strTitle% Help,%strMessage%
}



ShrinkEditControl(strEditHandle, intMaxRows, strGuiName)
{
	EM_GETLINECOUNT = 0xBA
	SendMessage, %EM_GETLINECOUNT%,,,, AHk_id %strEditHandle%
	intNbRows := ErrorLevel
	if (intNbRows > intMaxRows)
	{
		GuiControlGet, intPosEdit, %strGuiName%:Pos, %strEditHandle%
		intEditMargin := 8 ; top & bottom margin of the Edit control (regardless of the nb of rows)
		intOriginalHeight := intPosEditH
		intHeightOneRow := Round((intOriginalHeight - intEditMargin) / intNbRows)
		intNewHeight := (intHeightOneRow * intMaxRows) + intEditMargin
		; MsgBox, % "intNbRows: " . intNbRows . "`nintOriginalHeight: " . intOriginalHeight . "`nintHeightOneRow: " . intHeightOneRow . "`nintNewHeight: " . intNewHeight
		GuiControl, %strGuiName%:Move, %strEditHandle%, h%intNewHeight%
	}
}
