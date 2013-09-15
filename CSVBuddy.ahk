;===============================================
/*
CSV Buddy
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By JnLlnd on AHK forum
This script uses the library ObjCSV v0.2 (https://github.com/JnLlnd/ObjCSV)
*/ 

#NoEnv
#SingleInstance force
#Include %A_ScriptDir%\..\ObjCSV\lib\ObjCSV.ahk

; --------------------- GLOBAL AND DEFAULT VALUES --------------------------

global strApplicationName := "CSV Buddy"
global strApplicationVersion := "v0.2.1 ALPHA" ; (avec pull de v0.1.2 ALPHA)

intDefaultWidth := 16 ; used when export to fixed width format
strTemplateDelimiter := "¤" ; Chr(164)


; --------------------- GUI1 --------------------------


Gui, 1:New, +Resize, %strApplicationName%

Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Text, x10, %strApplicationName%

Gui, 1:Font, s10 w700, Verdana
Gui, 1:Add, Tab2, w950 r4 vtabCSVBuddy gChangedTabCSVBuddy, % " 1) Load CSV File     ||     2) Edit Columns     |     3) Save CSV File     |     4) Export     |     About     "
Gui, 1:Font

Gui, 1:Tab, 1
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToLoad w85 right, CSV &file to load:
Gui, 1:Add, Edit,		yp		x100	vstrFileToLoad disabled gChangedFileToLoad
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFileToLoad gButtonHelpFileToLoad, ?
Gui, 1:Add, Button,		yp		x+5		vbtnSelectFileToLoad gButtonSelectFileToLoad default, &Select
Gui, 1:Add, Text,		y+10	x10 	vlblHeader w85 right, CSV file &Header:
Gui, 1:Add, Edit,		yp		x100	vstrFileHeaderEscaped disabled
Gui, 1:Add, Button,		yp		x+5		vbtnHelpHeader gButtonHelpHeader, ?
Gui, 1:Add, Button,		yp		x+5		vbtnPreviewFile gButtonPreviewFile hidden, &Preview
Gui, 1:Add, Radio,		y+10	x100	vradGetHeader gClickRadGetHeader checked, &Get header from file
Gui, 1:Add, Radio,		yp		x+5		vradSetHeader gClickRadSetHeader, Set header
Gui, 1:Add, Button,		yp		x+0		vbtnHelpSetHeader gButtonHelpSetHeader, ?
Gui, 1:Add, Text,		xp		x+27	vlblFieldDelimiter1, Field &delimiter:
Gui, 1:Add, Edit,		yp		x+5		vstrFieldDelimiter1 gChangedFieldDelimiter1 w20 limit1 center, `, 
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFieldDelimiter1 gButtonHelpFieldDelimiter1, ?
Gui, 1:Add, Text,		yp		x+27	vlblFieldEncapsulator1, Field e&ncapsulator:
Gui, 1:Add, Edit,		yp		x+5		vstrFieldEncapsulator1 gChangedFieldEncapsulator1 w20 limit1 center, `"
Gui, 1:Add, Button,		yp		x+5		vbtnHelpEncapsulator1 gButtonHelpEncapsulator1, ?
Gui, 1:Add, Checkbox,	yp		x+27	vblnMultiline1 gChangedMultiline1, &Multi-line fields
Gui, 1:Add, Button,		yp		x+0		vbtnHelpMultiline1 gButtonHelpMultiline1, ?
Gui, 1:Add, Text,		yp		x+5		vlblEndoflineReplacement1 hidden, EOL replacement:
Gui, 1:Add, Edit,		yp		x+5		vstrEndoflineReplacement1 w30 center hidden
Gui, 1:Add, Button,		yp		x+5		vbtnLoadFile gButtonLoadFile hidden, &Load

Gui, 1:Tab, 2
Gui, 1:Add, Text,		y+10	x10		vlblRenameFields w85 right, Ren&ame fields:
Gui, 1:Add, Edit,		yp		x100	vstrRenameEscaped
Gui, 1:Add, Button,		yp		x+0		vbtnSetRename gButtonSetRename, &Rename
Gui, 1:Add, Button,		yp		x+5		vbtnHelpRename gButtonHelpRename, ?
Gui, 1:Add, Text,		y+10	x10		vlblSelectFields w85 right, Selec&t fields:
Gui, 1:Add, Edit,		yp		x100	vstrSelectEscaped
Gui, 1:Add, Button,		yp		x+0		vbtnSetSelect gButtonSetSelect, S&elect
Gui, 1:Add, Button,		yp		x+5		vbtnHelpSelect gButtonHelpSelect, ?
Gui, 1:Add, Text,		y+10	x10		vlblOrderFields w85 right, Order f&ields:
Gui, 1:Add, Edit,		yp		x100	vstrOrderEscaped
Gui, 1:Add, Button,		yp		x+0		vbtnSetOrder gButtonSetOrder, &Order
Gui, 1:Add, Button,		yp		x+5		vbtnHelpOrder gButtonHelpOrder, ?

Gui, 1:Tab, 3
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToSave w85 right, CS&V file to save:
Gui, 1:Add, Edit,		yp		x100	vstrFileToSave gChangedFileToSave
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFileToSave gButtonHelpFileToSave, ?
Gui, 1:Add, Button,		yp		x+5		vbtnSelectFileToSave gButtonSelectFileToSave default, &Select
Gui, 1:Add, Text,		y+10	x100	vlblFieldDelimiter3, Field delimiter:
Gui, 1:Add, Edit,		yp		x200	vstrFieldDelimiter3 gChangedFieldDelimiter3 w20 limit1 center, `, 
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFieldDelimiter3 gButtonHelpFieldDelimiter3, ?
Gui, 1:Add, Text,		y+10	x100	vlblFieldEncapsulator3, Field encaps&ulator:
Gui, 1:Add, Edit,		yp		x200	vstrFieldEncapsulator3 gChangedFieldEncapsulator3 w20 limit1 center, `"
Gui, 1:Add, Button,		yp		x+5		vbtnHelpEncapsulator3 gButtonHelpEncapsulator3, ?
Gui, 1:Add, Radio,		y100	x300	vradSaveWithHeader checked, Save &with CSV header
Gui, 1:Add, Radio,		y+10	x300	vradSaveNoHeader, Save without CSV header
Gui, 1:Add, Button,		y100	x450	vbtnHelpSaveHeader gButtonHelpSaveHeader, ?
Gui, 1:Add, Radio,		y100	x500	vradSaveMultiline gClickRadSaveMultiline checked, Save multi-line
Gui, 1:Add, Radio,		y+10	x500	vradSaveSingleline gClickRadSaveSingleline, Save single-line
Gui, 1:Add, Button,		y100	x620	vbtnHelpMultiline gButtonHelpSaveMultiline, ?
Gui, 1:Add, Text,		y+25	x500	vlblEndoflineReplacement3 hidden, End-of-line replacement:
Gui, 1:Add, Edit,		yp		x620	vstrEndoflineReplacement3 hidden w50 center, % chr(182)
Gui, 1:Add, Button,		y105	x+5		vbtnSaveFile gButtonSaveFile hidden, Save
Gui, 1:Add, Button,		y137	x+5		vbtnCheckFile hidden gButtonCheckFile, Check

Gui, 1:Tab, 4
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToExport w85 right, Export data to file:
Gui, 1:Add, Edit,		yp		x100	vstrFileToExport gChangedFileToExport
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFileToExport gButtonHelpFileToExport, ?
Gui, 1:Add, Button,		yp		x+5		vbtnSelectFileToExport gButtonSelectFileToExport default, &Select
Gui, 1:Add, Text,		y+10	x10		vlblCSVExportFormat w85 right, Export format:
Gui, 1:Add, Radio,		yp		x100	vradFixed gClickRadFixed, Fixed width
Gui, 1:Add, Radio,		yp		x+15	vradHTML gClickRadHTML, HTML
Gui, 1:Add, Radio,		yp		x+15	vradXML gClickRadXML, XML
Gui, 1:Add, Radio,		yp		x+15	vradExpress gClickRadExpress, Express
Gui, 1:Add, Button,		yp		x+15	vbtnHelpExportFormat gButtonHelpExportFormat, ?
Gui, 1:Add, Text,		y+10	x10		vlblMultiPurpose w85 right hidden, Hidden Label:
Gui, 1:Add, Edit,		yp		x100	vstrMultiPurpose gChangedMultiPurpose hidden
Gui, 1:Add, Button,		yp		x+5		vbtnMultiPurpose gButtonMultiPurpose hidden, Lorem ipsum dolor sitm ; conserver texte important pour la largeur du bouton
Gui, 1:Add, Button,		y105	x+5		vbtnExportFile gButtonExportFile hidden, Export
Gui, 1:Add, Button,		y137	x+5		vbtnCheckExportFile gButtonCheckExportFile hidden, Check

Gui, 1:Tab, 5
Gui, 1:Add, Link,		y+10	x10		vlblAboutText, <a href="https://bitbucket.org/JnLlnd/csvbuddy">%strApplicationName% %strApplicationVersion%</a>`nby Jean Lalonde (<a href="http://www.autohotkey.com/board/user/4880-jnllnd/">JnLlnd</a> on AHK forum)`nAll rights reserved (c)2013 - DO NOT DISTRIBUTE WITHOUT AUTHOR AUTORIZATION`n`nUsing AHK library: <a href="https://www.github.com/JnLlnd/ObjCSV">ObjCSV v0.2</a>`nUsing icon by: <a href="http://www.visualpharm.com">Visual Pharm</a>

Gui, 1:Tab

Gui, 1:Add, ListView, 	x10 r24 w200 vlvData -ReadOnly NoSort gListViewEvents -LV0x10

GuiControl, 1:Focus, btnSelectFileToLoad
GuiControl, 1:+Default, btnSelectFileToLoad
Gui, 1:Show, Autosize

return



ChangedTabCSVBuddy:
Gui, 1:Submit, NoHide
;  " 1) Load CSV File     ||     2) Edit Columns     |     3) Save CSV File     |     About     "
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
else if InStr(tabCSVBuddy, "Export")
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnSelectFileToExport
	else
	{
		MsgBox, 48, %strApplicationName%, First load a CSV file in the first tab.
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
else if InStr(tabCSVBuddy, "About")
{
	; do nothing
}
else
	###_D(tabCSVBuddy . " !?!")
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
GuiControl, 1:, strFileToSave, % NewFileName(strFileToLoad)
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
if FileExist(strFileToLoad)
{
	GuiControl, 1:Show, btnPreviewFile
	GuiControl, 1:Show, btnLoadFile
}
else
{
	GuiControl, 1:Hide, btnPreviewFile
	GuiControl, 1:Hide, btnLoadFile
}
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
GuiControl, 1:, lblHeader, File CSV &Header:
return



ClickRadSetHeader:
GuiControl, 1:Enable, strFileHeaderEscaped
GuiControl, 1:, lblHeader, Custom &Header:
GuiControl, 1:Focus, strFileHeaderEscaped
return



ButtonHelpSetHeader:
Gui, 1:Submit, NoHide
Help("CSV Get/Set CSV Header", "If the first line of the CSV file contains the list of data columns field names, click ""Get header from CSV file"". If not, click ""Set CSV header"" and enter the list of field names separated by the Field delimiter.")
return



ChangedFieldDelimiter1:
Gui, 1:Submit, NoHide
GuiControl, 1:, strFieldDelimiter3, %strFieldDelimiter1%
Gosub, UpdateCurrentHeader
return



ButtonHelpFieldDelimiter1:
Help("Field Delimiter", "Each field in the CSV header or in data rows of the file must be separated by a field delimiter. This is often comma ( , ), semicolon ( `; ) or Tab.`n`nEnter any single character such as comma, semicolon or one of these letters for special characters:`n`nt`tTab (HT)`nn`tLinefeed (LF)`nr`tCarriage return (CR)`nf`tFormfeed (FF)`n`nUse the ""Preview"" button to find what is the field delimiter in this file.")
return



ChangedFieldEncapsulator1:
Gui, 1:Submit, NoHide
GuiControl, 1:, strFieldEncapsulator3, %strFieldEncapsulator1%
Gosub, UpdateCurrentHeader
return



ButtonHelpEncapsulator1:
Help("Field Encapsulator", "When data fields in a CSV file contain characters used as delimiter or end-of-line, they must be enclosed in a field encapsulator. This encapsulator is often double-quotes ( ""..."" ) or single quotes ( '...' ). For example, if comma is used as field delimiter in a CSV file, the data field ""Smith, John"" must be encapsulated because it contains a comma.`n`nIf a field contains a character used as encapsulator, this character must be doubled. For example, the data ""John ""Junior"" Smith"" must be stated as ""John """"Junior"""" Smith"").`n`nUse the ""Preview"" button to find what is the field encapsulator in this file.")
return



ChangedMultiline1:
Gui, 1:Submit, NoHide
if (blnMultiline1)
{
	GuiControl, 1:Show, lblEndoflineReplacement1
	GuiControl, 1:Show, strEndoflineReplacement1
}
else
{
	GuiControl, 1:Hide, lblEndoflineReplacement1
	GuiControl, 1:Hide, strEndoflineReplacement1
}
return



ButtonHelpMultiline1:
Help("Multi-line Fields", "Most CSV files do not contain line breaks inside text field. But some do. For example, you can find multi-lines ""Notes"" fields in Google or Outlook contacts exported files.`n`nIf text fields in your CSV file contain line breaks, select this checkbox to turn this option ON. If not, keep it OFF since this will improve loading performance.`n`nIf you turn Multi-line ON, you have the additional option to choose a character (or string) that will be converted to line-breaks if found in the CSV file.")
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
strRealFieldDelimiter1 := StrMakeRealFieldDelimiter(strFieldDelimiter1)
; ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames [, blnHeader = 1, blnMultiline = 1, blnProgress = 0, strFieldDelimiter = ",", strEncapsulator = """", strRecordDelimiter = "`n", strOmitChars = "`r", strEolReplacement = ""])
obj := ObjCSV_CSV2Collection(strFileToLoad, strCurrentHeader, radGetHeader, blnMultiline1, 1, strRealFieldDelimiter1, strFieldEncapsulator1, , , strEndoflineReplacement1)
; ObjCSV_Collection2ListView(objCollection [, strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", strSortFields = "", strSortOptions = "", blnProgress = 0])
ObjCSV_Collection2ListView(obj, "1", "lvData", strCurrentHeader, strRealFieldDelimiter1, strFieldEncapsulator1, , , 1)
if !LV_GetCount()
{
	MsgBox, 16, %strApplicationName%, CSV file not loaded.`n`nNote that %strApplicationName% support files with a maximum of 200 fields.
	return
}
else
{
	gosub UpdateCurrentHeader
	Help("Ready to edit","Your CSV file is loaded.`n`nYou can sort rows by clicking on column headers. Choose sorting type: alphabetical, numeric integer or numeric float, ascending or descending.`n`nDouble-click on a row to edit a record.  Right-click anywhere in the list view to select all rows, deselect all rows or reverse selection.`n`nYou can use the ""2) Edit Columns"" tab to edit field names, select fields to keep or change fields order.`n`nWhen you will be ready, change to the ""3) Save CSV File"" tab to save all or selected rows in a new CSV file.")
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
	MsgBox, 52, %strApplicationName%, In "Rename fields:", enter the list of field names separated by the field delimiter ( %strFieldDelimiter1% ). Field names are automatically filled when you load a CSV file in the first tab.`n`nIf no field names are provided, numbers are used as field names. Do you want to use numbers as field names? 
	IfMsgBox, No
		return
}
objNewHeader := ReturnDSVObjectArray(StrUnEscape(strRenameEscaped), StrMakeRealFieldDelimiter(strFieldDelimiter1), strFieldEncapsulator1)
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
Help("Rename Fields", "To change field names (column headers), enter a new name for each fields, in the order they actually appear in the list, separated by the field delimiter ( " . strFieldDelimiter1 . " ) and click ""Rename"".`n`nIf you enter less names than the number of fields (or no field name at all), numbers are used as field names for remaining columns.`n`nField names including the separator character ( " . strFieldDelimiter1 . " ) must be enclosed by the encapsulator character ( " . strFieldEncapsulator1 . " ).`n`nTo save the file, click on the last tab ""3) Save CSV File"".")
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
strRealFieldDelimiter1 := StrMakeRealFieldDelimiter(strFieldDelimiter1)
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strRealFieldDelimiter1, strFieldEncapsulator1)
objNewHeader := ReturnDSVObjectArray(StrUnEscape(strSelectEscaped), strRealFieldDelimiter1, strFieldEncapsulator1)
intMaxCurrent := objCurrentHeader.MaxIndex()
intMaxNew := objNewHeader.MaxIndex()
intIndexCurrent := 1
intIndexNew := 1
intDeleted := 0
Loop
{
	if (objCurrentHeader[intIndexCurrent] = objNewHeader[intIndexNew])
	{
		intIndexCurrent := intIndexCurrent + 1
		intIndexNew := intIndexNew + 1
	}
	else
	{
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
ObjCSV_Collection2ListView(objCollection, strGuiID := "", strListViewID := "", strFieldOrder := "", strFieldDelimiter := ",", strEncapsulator := """", strSortFields := "", strSortOptions := "", blnProgress := 0)
ObjCSV_ListView2Collection(strGuiID := "", strListViewID := "", strFieldOrder := "", strFieldDelimiter := ",", strFieldEncapsulator := """", blnProgress := 0)
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
strRealFieldDelimiter1 := StrMakeRealFieldDelimiter(strFieldDelimiter1)
objNewCollection := ObjCSV_ListView2Collection("1", "lvData", StrUnEscape(strOrderEscaped), strRealFieldDelimiter1, strFieldEncapsulator1, 1)
LV_Delete() ;  better performance on large files when we delete rows before columns
loop, % LV_GetCount("Column")
	LV_DeleteCol(1) ; delete all rows
ObjCSV_Collection2ListView(objNewCollection, "1", "lvData", StrUnEscape(strOrderEscaped), strRealFieldDelimiter1, strFieldEncapsulator1, , , 1)
gosub UpdateCurrentHeader
objNewCollection := ; release object
return



ButtonHelpOrder:
Gui, 1:Submit, NoHide
Help("Order Fields", "To change the order of fields (columns) in the list, enter the name of fields in the new order you want to apply, separated by the field delimiter ( " . strFieldDelimiter1 . " ) and click ""Order"".`n`nField names including the separator character ( " . strFieldDelimiter1 . " ) must be enclosed by the encapsulator character ( " . strFieldEncapsulator1 . " ).`n`nIf you enter less fields than in the original header, fields not included in the new order are removed from the list. However, if you only want to remove fields from the list (without changing the order), the ""Select"" button gives better performance on large files.`n`nTo save the file, click on the last tab ""3) Save CSV File"" and select the destination file.")
return



; --------------------- TAB 3 --------------------------


ButtonHelpFileToSave:
Help("CSV File To Save", "Enter the name of the CSV file destination file (the current program's directory is used if an absolute path isn't specified) or hit ""Select"" to choose the CSV destination file. When other options are OK, hit ""Save"" to save all or selected rows to the CSV file.`n`nNote that all rows are saved by default. You can select one row (using Click), a series of adjacent rows (using Shift-Click) or non contiguous rows (using Ctrl-Click or Shift-Ctrl-Click).`n`nNote that fields are saved in the order they appear in the list and rows are saved according to the current sorting order (click on a column name to sort rows).")
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
SplitPath, strFileToSave, , strOutDir, strOutExtension, strOutNameNoExt
if FileExist(strFileToSave)
	GuiControl, 1:Show, btnCheckFile
else
	GuiControl, 1:Hide, btnCheckFile
if StrLen(strOutNameNoExt)
	GuiControl, 1:Show, btnSaveFile
else
	GuiControl, 1:Hide, btnSaveFile
return



ButtonHelpFieldDelimiter3:
Help("Field Delimiter", "Each field in the CSV header or in data rows of the file must be separated by a field delimiter. Enter the field delimiter character to use in the saved file.`n`nIt can be comma ( , ), semicolon ( `; ), Tab or any single character.`n`nFor the following special characters, use the letters on the left:`n`nt`tTab (HT)`nn`tLinefeed (LF)`nr`tCarriage return (CR)`nf`tFormfeed (FF)")
return



ChangedFieldDelimiter3:
Gui, 1:Submit, NoHide
Gosub, UpdateCurrentHeader
return



ChangedFieldEncapsulator3:
Gui, 1:Submit, NoHide
Gosub, UpdateCurrentHeader
return



ButtonHelpEncapsulator3:
Help("Field Encapsulator", "When data fields in a CSV file contain characters used as delimiter or end-of-line, they must be enclosed in a field encapsulator. Enter the field encapsulator character to use in the saved file.`n`nThe encapsulator is often double-quotes ( ""..."" ) or single quotes ( '...' ). For example, if comma is used as field delimiter in the saved CSV file, the data field ""Smith, John"" is encapsulated because it contains a comma.`n`nIf a field contains the character used as encapsulator, this character is doubled. For example, the data ""John ""Junior"" Smith"" will be entered as ""John """"Junior"""" Smith"").")
return



ButtonHelpSaveHeader:
Gui, 1:Submit, NoHide
Help("CSV Get/Set CSV Header", "To save the field names as the first line of the CSV file, select ""Save with CSV header"".`n`nIf you select ""Save without CSV header"", the first line of the file will contain the data of the first row to save.")
return



ClickRadSaveMultiline:
Gui, 1:Submit, NoHide
GuiControl, 1:Hide, lblEndoflineReplacement3
GuiControl, 1:Hide, strEndoflineReplacement3
return



ClickRadSaveSingleline:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, lblEndoflineReplacement3
GuiControl, 1:Show, strEndoflineReplacement3
return



ButtonHelpSaveMultiline:
Gui, 1:Submit, NoHide
Help("Saving multi-line fields", "If text fields contain line breaks, you can decide if line breaks will be saved as is or replaced with a character (or a sequence of characters) in order to keep these fields on a single line.`n`nIf you select ""Save multi-line"", line breaks are saved unchanged.`n`nIf you select ""Save single-line"", enter the replacement sequence for line breaks in the ""End-of-line replacement:"" zone. By default, the replacement character is """ . chr(182) . """ (ASCII code 182).")
return



ButtonSaveFile:
Gui, 1:Submit, NoHide
blnOverwrite := CheckIfFileExistOverwrite(strFileToSave)
if (blnOverwrite < 0)
	return
gosub, CheckOneRow
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = 0])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
if (radSaveMultiline)
	strEolReplacement := ""
else
	strEolReplacement := strEndoflineReplacement3
; ObjCSV_Collection2CSV(objCollection, strFilePath [, blnHeader = 0, strFieldOrder = "", blnProgress = 0, blnOverwrite = 0, strFieldDelimiter = ",", strEncapsulator = """", strEndOfLine = "`n", strEolReplacement = ""])
ObjCSV_Collection2CSV(obj, strFileToSave, radSaveWithHeader, GetHeader(strFieldDelimiter3, strFieldEncapsulator3), 1, blnOverwrite, StrMakeRealFieldDelimiter(strFieldDelimiter3), strFieldEncapsulator3, , strEolReplacement)
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
run notepad.exe %strFileToSave%
return



; --------------------- TAB 4 --------------------------

ButtonHelpFileToExport:
Help("Export data", "Enter the name of the destination file of the export (the current program's directory is used if an absolute path isn't specified) or hit ""Select"" to choose the destination file. When other options are OK, hit ""Export"" to export all or selected rows to the destination file.`n`nNote that all rows are saved by default. You can select one row (using Click), a series of adjacent rows (using Shift-Click) or non contiguous rows (using Ctrl-Click or Shift-Ctrl-Click).`n`nNote that fields are exported in the order they appear in the list and rows are saved according to the current sorting order (click on a column name to sort rows).")
return



ButtonSelectFileToExport:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
FileSelectFile, strOutputFile, 2, %A_ScriptDir%, Select export file
if !(StrLen(strOutputFile))
	return
GuiControl, 1:, strFileToExport, %strOutputFile%
GuiControl, 1:+Default, btnExportFile
GuiControl, 1:Focus, btnExportFile
return



ChangedFileToExport:
Gui, 1:Submit, NoHide
SplitPath, strFileToExport, , strOutDir, strOutExtension, strOutNameNoExt
if FileExist(strFileToExport)
	GuiControl, 1:Show, btnCheckExportFile
else
	GuiControl, 1:Hide, btnCheckExportFile
if StrLen(strOutNameNoExt)
	GuiControl, 1:Show, btnExportFile
else
	GuiControl, 1:Hide, btnExportFile
return



ClickRadFixed:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, Fields width:
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Show, btnMultiPurpose
GuiControl, 1:, btnMultiPurpose, Change default width
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
strRealFieldDelimiter1 := StrMakeRealFieldDelimiter(strFieldDelimiter1)
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strRealFieldDelimiter1, strFieldEncapsulator1)
; strFieldDelimiter1 et strFieldEncapsulator1 pour la lecture de strCurrentHeader
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
strMultiPurpose := ""
Loop, % objCurrentHeader.MaxIndex()
	strMultiPurpose := strMultiPurpose . Format4CSV(objCurrentHeader[A_Index], strRealFieldDelimiter3, strFieldEncapsulator3) . strRealFieldDelimiter3 . intDefaultWidth . strRealFieldDelimiter3
	; strFieldDelimiter3 et strFieldEncapsulator3 pour l'écriture
StringTrimRight, strMultiPurpose, strMultiPurpose, 1 ; remove extra delimiter
GuiControl, 1:, strMultiPurpose, % StrEscape(strMultiPurpose)
return



ClickRadHTML:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, HTML template:
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Show, btnMultiPurpose
GuiControl, 1:, btnMultiPurpose, Select HTML template
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "html")
return



ClickRadXML:
Gui, 1:Submit, NoHide
GuiControl, 1:Hide, lblMultiPurpose
GuiControl, 1:Hide, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Hide, btnMultiPurpose
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "xml")
return



ClickRadExpress:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, Express template:
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Hide, btnMultiPurpose
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
strRealFieldDelimiter1 := StrMakeRealFieldDelimiter(strFieldDelimiter1)
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strRealFieldDelimiter1, strFieldEncapsulator1)
; strFieldDelimiter1 et strFieldEncapsulator1 pour la lecture de strCurrentHeader
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
strMultiPurpose := ""
Loop, % objCurrentHeader.MaxIndex()
	strMultiPurpose := strMultiPurpose . strTemplateDelimiter . Format4CSV(objCurrentHeader[A_Index], strRealFieldDelimiter3, strFieldEncapsulator3) . strTemplateDelimiter . A_Space
	; strFieldDelimiter3 et strFieldEncapsulator3 pour l'écriture
StringTrimRight, strMultiPurpose, strMultiPurpose, 1 ; remove extra delimiter
GuiControl, 1:, strMultiPurpose, % StrEscape(strMultiPurpose)
return



ButtonHelpExportFormat:
/*
pour help, ré.écrire: Fixed width files are text files files where data is presented in lines and fields. The fields themselves are placed at fixed offsets.
MS ACCESS Fixed-width files     In a fixed-width file, each record appears on a separate line, and the width of each field remains consistent across records. In other words, the length of the first field of every record might always be seven characters, the length of the second field of every record might always be 12 characters, and so on. If the actual values of a field vary from record to record, the values that fall short of the required width will be padded with trailing spaces.

Pour fix: header selon tab 3
Pour tous (sauf HTML): eol remplacement du tab 3
Pour quels? Delimiters?

Pour HTML: délimiteur du template: ¤ (ASCII 164)


Pour Express
`t`r`n`f

*/
return



ChangedMultiPurpose:
return



ButtonMultiPurpose:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
if (radFixed)
{
	InputBox, intNewDefaultWidth, %strApplicationName% (%strApplicationVersion%) - Default fixed width, Enter the new default width:, , , 120, , , , , %intDefaultWidth%
	if !ErrorLevel
		if (intNewDefaultWidth > 0)
			intDefaultWidth := intNewDefaultWidth
		else
			MsgBox, 48, %strApplicationName%, Default fixed width must be greater than 0.
	Gosub, ClickRadFixed
}
else if (radHTML)
{
	FileSelectFile, strHtmlTemplateFile, 3, %A_ScriptDir%, Select HTML template
	if !(StrLen(strHtmlTemplateFile))
		return
	GuiControl, 1:, strMultiPurpose, %strHtmlTemplateFile%
}
; else if (radXML)
; else if (radExpress)
return



ButtonExportFile:
Gui, 1:Submit, NoHide
blnOverwrite := CheckIfFileExistOverwrite(strFileToExport)
if (blnOverwrite < 0)
	return
gosub, CheckOneRow
if (radFixed)
	if StrLen(strMultiPurpose)
		Gosub, ExportFixed
	else
	{
		MsgBox, 48, %strApplicationName%, Fill the "Fields width" zone with fields names and width separated by the field delimiter ( %strFieldDelimiter3% ).
		return
	}
else if (radHTML)
	if StrLen(strMultiPurpose)
		Gosub, ExportHTML
	else
	{
		MsgBox, 48, %strApplicationName%, First use the "Select HTML template" button to choose the HTML template file.
		return
	}
else if (radXML)
	Gosub, ExportXML
else if (radExpress)
	if StrLen(strMultiPurpose)
		Gosub, ExportExpress
	else
	{
		MsgBox, 48, %strApplicationName%, First enter the row template in the "Express template:" zone.
		return
	}
else
	MsgBox, 48, %strApplicationName%, Select the Export format.
return



ButtonCheckExportFile:
Gui, 1:Submit, NoHide
if InStr(strFileToExport, ".htm")
	run %strFileToExport%
else
	run notepad.exe %strFileToExport%
return




; --------------------- LISTVIEW EVENTS --------------------------


ListViewEvents:
if (A_GuiEvent = "ColClick")
{
	intColNumber := A_EventInfo
	Menu, SortMenu, Add, &Sort alphabetical, MenuSortText
	Menu, SortMenu, Add, Sort &numeric (integer), MenuSortInteger
	Menu, SortMenu, Add, Sort numeric (&float), MenuSortFloat
	Menu, SortMenu, Add
	Menu, SortMenu, Add, Sort descending &alphabetical, MenuSortDescText
	Menu, SortMenu, Add, Sort &descending numeric (integer), MenuSortDescInteger
	Menu, SortMenu, Add, Sort descending numeric (f&loat), MenuSortDescFloat
	Menu, SortMenu, Show
}
if (A_GuiEvent = "DoubleClick")
{
	intRowNumber := A_EventInfo
	Gui, 1:Submit, NoHide
	intGui1WinID := WinExist("A")
	Gui, 2:New, +Resize , %strApplicationName% - Edit row
	Gui, 2:+Owner1
	Gui, 1:Default
	SysGet, intMonWork, MonitorWorkArea 
	intColWidth := 380
	intEditWidth := intColWidth - 20
	intMaxNbCol := Floor(intMonWorkRight / intColWidth)
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
		}
		intYLabel := intY
		intYEdit := intY + 15
		LV_GetText(strColHeader, 0, A_Index)
		LV_GetText(strColData, intRowNumber, A_Index)
		Gui, 2:Add, Text, y%intYLabel% x%intX% vstrLabel%A_Index%, %strColHeader%
		Gui, 2:Add, Edit, y%intYEdit% x%intX% w%intEditWidth% vstrEdit%A_Index% +HwndstrEditHandle, %strColData%
		ShrinkEditControl(strEditHandle, 2, "2")
		GuiControlGet, intPosEdit, 2:Pos, %strEditHandle%
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



MenuSortText:
MenuSortInteger:
MenuSortFloat:
MenuSortDescText:
MenuSortDescInteger:
MenuSortDescFloat:
StringReplace, strOption, A_ThisLabel, Menu,
StringReplace, strOption, strOption, Sort, % "Sort "
StringReplace, strOption, strOption, Sort Desc, % "SortDesc "
LV_ModifyCol(intColNumber, strOption)
if InStr(strOption, "Text")
	LV_ModifyCol(intColNumber, "Left")
Menu, SortMenu, Delete
return



GuiContextMenu:  ; Launched in response to a right-click or press of the Apps key.
if A_GuiControl <> lvData  ; This check is optional. It displays the menu only for clicks inside the ListView.
    return
Menu, SelectMenu, Add, Select &All, MenuSelectAll
Menu, SelectMenu, Add, D&eselect All, MenuSelectNone
Menu, SelectMenu, Add, &Reverse Selection, MenuSelectReverse
; Show the menu at the provided coordinates, A_GuiX and A_GuiY.  These should be used
; because they provide correct coordinates even if the user pressed the Apps key:
Menu, SelectMenu, Show, %A_GuiX%, %A_GuiY%
return



MenuSelectAll:
GuiControl, Focus, lvData
LV_Modify(0, "Select")
Menu, SelectMenu, Delete
return



MenuSelectNone:
GuiControl, Focus, lvData
LV_Modify(0, "-Select")
Menu, SelectMenu, Delete
return



MenuSelectReverse:
GuiControl, Focus, lvData
Loop, % LV_GetCount()
	if IsRowSelected(A_Index)
		LV_Modify(A_Index, "-Select")
	else
		LV_Modify(A_Index, "Select")
Menu, SelectMenu, Delete
return



IsRowSelected(intRow)
{
	intNextSelectedRow := LV_GetNext(intRow - 1)
	return (intNextSelectedRow = intRow)
}



; --------------------- GUI2  --------------------------


ButtonSaveRecord:
if (intRowNumber < 1)
	###_D("Pas normal! intRowNumber: " . intRowNumber)
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



2GuiSize:  ; Expand or shrink the ListView in response to the user's resizing of the window.
; ###_D("intNbFieldsOnScreen: " . intNbFieldsOnScreen . " / intWidthSize: " . intWidthSize)
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



; --------------------- OTHER PROCEDURES --------------------------


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

GuiControl, 1:Move, strFileToExport, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpFileToExport, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnSelectFileToExport, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnExportFile, % "X" . (A_GuiWidth - 65) ; ###
GuiControl, 1:Move, btnCheckExportFile, % "X" . (A_GuiWidth - 65) ; ###
GuiControl, 1:Move, strMultiPurpose, % "W" . (A_GuiWidth - 305) ;  ### était 205
GuiControl, 1:Move, btnMultiPurpose, % "X" . (A_GuiWidth - 190) ; ### était 90

GuiControl, 1:Move, lvData, % "W" . (A_GuiWidth - 20) . " H" . (A_GuiHeight - 190)

return



UpdateCurrentHeader:
Gui, 1:Submit, NoHide
strCurrentHeader := GetHeader(strFieldDelimiter1, strFieldEncapsulator1)
strEscapedCurrentHeader := StrEscape(strCurrentHeader)
GuiControl, 1:, strFileHeaderEscaped, %strEscapedCurrentHeader%
GuiControl, 1:, strRenameEscaped, %strEscapedCurrentHeader%
GuiControl, 1:, strSelectEscaped, %strEscapedCurrentHeader%
GuiControl, 1:, strOrderEscaped, %strEscapedCurrentHeader%
if (radFixed)
	Gosub, ClickRadFixed
else if (radHTML)
	Gosub, ClickRadHTML
else if (radExpress)
	Gosub, ClickRadExpress
return



CheckOneRow:
if (LV_GetCount("Selected") = 1)
{
	MsgBox, 35, %strApplicationName% - One record selected, Only one record is selected. Do you want to save only this record?`n`nYes: Only one record will be saved.`nNo: All records will be saved.
	IfMsgBox, No
		LV_Modify(0, "Select") ; select all records
	IfMsgBox, Cancel
		return
}
return



ExportFixed:
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = 0])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3) ; strFieldDelimiter3 et strFieldEncapsulator3 pour l'écriture de l'entête seulement
objFieldsArray := ReturnDSVObjectArray(StrUnEscape(strMultiPurpose), strRealFieldDelimiter3, strFieldEncapsulator3)
strFieldsName := ""
strFieldsWidth := ""
loop, % objFieldsArray.MaxIndex() / 2
{
	strThisName := objFieldsArray[(A_Index * 2) - 1]
	strFieldsName := strFieldsName . Format4CSV(strThisName, strRealFieldDelimiter3, strFieldEncapsulator3) . strRealFieldDelimiter3
	intThisWidth := objFieldsArray[(A_Index * 2)]
	if intThisWidth  is integer
		strFieldsWidth := strFieldsWidth . intThisWidth . strRealFieldDelimiter3
	else
	{
		MsgBox, 48, %strApplicationName%, "%intThisWidth%" in field # %A_Index% "%strThisName%" must be an integer number.
		return
	}
}
StringTrimRight, strFieldsName, strFieldsName, 1 ; remove extra delimiter
StringTrimRight, strFieldsWidth, strFieldsWidth, 1 ; remove extra delimiter
if (radSaveMultiline)
	strEolReplacement := ""
else
	strEolReplacement := strEndoflineReplacement
; ObjCSV_Collection2Fixed(objCollection, strFilePath, strFieldsWidth, blnHeader := 0, strFieldOrder := "", blnProgress := 0, blnOverwrite := 0, strFieldDelimiter := ",", strEncapsulator := """", strEndOfLine := "`r`n", strEolReplacement := "")
ObjCSV_Collection2Fixed(obj, strFileToExport, strFieldsWidth, radSaveWithHeader, strFieldsName, 1, blnOverwrite, strRealFieldDelimiter3, strFieldEncapsulator3, , strEolReplacement)
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return



ExportHTML:
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = 0])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
; ObjCSV_Collection2HTML(objCollection, strFilePath, strTemplateFile [, strTemplateEncapsulator = ~, blnProgress = 0, blnOverwrite = 0])
ObjCSV_Collection2HTML(obj, strFileToExport, strMultiPurpose, strTemplateDelimiter, 0, 1)
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return



ExportXML:
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = 0])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
; ObjCSV_Collection2XML(objCollection, strFilePath [, blnProgress = 0, blnOverwrite = 0])
ObjCSV_Collection2XML(obj, strFileToExport, 0, 1)
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return



ExportExpress:
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = 0])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
; ObjCSV_Collection2HTML(objCollection, strFilePath, strTemplateFile [, strTemplateEncapsulator = ~, blnProgress = 0, blnOverwrite = 0])
strExpressTemplateTempFile := GUID() . ".TMP"
strExpressTemplate := strTemplateDelimiter . "ROWS" . strTemplateDelimiter
strExpressTemplate := strExpressTemplate  . strMultiPurpose
strExpressTemplate := strExpressTemplate  . strTemplateDelimiter . "/ROWS" . strTemplateDelimiter
strExpressTemplate := StrUnEscape(strExpressTemplate)
FileAppend, %strExpressTemplate%, %strExpressTemplateTempFile%
ObjCSV_Collection2HTML(obj, strFileToExport, strExpressTemplateTempFile, strTemplateDelimiter, 0, 1)
FileDelete, %strExpressTemplateTempFile%
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return



GuiClose:
ExitApp



; --------------------- FUNCTIONS --------------------------


GetHeader(strFieldDelimiter, strFieldEncapsulator)
{
	strDelimiter := StrMakeRealFieldDelimiter(strFieldDelimiter)
	strHeader := ""
	Loop, % LV_GetCount("Column")
	{
		LV_GetText(strColumnHeader, 0, A_Index)
		strHeader := strHeader . Format4CSV(strColumnHeader, strDelimiter, strFieldEncapsulator) . strDelimiter
	}
	StringTrimRight, strHeader, strHeader, 1 ; remove extra delimiter
	return strHeader
}



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



StrMakeRealFieldDelimiter(strConverted)
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
	MsgBox, 0, %strApplicationName% (%strApplicationVersion%) - %strTitle% Help,%strMessage%
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



CheckIfFileExistOverwrite(strFileName)
{
	if !StrLen(strFileName)
		return -1
	if !FileExist(strFileName)
		return True
	else
	{
		MsgBox, 35, %strApplicationName% - File exists, File exists:`n%strFileName%`n`nDo you want to overwrite this file?`n`nYes: The file will be overwritten.`nNo: Data will be added to the existing file.
		IfMsgBox, Yes
			return True
		IfMsgBox, No
			return False
		IfMsgBox, Cancel
			return -1
	}
}



NewFileName(strExistingFile, strNote := "", strExtension := "")
{
	SplitPath, strExistingFile, , strOutDir, strOutExtension, strOutNameNoExt
	if !StrLen(strExtension)
		strExtension := strOutExtension
	loop
	{
		strNewName := strOutDir . "\" . strOutNameNoExt . strNote . " (" . A_Index  . ")." . strExtension
		if !FileExist(strNewName)
			break
	}
	return strNewName
}



GUID()         ; 32 hex digits = 128-bit Globally Unique ID
; Source: Laszlo in http://www.autohotkey.com/board/topic/5362-more-secure-random-numbers/
{
   format = %A_FormatInteger%       ; save original integer format
   SetFormat Integer, Hex           ; for converting bytes to hex
   VarSetCapacity(A,16)
   DllCall("rpcrt4\UuidCreate","Str",A)
   Address := &A
   Loop 16
   {
      x := 256 + *Address           ; get byte in hex, set 17th bit
      StringTrimLeft x, x, 3        ; remove 0x1
      h = %x%%h%                    ; in memory: LS byte first
      Address++
   }
   SetFormat Integer, %format%      ; restore original format
   Return h
}




