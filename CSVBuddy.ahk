;===============================================
/*
CSV Buddy
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By JnLlnd on AHK forum
This script uses the library ObjCSV v0.2 (https://github.com/JnLlnd/ObjCSV)
*/ 

#NoEnv
#SingleInstance force
#LTrim ; omits spaces and tabs at the beginning of each line in continuation sections
#Include %A_ScriptDir%\..\ObjCSV\lib\ObjCSV.ahk

; --------------------- GLOBAL AND DEFAULT VALUES --------------------------

global strApplicationName := "CSV Buddy"
global strApplicationVersion := "v0.2.3 ALPHA" ; 2013-09-24

intDefaultWidth := 16 ; used when export to fixed-width format
strTemplateDelimiter := "¤" ; Chr(164)


; --------------------- GUI1 --------------------------


Gui, 1:New, +Resize, %strApplicationName%

Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Text, x10, %strApplicationName%

Gui, 1:Font, s10 w700, Verdana
Gui, 1:Add, Tab2, w950 r4 vtabCSVBuddy gChangedTabCSVBuddy
	, % " 1) Load CSV File     ||     2) Edit Columns     |     3) Save CSV File     |     4) Export     |     About     "
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
Gui, 1:Add, Radio,		y100	x300	vradSaveWithHeader checked, Save &with header
Gui, 1:Add, Radio,		y+10	x300	vradSaveNoHeader, Save without header
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
Gui, 1:Add, Radio,		yp		x100	vradFixed gClickRadFixed, Fixed-width
Gui, 1:Add, Radio,		yp		x+15	vradHTML gClickRadHTML, HTML
Gui, 1:Add, Radio,		yp		x+15	vradXML gClickRadXML, XML
Gui, 1:Add, Radio,		yp		x+15	vradExpress gClickRadExpress, Express
Gui, 1:Add, Button,		yp		x+15	vbtnHelpExportFormat gButtonHelpExportFormat, ?
Gui, 1:Add, Button,		yp		x+15	vbtnHelpExportMulti gButtonHelpExportMulti Hidden
	, Lorem ipsum dolor sitm ; conserver texte important pour la largeur du bouton
Gui, 1:Add, Text,		y+10	x10		vlblMultiPurpose w85 right hidden, Hidden Label:
Gui, 1:Add, Edit,		yp		x100	vstrMultiPurpose gChangedMultiPurpose hidden
Gui, 1:Add, Button,		yp		x+5		vbtnMultiPurpose gButtonMultiPurpose hidden
	, Lorem ipsum dolor sitm ; conserver texte important pour la largeur du bouton
Gui, 1:Add, Button,		y105	x+5		vbtnExportFile gButtonExportFile hidden, Export
Gui, 1:Add, Button,		y137	x+5		vbtnCheckExportFile gButtonCheckExportFile hidden, Check

Gui, 1:Tab, 5
Gui, 1:Font, s10 w700, Verdana
Gui, 1:Add, Link,		y+10	x10		vlblAboutText1,
(Join`s
<a href="https://bitbucket.org/JnLlnd/csvbuddy">%strApplicationName% %strApplicationVersion%</a>
)
Gui, 1:Font, s9 w500, Arial
Gui, 1:Add, Link,		y+4	x10		vlblAboutText2,
(Join`s
by Jean Lalonde (<a href="http://www.autohotkey.com/board/user/4880-jnllnd/">JnLlnd</a> on AHK forum)
)
Gui, 1:Font
Gui, 1:Add, Link,		y+4	x10		vlblAboutText3,
(Join`s
`nAll rights reserved (c)2013 - DO NOT DISTRIBUTE WITHOUT AUTHOR AUTORIZATION
`nUsing AHK library: <a href="https://www.github.com/JnLlnd/ObjCSV">ObjCSV v0.2</a>
`nUsing icon by: <a href="http://www.visualpharm.com">Visual Pharm</a>
)


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
		Oops("First load a CSV file in the first tab.")
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
else if InStr(tabCSVBuddy, "Save")
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnSelectFileToSave
	else
	{
		Oops("First load a CSV file in the first tab.")
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
else if InStr(tabCSVBuddy, "Export")
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnSelectFileToExport
	else
	{
		Oops("First load a CSV file in the first tab.")
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
strHelp =
(Join`s
Hit "Select" to choose the CSV file to load.

`n`nClick on the various Help (?) buttons to learn about the options offered by
%strApplicationName%. When setting are ready, hit "Load" to import the file.

`n`nNote that %strApplicationName% can load CSV files with up to 200 fields.
)
Help("CSV File To Load", strHelp)
return



DetectDelimiters:
Gui, 1:Submit, NoHide
strFileHeaderUnEscaped := StrUnEscape(strFileHeaderEscaped)
strCandidates := "`t;,:|~" ; check tab, semi-colon, comma, colon, pipe and tilde
strFieldDelimiterDetected := "," ; comma by default if no delimiter is detected
loop, Parse, strCandidates
	if InStr(strFileHeaderUnEscaped, A_LoopField)
	{
		strFieldDelimiterDetected := A_LoopField
		break
	}
GuiControl, 1:, strFieldDelimiter1, % StrMakeEncodedFieldDelimiter(strFieldDelimiterDetected)
strCandidates := """'~|" ; check double-quote, single-quote, tilde and pipe
strFieldEncapsulatorDetected := """" ; double-quotes by default if no encapsulator is detected
loop, Parse, strCandidates
	if (strFieldDelimiterDetected <> A_LoopField) and (InStr(strFileHeaderUnEscaped, strFieldDelimiterDetected . A_LoopField) or InStr(strFileHeaderUnEscaped, A_LoopField . strFieldDelimiterDetected))
	{
		strFieldEncapsulatorDetected := A_LoopField
		break
	}
GuiControl, 1:, strFieldEncapsulator1, %strFieldEncapsulatorDetected%
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
gosub, DetectDelimiters
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
strHelp =
(Join`s
Most of the time, the first line of a CSV file contains the CSV header, a list of field names separated by a field delimiter.
If your file contains a CSV Header, select the radio button "Get CSV Header". When you select a file (using the "Select" button),
the "CSV Header" zone displays the first line of the selected file.

`n`nNote that invisible characters used as delimiters (for example Tab) are displayed with an escape character. For example,
Tabs are shown as "``t".

`n`nIf the file does not contain a CSV header, select the radio button "Set CSV Header" and enter in the "CSV Header" zone the
field names for each column of data in the file, seperated by the field delimiter.
)
Help("CSV Header", strHelp)
return



ButtonPreviewFile:
Gui, 1:Submit, NoHide
if !StrLen(strFileToLoad)
{
	Oops("First use the ""Select"" button to choose the CSV file you want to load.")
	return
}
run, notepad.exe %strFileToLoad%
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
strHelp =
(Join`s
If the first line of the CSV file contains the list of field names, click "Get header from CSV file".
If not, click "Set CSV header" and enter the list of field names separated by the "Field delimiter".
)
Help("CSV Get/Set CSV Header", strHelp)
return



ChangedFieldDelimiter1:
return



ButtonHelpFieldDelimiter1:
strHelp =
(Join`s
Each field in the CSV header and in data rows of the file must be separated by a field delimiter.
This is often comma ( , ), semicolon ( `; ) or Tab.

`n`n%strApplicationName% will detect the delimiter if one of these characters is found in the first line
of the file: tab, semi-colon, comma, colon, pipe or tilde. If this is not the correct delimiter, enter
any single character or one of these replacement letters for invisible characters:

`n`nt`tTab (HT)
`nn`tLinefeed (LF)
`nr`tCarriage return (CR)
`nf`tForm feed (FF)

`n`nSpace can also be used as delimiter. Just enter a space in the text zone.

`n`nTip: Use the "Preview" button to find what is the field delimiter in the selected file.
)
Help("Field Delimiter", strHelp)
return



ChangedFieldEncapsulator1:
return



ButtonHelpEncapsulator1:
strHelp =
(Join`s
When data fields in a CSV file contain characters used as delimiter or end-of-line, they must be enclosed in a field encapsulator.
This encapsulator is often double-quotes ( "..." ) or single quotes ( '...' ). For example, if comma is used as field delimiter
in a CSV file, the data field "Smith, John" must be encapsulated because it contains a comma.

`n`nIf a field contains the character used as encapsulator, this character must be doubled. For example, the data "John "Junior" Smith"
must be stated as "John ""Junior"" Smith".

`n`n%strApplicationName% will detect the encapsulator if one of these characters is found in the first line of the file:
double-quote, single-quote, tilde or pipe. If this is not the correct encapsulator, enter any single character.

`n`nTip: Use the "Preview" button to find what is the field encapsulator in the selected file.
)
Help("Field Encapsulator", strHelp)
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
strHelp =
(Join`s
Most CSV files do not contain line breaks inside text field. But some do. For example, you can find multi-lines "Notes" fields in
Google or Outlook contacts exported files.

`n`nIf text fields in your CSV file contain line breaks, select this checkbox to turn this option ON.
If not, keep it OFF since this will improve loading performance.

`n`nIf you turn Multi-line ON, you have the additional option to choose a character (or string) that will be converted to
line-breaks if found in the CSV file.
)
Help("Multi-line Fields", strHelp)
return



ButtonLoadFile:
Gui, 1:+OwnDialogs
Gui, 1:Submit, NoHide
if !DelimitersOK(1)
	return
if !StrLen(strFileToLoad)
{
	Oops("First use the ""Select"" button to choose the CSV file you want to load.")
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
		MsgBox, 36, %strApplicationName%,
			(
			If the CSV file you want to load have the same fields, in the same order, you can add file data to the current list.
			
			Do you want to add to the content of this file to the list?
			)
		IfMsgBox, No
			return
	}
}
strCurrentHeader := StrUnEscape(strFileHeaderEscaped)
strCurrentFieldDelimiter := StrMakeRealFieldDelimiter(strFieldDelimiter1)
strCurrentVisibleFieldDelimiter := strFieldDelimiter1
strCurrentFieldEncapsulator := strFieldEncapsulator1
; ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames [, blnHeader = 1, blnMultiline = 1, blnProgress = 0, strFieldDelimiter = ","
; 	, strEncapsulator = """", strRecordDelimiter = "`n", strOmitChars = "`r", strEolReplacement = ""])
obj := ObjCSV_CSV2Collection(strFileToLoad, strCurrentHeader, radGetHeader, blnMultiline1, 1, strCurrentFieldDelimiter
	, strCurrentFieldEncapsulator, , , strEndoflineReplacement1)
; ObjCSV_Collection2ListView(objCollection [, strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", strSortFields = "", strSortOptions = "", blnProgress = 0])
ObjCSV_Collection2ListView(obj, "1", "lvData", strCurrentHeader, strCurrentFieldDelimiter
	, strCurrentFieldEncapsulator, , , 1)
if !LV_GetCount()
{
	Oops("CSV file not loaded.`n`nNote that " . strApplicationName . " support files with a maximum of 200 fields.")
	return
}
else
{
	Gosub, UpdateCurrentHeader
	strHelp =
	(Join`s
	Your CSV file is loaded.
	
	`n`nYou can sort rows by clicking on column headers. Choose sorting type: alphabetical, numeric integer or numeric float,
	ascending or descending.
	
	`n`nDouble-click on a row to edit a record.  Right-click anywhere in the list view to select all rows, deselect all rows
	or reverse selection.
	
	`n`nYou can use the "2) Edit Columns" tab to edit field names, select fields to keep or change fields order.
	
	`n`nWhen you will be ready, go to the "3) Save CSV File" tab to save all or selected rows in a new CSV file or
	to the "4) Export" tab to export your data to fixed-width, HTML or XML format.
	)
	Help("Ready to edit", strHelp)
	GuiControl, 1:, strFieldDelimiter3, %strCurrentVisibleFieldDelimiter%
	GuiControl, 1:, strFieldEncapsulator3, %strCurrentFieldEncapsulator%
}
obj := ; release object
return



; --------------------- TAB 2 --------------------------

ButtonSetRename:
Gui, 1:Submit, NoHide
if !LV_GetCount()
{
	Oops("First load a CSV file in the first tab.")
	return
}
objNewHeader := ReturnDSVObjectArray(StrUnEscape(strRenameEscaped), strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
intNbFieldNames := objNewHeader.MaxIndex()
intNbColumns := LV_GetCount("Column")
if !StrLen(strRenameEscaped)
{
	MsgBox, 52, %strApplicationName%,
	(
	In "Rename fields:", enter the list of field names separated by the field delimiter ( %strCurrentVisibleFieldDelimiter% ).
	
	Field names are automatically filled when you load a CSV file in the first tab.
	
	If no field names are provided, numbers are used as field names. Do you want to use numbers as field names? 
	)
	IfMsgBox, No
		return
}
else if (intNbFieldNames < intNbColumns)
{
	MsgBox, 52, %strApplicationName%,
	(
	There are less field names in the "Rename fields:" zone (%intNbFieldNames%) than the number of columns in the list (%intNbColumns%).

	Do you want to use numbers as field names for remaining columns? 
	)
	IfMsgBox, No
		return
}
Loop, % LV_GetCount("Column")
{
	if StrLen(objNewHeader[A_Index])
		LV_ModifyCol(A_Index, "", objNewHeader[A_Index])
	else
		LV_ModifyCol(A_Index, "", A_Index)
	LV_ModifyCol(A_Index, "AutoHdr")
}
Gosub, UpdateCurrentHeader
objNewHeader := ; release object
return



ButtonHelpRename:
Gui, 1:Submit, NoHide
strHelp =
(Join`s
To change field names (column headers), enter a new name for each fields, in the order they actually appear in the list,
separated by the field delimiter ( %strCurrentVisibleFieldDelimiter% ) and click "Rename".

`n`nIf you enter less names than the number of fields (or no field name at all), numbers are used as field names for remaining columns.

`n`nField names including the separator character ( %strCurrentVisibleFieldDelimiter% ) must be enclosed by the encapsulator character
( %strCurrentFieldEncapsulator% ).

`n`nTo save the file, click on the tab "3) Save CSV File".
)
Help("Rename Fields", strHelp)
return



ButtonSetSelect:
Gui, 1:Submit, NoHide
if !LV_GetCount()
{
	Oops("First load a CSV file in the first tab.")
	return
}
if !StrLen(strSelectEscaped)
{
	Oops(
	(Join`s
	"First enter the names of the fields you want to keep in the list, separated by the field delimiter ( " . strCurrentVisibleFieldDelimiter . " ),
	keeping their current order.
	
	`n`nField names are automatically filled when you load a CSV file in the first tab."
	))
	return
}
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
objNewHeader := ReturnDSVObjectArray(StrUnEscape(strSelectEscaped), strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
intPosPrevious := 0
for intKey, strVal in objNewHeader
{
	intPosThisOne := PositionInArray(strVal, objCurrentHeader)
	if !(intPosThisOne)
	{
		Oops(
		(Join`s
		"Field name """ . strVal . """ in the ""Select fields:"" zone not found in the list."
		))
		return
	}
	if (intPosThisOne <= intPosPrevious)
	{
		Oops("Field names in the ""Select fields:"" zone must be in the same order as the current list.")
		return
	}
	intPosPrevious := intPosThisOne
}
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
Gosub, UpdateCurrentHeader
objCurrentHeader := ; release object
objNewHeader := ; release object
return



ButtonHelpSelect:
Gui, 1:Submit, NoHide
strHelp =
(Join`s
To remove fields (columns) from the list, enter the name of fields you want to keep, in the order they actually appear in the list,
separated by the field delimiter ( %strCurrentVisibleFieldDelimiter% ) and click "Select".

`n`nTo save the file, click on the last tab "3) Save CSV File".
)
Help("Select Fields", strHelp)
return



ButtonSetOrder:
Gui, 1:Submit, NoHide
if !StrLen(strSelectEscaped)
{
	Oops(
	(Join`s
	"First enter the names of the fields you want to keep in the list, in the desired order,
	separated by the field delimiter ( " . strCurrentVisibleFieldDelimiter . " ).
	
	`n`nField names are automatically filled when you load a CSV file in the first tab."
	))
	return
}
if !LV_GetCount()
{
	Oops("First load a CSV file in the first tab.")
	return
}
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
objNewHeader := ReturnDSVObjectArray(StrUnEscape(strOrderEscaped), strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
for intKey, strVal in objNewHeader
{
	if !PositionInArray(strVal, objCurrentHeader)
	{
		Oops("Field name """ . strVal . """ in the ""Order fields:"" zone not found in the list.")
		return
	}
}
objNewCollection := ObjCSV_ListView2Collection("1", "lvData", StrUnEscape(strOrderEscaped), strCurrentFieldDelimiter
	, strCurrentFieldEncapsulator, 1)
LV_Delete() ;  better performance on large files when we delete rows before columns
loop, % LV_GetCount("Column")
	LV_DeleteCol(1) ; delete all rows
ObjCSV_Collection2ListView(objNewCollection, "1", "lvData", StrUnEscape(strOrderEscaped), strCurrentFieldDelimiter
	, strCurrentFieldEncapsulator, , , 1)
Gosub, UpdateCurrentHeader
objNewCollection := ; release object
return



ButtonHelpOrder:
Gui, 1:Submit, NoHide
strHelp =
(Join`s
To change the order of fields (columns) in the list, enter the field names in the new order you want to apply,
separated by the field delimiter ( %strCurrentVisibleFieldDelimiter% ) and click "Order".

`n`nIf you enter less field names than in the original header, fields not included in the new order are removed from the list.
However, if you only want to remove fields from the list (without changing the order), the "Select" button gives better
performance on large files.

`n`nTo save the file, click on the last tab "3) Save CSV File" and select the destination file.
)
Help("Order Fields", strHelp)
return



; --------------------- TAB 3 --------------------------


ButtonHelpFileToSave:
strHelp =
(Join`s
Enter the name of the destination CSV file (the current program's directory is used if an absolute path isn't specified)
or hit "Select" to choose the CSV destination file. When other options are OK, hit "Save" to save all or selected rows to the
CSV file.

`n`nNote that all rows are saved by default. You can select one row (using Click), a series of adjacent rows (using Shift-Click)
or non contiguous rows (using Ctrl-Click or Shift-Ctrl-Click). You can also Right-Click in the list to select or deselect all rows,
or to reverse the current row selection.

`n`nNote that fields are saved in the order they appear in the list and that rows are saved according to the current sorting order
(click on a column name to sort rows).
)
Help("CSV File To Save", strHelp)
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
strHelp =
(Join`s
Each field in the CSV header and in data rows of the file must be separated by a field delimiter.
Enter the field delimiter character to use in the destination file.

`n`nIt can be comma ( , ), semicolon ( `; ), Tab or any single character.

`n`nUse the letters on the left as replacement for the following invisible characters:

`n`nt`tTab (HT)
`nn`tLinefeed (LF)
`nr`tCarriage return (CR)
`nf`tForm feed (FF)
)
Help("Field Delimiter", strHelp)
return



ChangedFieldDelimiter3:
strPreviousDelimiter := strFieldDelimiter3
Gui, 1:Submit, NoHide
if StrLen(strFieldDelimiter3)
{
	if !NewDelimiterOrEncapsulatorOK(StrMakeRealFieldDelimiter(strFieldDelimiter3))
	{
		Oops("The new field delimiter ( " . strFieldDelimiter3 
			. " ) cannot be choosen because it is currently in use in the field names.")
		GuiControl, 1:, strFieldDelimiter3, %strPreviousDelimiter%
		return
	}
	Gosub, UpdateCurrentHeader
}
return



ChangedFieldEncapsulator3:
strPreviousEncapsulator := strFieldEncapsulator3
Gui, 1:Submit, NoHide
if StrLen(strFieldEncapsulator3)
{
	if !NewDelimiterOrEncapsulatorOK(strFieldEncapsulator3)
	{
		Oops("The new field encapsulator ( " . strFieldEncapsulator3 
			. " ) cannot be choosen because it is currently in use in the field names.")
		GuiControl, 1:, strFieldEncapsulator3, %strPreviousEncapsulator%
		return
	}
	Gosub, UpdateCurrentHeader
}
return



ButtonHelpEncapsulator3:
strHelp =
(Join`s
Data fields in a CSV file containing the character used as field delimiter or an end-of-line must be enclosed in a field encapsulator.
Enter the field encapsulator character to use in the destination file.

`n`nThe encapsulator is often double-quotes ( "..." ) or single quotes ( '...' ). In the example "Smith, John", the data field
containing a comma will be encapsulated because comma is also the field delimiter.

`n`nIf a field contains the character used as encapsulator, this encapsulator will be doubled. For example, the data "John "Junior" Smith"
will be entered as "John ""Junior"" Smith".
)
Help("Field Encapsulator", strHelp)
return



ButtonHelpSaveHeader:
Gui, 1:Submit, NoHide
strHelp =
(Join`s
To save the field names as the first line of the CSV file, select "Save with header".

`n`nIf you select "Save without header", the first line of the file will contain the data of the first row of the list.
)
Help("Save CSV Header", strHelp)
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
strHelp =
(Join`s
If a field contains line break, you can decide if this line break is saved as is or if it is replaced with a character (or a sequence of characters)
in order to keep the field on a single line. This can be useful if, later, you want to open this file in a software that do not support multi-line
fields (e.g. MS Excel).

`n`nIf you select "Save multi-line", line breaks are saved unchanged.

`n`nIf you select "Save single-line", enter the replacement sequence for line breaks in the "End-of-line replacement:" zone.
By default, the replacement character is "¶" (ASCII code 182).
)
Help("Saving multi-line fields", strHelp)
return



ButtonSaveFile:
Gui, 1:Submit, NoHide
if !DelimitersOK(3)
	return
blnOverwrite := CheckIfFileExistOverwrite(strFileToSave)
if (blnOverwrite < 0)
	return
gosub, CheckOneRow
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", blnProgress = 0])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
if (radSaveMultiline)
	strEolReplacement := ""
else
	strEolReplacement := strEndoflineReplacement3
; ObjCSV_Collection2CSV(objCollection, strFilePath [, blnHeader = 0, strFieldOrder = "", blnProgress = 0, blnOverwrite = 0
;	, strFieldDelimiter = ",", strEncapsulator = """", strEndOfLine = "`n", strEolReplacement = ""])
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
ObjCSV_Collection2CSV(obj, strFileToSave, radSaveWithHeader
	, GetListViewHeader(strRealFieldDelimiter3, strFieldEncapsulator3), 1, blnOverwrite
	, strRealFieldDelimiter3, strFieldEncapsulator3, , strEolReplacement)
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
run, notepad.exe %strFileToSave%
return



; --------------------- TAB 4 --------------------------

ButtonHelpFileToExport:
strHelp =
(Join`s
Enter the name of the destination file of the export (the current program's directory is used if an absolute path isn't specified)
or hit "Select" to choose the destination file. When other options are OK, hit "Export" to export all or selected rows to the
destination file.

`n`nNote that all rows are saved by default. You can select one row (using Click), a series of adjacent rows (using Shift-Click)
or non contiguous rows (using Ctrl-Click or Shift-Ctrl-Click). You can also Right-Click in the list to select or deselect all rows,
or to reverse the current row selection.

`n`nRows are saved according to the current sorting order (click on a column name to sort rows).
)
Help("Export data", strHelp)
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
if !DelimitersOK(3)
{
	GuiControl, , radFixed, 0
	return
}
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, Fixed-width Export Help
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, Fields width:
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Show, btnMultiPurpose
GuiControl, 1:, btnMultiPurpose, Change default width
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
; strCurrentFieldDelimiter (strFieldDelimiter1) et strCurrentFieldEncapsulator (strFieldEncapsulator1) pour la lecture de strCurrentHeader
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
strMultiPurpose := ""
Loop, % objCurrentHeader.MaxIndex()
	strMultiPurpose := strMultiPurpose . Format4CSV(objCurrentHeader[A_Index]
		, strRealFieldDelimiter3, strFieldEncapsulator3) . strRealFieldDelimiter3 . intDefaultWidth . strRealFieldDelimiter3
	; strFieldDelimiter3 et strFieldEncapsulator3 pour l'écriture
StringTrimRight, strMultiPurpose, strMultiPurpose, 1 ; remove extra delimiter
GuiControl, 1:, strMultiPurpose, % StrEscape(strMultiPurpose)
return



ClickRadHTML:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, HTML Export Help
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
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, XML Export Help
GuiControl, 1:Hide, lblMultiPurpose
GuiControl, 1:Hide, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Hide, btnMultiPurpose
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "xml")
return



ClickRadExpress:
Gui, 1:Submit, NoHide
if !DelimitersOK(3)
{
	GuiControl, , radExpress, 0
	return
}
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, Express Export Help
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, Express template:
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Hide, btnMultiPurpose
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
; strCurrentFieldDelimiter et strCurrentFieldEncapsulator pour la lecture de strCurrentHeader
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
strMultiPurpose := ""
Loop, % objCurrentHeader.MaxIndex()
	strMultiPurpose := strMultiPurpose . strTemplateDelimiter . Format4CSV(objCurrentHeader[A_Index]
		, strRealFieldDelimiter3, strFieldEncapsulator3) . strTemplateDelimiter . A_Space
	; strFieldDelimiter3 et strFieldEncapsulator3 pour l'écriture
StringTrimRight, strMultiPurpose, strMultiPurpose, 1 ; remove extra delimiter
GuiControl, 1:, strMultiPurpose, % StrEscape(strMultiPurpose)
return



ButtonHelpExportFormat:
strHelp =
(Join`s
Choose one of these export formats:

`n`n• Fixed-width: To export to a text file where each record appears on a separate line, and the width of each field remains consistent across records.
Field names can be optionaly inserted on the first line. Field names and data fields shorter than their width are padded with trailing spaces. Field
names and data fields longer than their width are truncated at their maximal width. Fields are exported in the order they appear in the list.

`n`n• HTML: To build an HTML file based on a template file specifying header and footer templates, and a row template where variable names are replaced with the content
of each record in the collection.

`n`n• XML: To build an XML file from the content of the collection. You must ensure that field names and field data comply with the rules of XML syntax.
Fields are exported in the order they appear in the list.

`n`n• Express: To build a text file based on a row template where variable names are replaced with the content of each record in the collection.

`n`nSelect the export format. An additional "... Export Help" button will provide more instructions about the selected format.

`n`nClick the "Export" button to export data and the "Check" button to see the result in the destination file.
)
Help("Export Format", strHelp)
return



ButtonHelpExportMulti:
if (radFixed)
{
	strHelp =
	(Join`s
	Transfer the selected fields from a collection of objects to a fixed-width file.
	
	`n`nIn the "Fields width:", enter each field name to include in the file, followed by the width of this field. Field names and width values are
	separated by the field delimiter ( %strFieldDelimiter3% ) specified in the tab "3) Save CSV File". Initialy, the "Fields width:" zone includes all fields with
	a default width of %intDefaultWidth% characters. To change the default width, click the "Change default width" button.
	
	`n`nField names and data fields shorter than their width are padded with trailing spaces. Field names and data fields longer than their width
	are truncated at their maximal width.

	`n`nField names can be optionnaly included on the first line of the file according to the selected option "Save with header" or "Save without
	header" on the tab "3) Save CSV File".

	`n`nA fixed-width file should not include end-of-line within data. If it does and if a value is entered in the "End-of-line replacement:" on
	the tab "3) Save CSV File" (click "Save single-line" to see this option), end-of-line in multi-line fields are replaced by a character or string
	of your choice. This string is included in the fixed-width character count.

	`n`nClick "Export" button to export data and the "Check" button to see the result in the destination file.
	)
	Help("Fixed-width Export", strHelp)
}
else if (radHTML)
{
	strHelp =
	(Join`s
	Build an HTML file based on a template file specifying header and footer templates, and a row template where variable names are replaced with the content of each record in the collection.

	`n`nEnter the template file name in the "HTML template:" or click "Select HTML template" to choose it. The template is divided in three sections: the header template (from the start of the file to the start of the row template), the row template (delimited by the markups ¤ROWS¤ and ¤/ROWS¤) and the footer template (from the end of the row template to the end of the file).
	
	`n`nThe row template is repeated in the output file for each record in the collection. Field names encapsulated by the ¤ character (ASCII code 164) are replaced by the matching data in each record. Also, ¤ROWNUMBER¤ is replaced by the current row number.
	
	`n`nIn the header and footer, the following variables are replaced by parts of the destination file name:
	`n  ¤FILENAME¤ file name without its path, but including its extension
	`n  ¤DIR¤ drive letter or share name, if present, and directory of the file, final backslash excluded
	`n  ¤EXTENSION¤ file's extension, dot excluded
	`n  ¤NAMENOEXT¤ file name without its path, dot and extension
	`n  ¤DRIVE¤ drive letter or server name, if present

	`n`nThis simple example, where each record has two fields named "Field1" and "Field2", shows the use of the various markups and variables:

	`n`n<HEAD>
	`n<TITLE>¤NAMENOEXT¤</TITLE>
	`n</HEAD>
	`n<BODY>
	`n<H1>¤FILENAME¤</H1>
	`n<TABLE>
	`n<TR>
	`n<TH>Row #</TH><TH>Field One</TH><TH>Field Two</TH>
	`n</TR>
	`n¤ROWS¤
	`n<TR>
	`n<TD>¤ROWNUMBER¤</TD><TD>¤Field1¤</TD><TD>¤Field2¤</TD>
	`n</TR>
	`n¤/ROWS¤
	`n</TABLE>
	`nSource: ¤DIR¤\¤FILENAME¤
	`n</BODY>

	`n`nClick "Export" button to export data and the "Check" button to see the resulting HTML file in your default browser.
	)
	Help("HTML Export", strHelp)
}
else if (radXML)
{
	strHelp =
	(Join`s
	Build an XML file from the content of the collection. You must ensure that field names and field data comply with the rules of XML syntax.

	`n`nThis simple example, where each record has two fields named "Field1" and "Field2", shows the XML output format:

	`n`n<?xml version='1.0'?>
	`n<XMLExport>
	`n    <Record>
	`n        <Field1>Value Row 1 Col 1</Field1>
	`n        <Field2>Value Row 1 Col 2</Field1>
	`n    </Record>
	`n    <Record>
	`n        <Field1>Value Row 2 Col 1</Field1>
	`n        <Field2>Value Row 2 Col 2</Field1>
	`n    </Record>
	`n</XMLExport>

	`n`nClick "Export" button to export data and the "Check" button to see the result in the destination file.
	)
	Help("XML Export", strHelp)
}
else if (radExpress)
{
	strHelp =
	(Join`s
	Build a text file based on a row template where variable names are replaced with the content of each record in the collection.
	
	`n`nIn the "Express template:" zone, enter the template for each row of data in the collection. In this template, field names
	encapsulated by the character ¤ (ASCII code 164) are replaced by the matching data in each record. Also, ¤ROWNUMBER¤ is replaced
	by the current row number.
	
	`n`nAdditionaly, these special characters can be inserted in the template:
	`n`t``t`treplaced by Tab (HT)
	`n`t``n`treplaced by Linefeed (LF)
	`n`t``r`treplaced by Carriage return (CR)
	`n`t``f`treplaced by Form feed (FF)

	`n`nThe "Express template:" zone is initialized with all fields encapsulated by the character ¤ (ASCII code 164) and delimited with spaces.

	`n`nClick "Export" button to export data and the "Check" button to see the result in the destination file.
	)
	Help("Express Export", strHelp)
}
return



ChangedMultiPurpose:
return



ButtonMultiPurpose:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
if (radFixed)
{
	InputBox, intNewDefaultWidth, %strApplicationName% (%strApplicationVersion%) - Default fixed-width
		, Enter the new default width:, , , 120, , , , , %intDefaultWidth%
	if !ErrorLevel
		if (intNewDefaultWidth > 0)
			intDefaultWidth := intNewDefaultWidth
		else
			Oops("Default fixed-width must be greater than 0.")
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
		Oops(
		(Join`s
		"Fill the ""Fields width"" zone with fields names and width separated by the field delimiter ( "
		. strFieldDelimiter3 . ")."
		))
		return
	}
else if (radHTML)
	if StrLen(strMultiPurpose)
		Gosub, ExportHTML
	else
	{
		Oops("First use the ""Select HTML template"" button to choose the HTML template file.")
		return
	}
else if (radXML)
	Gosub, ExportXML
else if (radExpress)
	if StrLen(strMultiPurpose)
		Gosub, ExportExpress
	else
	{
		Oops("First enter the row template in the ""Express template:"" zone.")
		return
	}
else
	Oops("First, select the Export format.")
return



ButtonCheckExportFile:
Gui, 1:Submit, NoHide
if InStr(strFileToExport, ".htm")
	run, %strFileToExport%
else
	run, notepad.exe %strFileToExport%
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



; --------------------- GUI1  --------------------------


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
GuiControl, 1:Move, btnExportFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnCheckExportFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, strMultiPurpose, % "W" . (A_GuiWidth - 305)
GuiControl, 1:Move, btnMultiPurpose, % "X" . (A_GuiWidth - 190)

GuiControl, 1:Move, lvData, % "W" . (A_GuiWidth - 20) . " H" . (A_GuiHeight - 190)

return



GuiClose:
ExitApp



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


UpdateCurrentHeader:
Gui, 1:Submit, NoHide
strCurrentHeader := GetListViewHeader(strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
strCurrentHeaderEscaped := StrEscape(strCurrentHeader)
GuiControl, 1:, strRenameEscaped, %strCurrentHeaderEscaped%
GuiControl, 1:, strSelectEscaped, %strCurrentHeaderEscaped%
GuiControl, 1:, strOrderEscaped, %strCurrentHeaderEscaped%
if (radFixed)
	Gosub, ClickRadFixed
else if (radExpress)
	Gosub, ClickRadExpress
return



CheckOneRow:
if (LV_GetCount("Selected") = 1)
{
	MsgBox, 35, %strApplicationName% - One record selected,
	(
	Only one record is selected. Do you want to save only this record?
	
	Yes: Only one record will be saved.
	No: All records will be saved.
	)
	IfMsgBox, No
		LV_Modify(0, "Select") ; select all records
	IfMsgBox, Cancel
		return
}
return



ExportFixed:
if !DelimitersOK(3)
	return
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", blnProgress = 0])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
; strFieldDelimiter3 et strFieldEncapsulator3 pour l'écriture de l'entête seulement
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
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
		Oops("""" . intThisWidth . """ in field #" . A_Index . " """ . strThisName . """ must be an integer number.")
		return
	}
}
StringTrimRight, strFieldsName, strFieldsName, 1 ; remove extra delimiter
StringTrimRight, strFieldsWidth, strFieldsWidth, 1 ; remove extra delimiter
if (radSaveMultiline)
	strEolReplacement := ""
else
	strEolReplacement := strEndoflineReplacement
; ObjCSV_Collection2Fixed(objCollection, strFilePath, strFieldsWidth, blnHeader := 0, strFieldOrder := "", blnProgress := 0
;	, blnOverwrite := 0, strFieldDelimiter := ",", strEncapsulator := """", strEndOfLine := "`r`n", strEolReplacement := "")
ObjCSV_Collection2Fixed(obj, strFileToExport, strFieldsWidth, radSaveWithHeader, strFieldsName, 1, blnOverwrite
	, strRealFieldDelimiter3, strFieldEncapsulator3, , strEolReplacement)
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return



ExportHTML:
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
; ObjCSV_Collection2HTML(objCollection, strFilePath, strTemplateFile [, strTemplateEncapsulator = ~
;	, blnProgress = 0, blnOverwrite = 0])
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
obj := ObjCSV_ListView2Collection("1", "lvData", , , , 1)
SplitPath, strFileToExport, , strOutDir
strExpressTemplateTempFile := strOutDir . "\" . GUID() . ".TMP"
strExpressTemplate := strTemplateDelimiter . "ROWS" . strTemplateDelimiter
strExpressTemplate := strExpressTemplate  . strMultiPurpose
strExpressTemplate := strExpressTemplate  . strTemplateDelimiter . "/ROWS" . strTemplateDelimiter
strExpressTemplate := StrUnEscape(strExpressTemplate)
FileAppend, %strExpressTemplate%, %strExpressTemplateTempFile%
if FileExist(strExpressTemplateTempFile)
{
	run, notepad.exe %strExpressTemplateTempFile%
	; ObjCSV_Collection2HTML(objCollection, strFilePath, strTemplateFile [, strTemplateEncapsulator = ~
	;	, blnProgress = 0, blnOverwrite = 0])
	ObjCSV_Collection2HTML(obj, strFileToExport, strExpressTemplateTempFile, strTemplateDelimiter, 0, 1)
	FileDelete, %strExpressTemplateTempFile%
	if FileExist(strFileToExport)
	{
		GuiControl, 1:Show, btnCheckExportFile
		GuiControl, 1:+Default, btnCheckExportFile
		GuiControl, 1:Focus, btnCheckExportFile
	}
}
else
	Oops("An error occured while creating a temporary template file.")
obj := ; release object
return



; --------------------- FUNCTIONS --------------------------


GetListViewHeader(strRealFieldDelimiter, strFieldEncapsulator)
{
	strHeader := ""
	Loop, % LV_GetCount("Column")
	{
		LV_GetText(strColumnHeader, 0, A_Index)
		strHeader := strHeader . Format4CSV(strColumnHeader, strRealFieldDelimiter, strFieldEncapsulator) . strRealFieldDelimiter
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



StrMakeEncodedFieldDelimiter(strConverted)
{
	StringReplace, strConverted, strConverted, `t, t, All ; Tab (HT)
	StringReplace, strConverted, strConverted, `n, n, All ; Linefeed (LF)
	StringReplace, strConverted, strConverted, `r, r, All ; Carriage return (CR)
	StringReplace, strConverted, strConverted, `f, f, All ; Form feed (FF)
	return strConverted
}



Help(strTitle, strMessage)
{
	Gui, 1:+OwnDialogs 
	MsgBox, 0, %strApplicationName% (%strApplicationVersion%) - %strTitle% Help, %strMessage%
}



Oops(strMessage)
{
	Gui, 1:+OwnDialogs 
	MsgBox, 48, %strApplicationName% (%strApplicationVersion%), %strMessage%
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
		; MsgBox, % "intNbRows: " . intNbRows . "`nintOriginalHeight: " . intOriginalHeight 
		;	. "`nintHeightOneRow: " . intHeightOneRow . "`nintNewHeight: " . intNewHeight
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
		MsgBox, 35, %strApplicationName% - File exists, 
		(
		File exists:`n%strFileName%
		
		Do you want to overwrite this file?
		
		Yes: The file will be overwritten.
		No: Data will be added to the existing file.
		)
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
	strNewName := strOutDir . "\" . strOutNameNoExt . strNote . "." . strExtension
	if FileExist(strNewName)
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



DelimitersOK(intTab)
{
	if (strFieldDelimiter%intTab% = strFieldEncapsulator%intTab%) or (strFieldDelimiter%intTab% = "") or (strFieldEncapsulator%intTab% = "")
	{
		Oops("Field delimiter and field encapsulator in tab " . intTab . " cannot be the same character and cannot be empty.")
		GuiControl, 1:Choose, tabCSVBuddy, %intTab%
		if (strFieldDelimiter%intTab% = strFieldEncapsulator%intTab%) or (strFieldDelimiter%intTab% = "")
			GuiControl, Focus, strFieldDelimiter%intTab%
		else
			GuiControl, Focus, strFieldEncapsulator%intTab%
		return false
	}
	else
		return true
}



NewDelimiterOrEncapsulatorOK(strChecked)
{
	global strCurrentHeader
	global strCurrentFieldDelimiter
	global strCurrentFieldEncapsulator
	objCurrentHeader := ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
	Loop, % objCurrentHeader.MaxIndex()
		If InStr(objCurrentHeader[A_Index], strChecked)
			return false
	return true
}



PositionInArray(strChecked, objArray)
{
	Loop, % objArray.MaxIndex()
		If (objArray[A_Index] = strChecked)
			return A_Index
	return 0
}



