global lDocCopyrightText := "CSV Buddy - Copyright (C) 2013  Jean Lalonde`n`nThis software is provided 'as-is', without any express or implied warranty.  In no event will the authors be held liable for any damages arising from the use of this software.`n`nPermission is granted to anyone to use this software for any purpose,  including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:`n`n1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation would be appreciated but is not required.`n2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.`n3. This notice may not be removed or altered from any source distribution.`n`nJean Lalonde, <A HREF=""mailto:ahk@jeanlalonde.ca"">ahk@jeanlalonde.ca</A>"
global lDocCopyrightTitle := "Copyright"
global lDocDesc2000 := "Even if the CSV file format is a widely accepted standard, it is still found in multiple flavors. In some implementations, fields are separated by comma. Others are delimited with tab, semi-colon or a variety of characters depending on the OS. Most CSV file records stand on one line. However, some programs export multi-line data with line breaks inside fields (try to load in Excel a CSV export from Outlook or Gmail contacts with multi-line notes text fields). Many programs will have a hard time importing these various variations of the CSV format.<BR /><BR />The freeware CSV Buddy helps you make your CSV files ready to be imported by a variety of software. Load files with all sorts of field delimiters (auto-detection of comma, tab, semi-colon, etc.). Field containing delimiters or line breaks can be embedded in various encapsulators (double-quotes, single-quotes, pipes or any character). Get field names from the file's header (first line) or set your own column titles. Load data with line-breaks.<BR /><BR />Rename, select or reorder fields. In a grid, edit or delete records. Sort them on alphabetical or numeric values (integer or float). Save all or selected rows to a new file using any delimiters, with header or not.  Replace line breaks in data fields with a marker to make your file ready to load in software (like MS-Excel) that can only load single-line fields.<BR /><BR />Export your data to fixed-width files with specific width for each field, truncating data or padding it with spaces. Export to HTML using your own template with markers to insert your data fields in the web page. Also export to XML standard format.<BR /><BR />CSV Buddy can load files having up to 200 fields (columns) and cells with up to 8191 characters. With the 32-bits version, file loading is limited by available physical memory. Tests were successful with files over 100 MB. With the 64-bits version, there is no limitation to file size thanks to virtual RAM. However, loading and saving time will increase as files get huge (in the hundreds of megs)."
global lDocDesc450 := "CSV Buddy helps you make your CSV files ready to be imported by a variety of software. You can load files with all sort of field delimiters (comma, tad, semi-colon, etc.) and encapsulators (double/single-quotes or any other character). Convert line breaks in data field making your file XL ready. Rename/reorder fields, edit records, save with any delimiters and export to fixed-width, HTML templates or XML formats. Freeware."
global lDocDescription := "Description"
global lDocDocumentation := "Documentation"
global lDocFeatures := "Features"
global lDocFeaturesList := "1) Load CSV file to a list view`n- Select and preview the file to load`n- Get field names from the file header (first line of the file)`n- Set the header of your choice to customize field names`n- Use any single-character custom field delimiter (comma, tab, semi-colon, etc.)`n- Use any single-character custom field encapsulator (double-quotes, single-quoted, etc.) to embed field containing a delimiters or line breaks`n- Auto-detection of field delimiter (comma, tab, semi-colon, colon, pipe or tilde) and field encapsulator (double-quote, single-quote, tilde or pipe) from file's first line`n- Load multi-line fields (do not consider a line break between double-quotes as the end of a record)`n- Restore line breaks within fields by replacing a temporary character of your choice (like "") with line break`n- Load to file into a list allowing these features:`n`t- Sort rows on any field by clicking on column headers`n`t- Sorting type: alphabetical, numeric integer or numeric float, ascending or descending`n`t- Double-click on a row to edit a record in a dialog box (field names are uses as form labels)`n`t- Right-click anywhere in the list view to select all rows, deselect all rows, reverse selection, edit a record or delete selected rows.`n`n2) Edit columns`n- Rename fields by entering a delimited string with the new field names`n- Select fields to keep in the list view by entering a delimited string with the names of these fields`n- Order fields by entering a delimited string with the names of the fields in the desired order`n`n3) Save list view to CSV file`n- Choose destination file name (default to original name + " (1)" or " (2)", etc.)`n- Check the content of the destination file if it exists`n- Overwrite or append data if destination file exists`n- Set any single-character as field delimiter in the destination file`n- Set any single-character as field encapsulator in the destination file`n- Save the file with or without a CSV header (first line of the file with field names)`n- Save multi-line fields (embedded with the encapsulator character)`n- Convert multi-line fields to single-line by replacing line breaks within fields with a replacement character of your choice (like "")`n- Save rows in the order they appear in the list view`n- Save all rows or only selected rows`n`n4) Export`n- Export to fixed-with file`n`t- Choose fixed-width for each field`n`t- Truncate data or pad with space`n- Export to HTML using an HTML template`n- Export to XML`n- Export to other format using a custom row template`n"
global lDocHelp := "Help"
global lDocHelpIntro := "Throughout CSV Buddy, you'll find help capsules available by clicking the <CODE>?</CODE> button. You will find below a compilation of help messages for each function in their logical sequence of use. Read the whole thing now for a quick overview of CSV Buddy or, if you prefer, read it as you need it in each of the four tabs of the CSV Buddy."
global lDocInstallationDetail := "<P>Absolutely free to download and use, for personal or commercial use.</P><OL><LI>Download <A HREF=""http://www.jeanlalonde.ca/ahk/csvbuddy/csvbuddy.zip"">csvbuddy.zip</A></LI><LI>There is no software to install. Just extract the zip file content to the folder of your choice.</LI><LI>Run the .EXE file from this folder (choose the 32-bits or 64-bits version depending on your system). Make sure it will run with <EM>read/write</EM> access to this folder.</LI><LI>At your convenience, create a shorcut on your Desktop or your Start menu.</LI></OL><P>CSV Buddy can be freely distributed over the internet in an unchanged form.</P>"
global lDocInstallationTitle := "Installation"
global lDocIntro := "First things first..."
global lDocSupportText := "<ul><li>Online support: <a href=""https://www.facebook.com/CSVBuddy"" target=""_blank"">Facebook Page</a></li><li>Email: <a href=""mailto:ahk@jeanlalonde.ca"">ahk@jeanlalonde.ca</a></li><li>Bug reports: <a href=""https://github.com/JnLlnd/CSVBuddy/issues"" target=""_blank"">GitHub issues</a> (you will need a GitHub account)</li></ul>"
global lDocSupportTitle := "Getting support"
global lDocTab := "tab"
