# VBA-Macros
VBA Macros I Created

These are various VBA modules I created to streamline and increase efficiency of work tasks.   
VBA Macro (.bas file) would be contained in a template file and bound to a keyboard shortcut for the team to execute function.  
All confidential information has been modified.

## EXCEL_RENAME_FOLDERS.BAS
* Macro that accompanies an Excel spreadsheet that contains 3 columns (DIR Path, Old (current) folder names, New (desired) folder names).  When ran, it will rename folders in-bulk.  Easy to use if folders are collectively renamed to be successive titles.

## LAX_OWNING_FLIGHTS.BAS
* Macro that accompanies an Excel template or blank.  Data is pulled from a flight departure table and data is autmotically sorted, trimmed and conditionally formatted based on "focus flight" parameters specified.

## OpsDocs_TEMPLATE_OUTLOOK_AUTOMATION.BAS
* Macro that lives in a master Word .dotm template that is in root folder with other daily operational word templates.  Upon creation of a new word document, user is prompted to make a selection to import desired template.  File title is automatically formatted and saved to pre-defined directory.  When document is complete and ready to be sent, user uses a shortcut hotkey (which another VBA Macro is bound to), to export document content to Outlook email that's pre-formatted with correct distribution list, subject and attachments.
