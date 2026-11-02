#!/usr/bin/osascript
-- Runs the RemoveFirstWord macro from column_with_spaces.vba.
-- Usage: osascript column_with_spaces.applescript [workbook.xlsx]
-- The macro must be installed in Excel (e.g. Personal Macro Workbook). If a workbook path
-- is given it is opened first; the macro works on the active workbook.
on run argv
	tell application "Microsoft Excel"
		activate
		if (count of argv) > 0 then open (POSIX file (item 1 of argv))
		run VB macro "RemoveFirstWord"
	end tell
end run
