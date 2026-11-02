#!/usr/bin/osascript
-- Runs the ConvertPDFsToDocx macro from protocol_saver.vba.
-- The macro must be installed in Word (Normal.dotm or a .dotm in Word's Startup folder).
tell application "Microsoft Word"
	activate
	run VB macro macro name "ConvertPDFsToDocx"
end tell
