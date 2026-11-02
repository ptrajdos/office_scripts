#!/usr/bin/osascript
-- Runs the MakeDocsSummary macro from protocol_info_extractor.vba.
-- The macro must be installed in Word (Normal.dotm or a .dotm in Word's Startup folder).
tell application "Microsoft Word"
	activate
	run VB macro macro name "MakeDocsSummary"
end tell
