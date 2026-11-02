#!/usr/bin/osascript
-- Runs the StrikeThroughPhrasesInDocs macro from prepare_protocols.vba.
-- The macro must be installed in Word (Normal.dotm or a .dotm in Word's Startup folder).
tell application "Microsoft Word"
	activate
	run VB macro macro name "StrikeThroughPhrasesInDocs"
end tell
