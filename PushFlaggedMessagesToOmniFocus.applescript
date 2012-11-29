on run
	
	tell application "Mail"
		repeat with _acct in imap accounts
			--Look For Flagged Messages in the INBOX
			set _acct_name to name of _acct
			if _acct is enabled then
			
			set _inbox to _acct's mailbox "INBOX"
			if _acct is enabled then

				set _msgs_to_capture to (a reference to ¬
					(every message of _inbox ¬
						whose flagged status is true))
			
				repeat with eachMessage in _msgs_to_capture
					set theStart to missing value
					set theDue to missing value
					set theOmniTask to missing value
				
					set theTitle to the subject of eachMessage
					set theNotes to the content of eachMessage
				
					set theCombinedBody to "message://%3c" & message id of eachMessage & "%3e" & return & return & theNotes
				
					tell application "OmniFocus"
						tell default document
							set newTaskProps to {name:theTitle}
							if theStart is not missing value then set newTaskProps to newTaskProps & {start date:theStart}
							if theDue is not missing value then set newTaskProps to newTaskProps & {due date:theDue}
							if theCombinedBody is not missing value then set newTaskProps to newTaskProps & {note:theCombinedBody}
						
							set newTask to make new inbox task with properties newTaskProps
						end tell
					end tell
				
					set flagged status of eachMessage to false
				
				end repeat
			end if
		end repeat
	end tell
	
end run