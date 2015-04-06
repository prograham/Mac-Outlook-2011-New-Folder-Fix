tell application "Microsoft Outlook"
	
	display dialog "Do you want to create the folder on the Server or On My Computer?" with title "Prograham's Mac Outlook New Folder Fix" buttons {"Cancel", "On the Server", "On My Computer"} default button "On My Computer" cancel button "Cancel" with icon caution
	
	if button returned of result is "On the Server" then
		
		set mycustomfolderName to text returned of (display dialog "Enter New Folder Name:" with title "Prograham's Mac Outlook New Folder Fix" default answer "01 Brad's Untitled New Folder" buttons {"Cancel", "Take THAT Microsoft!!"} default button "Take THAT Microsoft!!" cancel button "Cancel" with icon caution)
		
		make new mail folder with properties {name:mycustomfolderName}
		
	else if button returned of result is "On My Computer" then
		
		set mycustomfolderName to text returned of (display dialog "Enter New Folder Name:" with title "Brad's Outlook New Folder Fix" default answer "01 Brad's Untitled New Folder" buttons {"Cancel", "Take THAT Microsoft!!"} default button "Take THAT Microsoft!!" cancel button "Cancel" with icon caution)
		
		make new mail folder in on my computer with properties {name:mycustomfolderName}
		
	end if
	
end tell
