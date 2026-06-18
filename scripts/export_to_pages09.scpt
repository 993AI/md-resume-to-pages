on run argv
	if (count of argv) is not 2 then
		error "Usage: osascript export_to_pages09.scpt input.docx output.pages"
	end if
	set inputPath to item 1 of argv
	set outputPath to item 2 of argv
	set inputFile to POSIX file inputPath
	set outputFile to POSIX file outputPath
	
	using terms from application "/Applications/Pages.app"
		tell application "/Applications/Pages.app"
			activate
			open inputFile
			delay 4
			set theDoc to front document
			export theDoc to outputFile as Pages 09
			close theDoc saving no
		end tell
	end using terms from
end run
