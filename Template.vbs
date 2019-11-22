'Feature: open disc filename
'Feature: open QC file in resources
'Feature: open files from GIT

sub include ( sFileSpec )
	with CreateObject("Scripting.FileSystemObject")
		executeGlobal .OpenTextFile(sFileSpec).ReadAll()
	end with
end sub

include "filename"