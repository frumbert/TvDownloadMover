dim skiplist, slist, directorylist, dlist, shares, tmp, found, dsplit, myloc, i, j, k
dim skip, yr, dirmatch, t, fileFrom, fileTo, copyCommand, filesys, oshellapp, kAdded
dim root

'' if you want to see the log and pause on completion
const debuggy = false

root = "\\QNAP\Qdownload\transmission\completed\"

Set oShellApp = WScript.CreateObject("WScript.Shell")

set filesys = CreateObject("Scripting.FileSystemObject")
t = timer
logmsg "process begun" 

'' the various folders where tv shows are stored
shares = Split("\\pvr\c\,\\pvr\d\tv\,\\pvr\e\tv\",",")

if filesys.FileExists(root & "skip.list") then
	set tmp = filesys.OpenTextFile(root & "skip.list", 1, false)
	slist = tmp.ReadAll
	skiplist = split(trim("" & slist), vbNewLine)
	tmp.close
else
	Set tmp = filesys.CreateTextFile(root & "skip.list", True)
	tmp.writeline "run.vbs"
	tmp.writeline "directory.list"
	tmp.writeline "skip.list"
	tmp.close
	redim preserve skiplist(3) 
	slist = tmp.ReadAll
	skiplist = split(trim("" & slist),vbNewLine)
	tmp.close
end if

'' make an array of all the directories we know about and their shortname to make matching easier
k = 0
kAdded = false
for i = 0 to ubound(shares)
	set fold = filesys.getfolder(shares(i))
	for each fol in fold.subfolders
		k = k + 1
	next
next
redim directorylist(k)
k = 0
for i = 0 to ubound(shares)
	set fold = filesys.getfolder(shares(i))
	for each fol in fold.subfolders
		directorylist(k) = shares(i) & fol.name & ":" & cleanDir(fol.name)
		k = k + 1
	next
next
logmsg "finished precaching " & ubound(directorylist) & " destinations" 

'' process the files in source to find their matching directories, if they exist
set myloc = filesys.getfolder(root)
for each fil in myloc.files
	skip = false
	for j = 0 to ubound(skiplist)
		if lcase(fil.name) = lcase(skiplist(0)) or left(fil.name, 1) = "." or trim(fil.name) = "" then
			skip = true
			exit for
		end if
 	next
	if not skip then
		yr = ""
		sname = ""
		name = getname(fil.name, sname, yr)
		if name > " " then
			dirmatch = ""
			for j = 0 to ubound(directorylist)
				if instr(directorylist(j),":") > 0 then
					dsplit = split(directorylist(j),":")
					' some shows have (2011) etc after their folder name for tvdb matching reasons
					if dsplit(1) = trim(name & yr) or dsplit(1) = trim(name) then
						dirmatch = dsplit(0)
						exit for
					end if
				end if
			next
			if dirmatch > "" then
				fileFrom = quote(unicodeToAscii(myloc)) ' wrap quotes around the path, if needed
				fileTo = quote(unicodeToAscii(dirmatch)) ' remove non-printing characters
				copyCommand = "robocopy " & fileFrom & " " & fileTo & " " & fil.name & " /R:3 /W:10 /MOV /NS /NC /NFL /NDL /NP /TEE /NJH /NJS"
				logmsg copyCommand
				oShellApp.run copyCommand
			else
				logmsg "no matching directory found for: " & fil.name
			end if
		else
			logmsg "file not processed: " & name & " (" & fil.name & ")"
		end if
	end if
next

logmsg "process complete"

if debuggy then
	wscript.echo "press return to quit"
	foo = WScript.StdIn.ReadLine
end if

' --------------------------------------------------------------------------------------------------------

'' if a string has a space, put double quotes around it
function quote(byval msg)
	if instr(msg," ") > 0 then
		quote = chr(34) & msg & chr(34)
	else
		quote = msg
	end if
end function

'' turn
''	666.Park.Avenue.(2012).S01E04.HDTV.x264-LOL.[VTV].mp4
'' into
''	666parkavenue
function getname(byval name, byref spaceName, byref yeer)
Dim ret, rex, s, z, c, q
	spaceName = ""
	yeer = ""
	q = ""

	ret = lcase(name)
	set rex = new regexp
	rex.global = true

	' replace name.name.2012.s01e02 with name.name.s01e02
	rex.pattern = "[.]\d{4}[.]"
	set matches = rex.execute(ret)
	if matches.count > 0 then
		yeer = replace(matches(0).value, ".", "")
	end if
	ret = rex.replace(ret,".")

	' replace .US. and .UK. with .
	rex.pattern = "[.]u[sk][.]"
	ret = rex.replace(ret,".")

	' get name before season/episode
	rex.pattern = "[s]\d{2}[e]\d{2}"
	if rex.test(ret) then
		s = left(ret, instr(ret, ".s"))
		spl = split(replace(s,".", " ")," ")
		for z = 0 to ubound(spl)
			spaceName = spaceName & ucase(left(spl(z),1)) & mid(spl(z),2) & " "
		next
		spaceName = trim(spaceName)
		getname = replace(s, ".","")
	end if
	set rex = nothing
end function

function cleanDir(byval name)
Dim ret, rex
	ret = name
	set rex = new regexp
	rex.global = true
	rex.pattern = "[(]\d{4}[)]" ' date
	ret = rex.replace(ret,"")
	rex.pattern = "[^a-zA-Z0-9]" ' alphanumeric
	ret = rex.replace(ret,"")
	cleanDir = lcase(ret)
	set rex = nothing
end function

sub logmsg(msg)
	if debuggy then wscript.echo hhmmss(Timer - t) & ": " & msg
end sub

Function hhmmss(byval i)
Dim hr, min, sec, remainder

	elap = Int(i) 'Just use the INTeger portion of the variable
	hr = elap \ 3600 '1 hour = 3600 seconds
	remainder = elap - hr * 3600
	min = remainder \ 60
	remainder = remainder - min * 60
	sec = cint(0 + remainder)
	min = right("00" & min,2)
	sec = right("00" & sec,2)
	If hr = 0 Then
		hhmmss = min & ":" & sec
	Else
		hhmmss = hr & ":" & min & ":" & sec
	End If
End Function

'' handle if fso has been reading unicode encodings for filenames
'' e.g. source is on a linux box with smb
Function unicodeToAscii(sText)
  Dim x, aAscii, ascval, l
  l = len(sText)
  If l = 0 Then Exit Function
  redim aAscii(l)
  For x = 1 To l
    ascval = AscW(Mid(sText, x, 1))
    If (ascval < 0) Then
      ascval = 65536 + ascval ' http://support.microsoft.com/kb/272138
    End If
    aAscii(x) = chr(ascval)
  Next
  unicodeToAscii = join(aAscii,"")
End Function

set filesys = nothing
