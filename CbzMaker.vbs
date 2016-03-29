Set shell = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

path = Wscript.ScriptFullName
Set this = FSO.GetFile(path)
parent = FSO.GetParentFolderName(this)
Set dir = FSO.GetFolder(parent)
'Set dir = FSO.GetFolder(parent & "\test\")
Set colFolders = dir.SubFolders

Set reg = New RegExp
reg.IgnoreCase = True
reg.Global = True

Function createCbz()
	For Each folder in colFolders
		name = Replace(folder.Name, " ", "_")
		name = Replace(name, "-", "_")
		comic = true
		For Each file in folder.Files
			'Wscript.Echo file.Name
			If InStr(file.Name, ".png") = 0 AND InStr(file.Name, ".gif") = 0 AND InStr(file.Name, ".jpg") = 0 AND InStr(file.Name, ".jpeg") = 0 Then
				'Wscript.Echo "False"
				comic = false
			End If
		Next
		If comic = true Then 
			shell.run "cmd.exe /C cd " & folder.Path & " && 7z a -tzip -mx7 " & Chr(34) & folder.Path & ".cbz" & Chr(34) & " *" & Chr(34), 0
		End If
	Next
End Function

Function Rename(collection)
	For Each item in collection
		'ensures not editing scripts title
		If InStr(item.Name, WScript.ScriptName) = 0 Then
			'changes all archives to comic book archives
			name = Replace(item.Name, ".zip", ".cbz")
			name = Replace(name, ".rar", ".cbr")
			name = Replace(name, ".7z", ".cb7")
			'removes all tags
			reg.Pattern = "\[.*?\]|\{.*?\}"
			name = reg.Replace(name, "")
			'removes comic events
			reg.Pattern = "\(C[0-9][0-9]?\)|\(COMIC.*?\)|\(SC.*?\)"
			name = reg.Replace(name, "")
			'removes english tags
			reg.Pattern = "\(eng.*?\)|\(ENG.*?\)|\(Eng.*?\)"
			name = reg.Replace(name, "")
			'removes certain uploader names
			reg.Pattern = "=.*=|~.*~"
			name = reg.Replace(name, "")
			'removes double+ spaces that result from previous edits
			reg.Pattern = "\s{2,}"
			name = reg.Replace(name, " ")
			reg.Pattern = "\s+\."
			name = reg.Replace(name, ".")
			name = Trim(name)
			'Wscript.Echo name
			
			'sets item name if not same
			On Error Resume Next
				If Not item.Name = name Then
					item.Name = name
				End If
			If Err.Number <> 0 Then
				name = Replace(name, ".", "(1).")
				If Not item.Name = name Then
					item.Name = name
				End If
			End If
		End If
	Next
End Function

createCbz()

Set colFiles = dir.Files

Rename(colFiles)
