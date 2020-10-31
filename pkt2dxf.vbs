main

sub main
	if Wscript.Arguments.Count = 0 then 'if no argument is passed
		MsgBox "Drag one or more PKT file(s) onto the script." & vbLf & "A DXF file will be created." & vbLf & "The PKT file remains unchanged." & vbLf & "The source path is the same as the target path."
		exit sub
	end if
	
	For Each Argument in Wscript.Arguments
		'determine file extension of the arguments
		ext = Mid(Argument, InstrRev(Argument, ".")+1, Len(Argument))
		
		'determine source path
		source_path = Mid(Argument, 1, InStrRev(Argument, "\"))
		
		if ext <> "pkt" AND ext <> "PKT" Then
			MsgBox """" & Argument & """ is not a PKT file"""
			exit sub
		end if
		
		export_name = Mid(Argument, InStrRev(Argument, "\") + 1, Len(Argument))
		export_name = Mid(export_name, 1, InStrRev(export_name, ".") - 1)
		
		'target path
		export_file_path = source_path & export_name & ".DXF"
		
		Set fs = CreateObject("Scripting.FileSystemObject")
		Set export_file = fs.CreateTextFile(export_file_path, True)
		
		'write header data to target file
		export_file.WriteLine("  0")
		export_file.WriteLIne("SECTION")
		export_file.WriteLine("  2")
		export_file.WriteLine("ENTITIES")
		
		'open PKT source file
		Set source_file = fs.OpenTextFile(Argument)
		
		ln = source_file.ReadLine
		
		'determine separator (spaces, semicolons or tabs)
		If InStr(ln, " ") > 0 Then
			seperator = " "
		end if
		If InStr(ln, ";") > 0 Then
			seperator = ";"
		end if
		If InStr(ln, vbTab) > 0 Then
			seperator = vbTab
		end if
		
		Do Until source_file.AtEndOfStream
			x1 = ""
			y1 = ""
			
			'get x/y values from current line
			items = Split (ln, seperator)
			x1 = items(0) 'first value in items
			y1 = items(UBound(items)) 'last value in items
			
			if x1 <> "" Then
				'write line data 
				export_file.WriteLine("  0")
				export_file.WriteLine("LINE")
				
				'optional
				'export_file.WriteLine("100")
				'export_file.WriteLine("AcDbEntity")
				
				'8 = Layer (also optional)
				export_file.WriteLine("  8")
				export_file.WriteLine("1") ' Layer name
				
				'optional
				'export_file.WriteLine("100")
				'export_file.WriteLine("AbDbLine")
				
				'coordinates
				export_file.WriteLine(" 10")
				export_file.WriteLine(x1)
				export_file.WriteLine(" 20")
				export_file.WriteLine(y1)
			end if
			
			'read next line
			ln = source_file.ReadLine
			
			x2 = ""
			y2 = ""
			
			items = Split(ln, seperator)
			x2 = items(0) 'first value in items
			y2 = items(UBound(items)) 'last value in items
			
			if x1 <> "" Then
				export_file.WriteLine(" 11")
				export_file.WriteLine(x2)
				export_file.WriteLine(" 21")
				export_file.WriteLine(y2)
			End if
		Loop
		
		'write file closure data
		export_file.WriteLine("  0")
		export_file.WriteLine("ENDSEC")
		export_file.WriteLine("  0")
		export_file.WriteLine("EOF")
		
		source_file.Close
		export_file.Close
		
		ln = ""
	Next
end sub