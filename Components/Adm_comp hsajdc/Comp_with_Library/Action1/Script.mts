Dim n, fso, file, folder, filePath, numberOfLines, threeN

num = Parameter("number")

If num = "10" Then
	Reporter.ReportEvent micFail, "Validation Failed", "Not a good number"
End If

If num = "" Then
	Reporter.ReportEvent micFail, "Validation Failed", "Cannot be empty"
End If

If num <> "10" And num <> "" Then
	n = CInt(num)
	fourN = 4 * (n - 1)
	folder = "C:\Users\Administrator\Documents\TestGroundUFT"
	filePath = folder & "\file" & CStr(n) & ".txt"
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.CreateTextFile(filePath, True)
	Set shell = CreateObject("WScript.Shell")
	
	
	PrintInFile n, file
	file.Close
	
	OpenFileInNotepad shell, filePath
	
	numberOfLines = GetNumberOfLines(fso, filePath)
	
	If numberOfLines <> fourN Then
		Reporter.ReportEvent micFail, "Validation Failed", "Expected: " & fourN & ", Actual: " & numberOfLines
	Else 
		Reporter.ReportEvent micPass, "Validation passed", "Values match: " & numberOfLines
	End If
End If


