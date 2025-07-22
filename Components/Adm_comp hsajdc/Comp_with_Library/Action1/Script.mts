Dim n, fso, file, folder, filePath, numberOfLines, threeN,prime

num = Parameter("number")

If num = "10" Then
	Reporter.ReportEvent micFail, "Validation Failed", "Not a good number"
End If

If num = "" Then
	Reporter.ReportEvent micFail, "Validation Failed", "Cannot be empty"
End If

If num <> "10" And num <> "" Then
	n = CInt(num)
	prime = IsPrime(n)
	If Not prime Then
		Reporter.ReportEvent micFail, "Validation Failed", "Expected a prime number" & ", Actual: " & n
	Else
		eporter.ReportEvent micPass, "Validation passed", "Ia prime number: " & n
	End If
End If


