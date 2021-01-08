
Function checkrc() {
	Param(
		[int] $rc,
		[int] $expected
	)
	
	if ($rc -eq $expected) {
		$result = @{
			resultColor = "Green"
			message = "$rc equals $expected"
			result = "PASSED"
			rc = $rc
		}
	} else {
		$result = @{
			resultColor = "Red"
			message = "Received : $rc. Expected : $expected"
			result = "FAILED"
			rc = $rc
		}
	}
	return $result
}

Function runTest()
{
	Param(
		[HashTable]$testArgs
	)
	# testName = $testName
	# argList = @{
		# AccountName = $user
		# AccountPassword = $pass
		# Environment = $stenv
		# mmdFrom = $mmdFrom
		# mmdFilePath = $mmdfile
	# }
	# expected = $expected
	
	Write-Host -ForegroundColor Yellow "`r`n# ----------------Starting  : $($testArgs.testName)"
	# Start-Process -FilePath powershell -ArgumentList "$ScriptFullName $argList" -Wait -WindowStyle Hidden
	$runTestArgs = $($testArgs.argList)
	if ($runTestArgs.ContainsKey('AccountPassword') -and (($null -eq $runTestArgs['AccountPassword']) -or ("" -eq $runTestArgs['AccountPassword']))) {
		$runTestArgs.Remove('AccountPassword')
	}
	# $runTestArgs | Out-String
	& $scriptFullName @runTestArgs
	$rc = $LASTEXITCODE
	$result = checkrc $rc $($testArgs.expected)
	Write-Host -ForegroundColor $result.resultColor "Finished '$($testArgs.testName)' with result : $($result.message) : $($result.result)"
	return( @{
		testName = $($testArgs.testName)
		rc = $rc
		result = $($result.result)
	})
}