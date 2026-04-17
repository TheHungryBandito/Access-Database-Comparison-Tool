function Compare-AccessDB {
	<#
		.SYNOPSIS
			Compares two Microsoft Access Databases (.mdb) Databases to ensure Tables. Columns, and Values are identical
			If differences are found or -IncludeEqual is specified, the results will be viewed in Grid View

		.DESCRIPTION
			1. Connects to two different .mdb databases.
			2. Queries and compares table names.
			3. Queries and compares all Column/value pairs in all tables.
			4. Outputs the differences.

		.PARAMETER ReferencePath
			Full path to .mdb file to compare
		
		.PARAMETER DifferencePath
			Full path to .mdb file to compare against
		
		.PARAMETER OutputCSVPath
			If specified, outputs the comparison results to this full path
		
		.PARAMETER IncludeEqual
			Passes IncludeEqual to comparison to show all comparisons that are equal

		.PARAMETER Configuration
			Specifies which PowerShell configuration to use for the comparison. 
			This is necessary as the Access DB drivers are only available in either 32-bit or 64-bit versions of PowerShell depending on how they were installed. 
			If you receive errors about missing drivers, try switching the configuration.
			Default: Runs in the current PowerShell configuration
			32-bit: Runs the comparison in a new 32-bit PowerShell process
			64-bit: Runs the comparison in a new 64-bit PowerShell process

		.EXAMPLE
			This example checks two DBs, outputs the results to a file and displays the results in a grid view.
			Compare-AccessDB -ReferencePath $PathA -DifferencePath $PathB -OutputCSVPath $OutPath

		.EXAMPLE
			This example checks two DBs, outputs the results to a file and displays the results in a grid view including values that are equal.
			Compare-AccessDB -ReferencePath $PathA -DifferencePath $PathB -OutputCSVPath $OutPath -IncludeEqual
		
		.EXAMPLE
			This example checks two DBs forcing 32-bit PowerShell to use 32-bit Access DB drivers, outputs the results to a file and displays the results in a grid view.
			Compare-AccessDB -ReferencePath $PathA -DifferencePath $PathB -OutputCSVPath $OutPath -Configuration "32-bit"

		.EXAMPLE
			This example checks two DBs forcing 64-bit PowerShell to use 64-bit Access DB drivers, outputs the results to a file and displays the results in a grid view.
			Compare-AccessDB -ReferencePath $PathA -DifferencePath $PathB -OutputCSVPath $OutPath -Configuration "64-bit"

		.OUTPUTS
			Displays comparison results in a grid view and outputs the results to the OutputCSVPath (If specified)
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true)]
		[string]$ReferencePath,
		[Parameter(Mandatory=$true)]
		[string]$DifferencePath,
		[Parameter(Mandatory=$false)]
		[string]$OutputCSVPath = "",
		[Parameter(Mandatory=$false)]
		[switch]$IncludeEqual=$false,
		[Parameter(Mandatory=$false)]
		[ValidateSet("Default","32-bit","64-bit")]
		[string]$Configuration = "Default"
	)

	$processPath = "None"
	if ($Configuration -eq "64-bit") {
		$processPath = "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe"
	} elseif ($Configuration -eq "32-bit") {
		$processPath = "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
	}

	# Run in specified configuration for PowerShell to use Access Engine DB drivers
	if ($Configuration -ne "Default") { 
		try {
			Write-warning "Configuration specified for $Configuration. Entering $Configuration PowerShell to use $Configuration Access DB drivers. Creating temporary script file to run comparison in new process."
			$functionDefinition = $(Get-Command Compare-AccessDB | Select-Object -ExpandProperty Definition)
			$tempFile = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.ps1'
			$functionDefinition | Set-Content -Path $tempFile
			$processParams = @{
				FilePath = $processPath;
				ArgumentList = "-ExecutionPolicy Bypass -File $tempFile -ReferencePath $ReferencePath -DifferencePath $DifferencePath"
				NoNewWindow = $true;
				Wait = $true;
			}

			if (![string]::IsNullOrEmpty($OutputCSVPath)) {
				$processParams.ArgumentList += " -OutputCSVPath $OutputCSVPath"
			}
			if ($IncludeEqual) {
				$processParams.ArgumentList += " -IncludeEqual"
			}

			Start-Process @processParams
			return
		} catch {
			Write-Error "Failed to start $Configuration PowerShell: $($_.Exception.Message)"
			return
		} finally {
			if (Test-Path -Path $tempFile) {
				Write-Host -ForegroundColor White "Cleaning up temporary script file at $tempFile"
				Remove-Item -Path $tempFile -Force
			} else {
				Write-Warning "Temporary script file not found at $tempFile for cleanup"
			}
			Write-Host -ForegroundColor Green "Returned from $Configuration PowerShell process"
		}
	}

	if ([Environment]::Is64BitProcess) {
		Write-Host -ForegroundColor Green "Running in 64-bit PowerShell configuration"
	} else {
		Write-Host -ForegroundColor Green "Running in 32-bit PowerShell configuration"
	}

	function Close-ComObjects {
		param (
			$conn,
			$rs
		)
		try {
			# Ensure connections are closed
			if ($rs -and $rs.state -eq 1) { $rs.Close() }
			if ($conn -and $conn.state -eq 1) { $conn.Close() }
			
			# Ensure ComObjects are released
			if ($rs) {
				[System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs) | Out-Null
				$rs = $null
			}
			if ($conn) {
				[System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) | Out-Null
				$conn = $null
			}
		} catch {
			Write-Warning "COM Object Cleanup Failed: $_"
		}
	}

	# Database Connection
	try {
		$connStringA = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=$ReferencePath"
		$connStringB = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=$DifferencePath"

		$connA = New-Object -ComObject ADODB.Connection
		$connB = New-Object -ComObject ADODB.Connection

		$connA.ConnectionString = $connStringA
		$connB.ConnectionString = $connStringB

		Write-Host -ForegroundColor White "Connecting to Reference DB: $connStringA"
		$connA.Open($connStringA)
		Write-Host -ForegroundColor White "Connecting to Difference DB: $connStringB"
		$connB.Open($connStringB)
	} catch {
		Write-Error "Failed to connect to databases: $($_.Exception.Message)"
		Close-ComObjects -conn $connA
		Close-ComObjects -conn $connB
		return $null
	}

	# Table Collection
	try {
		# see https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/schemaenum for more info on this enum
		$adSchemaTables = 20 
		Write-Host -ForegroundColor White "Getting tables for the Reference DB"
		$schemaTableA = $connA.OpenSchema($adSchemaTables)
		Write-Host -ForegroundColor White "Getting tables for the Difference DB"
		$schemaTableB = $connB.OpenSchema($adSchemaTables)
		
		# Collect all tables from reference DB
		$allTablesA = @()
		while (-not $schemaTableA.EOF) {
			# Filter for standard user tables
			if ($schemaTableA.Fields.Item("TABLE_TYPE").Value -eq "TABLE") {
				$allTablesA += $schemaTableA.Fields.Item("TABLE_NAME").Value
			}
			$schemaTableA.MoveNext()
		}

		Close-ComObjects -rs $schemaTableA

		# Collect all tables from Difference DB
		$allTablesB = @()
		while (-not $schemaTableB.EOF) {
			# Filter for standard user tables
			if ($schemaTableB.Fields.Item("TABLE_TYPE").Value -eq "TABLE") {
				$allTablesB += $schemaTableB.Fields.Item("TABLE_NAME").Value
			}
			$schemaTableB.MoveNext()
		}

		Close-ComObjects -rs $schemaTableB
	} catch {
		Write-Error "Failed to collect Tables from databases: $($_.Exception.Message)"
		Close-ComObjects $connA $schemaTableA
		Close-ComObjects $connB $schemaTableB
		return $null
	}

	# Collect Table/Column/Values
	try {
		Write-Host -ForegroundColor White "Getting Table/Column/Value pairs from Reference DB"
		# Collect all Column/value pairs from Reference DB
		$allResultsA = @()
		$tableIteration = 1
		foreach ($table in $allTablesA) {
			$query = "SELECT * FROM $table"
			$progressPercent = ($tableIteration/$allTablesB.Count) * 100
			Write-Progress -Activity "Querying all tables in Reference DB..." -Status "Executing Query: $query" -ID 0 -PercentComplete $progressPercent
			$recordSetA = New-Object -ComObject ADODB.Recordset
			$recordSetA.Open($query, $connA)
			while (-not $recordSetA.EOF) {
				foreach ($field in $recordSetA.Fields) {
					$row = [PSCustomObject]@{
						Table = $table;
						Column = $field.Name;
						Value = $field.Value;
					}
					$allResultsA += $row			
				}
				$recordSetA.MoveNext()
			}
			$tableIteration += 1
			Close-ComObjects -rs $recordSetA
		}

		# Collect all Column/value pairs from Difference DB
		Write-Host -ForegroundColor White "Getting Table/Column/Value pairs from Difference DB"
		$allResultsB = @()
		$tableIteration = 1
		foreach ($table in $allTablesB) {
			$query = "SELECT * FROM $table"
			$progressPercent = ($tableIteration/$allTablesB.Count) * 100
			Write-Progress -Activity "Querying all tables in Difference DB..." -Status "Executing Query: $query" -ID 0 -PercentComplete $progressPercent
			$recordSetB = New-Object -ComObject ADODB.Recordset
			$recordSetB.Open($query, $connB)
			while (-not $recordSetB.EOF) {
				foreach ($field in $recordSetB.Fields) {
					$row = [PSCustomObject]@{
						Table = $table;
						Column = $field.Name;
						Value = $field.Value;
					}
					$allResultsB += $row
				}
				$recordSetB.MoveNext()
			}
			$tableIteration += 1
			Close-ComObjects -rs $recordSetB
		}
	} catch {
		Write-Error "Failed to collect Column/Values from databases: $($_.Exception.Message)"
		Close-ComObjects $connA $recordSetA
		Close-ComObjects $connB $recordSetB
		return $null
	}

	# Compare Table/Column/Values
	try {
		Write-Host -ForegroundColor White "Comparing Tables, Columns, and Values..."
		$resultComparison = Compare-Object -ReferenceObject $allResultsA -DifferenceObject $allResultsB -Property Table,Column,Value -IncludeEqual:$IncludeEqual -ErrorAction Stop -PassThru
		$differenceCount = ($resultComparison | Where-Object {$_.SideIndicator -ne "=="}).Count
		if ($differenceCount -gt 0) {
			Write-Host -ForegroundColor Red "Comparison detected $($differenceCount) differences"
		} else {
			Write-Host -ForegroundColor Green "All fields are identical"
		}
	} catch {
		Write-Error "Failed to compare Table/Column/Values: $($_.Exception.Message)"
		return $null
	} finally {
		Write-Host -ForegroundColor White "Cleaning up DB connections.."
		Close-ComObjects $connA $recordSetA
		Close-ComObjects $connB $recordSetB
	}

	Write-Host -ForegroundColor Green "Comparison Complete."

	# Output Results to file
	try {
		if (![string]::IsNullOrEmpty($OutputCSVPath) -and ($differenceCount -gt 0)) {
			$resultComparison | ConvertTo-Csv -NoTypeInformation | Out-File -FilePath $OutputCSVPath -ErrorAction Stop
			Write-Host -ForegroundColor White "Results saved to $OutputCSVPath"
		}
	} catch {
		Write-Error "Failed to output results to $($OutputCSVPath): $($_.Exception.Message)"
	}

	$resultComparison | Out-GridView -Title "Microsoft Access DB Comparison Results" -Wait
}