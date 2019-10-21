#Import the Excel File
$dataFile = Import-Excel "D:\\Input-modifier-testing\\play\\Regression Test - R11-1 - 2019 06 13 - in UNIT-play\\kathy'sFiles\\D 4 - OL - Regr-old.xlsx"
$Key = @();
$Value = @();
$dataRow = @();
$Body;
$startDate = Get-Date "2019-10-19";

##### Loop through the data File
Function InvokeWebRequest($dataFile) {
	write-host " This is the data file " + $dataFile + " \n"
	foreach ($row in $dataFile){
		write-host $row
		foreach($property in $row.psobject.Properties){
			$Key = ($property | Select-Object -Property name | % {$_.Name})
			$Value = ($property | Select-Object -Property value | % {$_.Value})
			if([string]$Value -as [DateTime]){
				$Value = Get-Date $Value -Format 'MM/dd/yyyy'
			}
			$Body += [string]([string]$Key + "=" + [string]$Value + "&")
		}
		$dataRow +=$Body
	}
	#Write-Host $dataRow
	$arr = $dataRow -split '&environment'
	#Write-Host $arr
	for($i=1; $i -lt $arr.Count; $i++){
		$arr[$i] = "environment" + $arr[$i]
	}
	for($count=1; $count -lt $arr.count; $count++){
		$Response = Invoke-WebRequest -Uri "http://bltmlu1:8080/testbillingui/services" -Method POST -Body $arr[$count] | convertFrom-Json
		write-host $Response
	}
}

#### This executes the aging tool, I should know when the batch file is done executing.
Function executeAging($date){
	powershell.exe D:\Tools\shelter-batch-tool\shelter-batch-tool-unit.bat $date -Wait -NoNewWindow
}

#### This gets the next file to process
Function getDataFile(){
	$transactionFile =  gci "D:\Input-modifier-testing\play\Regression Test - R11-1 - 2019 06 13 - in UNIT-play\kathy'sFiles" | Sort-Object -Property Name | Select-Object -First 1
	return "D:\Input-modifier-testing\play\Regression Test - R11-1 - 2019 06 13 - in UNIT-play\kathy'sFiles\" + $transactionFile.Name
}

##### This executes the Regression Test
##### 1. This gets the first,next file to process(Day of File to execute)
##### 2. Moves the file to the DONE folder
##### 3. Invokes Web request to create the transactions
##### 4. Executes Aging

Function executeRegression(){
	$transactionFile = getDataFile
	$dataFile = Import-Excel $transactionFile
	#InvokeWebRequest($dataFile)
	#MoveTransactionFile
	$NextFile =  gci "D:\Input-modifier-testing\play\Regression Test - R11-1 - 2019 06 13 - in UNIT-play\kathy'sFiles" | Sort-Object -Property Name | Select-Object -First 1
	$Days = ($NextFile.Name -replace "\D+(\d+)\D+", '$1')
	write-host $Days
	$ageDateTill = (Get-Date $startDate).AddDays($Days)
	write-host $ageDateTill
	executeAging(Get-Date $ageDateTill -Format 'yyyy-MM-dd')
	
}

Function loopThroughRegression(){
	'Execute Regression must be ivoked in a loop'
	Script:executeRegression {
	Return
	}
}

loopThroughRegression