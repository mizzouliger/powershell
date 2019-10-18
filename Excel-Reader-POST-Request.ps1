#Import the Excel File
$dataFile = Import-Excel "D:\\Input-modifier-testing\\play\\Regression Test - R11-1 - 2019 06 13 - in UNIT-play\\kathy'sFiles\\D 4 - OL - Regr-old.xlsx"
$Key = @();
$Value = @();
$dataRow = @();
$Body;

##### Loop through the data File
Function InvokeWebRequest($dataFile) {
	foreach ($row in $dataFile){
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

#### This executes the aging tool
Function executeAging($date){
	powershell.exe D:\Tools\shelter-batch-tool\shelter-batch-tool-unit.bat $date
}

#### This checks the activity file
Function checkActivityFile(){

}

##### This executes the Regression Test
##### 1. Looks at the activity File to determine the next activity
##### 2. Invokes Web request to create the transactions
##### 3. Executes Aging
Function executeRegression(){
	InvokeWebRequest($dataFile);
}

executeRegression;