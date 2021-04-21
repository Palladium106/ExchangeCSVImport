#Sript for import mail contacts into AD using Exchange Management Console.
#ExchangeCSVImport.conf file must be stored in the one folder with the script.

if((Get-Command New-MailContact -errorAction SilentlyContinue) -eq $null)
{
	Write-Host "This script must be executed on the Exchange Management Console!" -foregroundcolor red
	return
}

#Exit if ExchangeCSVImport.conf file not found.
if((Test-Path ExchangeCSVImport.conf) -ne $true) 
{
	Write-Host "'ExchangeCSVImport.conf' must be stored in the one folder with the script" -foregroundcolor red
	return
}

#Parsing ExchangeCSVImport.conf file.
Get-Content ".\ExchangeCSVImport.conf" | foreach-object -begin {$config=@{}} `
	-process {
	#Parse only strings that contains "="
	if($_.Contains("=") -eq $True)
	{
		$k = $_.Split("=")
		#Parse only strings not beginned with "[" and "#"
		if  (($k[0].StartsWith("[") -ne $True) `
		-and ($k[0].StartsWith("#") -ne $True))
		{
			$config.Add($k[0].Trim(), $k[1].Trim())
		}
	}
}

Write-Host "Initial paremeters:"
$config
$choice=Read-Host "`nDo you want to continue? (Y/n)"

#Proceed to import if user typed "y" or nothing.
switch ($choice)
{
	"y" {break}
	"" {break}
	Default {return}
}

#Exit script if CSV file not found.
if((Test-Path $config.CSVFileName) -ne $true) 
{
	Write-Host "File '$($config.CSVFileName)' does not exists"
	return
}

#Main import procedure
Import-CSV -Path $config.CSVFileName -Delimiter "," | ForEach-Object `
-process{
	$FullName = $_.DisplayName
	#Split Full User Name to place it in the right fields.
	$SplittedlName = $FullName.Split(" ",2)
	
	#Create contact and set basic attributes.
	try
	{
		New-MailContact -Name $FullName `
			-ExternalEmailAddress $_.EmailAddress `
			-Alias $_.SamAccountName `
			-OrganizationalUnit $config.TargetOU `
			-FirstName $SplittedlName[1] `
			-LastName $SplittedlName[0]

	}
	catch [Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException]
	{
		Write-Host "Specified OU: '$($config.TargetOU)' does not exists" -foregroundcolor red
		exit
	}
	catch [Microsoft.Exchange.Data.Directory.ADObjectAlreadyExistsException]
	{
		Write-Host "Specified User: '$($FullName)' already exists" -foregroundcolor yellow
	}
	
	#Check if the Custom Company Name is set.
	if($config.CompanyName -ne $null) {$Company = $config.CompanyName}
	else {$Company = $_.Company}
	
	#Set other attributes.
	try
	{	
		Set-Contact $_.SamAccountName `
			-Company $Company `
			-Phone $_.telephoneNumber `
			-MobilePhone $_.mobile `
			-Department $_.Department `
			-Title $_.Title
	}
	catch
	{
		$_.FullyQualifiedErrorID
	}
}
