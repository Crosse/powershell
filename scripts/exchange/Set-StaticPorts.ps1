<#
	.SYNOPSIS
		This script will configure static ports for RPC Client Access and Address Book Service on Exchange 2010 CAS servers.

	.DESCRIPTION
		This script configures static RPC ports on specified Client Access Server. It also restarts services is specified.

	.PARAMETER  Server
		Name of the Client Access Server. This could also be a Mailbox server hosting Public Folder database.
		
		Server and Auto are mutually exclusive parameters. You must specify one or the other but cannot specify not both.

	.PARAMETER  RPCPort
		Static port for RPC Client Access Service. If not specified, defaults to 59531. Microsoft recommends you set this to a unique value between 59531 and 60554 and use the same value on all CAS in any one AD site.

	.PARAMETER  ABPort
		Static port for Exchange Address Book Service. If not specified, defaults to 59532. Microsoft recommends you set this to a unique value between 59531 and 60554 and use the same value on all CAS in any one AD site. Ensure this value is different from value of RPCPort.
	
	.PARAMETER  Auto
		Discover Exchange 2010 Client Access Servers and Mailbox Servers hosting Public Folder databases. Change ports on all discovered servers.
		
		Server and Auto are mutually exclusive parameters. You must specify one or the other but cannot specify both.
	
	.PARAMETER  Force
		Suppresses all user prompts. This could have undesired result if user doesn't understand the impact and outcome when using this parameter.
	
	.EXAMPLE
		PS C:\> .\Set-StaticPorts.ps1 -Server ServerA
		
		This example will set default ports 59531 for RPC Client Access Service and 59532 for Exchange Address Book Service on ServerA.

	.EXAMPLE
		PS C:\> .\Set-StaticPorts.ps1 -RPCPort 59533 -ABPort 59535 -Force
		
		This example will set specified ports on server where script is run. All confirmation prompts will be suppressed.

	.EXAMPLE
		PS C:\> .\Set-StaticPorts.ps1 -Auto -Force
		
		This example will set default ports 59531 for RPC Client Access Service and 59532 for Exchange Address Book Service on all discovered Client Access Servers and Mailbox Servers hosting Public Folder databases. All confirmation prompts will be suppressed.

	.INPUTS
		System.String

	.OUTPUTS
		System.String
		
	.NOTES
		Created and maintainted by Bhargav Shukla (MSFT). Please report errors through contact form at http://blogs.technet.com/b/bshukla/contact.aspx. Do not remove original author credits or reference.

	.LINK
		http://social.technet.microsoft.com/wiki/contents/articles/configuring-static-rpc-ports-on-an-exchange-2010-client-access-server.aspx

	.LINK
		http://blogs.technet.com/bshukla
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param
(
	[Parameter(Position=0, Mandatory=$true, ParameterSetName = "Set01")]
	[System.String]
	$Server,
	[Parameter(Position=1, Mandatory=$false, ParameterSetName = "Set01")]
	[Parameter(ParameterSetName = "Set02")]
	[ValidateRange(59531,60554)]
	[System.String]
	$RPCPort = "59531",
	[Parameter(Position=2, Mandatory=$false, ParameterSetName = "Set01")]
	[Parameter(ParameterSetName = "Set02")]
	[ValidateScript({($_) -ne $RPCPort -and ($_) -ge "59531" -and ($_) -le "60554"})]
	[System.String]
	$ABPort = "59532",
	[Parameter(Position=3, Mandatory=$false, ParameterSetName = "Set02")]
	[Switch]
	$Auto,
	[Parameter(Position=4, Mandatory=$false, ParameterSetName = "Set01")]
	[Parameter(ParameterSetName = "Set02")]
	[Switch]
	$Force
)
	
# Before we proceed, ensure Exchange cmdlets are loaded.
try 
{
	Get-Command Get-ExchangeServer | Out-Null
}
catch 
{
	Write-Error "Exchange cmdlets are not loaded. Please connect to Exchange 2010 Server remotely or load Exchange Management Shell before running this script."
	Return
}

# Define all required functions
function Set-RemoteRegistry
{
	[CmdletBinding(SupportsShouldProcess=$true)]
	param
	(
		[Parameter(Position=0, Mandatory=$false)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Server = $Env:COMPUTERNAME,
		[Parameter(Position=1, Mandatory=$false)]
		[ValidateSet("ClassesRoot","CurrentConfig","CurrentUser","DynData","LocalMachine","PerformanceData","Users")]
		[System.String]
		$Hive = "LocalMachine",
		[Parameter(Position=2, Mandatory=$true, HelpMessage="Enter Registry key in format System\CurrentControlSet\Services")]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Key,
		[Parameter(Position=3, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Name,
		[Parameter(Position=4, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Value,		
		[Parameter(Position=5, Mandatory=$true)]
		[ValidateSet("String","ExpandString","Binary","DWord","MultiString","QWord")]
		[System.String]
		$Type,
		[Parameter(Position=6, Mandatory=$false)]
		[Switch]
		$Force
	)
	
	if ($pscmdlet.ShouldProcess($Server, "Open registry $Hive"))
	{
	#Open remote registry
	try
	{
			$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Server)
		
	}
	catch 
	{
		Write-Error "The Server $Server is inaccessible. Please check servername. Please ensure remote registry service is running and you have administrative access to $server."
		Return
	}
	}

	if ($pscmdlet.ShouldProcess($Server, "Check existense of $Key"))
	{
	#Open the targeted remote registry key/subkey as read/write
	$regKey = $reg.OpenSubKey($Key,$true)
		
	#Since trying to open a regkey doesn't error for non-existent key, let's sanity check
	#Create subkey if parent exists. If not, exit.
	If ($regkey -eq $null)
	{	
		Write-Warning "Specified key $Key does not exist in $Hive."
		$Key -match ".*\x5C" | Out-Null
		$parentKey = $matches[0]
		$Key -match ".*\x5C(\w*\z)" | Out-Null
		$childKey = $matches[1]

		try
		{
			$regtemp = $reg.OpenSubKey($parentKey,$true)
		}
		catch
		{
			Write-Error "$parentKey doesn't exist in $Hive or you don't have access to it. Exiting."
			Return
		}
		If ($regtemp -ne $null)
		{
			Write-Output "$parentKey exists. Creating $childKey in $parentKey."
			try
			{
				$regtemp.CreateSubKey($childKey) | Out-Null
			}
			catch 
			{
				Write-Error "Could not create $childKey in $parentKey. You  may not have permission. Exiting."
				Return
			}

			$regKey = $reg.OpenSubKey($Key,$true)
		}
		else
		{
			Write-Error "$parentKey doesn't exist. Exiting."
			Return
		}
	}
	
	#Cleanup temp operations
	try
	{
		$regtemp.close()
		Remove-Variable $regtemp,$parentKey,$childKey
	}
	catch
	{
		#Nothing to do here. Just suppressing the error if $regtemp was null
	}
	}
	
	#If we got this far, we have the key, create or update values
	If ($Force)
	{
		If ($pscmdlet.ShouldProcess($Server, "Create or change $Name's value to $Value in $Key. Since -Force is in use, no confirmation needed from user"))
		{
			$regKey.Setvalue("$Name", "$Value", "$Type")
		}
	}
	else
	{
		If ($pscmdlet.ShouldProcess($Server, "Create or change $Name's value to $Value in $Key. No -Force specified, user will be asked for confirmation"))
		{
		$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
		$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""
		$choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
		$caption = "Warning!"
		$message = "Value of $Name will be set to $Value. Current value `(If any`) will be replaced. Do you want to proceed?"
		Switch ($result = $Host.UI.PromptForChoice($caption,$message,$choices,0))
		{
			1
			{
				Return
			}
			0
			{
				$regKey.Setvalue("$Name", "$Value", "$Type")
			}
		}
		}
	}
	
	#Cleanup all variables
	try
	{
		$regKey.close()
		Remove-Variable $Server,$Hive,$Key,$Name,$Value,$Force,$reg,$regKey,$yes,$no,$caption,$message,$result
	}
	catch
	{
		#Nothing to do here. Just suppressing the error if any variable is null
	}
}

function Get-InstallPath
{
	# Set Exchange base key and value to read
	$Reg_ExSetup = "SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
	$VALUE = "MsiInstallPath"

	# Set regKey for MsiInstallPath
	$regKey = $reg.OpenSubKey($REG_ExSetup)

	# Get Install Path from Registry and replace : with $
	$installPath = ($regkey.getvalue($VALUE) | foreach {$_ -replace (":","`$")})
	
	# Set Address Book Service config file path
	$Script:ABFile = "\\$Server\$installPath"+"Bin\microsoft.exchange.addressbook.service.exe.config"
	
	# Close registry key
	$regKey.Close()
	
	#Cleanup all variables
	try
	{
		Remove-Variable $Reg_ExSetup,$Value,$regKey,$installPath
	}
	catch
	{
		#Nothing to do here. Just suppressing the error if any variable is null
	}
}

function Restart-Services
{
	# Restart Microsoft Exchange RPC Client Access and Microsoft Exchange Address Book service
	#### You must specify a timeout (in seconds) or the script could potentially never end
	$TimeOut = 30
		
	#### This will stop a single service in sequence
	$ServiceFilters = "(name = 'msexchangerpc')","(name = 'msexchangeab')"
		
		
	$Locator = new-object -com "WbemScripting.SWbemLocator" 
	$WMI = $Locator.ConnectServer($Server, "root\cimv2")
	# Stop Service and check for timeout or sucessful stop
		$ServiceFilters | %{ 
			$ThisFilter = $_
			(Get-WmiObject -Class Win32_Service -ComputerName $Server -filter "$ThisFilter AND state='running'") | %{
				$Service = $_
				$Refresher = new-object -comobject "WbemScripting.SWbemRefresher"
				$FreshObject = $Refresher.Add($WMI,$Service.__RELPATH)
				$Refresher.Refresh()
				$Then = Get-Date
				:Checking Do {
					$Service.StopService() | out-null
					$Refresher.Refresh()
					if (($FreshObject.Object.properties_ | ?{$_.name -eq "state"}).value -eq "Stopped") 
					{
						Write-Warning "Service $($Service.Name) is stopped on server $Server."
						break :Checking;
					} Else { 
						If (((Get-Date) - $Then).seconds -ge $TimeOut)
						{							
							Write-Warning "Service $($Service.Name) timed out while trying to stop on server $Server. Please restart service manually."
							break :Checking;
						} 
					}
				} While ($True)
			}
		}
		
	# Start Service and check for timeout or sucessful start
		$ServiceFilters | %{ 
			$ThisFilter = $_
			(Get-WmiObject -Class Win32_Service -ComputerName $Server -filter "$ThisFilter AND state='Stopped'") | %{
				$Service = $_
				$Refresher = new-object -comobject "WbemScripting.SWbemRefresher"
				$FreshObject = $Refresher.Add($WMI,$Service.__RELPATH)
				$Refresher.Refresh()
				$Then = Get-Date
				:Checking Do {
					$Service.StartService() | out-null
					$Refresher.Refresh()
					if (($FreshObject.Object.properties_ | ?{$_.name -eq "state"}).value -eq "running") 
					{
						Write-Host "Service $($Service.Name) started successfully on server $Server"												
						break :Checking;
					} Else { 
						If (((Get-Date) - $Then).seconds -ge $TimeOut)
						{
							Write-Warning "Service $($Service.Name) timed out while trying to start on server $Server. Please restart service manually."
							break :Checking;
						} 
					}
				} While ($True)
			}
		}
}

Function Main
{
If ((Get-ExchangeServer $server).AdminDisplayVersion -match "Version 14")
{
	Write-Output "Working on server $Server."

	# Set Registry Key for RPC Port
	$Key = 'System\CurrentControlSet\Services\MSExchangeRPC\ParametersSystem'
	$Name = 'TCP/IP Port'
	$Type = 'Dword'

	If ($Force)
	{
		If ($pscmdlet.ShouldProcess($Server, "Set RPC Port to $RPCPort in registry. -Force is in use, user won't be prompted."))
		{
			Set-RemoteRegistry -Server $Server -Key $Key -Name $Name -Value $RPCPort -Type $Type -Force
		}
	}
	else
	{
		If ($pscmdlet.ShouldProcess($Server, "Set RPC Port to $RPCPort in registry. -Force is not in use, user will be prompted."))
		{
			Set-RemoteRegistry -Server $Server -Key $Key -Name $Name -Value $RPCPort -Type $Type
		}
	}

	# Prepare choice object
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""
	$choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
	$caption = "Warning!"
	$message = "Change port from $($item.value) to $ABPort on server $server ?"

	# Make changes to ABPort in XML file or Regisry based on version of Exchange 2010 on server
	If ((Get-ExchangeServer $server).AdminDisplayVersion -match "Version 14.0" -and (Get-ExchangeServer $server).ServerRole -match "ClientAccess") 
	{
		If ($pscmdlet.ShouldProcess($Server, "Change Exchange Address Book port to $ABPort in config file since server is RTM `(14.0`)"))
		{
			# Get location of Address Book Service configuration file
			Get-InstallPath
	
			# Verify the file exists, exit if it doesn't
			If (Test-Path $ABFile -ErrorAction SilentlyContinue)
			{
				# Backup file before editing
				$currentDate = (get-date).tostring("MM_dd_yyyy-hh_mm_s")
				$ABbackup = $ABFile + ".$currentDate"
				$xml = [xml](get-content $ABFile)
				$xml.Save($ABbackup)
					
				# Edit RpcTcpPort if 0, ask user for approval if not 0
				$root = $xml.get_DocumentElement();
				ForEach ($item in $root.appSettings.add) 
				{
					if (($item.key -eq "RpcTcpPort") -and ($item.value -eq "0"))
					{
						$item.value="$abport"
						Write-Output "TCP/IP Port for Address Book Service is set to $ABPort on server $server."
					} 
					else
					{
						If (($item.key -eq "RpcTcpPort") -and ($item.value -ne "$abport"))
						{
							if (($force) -or ($result = $Host.UI.PromptForChoice($caption,$message,$choices,0)) -eq 0)
							{
								$item.value="$abport"
								Write-Output "TCP/IP Port for Address Book Service is set to $ABPort on server $server."
							}
							else
							{
								Write-Output "No changes are made to TCP/IP Port for Address Book Service on server $server."
							}
						}
					}
				}
				$xml.Save($ABFile)
			}
			else
			{
				Write-Error "Address Book Service configuration file does not exist for server $server. Please verify file and update manually."
			}
		}
	}
	else 
	{
		If ((Get-ExchangeServer $server).AdminDisplayVersion -match "Version 14.1" -or (Get-ExchangeServer $server).AdminDisplayVersion -match "Version 14.2"  -and (Get-ExchangeServer $server).ServerRole -match "ClientAccess")
		{
			# Set Registry Key for Exchange Address Book Port
			$Key = 'System\CurrentControlSet\Services\MSExchangeAB\Parameters'
			$Name = 'RpcTcpPort'
			$Type = 'String'
			
			If ($pscmdlet.ShouldProcess($Server, "Change Exchange Address Book port to $ABPort in registry since server is running SP1 or later"))
			{
				If ($Force)
				{
					Set-RemoteRegistry -Server $Server -Key $Key -Name $Name -Value $ABPort -Type $Type -Force
				}
				else
				{
					Set-RemoteRegistry -Server $Server -Key $Key -Name $Name -Value $ABPort -Type $Type
				}
			}
		}
	}

	# Ask for and restart services if requested
	If ($pscmdlet.ShouldProcess($Server, "Restart MSExchangeRPC and MSExchangeAB services"))
		{
			If (($force) -or ($result = $Host.UI.PromptForChoice($caption,"Restart services MSExchangeRPC and MSExchangeAB (if needed)?",$choices,0)) -eq 0)
			{
				restart-services
			}
			else
			{
				If (-not $Auto)
				{
					Write-Warning "Please restart the Microsoft Exchange RPC Client Access (MSExchangeRPC) service and Exchange Address Book (MSExchangeAB) service for changes to take effect."
				}
			}
		}
			
	# Reminder for Mailbox Servers
	If (-not $Auto -and -not ((Get-ExchangeServer $server).AdminDisplayVersion -match "Version 14" -and (Get-ExchangeServer $server).ServerRole -match "Mailbox"))
	{
		Write-Output "Please change ports on Mailbox servers hosting Public Folder databases if desired."
	}
}
else
{
	Write-Error "Server $server is not running Exchange Server 2010. Please specify an Exchange 2010 Server."
}
}
### End function definitions

### Run script
If ($Auto)
{
	$Servers = (Get-ExchangeServer | Where {$_.AdminDisplayVersion -match "Version 14" -and $_.ServerRole -match "ClientAccess"} | ForEach {$_.Name})
	$pf = Get-PublicFolderDatabase
	$Servers = $Servers + ($pf | Where {(get-exchangeserver $_.server).ServerRole -notmatch "ClientAccess"} | ForEach {$_.Server})

	ForEach ($Server in $Servers)
	{
		Main
	}
}
else
{
	Main
}
