
# Configuration	
############################################################################################
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!C O N F I G U R A T I O N!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!#
############################################################################################

# Snapins
Add-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction SilentlyContinue
Import-module grouppolicy -ErrorAction Continue
Import-Module CimCmdlets -ErrorAction Continue
Import-Module ActiveDirectory -ErrorAction Continue

# PowershellToolkit information
$ApplicationName = "Powershell Toolkit"
$ApplicationVersion = "3.0.0.2"
$ApplicationLastUpdate = "13.06.2016"

# Author Information
$AuthorName = "Renato Bacchi"
$AuthorEmail = "admin@renatobacchi.ch"
$AuthorWWW = "http://www.renatobacchi.ch"

# Text to show in the Status Bar when the form load
$StatusBarStartUp = "$ApplicationName - $ApplicationVersion - (c) Renato Bacchi - $AuthorWWW"

# Title of the MainForm / Mainform Titel
$domain = $env:userdomain.ToUpper()
$MainFormTitle = "$ApplicationName $ApplicationVersion - Last Update: $ApplicationLastUpdate - $domain\$env:username"

# Font Styles / Schrift Stile
$bold = New-Object Drawing.Font("Lucida Console", 8, [Drawing.Fontstyle]::Bold)
$norm = New-Object Drawing.Font("Lucida Console", 8, [Drawing.Fontstyle]::Regular)
$log = New-Object Drawing.Font("Lucida Console", 1, [Drawing.Fontstyle]::Regular)
[Drawing.Color]$gray = "Control"
[Drawing.Color]$green = "Green"
[Drawing.Color]$red = "Red"
[Drawing.Color]$black = "Black"
$global:Fillchar = 178
$Newline = "`n"
$Newline2 = "`n`n"

## Environment Variables / Umgebungsvariablen
if (Test-Path "C:\Program Files (x86)") { $global:Programfiles = "C:\Program Files (x86)" }
else { $global:Programfiles = "C:\Program Files" }
$cmd = "cmd.exe"

# Folder / Ordner
$global:Profilefolder = ""
$global:Homefolder = ""
$global:Outfile = $pwd
$global:Confpath = $env:APPDATA += "\Powershell Toolkit\"

# SCCM
$global:SCCMEnabled = "true"
$global:SiteName = ""
$global:SCCMServer = ""
$global:SCCMNameSpace = "root\sms\site_$SiteName"
$global:CmRCViewer = "$global:Programfiles\ConfigMgr\bin\i386\CmRcViewer.exe"

# External Tools
$global:Nirlauncher = "$global:Programfiles\Nirsoft"
$global:Sysinternals = "$global:Programfiles\Sysinternals"

# Loading Lang-Variables because $lang.xyz does not work in AddRichtTextbox -Text if there
# are multiple Variables, bc. those are not strings but hashtablekeys
# Maybe these should be change to something like global:langfolder and so on, so the code would be more readable
$global:ChangePasswordAtLogon = $lang.ChangePasswordAtLogon
$global:CheckComputerGroups = $lang.CheckComputerGroups
$global:CheckConn = $lang.CheckConn
$global:ComputerNotFound = $lang.ComputerNotFound
$global:ComputerOfflineOrWrong = $lang.ComputerOfflineOrWrong
$global:Cycle1 = $lang.Cycle1
$global:Cycle2 = $lang.Cycle2
$global:Cycle3 = $lang.Cycle3
$global:Cycle4 = $lang.Cycle4
$global:Cycle5 = $lang.Cycle5
$global:Cycle6 = $lang.Cycle6
$global:Cycle7 = $lang.Cycle7
$global:Cycle8 = $lang.Cycle8
$global:Cycle9 = $lang.Cycle9
$global:DestinationComputer = $lang.DestinationComputer
$global:DnsConf = $lang.DnsConf
$global:DoYouWantToTransfer = $lang.DoYouWantToTransfer
$global:EnterCommand = $lang.EnterCommand
$global:EnterDestinationComputer = $lang.EnterDestinationComputer
$global:EnterPassword = $lang.EnterPassword
$global:EnterSourceComputer = $lang.EnterSourceComputer
$global:EnterUsername = $lang.EnterUsername
$global:ErrorUnlocking = $lang.ErrorUnlocking
$global:FolderPathInputBoxMsg = $lang.FolderPathInputBoxMsg
$global:FolderPathInputBoxTitle = $lang.FolderPathInputBoxTitle
$global:FollowingLocked = $lang.FollowingLocked
$global:LockedUser = $lang.LockedUser
$global:NetConf = $lang.NetConf
$global:NoInputDetected = $lang.NoInputDetected
$global:NoUserUnlocked = $lang.NoUserUnlocked
$global:NoUsersLocked = $lang.NoUsersLocked
$global:NotExistinginAD = $lang.NotExistinginAD
$global:PSRnotEnabled = $lang.PSRnotEnabled
$global:PasswordResetOK = $lang.PasswordResetOK
$global:Please = $lang.Please
$global:RegKeySet = $lang.RegKeySet
$global:RemoteCommandSent = $lang.RemoteCommandSent
$global:RunRemoteCMD = $lang.RunRemoteCMD
$global:ShowFolderRights = $lang.ShowFolderRights
$global:ShowLocalAdminsOf = $lang.ShowLocalAdminsOf
$global:ShowingComputergroupsOf = $lang.ShowingComputergroupsOf
$global:ShowingLastPC = $lang.ShowingLastPC
$global:SourceComputer = $lang.SourceComputer
$global:TransferComputerGroups = $lang.TransferComputerGroups
$global:Transferring = $lang.Transferring
$global:TwoIdenticalComputers = $lang.TwoIdenticalComputers
$global:UnlockUser = $lang.UnlockUser
$global:UnlockedOK = $lang.UnlockedOK
$global:YouHaveEnteredTwoIdenticalComputers = $lang.YouHaveEnteredTwoIdenticalComputers
$global:checkHomeRights = $lang.checkHomeRights
$global:checkProfileRights = $lang.checkProfileRights
$global:configFolderExisting = $lang.configFolderExisting
$global:configFrom = $lang.configFrom
$global:createConfigError = $lang.createConfigError
$global:created = $lang.created
$global:existing = $lang.existing
$global:folder = $lang.folder
$global:loaded = $lang.loaded
$global:openPSRS = $lang.openPSRS
$global:starting = $lang.starting
# Languagefiles Language.psd1 in Folders, e.g. \de-DE\ with Variables and Strings

############################################################################################
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!#
############################################################################################
#endregion Configuration

# Adder Functions

#region Add-ListViewItem
function Add-ListViewItem
{
<#
	.SYNOPSIS
		Adds the item(s) to the ListView and stores the object in the ListViewItem's Tag property.

	.DESCRIPTION
		Adds the item(s) to the ListView and stores the object in the ListViewItem's Tag property.

	.PARAMETER ListView
		The ListView control to add the items to.

	.PARAMETER Items
		The object or objects you wish to load into the ListView's Items collection.
		
	.PARAMETER  ImageIndex
		The index of a predefined image in the ListView's ImageList.
	
	.PARAMETER  SubItems
		List of strings to add as Subitems.
	
	.PARAMETER Group
		The group to place the item(s) in.
	
	.PARAMETER Clear
		This switch clears the ListView's Items before adding the new item(s).
	
	.EXAMPLE
		Add-ListViewItem -ListView $listview1 -Items "Test" -Group $listview1.Groups[0] -ImageIndex 0 -SubItems "Installed"
#>
	
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ListView]$ListView,
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		$Items,
		[int]$ImageIndex = -1,
		[string[]]$SubItems,
		[System.Windows.Forms.ListViewGroup]$Group,
		[switch]$Clear)
	
	if ($Clear)
	{
		$ListView.Items.Clear();
	}
	
	if ($Items -is [Array])
	{
		$ListView.BeginUpdate()
		foreach ($item in $Items)
		{
			$listitem = $ListView.Items.Add($item.ToString(), $ImageIndex)
			#Store the object in the Tag
			$listitem.Tag = $item
			
			if ($SubItems -ne $null)
			{
				$listitem.SubItems.AddRange($SubItems)
			}
			
			if ($Group -ne $null)
			{
				$listitem.Group = $Group
			}
		}
		$ListView.EndUpdate()
	}
	else
	{
		#Add a new item to the ListView
		$listitem = $ListView.Items.Add($Items.ToString(), $ImageIndex)
		#Store the object in the Tag
		$listitem.Tag = $Items
		
		if ($SubItems -ne $null)
		{
			$listitem.SubItems.AddRange($SubItems)
		}
		
		if ($Group -ne $null)
		{
			$listitem.Group = $Group
		}
	}
}
#endregion

#region Add-RichTextBox
# Function - Add Text to RichTextBox
function Add-RichTextBox
{
	[CmdletBinding()]
	param ($text)
	$Fill = "-"
	$Fill = $Fill * $Fillchar
	#$richtextbox_output.Text += "`tCOMPUTERNAME: $ComputerName`n"
	$richtextbox_output.SelectionFont = $log
	$richtextbox_output.SelectionColor = $Gray
	$timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss`n"
	$richtextbox_output.AppendText($timestamp)
	$richtextbox_output.SelectionFont = $norm
	$richtextbox_output.SelectionColor = $Black
	$richtextbox_output.AppendText($text)
	$richtextbox_output.SelectionFont = $bold
	$richtextbox_output.AppendText($Newline)
	$richtextbox_output.AppendText($Fill)
	$richtextbox_output.AppendText($Newline)
}
#Set-Alias artb Add-RichTextBox -Description "Add content to the RichTextBox"
#endregion

#region Add-RichtextBoxOK
function Add-RichTextBoxOK
{
	[CmdletBinding()]
	param ($text)
	$Fill = "-"
	$Fill = $Fill * $Fillchar
	#$richtextbox_output.Text += "`tCOMPUTERNAME: $ComputerName`n"
	$richtextbox_output.SelectionFont = $log
	$richtextbox_output.SelectionColor = $Gray
	$timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss`n"
	$richtextbox_output.SelectionFont = $norm
	$richtextbox_output.SelectionColor = $Green
	$richtextbox_output.AppendText($text)
	$richtextbox_output.SelectionColor = $Black
	$richtextbox_output.AppendText($Newline)
	$richtextbox_output.AppendText($Fill)
	$richtextbox_output.AppendText($Newline)
}
#endregion RichtextBoxOK

#region Add-RichtextBoxTitle
function Add-RichTextBoxTitle
{
	[CmdletBinding()]
	param ($text)
	$Fill = "-"
	$Fill = $Fill * $Fillchar
	#$richtextbox_output.Text += "`tCOMPUTERNAME: $ComputerName`n"
	$richtextbox_output.SelectionFont = $log
	$richtextbox_output.SelectionColor = $Gray
	$timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss`n"
	$richtextbox_output.SelectionFont = $bold
	$richtextbox_output.SelectionColor = $Black
	$richtextbox_output.SelectionFont = $bold
	$richtextbox_output.AppendText($text)
	$richtextbox_output.SelectionFont = $bold
	$richtextbox_output.AppendText($Newline)
	$richtextbox_output.AppendText($Fill)
	$richtextbox_output.AppendText($Newline)
}
#endregion RichtextBoxWarn

#region Add-RichtextBoxWarn
function Add-RichTextBoxWarn
{
	[CmdletBinding()]
	param ($text)
	$Fill = "-"
	$Fill = $Fill * $Fillchar
	#$richtextbox_output.Text += "`tCOMPUTERNAME: $ComputerName`n"
	$richtextbox_output.SelectionFont = $log
	$richtextbox_output.SelectionColor = $Gray
	$timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss`n"
	$richtextbox_output.SelectionFont = $norm
	$richtextbox_output.SelectionColor = $Red
	$richtextbox_output.AppendText($text)
	$richtextbox_output.SelectionColor = $Black
	$richtextbox_output.SelectionFont = $bold
	$richtextbox_output.AppendText($Newline)
	$richtextbox_output.AppendText($Fill)
	$richtextbox_output.AppendText($Newline)
	
}
#endregion RichtextBoxWarn

# Clearer Functions

#region Clear-Chart
function Clear-Chart
{
<#
	.SYNOPSIS
		This function clears the contents of the chart

	.DESCRIPTION
		Use the function to remove contents from the chart control

	.PARAMETER  ChartControl
		The Chart Control to clear

	.PARAMETER  LeaveSingleChart
		Leaves the first chart and removes all others from the control
	
	.LINK
		http://www.sapien.com/blog/2011/05/05/primalforms-2011-designing-charts-for-powershell/
#>
	Param (
		[ValidateNotNull()]
		[Parameter(Position = 1, Mandatory = $true)]
		[System.Windows.Forms.DataVisualization.Charting.Chart]
		$ChartControl
		,
		[Parameter(Position = 2, Mandatory = $false)]
		[Switch]$LeaveSingleChart
	)
	
	$count = 0
	if ($LeaveSingleChart)
	{
		$count = 1
	}
	
	while ($ChartControl.Series.Count -gt $count)
	{
		$ChartControl.Series.RemoveAt($ChartControl.Series.Count - 1)
	}
	
	while ($ChartControl.ChartAreas.Count -gt $count)
	{
		$ChartControl.ChartAreas.RemoveAt($ChartControl.ChartAreas.Count - 1)
	}
	
	while ($ChartControl.Titles.Count -gt $count)
	{
		$ChartControl.Titles.RemoveAt($ChartControl.Titles.Count - 1)
	}
	
	if ($ChartControl.Series.Count -gt 0)
	{
		$ChartControl.Series[0].Points.Clear()
	}
}
#endregion Clear-Chart

# Getter Functions

#region Get-ComputerTxtBox
function Get-ComputerTxtBox
{ $global:ComputerName = $textbox_computername.Text }
#endregion

#region Get-DiskSpace

function Get-DiskSpace
{
	
	    <#
	        .Synopsis  
	            Gets the disk space for specified host
	            
	        .Description
	            Gets the disk space for specified host
	            
	        .Parameter ComputerName
	            Name of the Computer to get the diskspace from (Default is localhost.)
	            
	        .Example
	            Get-Diskspace
	            # Gets diskspace from local machine
	    
	        .Example
	            Get-Diskspace -ComputerName MyServer
	            Description
	            -----------
	            Gets diskspace from MyServer
	            
	        .Example
	            $Servers | Get-Diskspace
	            Description
	            -----------
	            Gets diskspace for each machine in the pipeline
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .Notes
	            NAME:      Get-DiskSpace 
	            AUTHOR:    YetiCentral\bshell
	            Website:   www.bsonposh.com
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	
	Begin
	{
		Write-Verbose " [Get-DiskSpace] :: Start Begin"
		$Culture = New-Object System.Globalization.CultureInfo("en-US")
		Write-Verbose " [Get-DiskSpace] :: End Begin"
	}
	
	Process
	{
		Write-Verbose " [Get-DiskSpace] :: Start Process"
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
			
		}
		Write-Verbose " [Get-DiskSpace] :: `$ComputerName - $ComputerName"
		Write-Verbose " [Get-DiskSpace] :: Testing Connectivity"
		if (Test-Host $ComputerName -TCPPort 135)
		{
			Write-Verbose " [Get-DiskSpace] :: Connectivity Passed"
			try
			{
				Write-Verbose " [Get-DiskSpace] :: Getting Operating System Version using - Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -Property Version"
				$OSVersionInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -Property Version -ea STOP
				Write-Verbose " [Get-DiskSpace] :: Getting Operating System returned $($OSVersionInfo.Version)"
				if ($OSVersionInfo.Version -gt 5.2)
				{
					Write-Verbose " [Get-DiskSpace] :: Version high enough to use Win32_Volume"
					Write-Verbose " [Get-DiskSpace] :: Calling Get-WmiObject -class Win32_Volume -ComputerName $ComputerName -Property `"Name`",`"FreeSpace`",`"Capacity`" -filter `"DriveType=3`""
					$DiskInfos = Get-WmiObject -class Win32_Volume                          `
											   -ComputerName $ComputerName                  `
											   -Property "Name", "FreeSpace", "Capacity"      `
											   -filter "DriveType=3" -ea STOP
					Write-Verbose " [Get-DiskSpace] :: Win32_Volume returned $($DiskInfos.count) disks"
					foreach ($DiskInfo in $DiskInfos)
					{
						$myobj = @{ }
						$myobj.ComputerName = $ComputerName
						$myobj.OSVersion = $OSVersionInfo.Version
						$Myobj.Drive = $DiskInfo.Name
						$Myobj.CapacityGB = [float]($DiskInfo.Capacity/1GB).ToString("n2", $Culture)
						$Myobj.FreeSpaceGB = [float]($DiskInfo.FreeSpace/1GB).ToString("n2", $Culture)
						$Myobj.PercentFree = "{0:P2}" -f ($DiskInfo.FreeSpace / $DiskInfo.Capacity)
						$obj = New-Object PSObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.DiskSpace')
						$obj
					}
				}
				else
				{
					Write-Verbose " [Get-DiskSpace] :: Version not high enough to use Win32_Volume using Win32_LogicalDisk"
					$DiskInfos = Get-WmiObject -class Win32_LogicalDisk                       `
											   -ComputerName $ComputerName                       `
											   -Property SystemName, DeviceID, FreeSpace, Size   `
											   -filter "DriveType=3" -ea STOP
					foreach ($DiskInfo in $DiskInfos)
					{
						$myobj = @{ }
						$myobj.ComputerName = $ComputerName
						$myobj.OSVersion = $OSVersionInfo.Version
						$Myobj.Drive = "{0}\" -f $DiskInfo.DeviceID
						$Myobj.CapacityGB = [float]($DiskInfo.Capacity/1GB).ToString("n2", $Culture)
						$Myobj.FreeSpaceGB = [float]($DiskInfo.FreeSpace/1GB).ToString("n2", $Culture)
						$Myobj.PercentFree = "{0:P2}" -f ($DiskInfo.FreeSpace / $DiskInfo.Capacity)
						$obj = New-Object PSObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.DiskSpace')
						$obj
					}
				}
			}
			catch
			{
				Write-Host " Host [$ComputerName] Failed with Error: $($Error[0])" -ForegroundColor Red
			}
		}
		else
		{
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
		Write-Verbose " [Get-DiskSpace] :: End Process"
		
	}
}

#endregion 

#region Get-InstalledSoftware

function Get-InstalledSoftware
{
	
	    <#
	        .Synopsis
	            Gets the installed software using Uninstall regkey for specified host.
	
	        .Description
	            Gets the installed software using Uninstall regkey for specified host.
	
	        .Parameter ComputerName
	            Name of the Computer to get the installed software from (Default is localhost.)
	
	        .Example
	            Get-InstalledSoftware
	            Description
	            -----------
	            Gets installed software from local machine
	
	        .Example
	            Get-InstalledSoftware -ComputerName MyServer
	            Description
	            -----------
	            Gets installed software from MyServer
	
	        .Example
	            $Servers | Get-InstalledSoftware
	            Description
	            -----------
	            Gets installed software for each machine in the pipeline
	
	        .OUTPUTS
	            PSCustomObject
	
	        .Notes
	            NAME:      Get-InstalledSoftware
	            AUTHOR:    YetiCentral\bshell
	            Website:   www.bsonposh.com
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	begin
	{
		
		Write-Verbose " [Get-InstalledPrograms] :: Start Begin"
		$Culture = New-Object System.Globalization.CultureInfo("en-US")
		Write-Verbose " [Get-InstalledPrograms] :: End Begin"
		
	}
	process
	{
		
		Write-Verbose " [Get-InstalledPrograms] :: Start Process"
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
			
		}
		Write-Verbose " [Get-InstalledPrograms] :: `$ComputerName - $ComputerName"
		Write-Verbose " [Get-InstalledPrograms] :: Testing Connectivity"
		if (Test-Host $ComputerName -TCPPort 135)
		{
			try
			{
				$RegKey = Get-RegistryKey -Path "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -ComputerName $ComputerName
				foreach ($key in $RegKey.GetSubKeyNames())
				{
					$SubKey = $RegKey.OpenSubKey($key)
					if ($SubKey.GetValue("DisplayName"))
					{
						$myobj = @{
							Name = $SubKey.GetValue("DisplayName")
							Version = $SubKey.GetValue("DisplayVersion")
							Vendor = $SubKey.GetValue("Publisher")
							Install = $SubKey.GetValue("InstallDate")
							#Uninstall = $SubKey.GetValue("UninstallString")
						}
						$obj = New-Object PSObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.SoftwareInfo')
						$obj
					}
				}
			}
			catch
			{
				Write-Host " Host [$ComputerName] Failed with Error: $($Error[0])" -ForegroundColor Red
			}
		}
		else
		{
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
		Write-Verbose " [Get-InstalledPrograms] :: End Process"
		
	}
}

#endregion 	

#region Get-IP 

function Get-IP
{
	
	    <#
	        .Synopsis 
	            Get the IP of the specified host.
	            
	        .Description
	            Get the IP of the specified host.
	            
	        .Parameter ComputerName
	            Name of the Computer to get IP (Default localhost.)
	                
	        .Example
	            Get-IP
	            Description
	            -----------
	            Get IP information the localhost
	            
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .INPUTS
	            System.String
	        
	        .Notes
	            NAME:      Get-IP
	            AUTHOR:    YetiCentral\bshell
	            Website:   www.bsonposh.com
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process
	{
		$NICs = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled='$True'" -ComputerName $ComputerName
		foreach ($Nic in $NICs)
		{
			$myobj = @{
				Name = $Nic.Description
				MacAddress = $Nic.MACAddress
				IP4 = $Nic.IPAddress | where{ $_ -match "\d+\.\d+\.\d+\.\d+" }
				IP6 = $Nic.IPAddress | where{ $_ -match "\:\:" }
				IP4Subnet = $Nic.IPSubnet | where{ $_ -match "\d+\.\d+\.\d+\.\d+" }
				DefaultGWY = $Nic.DefaultIPGateway | Select -First 1
				DNSServer = $Nic.DNSServerSearchOrder
				WINSPrimary = $Nic.WINSPrimaryServer
				WINSSecondary = $Nic.WINSSecondaryServer
			}
			$obj = New-Object PSObject -Property $myobj
			$obj.PSTypeNames.Clear()
			$obj.PSTypeNames.Add('BSonPosh.IPInfo')
			$obj
		}
	}
}

#endregion 

#region Get-MemoryConfiguration 

function Get-MemoryConfiguration
{
	
	    <#
	        .Synopsis 
	            Gets the Memory Config for specified host.
	            
	        .Description
	            Gets the Memory Config for specified host.
	            
	        .Parameter ComputerName
	            Name of the Computer to get the Memory Config from (Default is localhost.)
	            
	        .Example
	            Get-MemoryConfiguration
	            Description
	            -----------
	            Gets Memory Config from local machine
	    
	        .Example
	            Get-MemoryConfiguration -ComputerName MyServer
	            Description
	            -----------
	            Gets Memory Config from MyServer
	            
	        .Example
	            $Servers | Get-MemoryConfiguration
	            Description
	            -----------
	            Gets Memory Config for each machine in the pipeline
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .Notes
	            NAME:      Get-MemoryConfiguration 
	            AUTHOR:    YetiCentral\bshell
	            Website:   www.bsonposh.com
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	
	Process
	{
		
		Write-Verbose " [Get-MemoryConfiguration] :: Begin Process"
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		if (Test-Host $ComputerName -TCPPort 135)
		{
			Write-Verbose " [Get-MemoryConfiguration] :: Processing $ComputerName"
			try
			{
				$MemorySlots = Get-WmiObject Win32_PhysicalMemory -ComputerName $ComputerName -ea STOP
				foreach ($Dimm in $MemorySlots)
				{
					$myobj = @{ }
					$myobj.ComputerName = $ComputerName
					$myobj.Description = $Dimm.Tag
					$myobj.Slot = $Dimm.DeviceLocator
					$myobj.Speed = $Dimm.Speed
					$myobj.SizeGB = $Dimm.Capacity/1gb
					
					$obj = New-Object PSObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.MemoryConfiguration')
					$obj
				}
			}
			catch
			{
				Write-Host " Host [$ComputerName] Failed with Error: $($Error[0])" -ForegroundColor Red
			}
		}
		else
		{
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
		Write-Verbose " [Get-MemoryConfiguration] :: End Process"
		
	}
}

#endregion 

#region Get-MotherBoard

function Get-MotherBoard
{
	
	    <#
	        .Synopsis 
	            Gets the Mother Board info for specified host.
	            
	        .Description
	            Gets the Mother Board info for specified host.
	            
	        .Parameter ComputerName
	            Name of the Computer to get the Mother Board info from (Default is localhost.) 
	            
	        .Example
	            Get-MotherBoard
	            Description
	            -----------
	            Gets Mother Board info from local machine
	    
	        .Example
	            Get-MotherBoard -ComputerName MyOtherDesktop
	            Description
	            -----------
	            Gets Mother Board info from MyOtherDesktop
	            
	        .Example
	            $Windows7Machines | Get-MotherBoard
	            Description
	            -----------
	            Gets Mother Board info for each machine in the pipeline
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            N/A
	            
	        .Notes
	            NAME:      Get-MotherBoard
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	
	Process
	{
		
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		if (Test-Host -ComputerName $ComputerName -TCPPort 135)
		{
			try
			{
				$MBInfo = Get-WmiObject Win32_BaseBoard -ComputerName $ComputerName -ea STOP
				$myobj = @{
					ComputerName = $ComputerName
					Name = $MBInfo.Product
					Manufacturer = $MBInfo.Manufacturer
					Version = $MBInfo.Version
					SerialNumber = $MBInfo.SerialNumber
				}
				
				$obj = New-Object PSObject -Property $myobj
				$obj.PSTypeNames.Clear()
				$obj.PSTypeNames.Add('BSonPosh.Computer.MotherBoard')
				$obj
			}
			catch
			{
				Write-Host " Host [$ComputerName] Failed with Error: $($Error[0])" -ForegroundColor Red
			}
		}
		else
		{
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
		
	}
}

#endregion # Get-MotherBoard

#region Get-NicInfo 

function Get-NICInfo
{
	
	    <#
	        .Synopsis  
	            Gets the NIC info for specified host
	            
	        .Description
	            Gets the NIC info for specified host
	            
	        .Parameter ComputerName
	            Name of the Computer to get the NIC info from (Default is localhost.)
	            
	        .Example
	            Get-NicInfo
	            # Gets NIC info from local machine
	    
	        .Example
	            Get-NicInfo -ComputerName MyServer
	            Description
	            -----------
	            Gets NIC info from MyServer
	            
	        .Example
	            $Servers | Get-NicInfo
	            Description
	            -----------
	            Gets NIC info for each machine in the pipeline
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .Notes
	            NAME:      Get-NicInfo 
	            AUTHOR:    YetiCentral\bshell
	            Website:   www.bsonposh.com
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	
	Process
	{
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		
		if (Test-Host -ComputerName $ComputerName -TCPPort 135)
		{
			try
			{
				$NICS = Get-WmiObject -class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName
				
				foreach ($NIC in $NICS)
				{
					$Query = "Select Name,NetConnectionID FROM Win32_NetworkAdapter WHERE Index='$($NIC.Index)'"
					$NetConnnectionID = Get-WmiObject -Query $Query -ComputerName $ComputerName
					
					$myobj = @{
						ComputerName = $ComputerName
						Name = $NetConnnectionID.Name
						NetID = $NetConnnectionID.NetConnectionID
						MacAddress = $NIC.MacAddress
						IP = $NIC.IPAddress | ?{ $_ -match "\d*\.\d*\.\d*\." }
						Subnet = $NIC.IPSubnet | ?{ $_ -match "\d*\.\d*\.\d*\." }
						Enabled = $NIC.IPEnabled
						Index = $NIC.Index
					}
					
					$obj = New-Object PSObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.NICInfo')
					$obj
				}
			}
			catch
			{
				Add-RichTextBoxWarn -text "Host $ComputerName Failed"
			}
		}
		else
		{
			Add-RichTextBoxWarn -text "Host $ComputerName Failed Connectivity Test"
		}
	}
}

#endregion 

#region Get-Processor

function Get-Processor
{
	
	    <#
	        .Synopsis 
	            Gets the Computer Processor info for specified host.
	            
	        .Description
	            Gets the Computer Processor info for specified host.
	            
	        .Parameter ComputerName
	            Name of the Computer to get the Computer Processor info from (Default is localhost.)
	            
	        .Example
	            Get-Processor
	            Description
	            -----------
	            Gets Computer Processor info from local machine
	    
	        .Example
	            Get-Processor -ComputerName MyServer
	            Description
	            -----------
	            Gets Computer Processor info from MyServer
	            
	        .Example
	            $Servers | Get-Processor
	            Description
	            -----------
	            Gets Computer Processor info for each machine in the pipeline
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            N/A
	            
	        .Notes
	            NAME:      Get-Processor
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	
	Process
	{
		
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		if (Test-Host -ComputerName $ComputerName -TCPPort 135)
		{
			try
			{
				$CPUS = Get-WmiObject Win32_Processor -ComputerName $ComputerName -ea STOP
				foreach ($CPU in $CPUs)
				{
					$myobj = @{
						ComputerName = $ComputerName
						Name = $CPU.Name
						Manufacturer = $CPU.Manufacturer
						Speed = $CPU.MaxClockSpeed
						Cores = $CPU.NumberOfCores
						L2Cache = $CPU.L2CacheSize
						Stepping = $CPU.Stepping
					}
				}
				$obj = New-Object PSObject -Property $myobj
				$obj.PSTypeNames.Clear()
				$obj.PSTypeNames.Add('BSonPosh.Computer.Processor')
				$obj
			}
			catch
			{
				Write-Host " Host [$ComputerName] Failed with Error: $($Error[0])" -ForegroundColor Red
			}
		}
		else
		{
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
		
	}
}

#endregion

#region Get-RegistryHive 

function Get-RegistryHive
{
	param ($HiveName)
	Switch -regex ($HiveName)
	{
		"^(HKCR|ClassesRoot|HKEY_CLASSES_ROOT)$"               { [Microsoft.Win32.RegistryHive]"ClassesRoot"; continue }
		"^(HKCU|CurrentUser|HKEY_CURRENTt_USER)$"              { [Microsoft.Win32.RegistryHive]"CurrentUser"; continue }
		"^(HKLM|LocalMachine|HKEY_LOCAL_MACHINE)$"          { [Microsoft.Win32.RegistryHive]"LocalMachine"; continue }
		"^(HKU|Users|HKEY_USERS)$"                          { [Microsoft.Win32.RegistryHive]"Users"; continue }
		"^(HKCC|CurrentConfig|HKEY_CURRENT_CONFIG)$"          { [Microsoft.Win32.RegistryHive]"CurrentConfig"; continue }
		"^(HKPD|PerformanceData|HKEY_PERFORMANCE_DATA)$"    { [Microsoft.Win32.RegistryHive]"PerformanceData"; continue }
		Default { 1; continue }
	}
}

#endregion 

#region Get-RegistryKey 

function Get-RegistryKey
{
	
	    <#
	        .Synopsis 
	            Gets the registry key provide by Path.
	            
	        .Description
	            Gets the registry key provide by Path.
	                        
	        .Parameter Path 
	            Path to the key.
	            
	        .Parameter ComputerName 
	            Computer to get the registry key from.
	            
	        .Parameter Recurse 
	            Recursively returns registry keys starting from the Path.
	        
	        .Parameter ReadWrite
	            Returns the Registry key in Read Write mode.
	            
	        .Example
	            Get-registrykey HKLM\Software\Adobe
	            Description
	            -----------
	            Returns the Registry key for HKLM\Software\Adobe
	            
	        .Example
	            Get-registrykey HKLM\Software\Adobe -ComputerName MyServer1
	            Description
	            -----------
	            Returns the Registry key for HKLM\Software\Adobe on MyServer1
	        
	        .Example
	            Get-registrykey HKLM\Software\Adobe -Recurse
	            Description
	            -----------
	            Returns the Registry key for HKLM\Software\Adobe and all child keys
	                    
	        .OUTPUTS
	            Microsoft.Win32.RegistryKey
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            New-RegistryKey
	            Remove-RegistryKey
	            Test-RegistryKey
	        .Notes
	            NAME:      Get-RegistryKey
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Alias("Server")]
		[Parameter(ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:ComputerName,
		[Parameter()]
		[switch]$Recurse,
		[Alias("RW")]
		[Parameter()]
		[switch]$ReadWrite
		
	)
	
	Begin
	{
		
		Write-Verbose " [Get-RegistryKey] :: Start Begin"
		Write-Verbose " [Get-RegistryKey] :: `$Path = $Path"
		Write-Verbose " [Get-RegistryKey] :: Getting `$Hive and `$KeyPath from $Path "
		$PathParts = $Path -split "\\|/", 0, "RegexMatch"
		$Hive = $PathParts[0]
		$KeyPath = $PathParts[1..$PathParts.count] -join "\"
		Write-Verbose " [Get-RegistryKey] :: `$Hive = $Hive"
		Write-Verbose " [Get-RegistryKey] :: `$KeyPath = $KeyPath"
		
		Write-Verbose " [Get-RegistryKey] :: End Begin"
		
	}
	
	Process
	{
		
		Write-Verbose " [Get-RegistryKey] :: Start Process"
		Write-Verbose " [Get-RegistryKey] :: `$ComputerName = $ComputerName"
		
		$RegHive = Get-RegistryHive $hive
		
		if ($RegHive -eq 1)
		{
			Write-Host "Invalid Path: $Path, Registry Hive [$hive] is invalid!" -ForegroundColor Red
		}
		else
		{
			Write-Verbose " [Get-RegistryKey] :: `$RegHive = $RegHive"
			
			$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
			Write-Verbose " [Get-RegistryKey] :: `$BaseKey = $BaseKey"
			
			if ($ReadWrite)
			{
				try
				{
					$Key = $BaseKey.OpenSubKey($KeyPath, $true)
					$Key = $Key | Add-Member -Name "ComputerName" -MemberType NoteProperty -Value $ComputerName -PassThru
					$Key = $Key | Add-Member -Name "Hive" -MemberType NoteProperty -Value $RegHive -PassThru
					$Key = $Key | Add-Member -Name "Path" -MemberType NoteProperty -Value $KeyPath -PassThru
					$Key.PSTypeNames.Clear()
					$Key.PSTypeNames.Add('BSonPosh.Registry.Key')
					$Key
				}
				catch
				{
					Write-Verbose " [Get-RegistryKey] ::  ERROR :: Unable to Open Key:$KeyPath in $KeyPath with RW Access"
				}
				
			}
			else
			{
				try
				{
					$Key = $BaseKey.OpenSubKey("$KeyPath")
					if ($Key)
					{
						$Key = $Key | Add-Member -Name "ComputerName" -MemberType NoteProperty -Value $ComputerName -PassThru
						$Key = $Key | Add-Member -Name "Hive" -MemberType NoteProperty -Value $RegHive -PassThru
						$Key = $Key | Add-Member -Name "Path" -MemberType NoteProperty -Value $KeyPath -PassThru
						$Key.PSTypeNames.Clear()
						$Key.PSTypeNames.Add('BSonPosh.Registry.Key')
						$Key
					}
				}
				catch
				{
					Write-Verbose " [Get-RegistryKey] ::  ERROR :: Unable to Open SubKey:$Name in $KeyPath"
				}
			}
			
			if ($Recurse)
			{
				Write-Verbose " [Get-RegistryKey] :: Recurse Passed: Processing Subkeys of [$($Key.Name)]"
				$Key
				$SubKeyNames = $Key.GetSubKeyNames()
				foreach ($Name in $SubKeyNames)
				{
					try
					{
						$SubKey = $Key.OpenSubKey($Name)
						if ($SubKey.GetSubKeyNames())
						{
							Write-Verbose " [Get-RegistryKey] :: Calling [Get-RegistryKey] for [$($SubKey.Name)]"
							Get-RegistryKey -ComputerName $ComputerName -Path $SubKey.Name -Recurse
						}
						else
						{
							Get-RegistryKey -ComputerName $ComputerName -Path $SubKey.Name
						}
					}
					catch
					{
						Write-Verbose " [Get-RegistryKey] ::  ERROR :: Write-Host Unable to Open SubKey:$Name in $($Key.Name)"
					}
				}
			}
		}
		Write-Verbose " [Get-RegistryKey] :: End Process"
		
	}
}

#endregion 

#region Get-RegistryValue 

function Get-RegistryValue
{
	
	    <#
	        .Synopsis 
	            Get the value for given the registry value.
	            
	        .Description
	            Get the value for given the registry value.
	                        
	        .Parameter Path 
	            Path to the key that contains the value.
	            
	        .Parameter Name 
	            Name of the Value to check.
	            
	        .Parameter ComputerName 
	            Computer to get value.
	            
	        .Parameter Recurse 
	            Recursively gets the Values on the given key.
	            
	        .Parameter Default 
	            Returns the default value for the Value.
	        
	        .Example
	            Get-RegistryValue HKLM\SOFTWARE\Adobe\SwInstall -Name State 
	            Description
	            -----------
	            Returns value of State under HKLM\SOFTWARE\Adobe\SwInstall.
	            
	        .Example
	            Get-RegistryValue HKLM\Software\Adobe -Name State -ComputerName MyServer1
	            Description
	            -----------
	            Returns value of State under HKLM\SOFTWARE\Adobe\SwInstall on MyServer1
	            
	        .Example
	            Get-RegistryValue HKLM\Software\Adobe -Recurse
	            Description
	            -----------
	            Returns all the values under HKLM\SOFTWARE\Adobe.
	    
	        .Example
	            Get-RegistryValue HKLM\Software\Adobe -ComputerName MyServer1 -Recurse
	            Description
	            -----------
	            Returns all the values under HKLM\SOFTWARE\Adobe on MyServer1
	            
	        .Example
	            Get-RegistryValue HKLM\Software\Adobe -Default
	            Description
	            -----------
	            Returns the default value for HKLM\SOFTWARE\Adobe.
	                    
	        .OUTPUTS
	            PSCustomObject
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            New-RegistryValue
	            Remove-RegistryValue
	            Test-RegistryValue
	            
	        .Notes    
	            NAME:      Get-RegistryValue
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Parameter()]
		[string]$Name,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:ComputerName,
		[Parameter()]
		[switch]$Recurse,
		[Parameter()]
		[switch]$Default
	)
	
	Process
	{
		
		Write-Verbose " [Get-RegistryValue] :: Begin Process"
		Write-Verbose " [Get-RegistryValue] :: Calling Get-RegistryKey -Path $path -ComputerName $ComputerName"
		
		if ($Recurse)
		{
			$Keys = Get-RegistryKey -Path $path -ComputerName $ComputerName -Recurse
			foreach ($Key in $Keys)
			{
				if ($Name)
				{
					try
					{
						Write-Verbose " [Get-RegistryValue] :: Getting Value for [$Name]"
						$myobj = @{ } #| Select ComputerName,Name,Value,Type,Path
						$myobj.ComputerName = $ComputerName
						$myobj.Name = $Name
						$myobj.value = $Key.GetValue($Name)
						$myobj.Type = $Key.GetValueKind($Name)
						$myobj.path = $Key
						
						$obj = New-Object PSCustomObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
						$obj
					}
					catch
					{
						Write-Verbose " [Get-RegistryValue] ::  ERROR :: Unable to Get Value for:$Name in $($Key.Name)"
					}
					
				}
				elseif ($Default)
				{
					try
					{
						Write-Verbose " [Get-RegistryValue] :: Getting Value for [(Default)]"
						$myobj = @{ } #"" | Select ComputerName,Name,Value,Type,Path
						$myobj.ComputerName = $ComputerName
						$myobj.Name = "(Default)"
						$myobj.value = if ($Key.GetValue("")) { $Key.GetValue("") }
						else { "EMPTY" }
						$myobj.Type = if ($Key.GetValue("")) { $Key.GetValueKind("") }
						else { "N/A" }
						$myobj.path = $Key
						
						$obj = New-Object PSCustomObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
						$obj
					}
					catch
					{
						Write-Verbose " [Get-RegistryValue] ::  ERROR :: Unable to Get Value for:(Default) in $($Key.Name)"
					}
				}
				else
				{
					try
					{
						Write-Verbose " [Get-RegistryValue] :: Getting all Values for [$Key]"
						foreach ($ValueName in $Key.GetValueNames())
						{
							Write-Verbose " [Get-RegistryValue] :: Getting all Value for [$ValueName]"
							$myobj = @{ } #"" | Select ComputerName,Name,Value,Type,Path
							$myobj.ComputerName = $ComputerName
							$myobj.Name = if ($ValueName -match "^$") { "(Default)" }
							else { $ValueName }
							$myobj.value = $Key.GetValue($ValueName)
							$myobj.Type = $Key.GetValueKind($ValueName)
							$myobj.path = $Key
							
							$obj = New-Object PSCustomObject -Property $myobj
							$obj.PSTypeNames.Clear()
							$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
							$obj
						}
					}
					catch
					{
						Write-Verbose " [Get-RegistryValue] ::  ERROR :: Unable to Get Value for:$ValueName in $($Key.Name)"
					}
				}
			}
		}
		else
		{
			$Key = Get-RegistryKey -Path $path -ComputerName $ComputerName
			Write-Verbose " [Get-RegistryValue] :: Get-RegistryKey returned $Key"
			if ($Name)
			{
				try
				{
					Write-Verbose " [Get-RegistryValue] :: Getting Value for [$Name]"
					$myobj = @{ } # | Select ComputerName,Name,Value,Type,Path
					$myobj.ComputerName = $ComputerName
					$myobj.Name = $Name
					$myobj.value = $Key.GetValue($Name)
					$myobj.Type = $Key.GetValueKind($Name)
					$myobj.path = $Key
					
					$obj = New-Object PSCustomObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
					$obj
				}
				catch
				{
					Write-Verbose " [Get-RegistryValue] ::  ERROR :: Unable to Get Value for:$Name in $($Key.Name)"
				}
			}
			elseif ($Default)
			{
				try
				{
					Write-Verbose " [Get-RegistryValue] :: Getting Value for [(Default)]"
					$myobj = @{ } #"" | Select ComputerName,Name,Value,Type,Path
					$myobj.ComputerName = $ComputerName
					$myobj.Name = "(Default)"
					$myobj.value = if ($Key.GetValue("")) { $Key.GetValue("") }
					else { "EMPTY" }
					$myobj.Type = if ($Key.GetValue("")) { $Key.GetValueKind("") }
					else { "N/A" }
					$myobj.path = $Key
					
					$obj = New-Object PSCustomObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
					$obj
				}
				catch
				{
					Write-Verbose " [Get-RegistryValue] ::  ERROR :: Unable to Get Value for:$Name in $($Key.Name)"
				}
			}
			else
			{
				Write-Verbose " [Get-RegistryValue] :: Getting all Values for [$Key]"
				foreach ($ValueName in $Key.GetValueNames())
				{
					Write-Verbose " [Get-RegistryValue] :: Getting all Value for [$ValueName]"
					$myobj = @{ } #"" | Select ComputerName,Name,Value,Type,Path
					$myobj.ComputerName = $ComputerName
					$myobj.Name = if ($ValueName -match "^$") { "(Default)" }
					else { $ValueName }
					$myobj.value = $Key.GetValue($ValueName)
					$myobj.Type = $Key.GetValueKind($ValueName)
					$myobj.path = $Key
					
					$obj = New-Object PSCustomObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
					$obj
				}
			}
		}
		
		Write-Verbose " [Get-RegistryValue] :: End Process"
		
	}
}

#endregion 

#region Get-Routetable 

function Get-Routetable
{
	
	    <#
	        .Synopsis 
	            Gets the route table for specified host.
	            
	        .Description
	            Gets the route table for specified host.
	            
	        .Parameter ComputerName
	            Name of the Computer to get the route table from (Default is localhost.)
	            
	        .Example
	            Get-RouteTable
	            Description
	            -----------
	            Gets route table from local machine
	    
	        .Example
	            Get-RouteTable -ComputerName MyServer
	            Description
	            -----------
	            Gets route table from MyServer
	            
	        .Example
	            $Servers | Get-RouteTable
	            Description
	            -----------
	            Gets route table for each machine in the pipeline
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            N/A
	            
	        .Notes
	            NAME:      Get-RouteTable
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	process
	{
		
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		if (Test-Host $ComputerName -TCPPort 135)
		{
			$Routes = Get-WMIObject Win32_IP4RouteTable -ComputerName $ComputerName -Property Name, Mask, NextHop, Metric1, Type
			foreach ($Route in $Routes)
			{
				$myobj = @{ }
				$myobj.ComputerName = $ComputerName
				$myobj.Name = $Route.Name
				$myobj.NetworkMask = $Route.mask
				$myobj.Gateway = if ($Route.NextHop -eq "0.0.0.0") { "On-Link" }
				else { $Route.NextHop }
				$myobj.Metric = $Route.Metric1
				
				$obj = New-Object PSObject -Property $myobj
				$obj.PSTypeNames.Clear()
				$obj.PSTypeNames.Add('BSonPosh.RouteTable')
				$obj
			}
		}
		else
		{
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
		
	}
}

#endregion 

#region Get-SystemType 

function Get-SystemType
{
	
	    <#
	        .Synopsis 
	            Gets the system type for specified host
	            
	        .Description
	            Gets the system type info for specified host
	            
	        .Parameter ComputerName
	            Name of the Computer to get the System Type from (Default is localhost.)
	            
	        .Example
	            Get-SystemType
	            Description
	            -----------
	            Gets System Type from local machine
	    
	        .Example
	            Get-SystemType -ComputerName MyServer
	            Description
	            -----------
	            Gets System Type from MyServer
	            
	        .Example
	            $Servers | Get-SystemType
	            Description
	            -----------
	            Gets System Type for each machine in the pipeline
	            
	        .OUTPUTS
	            PSObject
	            
	        .Notes
	            NAME:      Get-SystemType 
	            AUTHOR:    YetiCentral\bshell
	            Website:   www.bsonposh.com
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	
	Begin
	{
		
		function ConvertTo-ChassisType($Type)
		{
			switch ($Type)
			{
				1    { "Other" }
				2    { "Unknown" }
				3    { "Desktop" }
				4    { "Low Profile Desktop" }
				5    { "Pizza Box" }
				6    { "Mini Tower" }
				7    { "Tower" }
				8    { "Portable" }
				9    { "Laptop" }
				10    { "Notebook" }
				11    { "Hand Held" }
				12    { "Docking Station" }
				13    { "All in One" }
				14    { "Sub Notebook" }
				15    { "Space-Saving" }
				16    { "Lunch Box" }
				17    { "Main System Chassis" }
				18    { "Expansion Chassis" }
				19    { "SubChassis" }
				20    { "Bus Expansion Chassis" }
				21    { "Peripheral Chassis" }
				22    { "Storage Chassis" }
				23    { "Rack Mount Chassis" }
				24    { "Sealed-Case PC" }
			}
		}
		function ConvertTo-SecurityStatus($Status)
		{
			switch ($Status)
			{
				1    { "Other" }
				2    { "Unknown" }
				3    { "None" }
				4    { "External Interface Locked Out" }
				5    { "External Interface Enabled" }
			}
		}
		
	}
	Process
	{
		
		Write-Verbose " [Get-SystemType] :: Process Start"
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		if (Test-Host $ComputerName -TCPPort 135)
		{
			try
			{
				Write-Verbose " [Get-SystemType] :: Getting System (Enclosure) Type info use WMI"
				$SystemInfo = Get-WmiObject Win32_SystemEnclosure -ComputerName $ComputerName
				$CSInfo = Get-WmiObject -Query "Select Model FROM Win32_ComputerSystem" -ComputerName $ComputerName
				
				Write-Verbose " [Get-SystemType] :: Creating Hash Table"
				$myobj = @{ }
				Write-Verbose " [Get-SystemType] :: Setting ComputerName   - $ComputerName"
				$myobj.ComputerName = $ComputerName
				
				Write-Verbose " [Get-SystemType] :: Setting Manufacturer   - $($SystemInfo.Manufacturer)"
				$myobj.Manufacturer = $SystemInfo.Manufacturer
				
				Write-Verbose " [Get-SystemType] :: Setting Module   - $($CSInfo.Model)"
				$myobj.Model = $CSInfo.Model
				
				Write-Verbose " [Get-SystemType] :: Setting SerialNumber   - $($SystemInfo.SerialNumber)"
				$myobj.SerialNumber = $SystemInfo.SerialNumber
				
				Write-Verbose " [Get-SystemType] :: Setting SecurityStatus - $($SystemInfo.SecurityStatus)"
				$myobj.SecurityStatus = ConvertTo-SecurityStatus $SystemInfo.SecurityStatus
				
				Write-Verbose " [Get-SystemType] :: Setting Type           - $($SystemInfo.ChassisTypes)"
				$myobj.Type = ConvertTo-ChassisType $SystemInfo.ChassisTypes
				
				Write-Verbose " [Get-SystemType] :: Creating Custom Object"
				$obj = New-Object PSCustomObject -Property $myobj
				$obj.PSTypeNames.Clear()
				$obj.PSTypeNames.Add('BSonPosh.SystemType')
				$obj
			}
			catch
			{
				Write-Verbose " [Get-SystemType] :: [$ComputerName] Failed with Error: $($Error[0])"
			}
		}
		
	}
	
}

#endregion 

#region Get-USB

function Get-USB
{
	    <#
	    .Synopsis
	        Gets USB devices attached to the system
	    .Description
	        Uses WMI to get the USB Devices attached to the system
	    .Example
	        Get-USB
	    .Example
	        Get-USB | Group-Object Manufacturer  
	    .Parameter ComputerName
	        The name of the computer to get the USB devices from
	    #>
	param ($computerName = "localhost")
	Get-WmiObject Win32_USBControllerDevice -ComputerName $ComputerName `
				  -Impersonation Impersonate -Authentication PacketPrivacy |
	Foreach-Object { [Wmi]$_.Dependent }
}
#endregion

#region Get-UserTxtBox
function Get-UserTxtBox
{ $global:UserName_Txt = $usertextbox.Text }
#endregion

#region Get-LocalAdmins
function get-localadmins
{
	[cmdletbinding()]
	Param (
		[string]$computerName
	)
	$group = get-wmiobject win32_group -ComputerName $computerName -Filter "LocalAccount=True AND SID='S-1-5-32-544'"
	$query = "GroupComponent = `"Win32_Group.Domain='$($group.domain)'`,Name='$($group.name)'`""
	$list = Get-WmiObject win32_groupuser -ComputerName $computerName -Filter $query
	$list | %{ $_.PartComponent } | % { $_.substring($_.lastindexof("Domain=") + 7).replace("`",Name=`"", "\") }
}
#endregion Get-LocalAdmins

#region Get-ComputerStats
function Get-ComputerStats
{
	param (
		[Parameter(Mandatory = $true, Position = 0,
				   ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNull()]
		[string[]]$ComputerNames
	)
	
	process
	{
		$avg = Get-WmiObject win32_processor -computername $computername |
		Measure-Object -property LoadPercentage -Average |
		Foreach { $_.Average }
		$mem = Get-WmiObject win32_operatingsystem -ComputerName $computername |
		Foreach { "{0:N2}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100)/ $_.TotalVisibleMemorySize) }
		new-object psobject -prop @{
			# Work on PowerShell V2 and below
			# [pscustomobject] [ordered] @{ # Only if on PowerShell V3
			AverageCpuLoad = $avg
			MemoryUsagePercent = $mem
		}
	}
}
#endregion Get-ComputerStats

# New Functions

#region New-RegistryKey 

function New-RegistryKey
{
	
	    <#
	        .Synopsis 
	            Creates a new key in the provide by Path.
	            
	        .Description
	            Creates a new key in the provide by Path.
	                        
	        .Parameter Path 
	            Path to create the key in.
	            
	        .Parameter ComputerName 
	            Computer to the create registry key on.
	            
	        .Parameter Name 
	            Name of the Key to create
	        
	        .Example
	            New-registrykey HKLM\Software\Adobe -Name DeleteMe
	            Description
	            -----------
	            Creates a key called DeleteMe under HKLM\Software\Adobe
	            
	        .Example
	            New-registrykey HKLM\Software\Adobe -Name DeleteMe -ComputerName MyServer1
	            Description
	            -----------
	            Creates a key called DeleteMe under HKLM\Software\Adobe on MyServer1
	                    
	        .OUTPUTS
	            Microsoft.Win32.RegistryKey
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            Get-RegistryKey
	            Remove-RegistryKey
	            Test-RegistryKey
	            
	        NAME:      New-RegistryKey
	        AUTHOR:    bsonposh
	        Website:   http://www.bsonposh.com
	        Version:   1
	        #Requires -Version 2.0
	    #>
	[Cmdletbinding(SupportsShouldProcess = $true)]
	Param (
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Parameter(mandatory = $true)]
		[string]$Name,
		[Alias("Server")]
		[Parameter(ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:ComputerName
	)
	Begin
	{
		
		Write-Verbose " [New-RegistryKey] :: Start Begin"
		$ReadWrite = [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree
		
		Write-Verbose " [New-RegistryKey] :: `$Path = $Path"
		Write-Verbose " [New-RegistryKey] :: Getting `$Hive and `$KeyPath from $Path "
		$PathParts = $Path -split "\\|/", 0, "RegexMatch"
		$Hive = $PathParts[0]
		$KeyPath = $PathParts[1..$PathParts.count] -join "\"
		Write-Verbose " [New-RegistryKey] :: `$Hive = $Hive"
		Write-Verbose " [New-RegistryKey] :: `$KeyPath = $KeyPath"
		
		Write-Verbose " [New-RegistryKey] :: End Begin"
		
	}
	Process
	{
		
		Write-Verbose " [Get-RegistryKey] :: Start Process"
		Write-Verbose " [Get-RegistryKey] :: `$ComputerName = $ComputerName"
		
		$RegHive = Get-RegistryHive $hive
		
		if ($RegHive -eq 1)
		{
			Write-Host "Invalid Path: $Path, Registry Hive [$hive] is invalid!" -ForegroundColor Red
		}
		else
		{
			Write-Verbose " [Get-RegistryKey] :: `$RegHive = $RegHive"
			$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
			Write-Verbose " [Get-RegistryKey] :: `$BaseKey = $BaseKey"
			$Key = $BaseKey.OpenSubKey($KeyPath, $True)
			if ($PSCmdlet.ShouldProcess($ComputerName, "Creating Key [$Name] under $Path"))
			{
				$Key.CreateSubKey($Name, $ReadWrite)
			}
		}
		Write-Verbose " [Get-RegistryKey] :: End Process"
		
	}
}

#endregion 

#region New-RegistryValue 

function New-RegistryValue
{
	
	    <#
	        .Synopsis 
	            Create a value under the registry key.
	            
	        .Description
	            Create a value under the registry key.
	                        
	        .Parameter Path 
	            Path to the key.
	            
	        .Parameter Name 
	            Name of the Value to create.
	            
	        .Parameter Value 
	            Value to for the new Value.
	            
	        .Parameter Type
	            Type for the new Value. Valid Types: Unknown, String (default,) ExpandString, Binary, DWord, MultiString, a
	    nd Qword
	            
	        .Parameter ComputerName 
	            Computer to create the Value on.
	            
	        .Example
	            New-RegistryValue HKLM\SOFTWARE\Adobe\MyKey -Name State -Value "Hi There"
	            Description
	            -----------
	            Creates the Value State and sets the value to "Hi There" under HKLM\SOFTWARE\Adobe\MyKey.
	            
	        .Example
	            New-RegistryValue HKLM\SOFTWARE\Adobe\MyKey -Name State -Value 0 -ComputerName MyServer1
	            Description
	            -----------
	            Creates the Value State and sets the value to "Hi There" under HKLM\SOFTWARE\Adobe\MyKey on MyServer1.
	            
	        .Example
	            New-RegistryValue HKLM\SOFTWARE\Adobe\MyKey -Name MyDWord -Value 0 -Type DWord
	            Description
	            -----------
	            Creates the DWORD Value MyDWord and sets the value to 0 under HKLM\SOFTWARE\Adobe\MyKey.
	                    
	        .OUTPUTS
	            System.Boolean
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            New-RegistryValue
	            Remove-RegistryValue
	            Get-RegistryValue
	            
	        NAME:      Test-RegistryValue
	        AUTHOR:    bsonposh
	        Website:   http://www.bsonposh.com
	        Version:   1
	        #Requires -Version 2.0
	    #>
	
	[Cmdletbinding(SupportsShouldProcess = $true)]
	Param (
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Parameter(mandatory = $true)]
		[string]$Name,
		[Parameter()]
		[string]$Value,
		[Parameter()]
		[string]$Type,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:ComputerName
	)
	Begin
	{
		
		Write-Verbose " [New-RegistryValue] :: Start Begin"
		Write-Verbose " [New-RegistryValue] :: `$Path = $Path"
		Write-Verbose " [New-RegistryValue] :: `$Name = $Name"
		Write-Verbose " [New-RegistryValue] :: `$Value = $Value"
		
		Switch ($Type)
		{
			"Unknown"       { $ValueType = [Microsoft.Win32.RegistryValueKind]::Unknown; continue }
			"String"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
			"ExpandString"  { $ValueType = [Microsoft.Win32.RegistryValueKind]::ExpandString; continue }
			"Binary"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::Binary; continue }
			"DWord"         { $ValueType = [Microsoft.Win32.RegistryValueKind]::DWord; continue }
			"MultiString"   { $ValueType = [Microsoft.Win32.RegistryValueKind]::MultiString; continue }
			"QWord"         { $ValueType = [Microsoft.Win32.RegistryValueKind]::QWord; continue }
			default { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
		}
		Write-Verbose " [New-RegistryValue] :: `$Type = $Type"
		Write-Verbose " [New-RegistryValue] :: End Begin"
		
	}
	
	Process
	{
		
		if (Test-RegistryValue -Path $path -Name $Name -ComputerName $ComputerName)
		{
			"Registry value already exist"
		}
		else
		{
			Write-Verbose " [New-RegistryValue] :: Start Process"
			Write-Verbose " [New-RegistryValue] :: Calling Get-RegistryKey -Path $path -ComputerName $ComputerName"
			$Key = Get-RegistryKey -Path $path -ComputerName $ComputerName -ReadWrite
			Write-Verbose " [New-RegistryValue] :: Get-RegistryKey returned $Key"
			Write-Verbose " [New-RegistryValue] :: Setting Value for [$Name]"
			if ($PSCmdlet.ShouldProcess($ComputerName, "Creating Value [$Name] under $Path with value [$Value]"))
			{
				if ($Value)
				{
					$Key.SetValue($Name, $Value, $ValueType)
				}
				else
				{
					$Key.SetValue($Name, $ValueType)
				}
				Write-Verbose " [New-RegistryValue] :: Returning New Key: Get-RegistryValue -Path $path -Name $Name -ComputerName $ComputerName"
				Get-RegistryValue -Path $path -Name $Name -ComputerName $ComputerName
			}
		}
		Write-Verbose " [New-RegistryValue] :: End Process"
		
	}
}

#endregion 

# Load Functions

#region Load-ComboBox
function Load-ComboBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ComboBox.

	.DESCRIPTION
		Use this function to dynamically load items into the ComboBox control.

	.PARAMETER  ComboBox
		The ComboBox control you want to add items to.

	.PARAMETER  Items
		The object or objects you wish to load into the ComboBox's Items collection.

	.PARAMETER  DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER  Append
		Adds the item(s) to the ComboBox without clearing the Items collection.
	
	.EXAMPLE
		Load-ComboBox $combobox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Load-ComboBox $combobox1 "Red" -Append
		Load-ComboBox $combobox1 "White" -Append
		Load-ComboBox $combobox1 "Blue" -Append
	
	.EXAMPLE
		Load-ComboBox $combobox1 (Get-Process) "ProcessName"
#>
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ComboBox]$ComboBox,
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [Array])
	{
		$ComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ComboBox.Items.Add($obj)
		}
		$ComboBox.EndUpdate()
	}
	else
	{
		$ComboBox.Items.Add($Items)
	}
	
	$ComboBox.DisplayMember = $DisplayMember
}
#endregion Load-Combobox

#region Load-Chart
function Load-Chart
{
<#
	.SYNOPSIS
		This functions helps you plot points on a chart

	.DESCRIPTION
		Use the function to plot points on a chart or add more charts to a chart control

	.PARAMETER  ChartControl
		The Chart Control you when to add points to

	.PARAMETER  XPoints
		Set the X Axis Points. These can be strings or numerical values.

	.PARAMETER  YPoints
		Set the Y Axis Points. These can be strings or numerical values.
	
	.PARAMETER  XTitle
		Set the Title for the X Axis.

	.PARAMETER  YTitle
		Set the Title for the Y Axis.
	
	.PARAMETER  Title
		Set the Title for the chart.
	
	.PARAMETER  ChartType
		Set the Style of the chart. See System.Windows.Forms.DataVisualization.Charting.SeriesChartType Enum

	.PARAMETER SeriesIndex
		Set the settings of a particular Series and corresponding ChartArea

	.PARAMETER TitleIndex
		Set the settings of a particular Title
	
	.PARAMETER SeriesName
		Set the settings of a particular Series using its name and corresponding ChartArea. 
		The Series will be created if not found.
		If SeriesIndex is set, it will replace the Series' name if the Series does not exist
	
	.PARAMETER Enable3D
		The chart will be rendered in 3D.
	
	.PARAMETER Disable3D
		The chart will be rendered in 2D.	
	
	.PARAMETER Append
		When this switch is used, a new ChartArea is added to Chart Control.

	.LINK
		http://www.sapien.com/blog/2011/05/05/primalforms-2011-designing-charts-for-powershell/
	
#>
	Param (#$XPoints, $YPoints, $XTitle, $YTitle, $Title, $ChartStyle)
		[ValidateNotNull()]
		[Parameter(Position = 1, Mandatory = $true)]
		[System.Windows.Forms.DataVisualization.Charting.Chart]
		$ChartControl
		,
		[ValidateNotNull()]
		[Parameter(Position = 2, Mandatory = $true)]
		$XPoints
		,
		[Parameter(Position = 3, Mandatory = $true)]
		$YPoints
		,
		[Parameter(Position = 4, Mandatory = $false)]
		[string]$XTitle
		,
		[Parameter(Position = 5, Mandatory = $false)]
		[string]$YTitle
		,
		[Parameter(Position = 6, Mandatory = $false)]
		[string]$Title
		,
		[Parameter(Position = 7, Mandatory = $false)]
		[System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
		$ChartType
		,
		[Parameter(Position = 8, Mandatory = $false)]
		$SeriesIndex = -1
		,
		[Parameter(Position = 9, Mandatory = $false)]
		$TitleIndex = 0,
		[Parameter(Mandatory = $false)]
		[string]$SeriesName = $null,
		[switch]$Enable3D,
		[switch]$Disable3D,
		[switch]$Append)
	
	$ChartAreaIndex = 0
	if ($Append)
	{
		$name = "ChartArea " + ($ChartControl.ChartAreas.Count + 1).ToString();
		$ChartArea = $ChartControl.ChartAreas.Add($name)
		$ChartAreaIndex = $ChartControl.ChartAreas.Count - 1
		
		$name = "Series " + ($ChartControl.Series.Count + 1).ToString();
		$Series = $ChartControl.Series.Add($name)
		$SeriesIndex = $ChartControl.Series.Count - 1
		
		$Series.ChartArea = $ChartArea.Name
		
		if ($Title)
		{
			$name = "Title " + ($ChartControl.Titles.Count + 1).ToString();
			$TitleObj = $ChartControl.Titles.Add($Title)
			$TitleIndex = $ChartControl.Titles.Count - 1
			$TitleObj.DockedToChartArea = $ChartArea.Name
			$TitleObj.IsDockedInsideChartArea = $false
		}
	}
	else
	{
		if ($ChartControl.ChartAreas.Count -eq 0)
		{
			$name = "ChartArea " + ($ChartControl.ChartAreas.Count + 1).ToString();
			[void]$ChartControl.ChartAreas.Add($name)
			$ChartAreaIndex = $ChartControl.ChartAreas.Count - 1
		}
		
		if ($ChartControl.Series.Count -eq 0)
		{
			if (-not $SeriesName)
			{
				$SeriesName = "Series " + ($ChartControl.Series.Count + 1).ToString();
			}
			
			$Series = $ChartControl.Series.Add($SeriesName)
			$SeriesIndex = $ChartControl.Series.Count - 1
			$Series.ChartArea = $ChartControl.ChartAreas[$ChartAreaIndex].Name
		}
		elseif ($SeriesName)
		{
			$Series = $ChartControl.Series.FindByName($SeriesName)
			
			if ($Series -eq $null)
			{
				if (($SeriesIndex -gt -1) -and ($SeriesIndex -lt $ChartControl.Series.Count))
				{
					$Series = $ChartControl.Series[$SeriesIndex]
					$Series.Name = $SeriesName
				}
				else
				{
					$Series = $ChartControl.Series.Add($SeriesName)
					$SeriesIndex = $ChartControl.Series.Count - 1
				}
				
				$Series.ChartArea = $ChartControl.ChartAreas[$ChartAreaIndex].Name
			}
			else
			{
				$SeriesIndex = $ChartControl.Series.IndexOf($Series)
				$ChartAreaIndex = $ChartControl.ChartAreas.IndexOf($Series.ChartArea)
			}
		}
	}
	
	if (($SeriesIndex -lt 0) -or ($SeriesIndex -ge $ChartControl.Series.Count))
	{
		$SeriesIndex = 0
	}
	
	$Series = $ChartControl.Series[$SeriesIndex]
	$Series.Points.Clear()
	$ChartArea = $ChartControl.ChartAreas[$Series.ChartArea]
	
	if ($Enable3D)
	{
		$ChartArea.Area3DStyle.Enable3D = $true
	}
	elseif ($Disable3D)
	{
		$ChartArea.Area3DStyle.Enable3D = $false
	}
	
	if ($Title)
	{
		if ($ChartControl.Titles.Count -eq 0)
		{
			#$name = "Title " + ($ChartControl.Titles.Count + 1).ToString();
			$TitleObj = $ChartControl.Titles.Add($Title)
			$TitleIndex = $ChartControl.Titles.Count - 1
			$TitleObj.DockedToChartArea = $ChartArea.Name
			$TitleObj.IsDockedInsideChartArea = $false
		}
		
		$ChartControl.Titles[$TitleIndex].Text = $Title
	}
	
	if ($ChartType)
	{
		$Series.ChartType = $ChartType
	}
	
	if ($XTitle)
	{
		$ChartArea.AxisX.Title = $XTitle
	}
	
	if ($YTitle)
	{
		$ChartArea.AxisY.Title = $YTitle
	}
	
	if ($XPoints -isnot [Array] -or $XPoints -isnot [System.Collections.IEnumerable])
	{
		$array = New-Object System.Collections.ArrayList
		$array.Add($XPoints)
		$XPoints = $array
	}
	
	if ($YPoints -isnot [Array] -or $YPoints -isnot [System.Collections.IEnumerable])
	{
		$array = New-Object System.Collections.ArrayList
		$array.Add($YPoints)
		$YPoints = $array
	}
	
	$Series.Points.DataBindXY($XPoints, $YPoints)
	
}
#endregion Load-Chart

#region Load-ListBox
function Load-ListBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ListBox or CheckedListBox.

	.DESCRIPTION
		Use this function to dynamically load items into the ListBox control.

	.PARAMETER  ListBox
		The ListBox control you want to add items to.

	.PARAMETER  Items
		The object or objects you wish to load into the ListBox's Items collection.

	.PARAMETER  DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER  Append
		Adds the item(s) to the ListBox without clearing the Items collection.
	
	.EXAMPLE
		Load-ListBox $ListBox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Load-ListBox $listBox1 "Red" -Append
		Load-ListBox $listBox1 "White" -Append
		Load-ListBox $listBox1 "Blue" -Append
	
	.EXAMPLE
		Load-ListBox $listBox1 (Get-Process) "ProcessName"
#>
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ListBox]$ListBox,
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$listBox.Items.Clear()
	}
	
	if ($Items -is [System.Windows.Forms.ListBox+ObjectCollection])
	{
		$listBox.Items.AddRange($Items)
	}
	elseif ($Items -is [Array])
	{
		$listBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$listBox.Items.Add($obj)
		}
		$listBox.EndUpdate()
	}
	else
	{
		$listBox.Items.Add($Items)
	}
	
	$listBox.DisplayMember = $DisplayMember
}
#endregion Load-ListBo

# Remove Functions

#region Remove-RegistryKey 

function Remove-RegistryKey
{
	
	    <#
	        .Synopsis 
	            Removes a new key in the provide by Path.
	            
	        .Description
	            Removes a new key in the provide by Path.
	                        
	        .Parameter Path 
	            Path to remove the registry key from.
	            
	        .Parameter ComputerName 
	            Computer to remove the registry key from.
	            
	        .Parameter Name 
	            Name of the registry key to remove.
	            
	        .Parameter Recurse 
	            Recursively removes registry key and all children from path.
	        
	        .Example
	            Remove-registrykey HKLM\Software\Adobe -Name DeleteMe
	            Description
	            -----------
	            Removes the registry key called DeleteMe under HKLM\Software\Adobe
	            
	        .Example
	            Remove-RegistryKey HKLM\Software\Adobe -Name DeleteMe -ComputerName MyServer1
	            Description
	            -----------
	            Removes the key called DeleteMe under HKLM\Software\Adobe on MyServer1
	            
	        .Example
	            Remove-RegistryKey HKLM\Software\Adobe -Name DeleteMe -ComputerName MyServer1 -Recurse
	            Description
	            -----------
	            Removes the key called DeleteMe under HKLM\Software\Adobe on MyServer1 and all child keys.
	                    
	        .OUTPUTS
	            $null
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            Get-RegistryKey
	            New-RegistryKey
	            Test-RegistryKey
	            
	        .Notes
	        NAME:      Remove-RegistryKey
	        AUTHOR:    bsonposh
	        Website:   http://www.bsonposh.com
	        Version:   1
	        #Requires -Version 2.0
	    #>
	
	[Cmdletbinding(SupportsShouldProcess = $true)]
	Param (
		
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Parameter(mandatory = $true)]
		[string]$Name,
		[Alias("Server")]
		[Parameter(ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:ComputerName,
		[Parameter()]
		[switch]$Recurse
	)
	Begin
	{
		
		Write-Verbose " [Remove-RegistryKey] :: Start Begin"
		
		Write-Verbose " [Remove-RegistryKey] :: `$Path = $Path"
		Write-Verbose " [Remove-RegistryKey] :: Getting `$Hive and `$KeyPath from $Path "
		$PathParts = $Path -split "\\|/", 0, "RegexMatch"
		$Hive = $PathParts[0]
		$KeyPath = $PathParts[1..$PathParts.count] -join "\"
		Write-Verbose " [Remove-RegistryKey] :: `$Hive = $Hive"
		Write-Verbose " [Remove-RegistryKey] :: `$KeyPath = $KeyPath"
		
		Write-Verbose " [Remove-RegistryKey] :: End Begin"
		
	}
	
	Process
	{
		
		Write-Verbose " [Remove-RegistryKey] :: Start Process"
		Write-Verbose " [Remove-RegistryKey] :: `$ComputerName = $ComputerName"
		
		if (Test-RegistryKey -Path $path\$name -ComputerName $ComputerName)
		{
			$RegHive = Get-RegistryHive $hive
			
			if ($RegHive -eq 1)
			{
				Write-Host "Invalid Path: $Path, Registry Hive [$hive] is invalid!" -ForegroundColor Red
			}
			else
			{
				Write-Verbose " [Remove-RegistryKey] :: `$RegHive = $RegHive"
				$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
				Write-Verbose " [Remove-RegistryKey] :: `$BaseKey = $BaseKey"
				
				$Key = $BaseKey.OpenSubKey($KeyPath, $True)
				
				if ($PSCmdlet.ShouldProcess($ComputerName, "Deleteing Key [$Name]"))
				{
					if ($Recurse)
					{
						Write-Verbose " [Remove-RegistryKey] :: Calling DeleteSubKeyTree($Name)"
						$Key.DeleteSubKeyTree($Name)
					}
					else
					{
						Write-Verbose " [Remove-RegistryKey] :: Calling DeleteSubKey($Name)"
						$Key.DeleteSubKey($Name)
					}
				}
			}
		}
		else
		{
			"Key [$path\$name] does not exist"
		}
		Write-Verbose " [Remove-RegistryKey] :: End Process"
		
	}
}

#endregion 

#region Remove-RegistryValue 

function Remove-RegistryValue
{
	
	    <#
	        .Synopsis 
	            Removes the value.
	            
	        .Description
	            Removes the value.
	                        
	        .Parameter Path 
	            Path to the key that contains the value.
	            
	        .Parameter Name 
	            Name of the Value to Remove.
	    
	        .Parameter ComputerName 
	            Computer to remove value from.
	            
	        .Example
	            Remove-RegistryValue HKLM\SOFTWARE\Adobe\MyKey -Name State
	            Description
	            -----------
	            Removes the value STATE under HKLM\SOFTWARE\Adobe\MyKey.
	            
	        .Example
	            Remove-RegistryValue HKLM\Software\Adobe\MyKey -Name State -ComputerName MyServer1
	            Description
	            -----------
	            Removes the value STATE under HKLM\SOFTWARE\Adobe\MyKey on MyServer1.
	                    
	        .OUTPUTS
	            $null
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            New-RegistryValue
	            Test-RegistryValue
	            Get-RegistryValue
	            Set-RegistryValue
	            
	        NAME:      Remove-RegistryValue
	        AUTHOR:    bsonposh
	        Website:   http://www.bsonposh.com
	        Version:   1
	        #Requires -Version 2.0
	    #>
	
	[Cmdletbinding(SupportsShouldProcess = $true)]
	Param (
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Parameter(mandatory = $true)]
		[string]$Name,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:ComputerName
	)
	Begin
	{
		
		Write-Verbose " [Remove-RegistryValue] :: Start Begin"
		
		Write-Verbose " [Remove-RegistryValue] :: `$Path = $Path"
		Write-Verbose " [Remove-RegistryValue] :: `$Name = $Name"
		
		Write-Verbose " [Remove-RegistryValue] :: End Begin"
		
	}
	
	Process
	{
		
		if (Test-RegistryValue -Path $path -Name $Name -ComputerName $ComputerName)
		{
			Write-Verbose " [Remove-RegistryValue] :: Start Process"
			Write-Verbose " [Remove-RegistryValue] :: Calling Get-RegistryKey -Path $path -ComputerName $ComputerName"
			$Key = Get-RegistryKey -Path $path -ComputerName $ComputerName -ReadWrite
			Write-Verbose " [Remove-RegistryValue] :: Get-RegistryKey returned $Key"
			Write-Verbose " [Remove-RegistryValue] :: Setting Value for [$Name]"
			if ($PSCmdlet.ShouldProcess($ComputerName, "Deleting Value [$Name] under $Path"))
			{
				$Key.DeleteValue($Name)
			}
		}
		else
		{
			"Registry Value is already gone"
		}
		
		Write-Verbose " [Remove-RegistryValue] :: End Process"
		
	}
}

#endregion 

# Runner Functions

#region Run-RemoteCMD
#http://gallery.technet.microsoft.com/scriptcenter/56962f03-0243-4c83-8cdd-88c37898ccc4
function Run-RemoteCMD
{
	param (
		[Parameter(Mandatory = $true, valuefrompipeline = $true)]
		[string]$ComputerName,
		[string]$Command)
	begin
	{
		
		[string]$cmd = "CMD.EXE /C " + $command
	}
	process
	{
		$newproc = Invoke-WmiMethod -class Win32_process -name Create -ArgumentList ($cmd) -ComputerName $ComputerName
		if ($newproc.ReturnValue -eq 0)
		{ Add-RichTextBoxOK "Command $($command) invoked Sucessfully on $($ComputerName)" }
		# if command is sucessfully invoked it doesn't mean that it did what its supposed to do 
		#it means that the command only sucessfully ran on the cmd.exe of the server 
		#syntax errors can occur due to user input  
	}
	End { Write-Output "Script ...END" }
}
#endregion

# Searcher Functions

#region Search-Registry 

function Search-Registry
{
	
	    <#
	        .Synopsis 
	            Searchs the Registry.
	            
	        .Description
	            Searchs the Registry.
	                        
	        .Parameter Filter 
	            The RegEx filter you want to search for.
	            
	        .Parameter Name 
	            Name of the Key or Value you want to search for.
	        
	        .Parameter Value
	            Value to search for (Registry Values only.)
	            
	        .Parameter Path
	            Base of the Search. Should be in this format: "Software\Microsoft\..." See the Examples for specific exampl
	    es.
	            
	        .Parameter Hive
	            The Base Hive to search in (Default to LocalMachine.)
	            
	        .Parameter ComputerName 
	            Computer to search.
	            
	        .Parameter KeyOnly
	            Only returns Registry Keys. Not valid with -value parameter.
	            
	        .Example
	            Search-Registry -Hive HKLM -Filter "Powershell" -Path "SOFTWARE\Clients"
	            Description
	            -----------
	            Searchs the Registry for Keys or Values that match 'Powershell" in path "SOFTWARE\Clients"
	            
	        .Example
	            Search-Registry -Hive HKLM -Filter "Powershell" -Path "SOFTWARE\Clients" -computername MyServer1
	            Description
	            -----------
	            Searchs the Registry for Keys or Values that match 'Powershell" in path "SOFTWARE\Clients" on MyServer1
	            
	        .Example
	            Search-Registry -Hive HKLM -Name "Powershell" -Path "SOFTWARE\Clients"
	            Description
	            -----------
	            Searchs the Registry keys and values with name 'Powershell' in "SOFTWARE\Clients"
	            
	        .Example
	            Search-Registry -Hive HKLM -Name "Powershell" -Path "SOFTWARE\Clients" -KeyOnly
	            Description
	            -----------
	            Searchs the Registry keys with name 'Powershell' in "SOFTWARE\Clients"
	        
	        .Example
	            Search-Registry -Hive HKLM -Value "Powershell" -Path "SOFTWARE\Clients"
	            Description
	            -----------
	            Searchs the Registry Values with Value of 'Powershell' in "SOFTWARE\Clients"
	            
	        .OUTPUTS
	            Microsoft.Win32.RegistryKey
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            Get-RegistryKey
	            Get-RegistryValue
	            Test-RegistryKey
	        
	        .Notes
	            NAME:      Search-Registry
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding(DefaultParameterSetName = "ByFilter")]
	Param (
		[Parameter(ParameterSetName = "ByFilter", Position = 0)]
		[string]$Filter = ".*",
		[Parameter(ParameterSetName = "ByName", Position = 0)]
		[string]$Name,
		[Parameter(ParameterSetName = "ByValue", Position = 0)]
		[string]$Value,
		[Parameter()]
		[string]$Path,
		[Parameter()]
		[string]$Hive = "LocalMachine",
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME,
		[Parameter()]
		[switch]$KeyOnly
	)
	Begin
	{
		
		Write-Verbose " [Search-Registry] :: Start Begin"
		
		Write-Verbose " [Search-Registry] :: Active Parameter Set $($PSCmdlet.ParameterSetName)"
		switch ($PSCmdlet.ParameterSetName)
		{
			"ByFilter"    { Write-Verbose " [Search-Registry] :: `$Filter = $Filter" }
			"ByName"    { Write-Verbose " [Search-Registry] :: `$Name = $Name" }
			"ByValue"    { Write-Verbose " [Search-Registry] :: `$Value = $Value" }
		}
		$RegHive = Get-RegistryHive $Hive
		Write-Verbose " [Search-Registry] :: `$Hive = $RegHive"
		Write-Verbose " [Search-Registry] :: `$KeyOnly = $KeyOnly"
		
		Write-Verbose " [Search-Registry] :: End Begin"
		
	}
	
	Process
	{
		
		Write-Verbose " [Search-Registry] :: Start Process"
		
		Write-Verbose " [Search-Registry] :: `$ComputerName = $ComputerName"
		switch ($PSCmdlet.ParameterSetName)
		{
			"ByFilter"    {
				if ($KeyOnly)
				{
					if ($Path -and (Test-RegistryKey "$RegHive\$Path"))
					{
						Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -match "$Filter" }
					}
					else
					{
						$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
						foreach ($SubKeyName in $BaseKey.GetSubKeyNames())
						{
							try
							{
								$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
								Get-RegistryKey -Path $SubKey.Name -ComputerName $ComputerName -Recurse | ?{ $_.Name -match "$Filter" }
							}
							catch
							{
								Write-Host "Access Error on Key [$SubKeyName]... skipping."
							}
						}
					}
				}
				else
				{
					if ($Path -and (Test-RegistryKey "$RegHive\$Path"))
					{
						Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -match "$Filter" }
						Get-RegistryValue -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -match "$Filter" }
					}
					else
					{
						$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
						foreach ($SubKeyName in $BaseKey.GetSubKeyNames())
						{
							try
							{
								$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
								Get-RegistryKey -Path $SubKey.Name -ComputerName $ComputerName -Recurse | ?{ $_.Name -match "$Filter" }
								Get-RegistryValue -Path $SubKey.Name -ComputerName $ComputerName -Recurse | ?{ $_.Name -match "$Filter" }
							}
							catch
							{
								Write-Host "Access Error on Key [$SubKeyName]... skipping."
							}
						}
					}
				}
			}
			"ByName"    {
				if ($KeyOnly)
				{
					if ($Path -and (Test-RegistryKey "$RegHive\$Path"))
					{
						$NameFilter = "^.*\\{0}$" -f $Name
						Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -match $NameFilter }
					}
					else
					{
						$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
						foreach ($SubKeyName in $BaseKey.GetSubKeyNames())
						{
							try
							{
								$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
								$NameFilter = "^.*\\{0}$" -f $Name
								Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -match $NameFilter }
							}
							catch
							{
								Write-Host "Access Error on Key [$SubKeyName]... skipping."
							}
						}
					}
				}
				else
				{
					if ($Path -and (Test-RegistryKey "$RegHive\$Path"))
					{
						$NameFilter = "^.*\\{0}$" -f $Name
						Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -match $NameFilter }
						Get-RegistryValue -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -eq $Name }
					}
					else
					{
						$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
						foreach ($SubKeyName in $BaseKey.GetSubKeyNames())
						{
							try
							{
								$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
								$NameFilter = "^.*\\{0}$" -f $Name
								Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Name -match $NameFilter }
								Get-RegistryValue -Path $SubKey.Name -ComputerName $ComputerName -Recurse | ?{ $_.Name -eq $Name }
							}
							catch
							{
								Write-Host "Access Error on Key [$SubKeyName]... skipping."
							}
						}
					}
				}
			}
			"ByValue"    {
				if ($Path -and (Test-RegistryKey "$RegHive\$Path"))
				{
					Get-RegistryValue -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Value -eq $Value }
				}
				else
				{
					$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
					foreach ($SubKeyName in $BaseKey.GetSubKeyNames())
					{
						try
						{
							$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
							Get-RegistryValue -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | ?{ $_.Value -eq $Value }
						}
						catch
						{
							Write-Host "Access Error on Key [$SubKeyName]... skipping."
						}
					}
				}
			}
		}
		
		Write-Verbose " [Search-Registry] :: End Process"
		
	}
}

#endregion 

# Sender Functions

#region Send-WOL
function Send-WOL
{
<#  
  .SYNOPSIS   
    Send a WOL packet to a broadcast address 
  .PARAMETER mac 
   The MAC address of the device that need to wake up 
  .PARAMETER ip 
   The IP address where the WOL packet will be sent to 
  .EXAMPLE  
   Send-WOL -mac 00:11:32:21:2D:11 -ip 192.168.8.255  
#>	
	
	param (
		[string]$mac,
		[string]$ip,
		[int]$port = 9
	)
	$broadcast = [Net.IPAddress]::Parse($ip)
	
	$mac = (($mac.replace(":", "")).replace("-", "")).replace(".", "")
	$target = 0, 2, 4, 6, 8, 10 | % { [convert]::ToByte($mac.substring($_, 2), 16) }
	$packet = (, [byte]255 * 6) + ($target * 16)
	
	$UDPclient = new-Object System.Net.Sockets.UdpClient
	$UDPclient.Connect($broadcast, $port)
	[void]$UDPclient.Send($packet, 102)
	
}
#endregion Send-WOL

# Setter Functions

#region Set-RegistryValue 

function Set-RegistryValue
{
	
	    <#
	        .Synopsis 
	            Sets a value under the registry key.
	            
	        .Description
	            Sets a value under the registry key.
	                        
	        .Parameter Path 
	            Path to the key.
	            
	        .Parameter Name 
	            Name of the Value to Set.
	            
	        .Parameter Value 
	            New Value.
	            
	        .Parameter Type
	            Type for the Value. Valid Types: Unknown, String (default,) ExpandString, Binary, DWord, MultiString, and Q
	    word
	            
	        .Parameter ComputerName 
	            Computer to set the Value on.
	            
	        .Example
	            Set-RegistryValue HKLM\SOFTWARE\Adobe\MyKey -Name State -Value "Hi There"
	            Description
	            -----------
	            Sets the Value State and sets the value to "Hi There" under HKLM\SOFTWARE\Adobe\MyKey.
	            
	        .Example
	            Set-RegistryValue HKLM\SOFTWARE\Adobe\MyKey -Name State -Value 0 -ComputerName MyServer1
	            Description
	            -----------
	            Sets the Value State and sets the value to "Hi There" under HKLM\SOFTWARE\Adobe\MyKey on MyServer1.
	            
	        .Example
	            Set-RegistryValue HKLM\SOFTWARE\Adobe\MyKey -Name MyDWord -Value 0 -Type DWord
	            Description
	            -----------
	            Sets the DWORD Value MyDWord and sets the value to 0 under HKLM\SOFTWARE\Adobe\MyKey.
	            
	        .OUTPUTS
	            PSCustomObject
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            New-RegistryValue
	            Remove-RegistryValue
	            Get-RegistryValue
	            Test-RegistryValue
	        
	        .Notes
	            NAME:      Set-RegistryValue
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding(SupportsShouldProcess = $true)]
	Param (
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Parameter(mandatory = $true)]
		[string]$Name,
		[Parameter()]
		[string]$Value,
		[Parameter()]
		[string]$Type,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:ComputerName
	)
	
	Begin
	{
		
		Write-Verbose " [Set-RegistryValue] :: Start Begin"
		
		Write-Verbose " [Set-RegistryValue] :: `$Path = $Path"
		Write-Verbose " [Set-RegistryValue] :: `$Name = $Name"
		Write-Verbose " [Set-RegistryValue] :: `$Value = $Value"
		
		Switch ($Type)
		{
			"Unknown"       { $ValueType = [Microsoft.Win32.RegistryValueKind]::Unknown; continue }
			"String"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
			"ExpandString"  { $ValueType = [Microsoft.Win32.RegistryValueKind]::ExpandString; continue }
			"Binary"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::Binary; continue }
			"DWord"         { $ValueType = [Microsoft.Win32.RegistryValueKind]::DWord; continue }
			"MultiString"   { $ValueType = [Microsoft.Win32.RegistryValueKind]::MultiString; continue }
			"QWord"         { $ValueType = [Microsoft.Win32.RegistryValueKind]::QWord; continue }
			default { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
		}
		Write-Verbose " [Set-RegistryValue] :: `$Type = $Type"
		
		Write-Verbose " [Set-RegistryValue] :: End Begin"
		
	}
	
	Process
	{
		
		Write-Verbose " [Set-RegistryValue] :: Start Process"
		
		Write-Verbose " [Set-RegistryValue] :: Calling Get-RegistryKey -Path $path -ComputerName $ComputerName"
		$Key = Get-RegistryKey -Path $path -ComputerName $ComputerName -ReadWrite
		Write-Verbose " [Set-RegistryValue] :: Get-RegistryKey returned $Key"
		Write-Verbose " [Set-RegistryValue] :: Setting Value for [$Name]"
		if ($PSCmdlet.ShouldProcess($ComputerName, "Creating Value [$Name] under $Path with value [$Value]"))
		{
			if ($Value)
			{
				$Key.SetValue($Name, $Value, $ValueType)
			}
			else
			{
				$Key.SetValue($Name, $ValueType)
			}
			Write-Verbose " [Set-RegistryValue] :: Returning New Key: Get-RegistryValue -Path $path -Name $Name -ComputerName $ComputerName"
			Get-RegistryValue -Path $path -Name $Name -ComputerName $ComputerName
		}
		Write-Verbose " [Set-RegistryValue] :: End Process"
		
	}
}

#endregion 

# Show Functions

#region Show-MsgBox
	<# 
	            .SYNOPSIS  
	            Shows a graphical message box, with various prompt types available. 
	 
	            .DESCRIPTION 
	            Emulates the Visual Basic MsgBox function.  It takes four parameters, of which only the prompt is mandatory 
	 
	            .INPUTS 
	            The parameters are:- 
	             
	            Prompt (mandatory):  
	                Text string that you wish to display 
	                 
	            Title (optional): 
	                The title that appears on the message box 
	                 
	            Icon (optional).  Available options are: 
	                Information, Question, Critical, Exclamation (not case sensitive) 
	                
	            BoxType (optional). Available options are: 
	                OKOnly, OkCancel, AbortRetryIgnore, YesNoCancel, YesNo, RetryCancel (not case sensitive) 
	                 
	            DefaultButton (optional). Available options are: 
	                1, 2, 3 
	 
	            .OUTPUTS 
	            Microsoft.VisualBasic.MsgBoxResult 
	 
	            .EXAMPLE 
	            C:\PS> Show-MsgBox Hello 
	            Shows a popup message with the text "Hello", and the default box, icon and defaultbutton settings. 
	 
	            .EXAMPLE 
	            C:\PS> Show-MsgBox -Prompt "This is the prompt" -Title "This Is The Title" -Icon Critical -BoxType YesNo -DefaultButton 2 
	            Shows a popup with the parameter as supplied. 
	 
	            .LINK 
	            http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.msgboxresult.aspx 
	 
	            .LINK 
	            http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.msgboxstyle.aspx 
	            #>
# By BigTeddy August 24, 2011 
# http://social.technet.microsoft.com/profile/bigteddy/. 

function Show-MsgBox
{
	
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]$Prompt,
		[Parameter(Position = 1, Mandatory = $false)]
		[string]$Title = "",
		[Parameter(Position = 2, Mandatory = $false)]
		[ValidateSet("Information", "Question", "Critical", "Exclamation")]
		[string]$Icon = "Information",
		[Parameter(Position = 3, Mandatory = $false)]
		[ValidateSet("OKOnly", "OKCancel", "AbortRetryIgnore", "YesNoCancel", "YesNo", "RetryCancel")]
		[string]$BoxType = "OkOnly",
		[Parameter(Position = 4, Mandatory = $false)]
		[ValidateSet(1, 2, 3)]
		[int]$DefaultButton = 1
	)
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
	switch ($Icon)
	{
		"Question" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Question }
		"Critical" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Critical }
		"Exclamation" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Exclamation }
		"Information" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Information }
	}
	switch ($BoxType)
	{
		"OKOnly" { $vb_box = [microsoft.visualbasic.msgboxstyle]::OKOnly }
		"OKCancel" { $vb_box = [microsoft.visualbasic.msgboxstyle]::OkCancel }
		"AbortRetryIgnore" { $vb_box = [microsoft.visualbasic.msgboxstyle]::AbortRetryIgnore }
		"YesNoCancel" { $vb_box = [microsoft.visualbasic.msgboxstyle]::YesNoCancel }
		"YesNo" { $vb_box = [microsoft.visualbasic.msgboxstyle]::YesNo }
		"RetryCancel" { $vb_box = [microsoft.visualbasic.msgboxstyle]::RetryCancel }
	}
	switch ($Defaultbutton)
	{
		1 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton1 }
		2 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton2 }
		3 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton3 }
	}
	$popuptype = $vb_icon -bor $vb_box -bor $vb_defaultbutton
	$ans = [Microsoft.VisualBasic.Interaction]::MsgBox($prompt, $popuptype, $title)
	return $ans
} #end
#endregion

#region Show-InputBox
#http://www.sapien.com/forums/scriptinganswers/forum_posts.asp?TID=2890
#$c=Show-Inputbox -message "Enter a computername" -title "Computername" -default $env:Computername
#
#if ($c.Trim()) {
#  Get-WmiObject win32_computersystem -computer $c
#  }
Function Show-InputBox
{
	Param ([string]$message = $(Throw "You must enter a prompt message"),
		[string]$title = "Input",
		[string]$default
	)
	
	[reflection.assembly]::loadwithpartialname("microsoft.visualbasic") | Out-Null
	[microsoft.visualbasic.interaction]::InputBox($message, $title, $default)
	
}
#endregion

# Sort Functions

#region Sort-ListViewColumn
function Sort-ListViewColumn
{
	<#
	.SYNOPSIS
		Sort the ListView's item using the specified column.

	.DESCRIPTION
		Sort the ListView's item using the specified column.
		This function uses Add-Type to define a class that sort the items.
		The ListView's Tag property is used to keep track of the sorting.

	.PARAMETER ListView
		The ListView control to sort.

	.PARAMETER ColumnIndex
		The index of the column to use for sorting.
		
	.PARAMETER  SortOrder
		The direction to sort the items. If not specified or set to None, it will toggle.
	
	.EXAMPLE
		Sort-ListViewColumn -ListView $listview1 -ColumnIndex 0
#>
	param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ListView]$ListView,
		[Parameter(Mandatory = $true)]
		[int]$ColumnIndex,
		[System.Windows.Forms.SortOrder]$SortOrder = 'None')
	
	if (($ListView.Items.Count -eq 0) -or ($ColumnIndex -lt 0) -or ($ColumnIndex -ge $ListView.Columns.Count))
	{
		return;
	}
	
	#region Define ListViewItemComparer
	try
	{
		$local:type = [ListViewItemComparer]
	}
	catch
	{
		Add-Type -ReferencedAssemblies ('System.Windows.Forms') -TypeDefinition  @" 
	using System;
	using System.Windows.Forms;
	using System.Collections;
	public class ListViewItemComparer : IComparer
	{
	    public int column;
	    public SortOrder sortOrder;
	    public ListViewItemComparer()
	    {
	        column = 0;
			sortOrder = SortOrder.Ascending;
	    }
	    public ListViewItemComparer(int column, SortOrder sort)
	    {
	        this.column = column;
			sortOrder = sort;
	    }
	    public int Compare(object x, object y)
	    {
			if(column >= ((ListViewItem)x).SubItems.Count)
				return  sortOrder == SortOrder.Ascending ? -1 : 1;
		
			if(column >= ((ListViewItem)y).SubItems.Count)
				return sortOrder == SortOrder.Ascending ? 1 : -1;
		
			if(sortOrder == SortOrder.Ascending)
	        	return String.Compare(((ListViewItem)x).SubItems[column].Text, ((ListViewItem)y).SubItems[column].Text);
			else
				return String.Compare(((ListViewItem)y).SubItems[column].Text, ((ListViewItem)x).SubItems[column].Text);
	    }
	}
"@ | Out-Null
	}
	#endregion
	
	if ($ListView.Tag -is [ListViewItemComparer])
	{
		#Toggle the Sort Order
		if ($SortOrder -eq [System.Windows.Forms.SortOrder]::None)
		{
			if ($ListView.Tag.column -eq $ColumnIndex -and $ListView.Tag.sortOrder -eq 'Ascending')
			{
				$ListView.Tag.sortOrder = 'Descending'
			}
			else
			{
				$ListView.Tag.sortOrder = 'Ascending'
			}
		}
		else
		{
			$ListView.Tag.sortOrder = $SortOrder
		}
		
		$ListView.Tag.column = $ColumnIndex
		$ListView.Sort() #Sort the items
	}
	else
	{
		if ($Sort -eq [System.Windows.Forms.SortOrder]::None)
		{
			$Sort = [System.Windows.Forms.SortOrder]::Ascending
		}
		
		#Set to Tag because for some reason in PowerShell ListViewItemSorter prop returns null
		$ListView.Tag = New-Object ListViewItemComparer ($ColumnIndex, $SortOrder)
		$ListView.ListViewItemSorter = $ListView.Tag #Automatically sorts
	}
}
#endregion

# Tester Functions

#region Test-Host 

function Test-Host
{
	
	    <#
	        .Synopsis 
	            Test a host for connectivity using either WMI ping or TCP port
	            
	        .Description
	            Allows you to test a host for connectivity before further processing
	            
	        .Parameter Server
	            Name of the Server to Process.
	            
	        .Parameter TCPPort
	            TCP Port to connect to. (default 135)
	            
	        .Parameter Timeout
	            Timeout for the TCP connection (default 1 sec)
	            
	        .Parameter Property
	            Name of the Property that contains the value to test.
	            
	        .Example
	            cat ServerFile.txt | Test-Host | Invoke-DoSomething
	            Description
	            -----------
	            To test a list of hosts.
	            
	        .Example
	            cat ServerFile.txt | Test-Host -tcp 80 | Invoke-DoSomething
	            Description
	            -----------
	            To test a list of hosts against port 80.
	            
	        .Example
	            Get-ADComputer | Test-Host -property dnsHostname | Invoke-DoSomething
	            Description
	            -----------
	            To test the output of Get-ADComputer using the dnshostname property
	            
	            
	        .OUTPUTS
	            System.Object
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            Test-Port
	            
	        NAME:      Test-Host
	        AUTHOR:    YetiCentral\bshell
	        Website:   www.bsonposh.com
	        LASTEDIT:  02/04/2009 18:25:15
	        #Requires -Version 2.0
	    #>
	
	[CmdletBinding()]
	Param (
		
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $True)]
		[string]$ComputerName,
		[Parameter()]
		[int]$TCPPort = 80,
		[Parameter()]
		[int]$timeout = 3000,
		[Parameter()]
		[string]$property
		
	)
	Begin
	{
		
		function PingServer
		{
			Param ($MyHost)
			$ErrorActionPreference = "SilentlyContinue"
			Write-Verbose " [PingServer] :: Pinging [$MyHost]"
			try
			{
				$pingresult = Get-WmiObject win32_pingstatus -f "address='$MyHost'"
				$ResultCode = $pingresult.statuscode
				Write-Verbose " [PingServer] :: Ping returned $ResultCode"
				if ($ResultCode -eq 0) { $true }
				else { $false }
			}
			catch
			{
				Write-Verbose " [PingServer] :: Ping Failed with Error: ${error[0]}"
				$false
			}
		}
		
	}
	
	Process
	{
		
		Write-Verbose " [Test-Host] :: Begin Process"
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		Write-Verbose " [Test-Host] :: ComputerName   : $ComputerName"
		if ($TCPPort)
		{
			Write-Verbose " [Test-Host] :: Timeout  : $timeout"
			Write-Verbose " [Test-Host] :: Port     : $TCPPort"
			if ($property)
			{
				Write-Verbose " [Test-Host] :: Property : $Property"
				$Result = Test-Port $_.$property -tcp $TCPPort -timeout $timeout
				if ($Result)
				{
					if ($_) { $_ }
					else { $ComputerName }
				}
			}
			else
			{
				Write-Verbose " [Test-Host] :: Running - 'Test-Port $ComputerName -tcp $TCPPort -timeout $timeout'"
				$Result = Test-Port $ComputerName -tcp $TCPPort -timeout $timeout
				if ($Result)
				{
					if ($_) { $_ }
					else { $ComputerName }
				}
			}
		}
		else
		{
			if ($property)
			{
				Write-Verbose " [Test-Host] :: Property : $Property"
				try
				{
					if (PingServer $_.$property)
					{
						if ($_) { $_ }
						else { $ComputerName }
					}
				}
				catch
				{
					Write-Verbose " [Test-Host] :: $($_.$property) Failed Ping"
				}
			}
			else
			{
				Write-Verbose " [Test-Host] :: Simple Ping"
				try
				{
					if (PingServer $ComputerName) { $ComputerName }
				}
				catch
				{
					Write-Verbose " [Test-Host] :: $ComputerName Failed Ping"
				}
			}
		}
		Write-Verbose " [Test-Host] :: End Process"
		
	}
	
}

#endregion 

#region Test-Port 

function Test-Port
{
	
	    <#
	        .Synopsis 
	            Test a host to see if the specified port is open.
	            
	        .Description
	            Test a host to see if the specified port is open.
	                        
	        .Parameter TCPPort 
	            Port to test (Default 135.)
	            
	        .Parameter Timeout 
	            How long to wait (in milliseconds) for the TCP connection (Default 3000.)
	            
	        .Parameter ComputerName 
	            Computer to test the port against (Default in localhost.)
	            
	        .Example
	            Test-Port -tcp 3389
	            Description
	            -----------
	            Returns $True if the localhost is listening on 3389
	            
	        .Example
	            Test-Port -tcp 3389 -ComputerName MyServer1
	            Description
	            -----------
	            Returns $True if MyServer1 is listening on 3389
	                    
	        .OUTPUTS
	            System.Boolean
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            Test-Host
	            Wait-Port
	            
	        .Notes
	            NAME:      Test-Port
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		[Parameter()]
		[int]$TCPport = 135,
		[Parameter()]
		[int]$TimeOut = 3000,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[String]$ComputerName = $env:COMPUTERNAME
	)
	Begin
	{
		Write-Verbose " [Test-Port] :: Start Script"
		Write-Verbose " [Test-Port] :: Setting Error state = 0"
	}
	
	Process
	{
		
		Write-Verbose " [Test-Port] :: Creating [system.Net.Sockets.TcpClient] instance"
		$tcpclient = New-Object system.Net.Sockets.TcpClient
		
		Write-Verbose " [Test-Port] :: Calling BeginConnect($ComputerName,$TCPport,$null,$null)"
		try
		{
			$iar = $tcpclient.BeginConnect($ComputerName, $TCPport, $null, $null)
			Write-Verbose " [Test-Port] :: Waiting for timeout [$timeout]"
			$wait = $iar.AsyncWaitHandle.WaitOne($TimeOut, $false)
		}
		catch [System.Net.Sockets.SocketException]
		{
			Write-Verbose " [Test-Port] :: Exception: $($_.exception.message)"
			Write-Verbose " [Test-Port] :: End"
			return $false
		}
		catch
		{
			Write-Verbose " [Test-Port] :: General Exception"
			Write-Verbose " [Test-Port] :: End"
			return $false
		}
		
		if (!$wait)
		{
			$tcpclient.Close()
			Write-Verbose " [Test-Port] :: Connection Timeout"
			Write-Verbose " [Test-Port] :: End"
			return $false
		}
		else
		{
			Write-Verbose " [Test-Port] :: Closing TCP Socket"
			try
			{
				$tcpclient.EndConnect($iar) | out-Null
				$tcpclient.Close()
			}
			catch
			{
				Write-Verbose " [Test-Port] :: Unable to Close TCP Socket"
			}
			$true
		}
	}
	End
	{
		Write-Verbose " [Test-Port] :: End Script"
	}
}
#endregion 

#region Test-PSRemoting

function Test-PSRemoting
{
	Param (
		[alias('dnsHostName')]
		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName
	)
	Process
	{
		Write-Verbose " [Test-PSRemoting] :: Start Process"
		if ($ComputerName -match "(.*)(\$)$")
		{
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		
		try
		{
			
			$result = Invoke-Command -ComputerName $computername { 1 } -ErrorAction SilentlyContinue
			
			if ($result -eq 1)
			{
				return $True
			}
			else
			{
				return $False
			}
		}
		catch
		{
			return $False
		}
	}
}

#endregion

#region Test-RegistryKey 

function Test-RegistryKey
{
	
	    <#
	        .Synopsis 
	            Test for given the registry key.
	            
	        .Description
	            Test for given the registry key.
	                        
	        .Parameter Path 
	            Path to the key.
	            
	        .Parameter ComputerName 
	            Computer to test the registry key on.
	            
	        .Example
	            Test-registrykey HKLM\Software\Adobe
	            Description
	            -----------
	            Returns $True if the Registry key for HKLM\Software\Adobe
	            
	        .Example
	            Test-registrykey HKLM\Software\Adobe -ComputerName MyServer1
	            Description
	            -----------
	            Returns $True if the Registry key for HKLM\Software\Adobe on MyServer1
	                    
	        .OUTPUTS
	            System.Boolean
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            New-RegistryKey
	            Remove-RegistryKey
	            Get-RegistryKey
	        
	        .Notes
	            NAME:      Test-RegistryKey
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding(SupportsShouldProcess = $true)]
	Param (
		
		[Parameter(ValueFromPipelineByPropertyName = $True, mandatory = $true)]
		[string]$Path,
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
		
	)
	
	Begin
	{
		
		Write-Verbose " [Test-RegistryKey] :: Start Begin"
		
		Write-Verbose " [Test-RegistryKey] :: `$Path = $Path"
		Write-Verbose " [Test-RegistryKey] :: Getting `$Hive and `$KeyPath from $Path "
		$PathParts = $Path -split "\\|/", 0, "RegexMatch"
		$Hive = $PathParts[0]
		$KeyPath = $PathParts[1..$PathParts.count] -join "\"
		Write-Verbose " [Test-RegistryKey] :: `$Hive = $Hive"
		Write-Verbose " [Test-RegistryKey] :: `$KeyPath = $KeyPath"
		
		Write-Verbose " [Test-RegistryKey] :: End Begin"
		
	}
	
	Process
	{
		
		Write-Verbose " [Test-RegistryKey] :: Start Process"
		
		Write-Verbose " [Test-RegistryKey] :: `$ComputerName = $ComputerName"
		
		$RegHive = Get-RegistryHive $hive
		
		if ($RegHive -eq 1)
		{
			Write-Host "Invalid Path: $Path, Registry Hive [$hive] is invalid!" -ForegroundColor Red
		}
		else
		{
			Write-Verbose " [Test-RegistryKey] :: `$RegHive = $RegHive"
			
			$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
			Write-Verbose " [Test-RegistryKey] :: `$BaseKey = $BaseKey"
			
			Try
			{
				$Key = $BaseKey.OpenSubKey($KeyPath)
				if ($Key)
				{
					$true
				}
				else
				{
					$false
				}
			}
			catch
			{
				$false
			}
		}
		Write-Verbose " [Test-RegistryKey] :: End Process"
		
	}
}

#endregion 

#region Test-RegistryValue 

function Test-RegistryValue
{
	
	    <#
	        .Synopsis 
	            Test the value for given the registry value.
	            
	        .Description
	            Test the value for given the registry value.
	                        
	        .Parameter Path 
	            Path to the key that contains the value.
	            
	        .Parameter Name 
	            Name of the Value to check.
	            
	        .Parameter Value 
	            Value to check for.
	            
	        .Parameter ComputerName 
	            Computer to test.
	            
	        .Example
	            Test-RegistryValue HKLM\SOFTWARE\Adobe\SwInstall -Name State -Value 0
	            Description
	            -----------
	            Returns $True if the value of State under HKLM\SOFTWARE\Adobe\SwInstall is 0
	            
	        .Example
	            Test-RegistryValue HKLM\Software\Adobe -ComputerName MyServer1
	            Description
	            -----------
	            Returns $True if the value of State under HKLM\SOFTWARE\Adobe\SwInstall is 0 on MyServer1
	                    
	        .OUTPUTS
	            System.Boolean
	            
	        .INPUTS
	            System.String
	            
	        .Link
	            New-RegistryValue
	            Remove-RegistryValue
	            Get-RegistryValue
	        
	        .Notes    
	            NAME:      Test-RegistryValue
	            AUTHOR:    bsonposh
	            Website:   http://www.bsonposh.com
	            Version:   1
	            #Requires -Version 2.0
	    #>
	
	[Cmdletbinding()]
	Param (
		
		[Parameter(mandatory = $true)]
		[string]$Path,
		[Parameter(mandatory = $true)]
		[string]$Name,
		[Parameter()]
		[string]$Value,
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
		
	)
	
	Process
	{
		
		Write-Verbose " [Test-RegistryValue] :: Begin Process"
		Write-Verbose " [Test-RegistryValue] :: Calling Get-RegistryKey -Path $path -ComputerName $ComputerName"
		$Key = Get-RegistryKey -Path $path -ComputerName $ComputerName
		Write-Verbose " [Test-RegistryValue] :: Get-RegistryKey returned $Key"
		if ($Value)
		{
			try
			{
				$CurrentValue = $Key.GetValue($Name)
				$Value -eq $CurrentValue
			}
			catch
			{
				$false
			}
		}
		else
		{
			try
			{
				$CurrentValue = $Key.GetValue($Name)
				if ($CurrentValue) { $True }
				else { $false }
			}
			catch
			{
				$false
			}
		}
		Write-Verbose " [Test-RegistryValue] :: End Process"
		
	}
}

#endregion 

#region Test-TcpPort
function Test-TcpPort ($ComputerName, [int]$port = 80)
{
	$socket = new-object Net.Sockets.TcpClient
	$socket.Connect($ComputerName, $port)
	if ($socket.Connected)
	{
		$status = "Open"
		$socket.Close()
	}
	else
	{
		$status = "Closed / Filtered"
	}
	$socket = $null
	Add-RichTextBox "ComputerName:$ComputerName`nPort:$port`nStatus:$status"
}
#endregion

# Taskmanager Functions
#region initialize-stuff
function initialize-stuff
{
	### Basically just an initialization routine for the hashes and the drawing objects.
	$server = $things["machine"];
	update-Status-Label "Initializing server information...";
	$label3.update();
	$available = load-os-info $server;
	update-Status-Label "Initializing processes...";
	
	$ysize = 0;
	$procs = return-win32_perfrawdata_perfproc_process $server;
	foreach ($proc in $procs)
	{
		if ($proc.IDProcess -eq 0)
		{
			$beforeprocs[0] = $proc.percentprocessortime;
		}
		else
		{
			$beforeprocs.Add($proc.IDProcess, $proc.percentprocessortime);
		}
	}
	
	update-Status-Label "Initializing CPU...";
	$t1 = return-win32_PerfRawData_PerfOS_processor $server;
	$ch = new-object system.drawing.drawing2d.HatchBrush([system.drawing.drawing2d.hatchstyle]::LargeGrid, $things["colors"][1], $things["colors"][2]);
	
	update-Status-Label "Initializing graphics...";
	$orderarray = New-Object -TypeName System.Collections.ArrayList;
	foreach ($cpu in $t1) { $orderarray.add($cpu.name); }
	$y = 100;
	$x = -550;
	$counter = 0;
	for ($j = 0; $j -lt $orderarray.count; $j++)
	{
		$t = "";
		$key = $orderarray[$j];
		foreach ($bob in $t1) { if ($bob.Name -eq $key) { $t = $bob; } }
		if (($counter % 8) -eq 0)
		{
			$y = 100;
			$x += 600;
			$xstrpt = $x - 50;
		}
		$keyhash.Add($key, @($t.percentprocessortime, $t.timestamp_sys100ns));
		$pointhash.Add($key, @());
		$ysize += 105;
		$rect = new-object system.drawing.rectangle(($x + 1), ($y - 100), 500, 99);
		$point = new-object system.drawing.pointf(($x - 50), ($y - 15));
		$point2 = new-object system.drawing.pointf(($x - 50), ($y - 50));
		$parms.Add($key, @($x, $y, 0.0, $rect, $ch, $point, $point2));
		$points = @(new-object system.drawing.point($x, $y));
		$hash.Add($key, @());
		$y = $y + 100;
		$counter++;
	}
	$keyhash.Add("Memory", @(0.0, 0.0));
	$pointhash.Add("Memory", @());
	$hash.Add("Memory", @());
	$xsize = [int32]((($counter/8) + 1) * 575);
	if ($ysize -gt 500) { $ysize = 500; }
	
	#### Memory
	if ((($counter % 8) -eq 0) -OR (($y + 300) -gt 900))
	{
		$y = 100;
		$x += 600;
		$xstrpt = $x - 50;
	}
	else { $ysize += 300; }
	$rect = new-object system.drawing.rectangle(($x + 1), $y, 500, 199);
	$y += 200;
	$ystrpt = $y - 15;
	$point = new-object system.drawing.pointf(($x - 50), ($y - 15));
	$xstrpt = $x - 50;
	$ystrpt = $y - 50;
	$point2 = new-object system.drawing.pointf(($x - 50), ($y - 50));
	$point3 = new-object system.drawing.pointf(($x - 50), ($y - 70));
	$parms.Add("Memory", @($x, $y, 0.0, $rect, $ch, $point, $point2, 0.0, $point3));
	
	update-Status-Label "Updating CPU...";
	
	get-allCPU;
	$parms.Add("Bitmap", @($xsize, $ysize));
	#$picturebox1.AutoScrollMargin = new-object System.Drawing.Size($xsize, $ysize);
	update-Status-Label "";
}
#endregion initialize-stuff

#region update-serverdatetime
function update-serverdatetime
{
	param ($server);
}
#endregion update-serverdatetime

#region load-os-info
function load-os-info
{
	param ($server);
	### This was an afterthought sort of like the services tab. It occurred to me that it might be nice to know some of the details
	### about the machine and the OS running on it. All this does is create a bunch of labels on $Tab4 and fill in the information.
	### The only one that's different is the label that shows the localdatetime. That is created in BuildTheForm() so it can be
	### updated every time we refresh the processes.
	while ($Tab4.Controls.count -gt 1) { foreach ($item in $Tab4.Controls) { if ($item.name -ne "ServerTime_Label") { $item.dispose(); } } }
	$tm = 0;
	$cs = CIM-Stuff win32_computersystem;
	$prcsrs = CIM-Stuff win32_processor;
	$srv = CIM-Stuff Win32_OperatingSystem;
	$mem = CIM-Stuff CIM_PhysicalMemory;
	$mem | % { $tm += $_.capacity; }
	$caption = $srv.caption;
	$y = 20;
	$x = 20;
	
	
	$arch = "32-bit";
	if (($srv.OSArchitecture).length -gt 0) { $arch = $srv.OSArchitecture; }
	elseif ($srv.caption -match "x64") { $arch = "64-bit"; }
	foreach ($m in $mem)
	{
		$l = $m.tag + " (" + $m.devicelocator.trimend() + ") | " + (dsize $m.capacity);
		if ($m.speed -gt $null) { $l += " | Speed = " + $m.speed.tostring() + " ns"; }
		$st = "OK";
		if ($m.status -gt $null) { $st = $m.status; }
		$l += " | Status = " + $st;
	}
	
	foreach ($p in $prcsrs)
	{
		$noc = "";
		if ($p.numberofcores -ne $null) { $noc = " -- " + $p.numberofcores + " Cores"; }
	}
	
	
	### If this is an older OS, we need to use Win32_LogicalDisk to get the list of disks. If it's running a later OS, we can use
	### Win32_Volume which will also list the mount points (if any).
	$vlen = 20;
	if ($caption -match "2000")
	{
		$hds = CIM-Stuff Win32_LogicalDisk "DriveType=3" |
		select-object -property @{ expression = { $_.deviceid }; name = "VolumeName" }, @{ expression = { $_.size }; name = "Capacity" },
					  FreeSpace, @{ expression = { $_.volumename }; name = "Label" };
	}
	else
	{
		$hds = CIM-Stuff Win32_Volume "DriveType=3" |
		select-object -property @{ expression = { $_.name }; name = "VolumeName" }, Capacity, FreeSpace, Label;
	}
	
	$hds = $hds | sort-object -property VolumeName;
	foreach ($d in $hds) { if ($d.VolumeName.length -gt $vlen) { $vlen = $d.VolumeName.length; } }
	
	$x = ($vlen * 9);
	$cs.totalphysicalmemory;
}
#endregion load-os-info

#region display-myMessageBox
function display-myMessageBox
{
	param ($msg);
	$myMessageBox = new-object System.Windows.Forms.Form;
	$myMB_TextBox = new-object System.Windows.Forms.TextBox;
	$myMessageBox.cancelbutton = $Cancel_Button;
	$myMB_TextBox.Anchor = "Left, Top, Right, Bottom";
	$myMB_TextBox.Location = new-object system.drawing.point(0, 0);
	$myMB_TextBox.font = $fonts["cn8"];
	$myMB_TextBox.Name = "myMB_TextBox";
	$myMB_TextBox.multiline = $true;
	$myMB_TextBox.Text = "";
	$myMessageBox.Controls.Add($myMB_TextBox);
	$array = $msg.split("`n");
	$count = $array.count;
	$w = 0;
	foreach ($line in $array) { if ($line.length -gt $w) { $w = $line.length; } }
	$height = $count * 14.0;
	$width = $w * 8.25;
	$myMessageBox.ClientSize = new-object System.Drawing.Size($width, $height);
	$myMB_TextBox.ClientSize = new-object System.Drawing.Size($width, $height);
	$myMB_TextBox.text = $msg;
	$IFWS = new-object System.Windows.Forms.FormWindowState;
	$IFWS = $myMessageBox.WindowState;
	$myMessageBox.TopMost = $true;
	$myMessageBox.Refresh();
	$myMessageBox.BringToFront();
	$myMessageBox.add_Load($OnLoadForm_StateCorrection);
	$myMessageBox.Show() | Out-Null;
}
#endregion display-myMessageBox

#region get-services
function get-services
{
	### Simply loads the services on the Services tab ($Tab3).
	$server = $things["machine"];
	$listview2.Items.Clear();
	foreach ($s in (return-win32_Service $server))
	{
		$lvi = new-object system.windows.forms.ListViewItem($s.displayname);
		if ($s.description -eq $null) { $s.description = ""; }
		foreach ($c in (1..($listview2.columns.count - 1)))
		{
			$name = $listview2.columns[$c].name;
			$lvi.subitems.add($s.$name);
		}
		
		$listview2.Items.Add($lvi);
	}
	
	$error.clear;
}
#endregion get-services

#region get-processes
function get-processes
{
	param ($update);
	### Loads and updates the $listview1 listview on $Tab1.
	$server = $things["machine"];
	$procs = return-win32_perfrawdata_perfproc_process $server;
	$idle = $cpu = $totalcpu = $totalcpuUsed = 0;
	$procs | % { if ($_.name -eq "_Total") { $totalcpu = [long]$_.percentprocessortime - [long]$beforeprocs[$_.IDProcess]; } };
	if ($update -eq $false)
	{
		$listview1.items.clear();
		$users = @{ };
		foreach ($proc in (CIM-Stuff win32_process))
		{
			$users.add($proc.ProcessID, ($proc |
			Invoke-CimMethod -CimSession $things["session"] -MethodName GetOwner).user);
		}
	}
	
	drop-dead-procs $procs;
	foreach ($proc in $procs)
	{
		$idproc = $proc.IDProcess;
		if ($proc.Name -eq "Idle")
		{
			$idle = kbytes $proc.WorkingSet;
			if ($update -eq $false) { make-listviewitem $proc $null; }
		}
		elseif ($proc.Name -ne "_Total")
		{
			if ($update -eq $false)
			{
				$cpu = pcnt-cpu $proc.percentprocessortime $beforeprocs[$idproc] $totalcpu;
				$totalcpuUsed += $cpu;
				make-listviewitem $proc $users[$idproc];
			}
			elseif ($beforeprocs[$idproc] -eq $null)
			{
				### If this is a new process, create a ListViewItem for it.
				$beforeprocs.Add($idproc, $proc.percentprocessortime);
				$cpu = 0;
				make-listviewitem $proc (CIM-Method win32_process ("ProcessID='" + $idproc + "'") GetOwner).user;
			}
			else
			{
				### Otherwise, just calculate the CPU for it.
				$cpu = pcnt-cpu $proc.percentprocessortime $beforeprocs[$idproc] $totalcpu;
				$totalcpuUsed += $cpu;
			}
			
			### Update the memory and CPU for the process in its ListView entry
			$lvi = $listview1.FindItemWithText($idproc)
			if ($lvi.SubItems[3].Text -ne [int32]($cpu)) { $lvi.SubItems[3].Text = [int32]($cpu); }
			if ($lvi.SubItems[4].Text -ne (kbytes $proc.WorkingSet)) { $lvi.SubItems[4].Text = kbytes $proc.WorkingSet; }
		}
		
		$beforeprocs[$idproc] = $proc.percentprocessortime;
	}
	
	$indx = $listview1.FindItemWithText("Idle").index;
	$cpu = [int32](100 - $totalcpuUsed);
	$listview1.Items[$indx].SubItems[3].Text = [int32]($cpu);
	$listview1.Items[$indx].SubItems[4].Text = $idle;
	$listview1.refresh();
	$things["procs"] = ($procs.count - 1);
	$things["cpu"] = [int32]$totalcpuUsed;
	update-Procs-Label;
	update-serverdatetime $server;
	if ($things["LVCols"] -ne $null)
	{
		if (($things["LVCols"] -eq 0) -OR ($things["LVCols"] -eq 2))
		{
			$listview1.ListViewItemSorter = new-object ListViewItemComparer($things["LVCols"], $listview1.Sorting);
		}
		else
		{
			$listview1.ListViewItemSorter = new-object ListViewItemIntComparer($things["LVCols"], $listview1.Sorting);
		}
	}
	
}
#endregion get-processes

#region make-listviewitem
function make-listviewitem
{
	param ($proc,
		$user);
	$idproc = $proc.IDProcess;
	$lvi = new-object system.windows.forms.ListViewItem($proc.Name);
	$lvi.SubItems.Add($idproc);
	if ($user -eq $null) { $user = "SYSTEM"; }
	$lvi.SubItems.Add($user);
	$lvi.SubItems.Add(0);
	$mem = kbytes $proc.WorkingSet;
	$lvi.SubItems.Add($mem);
	$listview1.Items.Add($lvi);
}
#endregion make-listviewitem

#region drop-dead-procs
function drop-dead-procs
{
	param ($procs);
	$temp = @{ };
	$currp = @();
	foreach ($proc in $procs) { $currp += $proc.IDProcess; }
	foreach ($idproc in $beforeprocs.Keys) { if ($currp -notcontains $idproc) { $temp.Add($idproc, 0); } }
	foreach ($procid in $temp.Keys)
	{
		$indx = $listview1.FindItemWithText($procid).index;
		$listview1.Items[$indx].Remove();
		$beforeprocs.Remove($procid);
	}
	
	$temp.clear()
}
#endregion drop-dead-procs

#region set-context-menu
function set-context-menu
{
	foreach ($si in $listview2.SelectedItems)
	{
		if ($si.subitems[($listview2.columns["state"].index)].text -eq "Running")
		{
			$start_svc.enabled = $false;
			$stop_svc.enabled = $true;
			$cycle_svc.enabled = $true;
		}
		elseif ($si.subitems[($listview2.columns["state"].index)].text -eq "Stopped")
		{
			$start_svc.enabled = $true;
			$stop_svc.enabled = $false;
			$cycle_svc.enabled = $false;
		}
	}
}
#endregion set-context-menu

#region stop-related-services
function stop-related-services
{
	param ($sname,
		$dependencies);
	if ($dependencies.count -gt 0)
	{
		foreach ($d in $dependencies) { stop-start-service $d.name "Stop"; }
	}
	
	stop-start-service $sname "Stop";
}
#endregion stop-related-services

#region start-related-services
function start-related-services
{
	param ($sname,
		$dependencies);
	stop-start-service $sname "Start";
	if ($dependencies.count -gt 0)
	{
		foreach ($d in $dependencies) { stop-start-service $d.name "Start"; }
	}
}
#endregion start-related-services

#region stop-start-service
function stop-start-service
{
	param ($sname,
		$whattodo);
	$vars = @{
		"Stop" = @("Stopped", "Stopping", "1", { $svc.Stop() }, { $svc.StopService() });
		"Start" = @("Running", "Starting", "4", { $svc.Start() }, { $svc.StartService() });
	};
	
	$svc = CIM-Stuff win32_service ("name = '" + $sname + "'");
	if (($svc.state -eq $vars[$whattodo][0]) -OR ($svc.state -eq $null)) { return; }
	if ($svc.StartMode -eq "Disabled")
	{
		[system.windows.forms.messagebox]::Show("Cannot start or stop a disabled service, and I'm not enabling it just for you.");
		return;
	}
	
	if ($things["adsi"])
	{
		$cmd = "[ADSI](""WinNT://" + $things["machine"] + "/" + $sname + ",service"")";
		$svc = invoke-expression $cmd;
		if ($svc.status -ne $vars[$whattodo][2]) { &$vars[$whattodo][3]; }
	}
	else
	{
		$svc = CIM-Stuff win32_Service "Name='$sname'";
		if ($svc.state -ne $vars[$whattodo][0])
		{
			$r = &$vars[$whattodo][4];
			if ($r.returnvalue -ne 0)
			{
				[system.windows.forms.messagebox]::Show("Unable to $whattodo the $sname service.");
				return;
			}
		}
	}
	
	if ((GetStatus $sname $vars[$whattodo][0]) -eq 1)
	{
		### Update the status of the service
		($listview2.items[($listview2.FindItemWithText($sname).index)]).subitems[($listview2.columns["state"].index)].text = $vars[$whattodo][0];
		set-context-menu;
	}
}
#endregion stop-start-service

#region reset-iis
function reset-iis
{
	$server = $things["machine"];
	$test = iisreset $server;
	$outcome = "Failed";
	if ($test -match "successfully restarted") { $outcome = "Succeeded"; }
}
#endregion reset-iis

#region pcnt-cpu
### CPU percentage calculation. I picked this up from an article on SQL server long ago. It seems to be the
### same one used for the OS.
function pcnt-cpu
{
	(([long]$args[0] - [long]$args[1]) / [system.double]$args[2]) * 100;
}
#endregion pcnt-cpu

#region kbytes
function kbytes
{
	param ($dsize);
	[Math]::round($dsize / 1kb, 2);
}
#endregion  

#region dsize
function dsize
{
	param ($dsize);
	$size = "";
	if ($dsize -ge 1gb) { $size = [Math]::round($dsize / 1gb, 2).tostring() + " GB"; }
	elseif ($dsize -ge 1mb) { $size = [Math]::round($dsize / 1mb, 2).tostring() + " MB"; }
	elseif ($dsize -ge 1kb) { $size = [Math]::round($dsize / 1kb, 2).tostring() + " KB"; }
	else { $size = $dsize.tostring() + " B"; }
	$size;
}
#endregion

#region return-Win32_PerfFormattedDAte_PerfProc_Process
### These two functions get their data depending on how new the OS is. In the most recent version of Task Manager,
### it uses WorkingSetPrivate (that's the default) for the memory, but that isn't a property on older versions.
### Whether it's available or not is determined during initialization.
function return-Win32_PerfFormattedData_PerfProc_Process
{
	param ($server);
	if ($things["wsp"])
	{
		(CIM-Stuff Win32_PerfFormattedData_PerfProc_Process) |
		select-object -property idprocess, name, @{ expression = { $_.workingsetprivate }; name = "workingset" }, percentprocessortime;
	}
	else
	{
		(CIM-Stuff Win32_PerfFormattedData_PerfProc_Process) | select-object -property idprocess, name, workingset, percentprocessortime;
	}
}
#endregion

#region return-win32_perfrawdata_perfproc_process
function return-win32_perfrawdata_perfproc_process
{
	param ($server);
	if ($things["wsp"])
	{
		(CIM-Stuff win32_perfrawdata_perfproc_process) |
		select-object -property idprocess, name, @{ expression = { $_.workingsetprivate }; name = "workingset" }, percentprocessortime;
	}
	else
	{
		(CIM-Stuff win32_perfrawdata_perfproc_process) | select-object -property idprocess, name, workingset, percentprocessortime;
	}
}
#endregion

#region return-win32_service
function return-win32_Service
{
	CIM-Stuff win32_Service | select-object -property name, displayname, processid, description, state, startmode, startname | Sort-Object -property displayname;
}

#endregion

#region return-win32_PerfRawData_PerfOS_processor
function return-win32_PerfRawData_PerfOS_processor
{
	CIM-Stuff win32_PerfRawData_PerfOS_processor | select-object -property name, percentprocessortime, timestamp_sys100ns;
}
#endregion

#region CIM-Stuff
function CIM-Stuff
{
	param ($class,
		$filter);
	if ($filter -eq $null)
	{
		Get-CimInstance -class $class -CimSession $things["session"];
	}
	else
	{
		Get-CimInstance -class $class -filter $filter -CimSession $things["session"];
	}
}
#endregion

#region CIM-Method
function CIM-Method
{
	param ($class,
		$filter,
		$method);
	CIM-Stuff $class $filter | Invoke-CimMethod -CimSession $things["session"] -MethodName $method;
}
#endregion

#region get-allCPU
function get-allCPU
{
	### This calculates the CPU for the individual processors and adds them into $hash.
	$server = $things["machine"];
	$p2 = return-win32_PerfRawData_PerfOS_processor $server;
	for ($i = 0; $i -lt $p2.length; $i++)
	{
		$key = $p2[$i].Name;
		$cpu = 100.0 - (pcnt-cpu $p2[$i].percentprocessortime $keyhash[$key][0] ([system.double]$p2[$i].timestamp_sys100ns - [system.double]$keyhash[$key][1]));
		$count = $hash[$key].count;
		if ($cpu -lt 0.0) { $cpu = 0.0; }
		$x = ($count * 5) + $parms[$key][0];
		$y = $parms[$key][1] - $cpu;
		$parms[$key][2] = $cpu;
		$point = new-object system.drawing.point($x, $y);
		$pointhash[$key] += $y;
		$hash[$key] += $point;
		$keyhash[$key] = @([system.double]$p2[$i].percentprocessortime, [system.double]$p2[$i].timestamp_sys100ns);
	}
	
	##### Memory
	$tpm = (CIM-Stuff win32_computersystem).totalphysicalmemory;
	$avb = (CIM-Stuff Win32_PerfRawData_PerfOS_Memory).availablebytes;
	$newy = (1 - ([system.double]$avb / [system.double]$tpm)) * 200;
	$count = $hash["Memory"].count;
	$x = ($count * 5) + $parms["Memory"][0];
	$y = $parms["Memory"][1] - $newy;
	$parms["Memory"][2] = $newy / 2;
	$parms["Memory"][7] = [system.double]$tpm - [system.double]$avb;
	$point = new-object system.drawing.point($x, $y);
	$pointhash["Memory"] += $y;
	$hash["Memory"] += $point;
}
#endregion

#region Plot
function Plot
{
	### Draw the pretty pictures of CPU and Memory usage
	param ($old_btmp);
	if ($old_btmp -ne $null) { $old_btmp.Dispose(); }
	$btmp = new-object system.drawing.bitmap($parms["Bitmap"][0], $parms["Bitmap"][1]);
	$grfx = [system.drawing.graphics]::fromimage($btmp);
	
	### Coordinates drawing the graphs for the CPU and memory.
	$orderarray = build-order;
	build-axes $grfx;
	for ($i = 0; $i -lt $orderarray.count; $i++)
	{
		$key = $orderarray[$i];
		$ptarray = $hash[$key];
		for ($j = 1; $j -lt $ptarray.count; $j++)
		{
			$grfx.DrawLine($things["plotpens"][$j - 1], $ptarray[$j - 1], $ptarray[$j]);
		}
	}
	
	$Picturebox1.image = $btmp;
	$grfx.Dispose();
	$pointhash = shift-arrays $pointhash;
	reload-points;
	$error.clear();
	$btmp;
}
#endregion

#region reload-points
function reload-points
{
	### This may seem like a lot of trouble for nothing, but if you go ahead and turn everything into drawing points and
	### store them in an array to pass into DrawLines, it works much more smoothly than passing the coordinates into
	### DrawLine one at a time and having it do the conversion. Take my word for it, watching it draw a bunch of line
	### segments one at a time is entertaining as hell, but this gives better performance.
	$pts = @{ };
	foreach ($key in $keyhash.Keys) { $points = @(new-object system.drawing.point($parms[$key][0], $pointhash[$key][0])); $pts.Add($key, $points); }
	for ($i = 1; $i -lt $pointhash["_Total"].count; $i++)
	{
		$x = ($i * 5);
		foreach ($key in $keyhash.Keys)
		{
			$pts[$key] += new-object system.drawing.point(($x + $parms[$key][0]), $pointhash[$key][$i]);
		}
	}
	
	foreach ($key in $keyhash.Keys) { $hash[$key] = $pts[$key]; }
	$pts = $null;
}
#endregion

#region shift-arrays
### We only maintain 100 sets of data for each CPU and the memory. When the array gets to 100, we
### pop off the top one and the new one gets added to the end.
function shift-arrays ($myhash)
{
	if ($myhash["_Total"].length -ge 100)
	{
		$null, $things["plotpens"] = $things["plotpens"];
		foreach ($key in $keyhash.Keys)
		{
			$null, $myhash[$key] = $myhash[$key];
		}
	}
	
	$myhash;
}
#endregion

#region build-order
### Probably not needed, but I want to make sure that the CPUs are ordered numerically. It just makes things neater.
function build-order
{
	$count = $keyhash.count;
	$count = $count - 1;
	$orderarray = @(0..$count);
	for ($i = 0; $i -lt $count - 1; $i++) { $orderarray[$i] = [system.string]$i; }
	$orderarray[$count - 1] = "_Total";
	$orderarray[$count] = "Memory";
	$orderarray;
}
#endregion

#region build-axes
function build-axes
{
	param ($grfx);
	### The various drawing surfaces for each processor and memory are stored in the $parms hash. That
	### way we don't have to keep recalculating them for each refresh.
	$orderarray = build-order;
	$mypen = $pens["white"];
	$mypen.Width = 2;
	$font = $fonts["verdana8"];
	$brush = $brushes["red"];
	$y = $add = 100;
	$x = -550;
	$xstrpt = $x - 50;
	for ($i = 0; $i -lt $orderarray.count; $i++)
	{
		$key = $orderarray[$i];
		$x = $parms[$key][0];
		$y = $parms[$key][1];
		if ($key -eq "Memory") { $add = 200; }
		$grfx.FillRectangle($parms[$key][4], $parms[$key][3]);
		$grfx.Drawline($mypen, $x, $y, $x, $y - $add);
		$grfx.Drawline($mypen, $x, $y, $x + 500, $y);
		$grfx.DrawString($key, $font, $brush, $parms[$key][5]);
		$pct = "{0:#.##}%" -f $parms[$key][2];
		$grfx.DrawString($pct, $font, $brush, $parms[$key][6]);
	}
	
	$newgb = "{0:#.##}GB" -f ($parms["Memory"][7] / 1gb);
	$grfx.DrawString($newgb, $font, $brush, $parms["Memory"][8]);
	
}
#endregion

#region update-procs-label
function update-Procs-Label
{
	$label2.Text = "Updating every " + $things["timer"].interval.ToString() + " ms -- Processes: " +
	($things["procs"]).ToString() + "  |  CPU Usage: " + ($things["cpu"]).ToString() + "%";
}
#endregion

#region update-status-label
function update-Status-Label
{
	$label3.Text = $args[0];
	$label3.update();
}
#endregion

#region restart-timer
### Called when the "Pause"/"Restart" button is pushed.
function Restart-timer
{
	$button2.Text = "Pause";
	$button2.add_click({ Stop-timer; });
	$things["timer"].Enabled = $true;
	$things["timer"].Start();
}
#endregion

#region stop-timer
function Stop-timer
{
	$things["timer"].Enabled = $false;
	$things["timer"].Stop();
	$button2.Text = "Restart";
	$button2.add_click({ Restart-timer; });
}
#endregion

#region Stop-stuff
### Try to shut down in an orderly fashion. Called when the "Quit" button is pressed.
function Stop-stuff
{
	Stop-timer;
	$things["timer"].Dispose();
	if ($things["session"] -ne $null) { remove-cimsession -cimsession $things["session"] }
	$things["session"].Close();
	$things["session"].Dispose();
}
#endregion

#region getStatus
function GetStatus
{
	param ($service,
		$check_status)
	### This is supposed to sit and wait until a service has been stopped or started. It tests the
	### service status until it matches what we want it to be. If it hasn't done what we requested
	### after 30 seconds, we flag an error and go on.
	$server = $things["machine"];
	$counter = 0;
	$results = 1;
	$test_status = (CIM-Stuff win32_service ("name='" + $service + "'")).State;
	### The thinking behind this is that if it can't kill the service withing 30 seconds, it isn't going to die. So
	### we won't leave ourself hanging out in here. We'll just flag an error and go on with this tedium we call life.
	while (($check_status -ne $test_status) -AND ($counter -lt 60))
	{
		start-sleep -m 500;
		$test_status = (CIM-Stuff win32_service ("name='" + $service + "'")).State;
		$counter++;
		waitingtodie $counter "-";
		if ($counter -eq 60) { $results = 0; };
	}
	
	$results;
}
#endregion

#region initialize-the-hashes
function initialize-the-hashes
{
	$things["timer"].Dispose();
	if ($things["session"] -ne $null)
	{
		remove-cimsession -cimsession $things["session"];
		$things["session"].Close();
		$things["session"].Dispose();
	}
	
	$keyhash.clear();
	$hash.clear();
	$pointhash.clear();
	$parms.clear();
	$beforeprocs.clear();
	$things.clear();
	initialize-things;
}
#endregion

#region initialize-things
function initialize-things
{
	$things.add("wsp", $false);
	$things.add("adsi", $true);
	$things.add("LVCols", $null);
	$things.add("SDGCols", $null);
	$things.add("colors", (.{$args} red darkgreen black lightgreen white blue));
	$things.add("timer", (new-object System.Windows.Forms.timer));
	$things.add("machine", $textbox1.text);
	$things.add("plotpens", @((new-object system.drawing.pen("lightgreen")), (new-object system.drawing.pen("lightgreen"))));
	$things.add("procs", 0);
	$things.add("cpu", 0);
	$things.add("session", $null);
	$things.add("response", $null);
	if ($things["machine"] -gt "")
	{
		$things["session"] = New-CimSession -ComputerName $things["machine"] -SessionOption (New-CimSessionOption -Protocol Dcom);
	}
}
#endregion

#region test-interval
function test-interval
{
	param ($span)
	if ($span -gt $things["timer"].interval) { $things["plotpens"] += $pens["red"]; }
	else { $things["plotpens"] += $pens["lightgreen"] };
}
#endregion

#region do-it-to-it
function do-it-to-it
{
	### Test the status of the telephony service on the remote machine. We don't really care what the status
	### is, we just want to know that we can get to it. Through playing around with some of this stuff, I've
	### found that this is a way to test that a server is available without generating a bunch of errors.
	$server = $textbox1.text;
	#   if ((new-object system.serviceprocess.servicecontroller("telephony", $server)).status -eq $null) {
	#     $rtrn = [system.windows.forms.messagebox]::Show("Cannot find the $server machine. Make sure it exists and you have permissions to it.");
	#    $main.Cursor = [System.Windows.Forms.Cursors]::Default;
	#   return;
	#}
	
	initialize-the-hashes;
	
	try
	{
		$props = CIM-Stuff win32_perfrawdata_perfproc_process;
		if ($props[0].__property_count -gt 36) { $things["wsp"] = $true; }
	}
	catch
	{
		$rtrn = [system.windows.forms.messagebox]::Show("It looks as though the $server machine does not allow remote WMI calls. We won't be able to monitor it.");
		return;
	}
	
	### Start initializing things.
	initialize-stuff;
	get-services;
	get-allCPU;
	$btmp = new-object system.drawing.bitmap($parms["Bitmap"][0], $parms["Bitmap"][1]);
	get-processes $false;
	$btmp = Plot $btmp;
	
	### The $handler is a list of what to do when the timer fires off.
	### Notice that timer is stopped until the remote machine responds and is then restarted. This prevents
	### putting more stress on a remote box that's already overloaded.
	$handler = {
		$things["timer"].Stop();
		$tmthen = get-date;
		get-allCPU;
		get-processes $true;
		$tmspn = [math]::round((new-timespan -start $tmthen).totalmilliseconds, 0);
		test-interval $tmspn;
		$btmp = Plot $btmp;
		$things["timer"].Start();
	}
	
	
	### This is how I've implemented the update interval. The regular TaskManager refreshes about once a
	### second, but this can run into problems if you try that, particularly if you're going for a machine that
	### is pretty busy. If things are pegged on a machine, it may not have the resources to get back to you every
	### second, so after we've initialized everything, we take a reading of how long it takes to run through the
	### three things it's going to have to do each time the timer fires off. It may take a long time to
	### get through with the initial stuff, but seems to work pretty well after that. To be honest, I haven't run
	### into many cases where the machine resources have slowed things down. It seems mostly limited by
	### the network. You can probably hardcode this to run once a second and it will do great 99 times out of
	### a 100, but that one time when you really need it will be the one that messes up.
	$et = [System.Diagnostics.Stopwatch]::StartNew();
	invoke-command -scriptblock $handler;
	$et.Stop();
	$intrvl = [Math]::round($et.Elapsed.TotalMilliseconds, 0);
	if ($intrvl -lt 1000) { $intrvl = 1000; }
	$things["timer"].interval = $intrvl;
	$trackbar1.value = $intrvl;
	$things["timer"].add_tick($handler);
	$things["timer"].Start();
}
#endregion