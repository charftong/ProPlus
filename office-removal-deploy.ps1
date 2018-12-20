<#
  .SYNOPSIS
  Uninstalls Microsoft Office ProPlus as well as stand-alone products Visio and Project (supports 2007, 2010 and 2013 as well as both 32-bit or 64-bit installs).  
  (Optional) Can also Install Office Click-to-run.  

  .DESCRIPTION
  Uninstalls the Microsoft Office Products ProPlus as well as Visio and Project from a specified local source or share and using
  configuration XML files to configure the uninstallation process.  With the InstallOffice2016 switch, you can proceed to install Office 2016 click-to-run after the uninstall process completes.   
        
  .PARAMETER Products
  The Microsoft Office Products to uninstall. This must match the installation files for the product referred to in the SourcePath parameter.
	List of known values for Products
	Office12.ENTERPRISE - Office 2007 Enterprise
	Office12.PROPLUS - Office 2007 ProPlus
	Office14.PROPLUS - Office 2010 ProPlus
	Office14.PROPLUSR - Office 2010 ProPlus Trial
	Office14.STANDARD - Office Standard 2010
	Office15.PROPLUS - Office 2013 ProPlus
	Office15.STANDARD - Office Standard 2013
	Office16.PROPLUS - Office 2016 ProPlus	
	Office12.Visio - Visio Premium 2007
	Office14.Visio - Visio Premium 2010
	Office12.VisPro - Visio Professional 2007
	Office14.VisPro - Visio Professional 2010
	Office15.VisPro - Visio Professional 2013
	Office15.VisProR - Visio Professional 2013 Trial
	Office16.VisPro - Visio Professional 2016
	Office12.PrjPro - Project Professional 2007
	Office14.PrjPro - Project Professional 2010
	Office15.PrjPro - Project Professional 2013
	Office15.PrjProR - Project Professional 2013 Trial
	Office16.PrjPro - Project Professional 2016
  Default value is set to try and remove all known legacy products. 

  .PARAMETER UninstallSourcePath (optional)
  The location of the installation source files for the products to be removed. Can be a local or network path.   If not specified, the script will try to use the default installation location.  

  .PARAMETER WorkingDir (optional)
  The location of the config XML files that the script will create to control the product uninstallation as well as the script log files if a LoggingDir is not specified
  Defaults to %temp%

  .PARAMETER LoggingDir (optional)

  
  .PARAMETER InstallOffice (optional)
  If you pass this switch, the script will proceed to install Office Click-to-Run after the removal of the older versions of Office.  
  The default value is no

  .PARAMETER InstallSourcePath (optional)
  The location of the installation source files for Click-To-Run. Can be a local or network path.
  The parameter needs to be populated when using the "InstallOffice2016" switch. 

  .PARAMETER InstallConfigFile (optional)
  The file name of the configuration file needed for click-to-run install.   If not specified, it will default to 'configuration.xml'

  .PARAMETER SkipRemoval (optional)
  If you pass this switch, the script will skip over the Office Removal part of the script.   
  The default value is no

  .Requirements
   The account running this script will need the following permissions.
   local admin to the system that needs office uninstalled/installed
   At least NTFS 'read' access to the "UninstallSourcePath" location (if used)
   At least NTFS 'read' access to the "InstallSourcePath" location (if used)
   At least NTFS 'read/write' access to the "LoggingDir" location (if used)
   At least NTFS 'read/write' access to the "WorkingDir" location (if used)

  .TODO
	-Create a 'report' option to simply detect what the script will do prior to running it.
	-Check if uninstall XML Doc already exists in WorkingDir (if WorkingDir is not the default location)
	-add verbose log to uninstall xml file <Logging Type="Verbose" Path="%Temp% Template="Office_2016_InstallLog.txt" />
	-implement reboot after uninstall but before install

  .EXAMPLEs
   office-removal-deploy -Products 'Office15.PROPLUS' -WorkingDir 'c:\temp' -InstallOffice -ClickToRunSourcePath '\\server\share\office'
   office-removal-deploy -LoggingDir '\\server\share\office\logs' -InstallOffice -ClickToRunSourcePath '\\server\share\office' -SkipRemoval

  .NOTES
	Version:       6.0
	Author:        Joe Engelman
	Creation Date: Aug 3rd, 2016
	Purpose:       Remove Legacy Office installs(msi) and install Office Click-to-Run
	
#>

[CmdLetBinding(SupportsShouldProcess=$true)]

param( 
    [Alias("SourcePath")]
	[String]
    $UninstallSourcePath,

	[String]	
    [ValidateScript({ ($_ -ne '') -and ($_ -split '\.').Count -gt 1 })]
    $Products='Office12.ENTERPRISE; Office12.PROPLUS; Office14.PROPLUS; Office14.PROPLUSR; Office14.STANDARD; Office15.PROPLUS; Office15.STANDARD; Office16.PROPLUS; Office12.Visio; Office14.Visio; Office12.VisPro; Office14.VisPro; Office15.VisPro; Office15.VisProR; Office16.VisPro; Office12.PrjPro; Office14.PrjPro; Office15.PrjPro; Office15.PrjProR; Office16.PrjPro',

	[String]
    [ValidateScript({ ($_ -ne '') -and (Test-Path $_) })]
    $WorkingDir,

	[String]
    [ValidateScript({ ($_ -ne '') -and (Test-Path $_) })]
    $LoggingDir,

	[Alias("InstallOffice2016")]
	[Switch]
	$InstallOffice,

	[Alias("ClickToRunSourcePath")]
	[String]
    [ValidateScript({ ($_ -ne '') -and (Test-Path $_) })]
    $InstallSourcePath,

	[String]
	$InstallConfigFile,

	[Switch]
	$SkipRemoval
)

#variables
$timeStamp = (Get-Date -Format yyyyMMdd_HHMMss)
$scriptName = $myInvocation.MyCommand.Name
$machinename = $env:COMPUTERNAME
#if working dir is empty set to local temp
$WorkingDir = if($WorkingDir) { $WorkingDir.TrimEnd("\") } else { $env:TEMP }
#if Logging dir is empty set to working dir
$logFileLocation = if($LoggingDir) { $LoggingDir.TrimEnd("\") } else { $WorkingDir }
$logFile = "$($logFileLocation)\$($machinename)-$($scriptName)-$($timeStamp).log"
#office click-to-run installation configuraiton file
$installConfig = if($InstallConfigFile) { $InstallConfigFile } else { "configuration.xml" }
#variable to check if products were uninstalled
$uninstallerrors = @()
#variables to keep track of Step
$stepfile

Function LogWrite
{
	Param
	(
		[string]$LogString,
		[switch]$display,
		[string]$color=(get-host).ui.rawui.ForegroundColor
	)
	$LogEntry = (Get-Date -Format o) + ":  " + $LogString
	Add-content $LogFile -value $LogEntry
	if($display) { Write-Host $LogString }	
}

Function WriteXmlConfig
{
	Param ([string]$productcode)
	
	<# This Funciton creates an XML Document used for the uninstall of an Office Product
	Example of Output
	<Configuration Product=$ProductID>
	<Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
	<Setting Id="SETUP_REBOOT" Value="never" />
	</Configuration>
	#>

	#create new xml doc
	LogWrite "Creating new XML Document....."
	$filepath = "$workingdir\uninstall-$productcode.xml"
	$XmlWriter = New-Object System.XMl.XmlTextWriter($filePath,$Null)
		
	#set formatting
	$xmlWriter.Formatting = "Indented"
	$xmlWriter.Indentation = "4"

	#start writing xml
	$XmlWriter.WriteStartDocument()
	$XmlWriter.WriteStartElement("Configuration")
	$XmlWriter.WriteAttributeString("Product",$ProductID)
	$XmlWriter.WriteStartElement("Display")
	$XmlWriter.WriteAttributeString("Level","none")
	$XmlWriter.WriteAttributeString("CompletionNotice","no")
	$XmlWriter.WriteAttributeString("SuppressModal","yes")
	$XmlWriter.WriteAttributeString("AcceptEula","yes")
	$XmlWriter.WriteEndElement() #end Display
	$XmlWriter.WriteStartElement("Setting")
	$XmlWriter.WriteAttributeString("Id","SETUP_REBOOT")
	$XmlWriter.WriteAttributeString("Value","never")
	$XmlWriter.WriteEndElement() #end Setting
	$XmlWriter.WriteStartElement("Logging")
	$XmlWriter.WriteAttributeString("Type","Verbose")
	$XmlWriter.WriteAttributeString("Path",$workingdir)
	$XmlWriter.WriteEndElement() #end Logging
	$XmlWriter.WriteEndElement() #end Configuration
	$XmlWriter.WriteEndDocument()
	$xmlWriter.Finalize
	$xmlWriter.Flush()
	$xmlWriter.Close()
	LogWrite "Successfully created new XML Document at $filepath"

}

LogWrite "********* script start *********"

if($SkipRemoval)
{
	LogWrite "The Skip Removal switch was specified.   Skipping over the removal of older versions of Office"
}
else
{
	LogWrite "Checking for installtions of the specified versions of Office"	
	$productslist = $products -split "; "
	foreach($product in $productslist)
	{
		#split product into product version and productID
		$productversion,$productID = $product.split('.',2)
		
		if($product -like "Office12*") 
		{ 
			#if product contains Office12 set product lookup value to just the product name($productID) as the version does not appear in the registry
			$productlookup = $productID
		}
		else
		{ 
			#else set product lookup value to product (this is because Office 2007(office12) does not have version in the uninstall part of the registry)
			$productlookup = $product
		}
		
		if ( $env:PROCESSOR_ARCHITECTURE -eq 'AMD64' )
		{ 
			#os is 64-bit
			$system = "64bit" 			
			#check if product is installed (32-bit or 64-bit)
			$productinstalled32 = Test-Path -Path "HKLM:\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall\$productlookup"
			$productinstalled64 = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$productlookup"
		}
		else
		{ 
			#os is 32-bit
			$system = "32bit"
			#check if product is installed
			$productinstalled = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$productlookup"
		}
	
		if ($productinstalled64 -or $productinstalled32 -or $productinstalled)
		{ 
			#product is installed
			LogWrite "$product found.  Proceeding to remove."
			
			#create an uninstall XML File for each product we want to uninstall
			WriteXmlConfig -productcode $productID
			#check if 'Uninstall Source' parameter was filed otherwise use dynamic source location
			if(!$UninstallSourcePath){
				if($productinstalled32){
					$setupsourcepath = "C:\Program Files (x86)\Common Files\Microsoft Shared\$Productversion\Office Setup Controller\setup.exe"	
				}
				else{
					$setupsourcepath = "C:\Program Files\Common Files\Microsoft Shared\$Productversion\Office Setup Controller\setup.exe"
				}
			}
			else{
				$setupsourcepath = $UninstallSourcePath.TrimEnd("\") + "\setup.exe"
			}
			
			#Check uninstall setup.exe location
			$uninstallpathexist = Test-Path $setupsourcepath
			if($uninstallpathexist)
			{
				LogWrite "Verified Uninstall Path: $setupsourcepath"
				#create argument list to pass to setup.exe	
				$arglist = "/uninstall $productID /config $workingdir\uninstall-$productID.xml"
		
				try
				{
					LogWrite "removing $product..."
					$process = start-process $setupsourcepath -ArgumentList $arglist -Wait -PassThru -NoNewWindow
					$exitcode = $process.ExitCode
					if($exitcode -eq 3010){
						LogWrite "$product has been successfully removed but a reboot should be performed. "
					}
					elseif($exitcode -eq 30066){
						LogWrite "$product failed to uninstall.  Process Exit Code:  $exitcode"
						$uninstallerrors += "$product failed to uninstall."
					}
					else{
						LogWrite "Process Exit Code:  $exitcode"
					}
				
				}
				catch
				{
					$Exception = $_.Exception.Message
					LogWrite "Error: $Exception"
					LogWrite "$product was not removed succesfully.  Further investigation is required"
					$uninstallerrors += "$product failed to uninstall."
				}
			}
			else{
				LogWrite "Cannot verify Uninstall Path: $setupsourcepath"
				LogWrite "Script will not proceed to uninstall $product"
				$uninstallerrors += "$product failed to uninstall."
			}
		
		}
		else{
			LogWrite "$product is not installed to this system"
		}	
	}	
}

#check if Install Office Switch was specified
if($InstallOffice)
{
	#Verify there were no uninstall errors
	if($uninstallerrors)
	{
		LogWrite "There were errors during uninstallation.   Script will not proceed to install Office Click to Run."
	}
	else{
		#Verify a Click-to-Run Source Path has been specified.
		if(!$InstallSourcePath)
		{
			LogWrite "Click-to-Run Source location was not properly specified.   Office will not install"
		}
		else
		{
			LogWrite "Checking if Office Click-to-Run is installed."
			$office201632 = Test-Path -Path "HKLM:\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us"
			$office201664 = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us"
			if($office201632 -or $office201664)
			{
				LogWrite "Office Click-to-Run is already installed.   Going to End of Script."
			}
			else
			{
				LogWrite "Office Click-to-Run is not installed.   Proceeding to install Office Click-to-Run..."
				try
				{
					$clickToRunExe = $InstallSourcePath.TrimEnd("\") + "\setup.exe"
					$clickToRunArgs = "/configure " + $InstallSourcePath.TrimEnd("\") + "\" + $installConfig
					$installprocess = start-process $clickToRunExe -ArgumentList $clickToRunArgs -Wait -PassThru -NoNewWindow
					$installexitcode = $installprocess.ExitCode
					LogWrite "Office Click-to-Run installation has finished running with Exit Code $installexitcode"
				}
				catch
				{
					$Exception = $_.Exception.Message
					LogWrite "Error: $Exception"
					LogWrite "Offce Click-to-Run was not installed succesfully.  Further investigation is required"			
				}
			}
		}
	}
}
else
{
	LogWrite "Install Office Switch was not specified.   Skipping Office Click-to-Run Installation."	
}

LogWrite "********** script end **********"
